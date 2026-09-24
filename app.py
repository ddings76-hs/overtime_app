import re
from pathlib import Path
import uuid
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import io
import base64
import hashlib
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from supabase import create_client, Client

# 페이지 기본 설정

TRIP_STORAGE_BUCKET = "business-trip-files"
APP_ASSET_BUCKET = "app-assets"
APP_VERSION = "v20.5"
COMPANY_LOGO_PATH = "branding/company_logo.png"
st.set_page_config(page_title="화성시장기요양지원센터 통합 업무관리 시스템 · v20.5", layout="wide")

# -------------------------------------------------------------------
# Supabase 클라우드 DB 연결 설정 (Secrets 참조)
# -------------------------------------------------------------------
try:
    url: str = st.secrets["SUPABASE_URL"]
    key: str = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(url, key)
    # 관리자 계정 생성/권한/비밀번호 초기화 전용 Secret Key.
    # 반드시 Streamlit Secrets에만 저장하고 사용자 화면/코드에 직접 노출하지 않습니다.
    admin_secret = st.secrets.get("SUPABASE_SECRET_KEY", "")
    supabase_admin: Client | None = create_client(url, admin_secret) if admin_secret else None
except Exception as e:
    st.error("Supabase 연결 실패! Streamlit Secrets에 SUPABASE_URL과 SUPABASE_KEY가 정상 설정되었는지 확인이 필요하다.")
    st.stop()


# -------------------------------------------------------------------
# 로그인 / 사용자 권한 (Supabase Authentication)
# -------------------------------------------------------------------
def _auth_user_obj(response):
    if response is None:
        return None
    if hasattr(response, "user"):
        return response.user
    if isinstance(response, dict):
        return response.get("user")
    return None

def _user_value(user, key, default=""):
    if user is None:
        return default
    if isinstance(user, dict):
        return user.get(key, default)
    return getattr(user, key, default)

def _user_metadata(user):
    meta = _user_value(user, "user_metadata", {})
    return meta if isinstance(meta, dict) else {}

def _role_label(role):
    return {"admin":"관리자", "manager":"담당자", "employee":"직원", "viewer":"조회자"}.get(role, role or "사용자")

if "auth_user" not in st.session_state:
    st.session_state.auth_user = None
if "auth_role" not in st.session_state:
    st.session_state.auth_role = "viewer"

# Streamlit 재실행 시 Supabase Auth 세션 복구
if st.session_state.auth_user is None:
    try:
        session_res = supabase.auth.get_session()
        session = getattr(session_res, "session", session_res)
        session_user = getattr(session, "user", None) if session else None
        if session_user:
            st.session_state.auth_user = session_user
            st.session_state.auth_role = str(_user_metadata(session_user).get("role", "viewer"))
    except Exception:
        pass

# 인증 전에는 업무화면/DB 내용을 표시하지 않음
if st.session_state.auth_user is None:
    st.markdown("""
    <style>
    .block-container {max-width:560px!important;padding-top:7vh!important;}
    .login-brand{text-align:center;margin:0 0 24px;}
    .login-brand h1{font-size:1.75rem!important;color:#172033;margin-bottom:5px;}
    .login-brand p{color:#667085;margin:0;}
    div[data-testid="stForm"]{
      background:#fff;border:1px solid #e4e9f0;border-radius:16px;
      padding:24px;box-shadow:0 8px 28px rgba(16,24,40,.07);
    }
    </style>
    <div class="login-brand">
      <h1>🏢 화성시장기요양지원센터</h1>
      <p>통합 업무관리 시스템</p>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        st.subheader("로그인")
        login_email = st.text_input("이메일", placeholder="name@example.com")
        login_password = st.text_input("비밀번호", type="password")
        submitted = st.form_submit_button("로그인", type="primary", use_container_width=True)

    if submitted:
        if not login_email.strip() or not login_password:
            st.warning("이메일과 비밀번호를 입력해 주세요.")
        else:
            try:
                result = supabase.auth.sign_in_with_password({
                    "email": login_email.strip(),
                    "password": login_password
                })
                user = _auth_user_obj(result)
                if user is None:
                    raise RuntimeError("사용자 정보 없음")
                st.session_state.auth_user = user
                st.session_state.auth_role = str(_user_metadata(user).get("role", "viewer"))
                st.rerun()
            except Exception:
                st.error("로그인에 실패했습니다. 이메일 또는 비밀번호를 확인해 주세요.")

    st.caption("계정 발급 및 권한 변경은 시스템 관리자가 처리합니다.")
    st.stop()

CURRENT_USER = st.session_state.auth_user
CURRENT_EMAIL = str(_user_value(CURRENT_USER, "email", ""))
CURRENT_META = _user_metadata(CURRENT_USER)
CURRENT_ROLE = str(st.session_state.get("auth_role", CURRENT_META.get("role", "viewer")))
CURRENT_NAME = str(CURRENT_META.get("name", CURRENT_EMAIL.split("@")[0] if CURRENT_EMAIL else "사용자"))

@st.cache_data(ttl=300, show_spinner=False)
def load_company_logo_bytes():
    """Supabase Storage의 고정 경로에서 회사 로고를 불러옵니다."""
    try:
        return supabase.storage.from_(APP_ASSET_BUCKET).download(COMPANY_LOGO_PATH)
    except Exception:
        return None

def save_company_logo(uploaded_file):
    """회사 로고를 Storage 고정 경로에 upsert하여 앱 재시작 후에도 유지합니다."""
    content = uploaded_file.getvalue()
    content_type = uploaded_file.type or "image/png"
    try:
        supabase.storage.from_(APP_ASSET_BUCKET).upload(
            COMPANY_LOGO_PATH,
            content,
            {"content-type": content_type, "upsert": "true"}
        )
    except Exception:
        # SDK/Storage 버전에 따라 upload upsert가 다를 수 있어 update로 재시도
        try:
            supabase.storage.from_(APP_ASSET_BUCKET).update(
                COMPANY_LOGO_PATH,
                content,
                {"content-type": content_type, "upsert": "true"}
            )
        except Exception:
            # 파일이 아직 없을 경우 최종 upload
            supabase.storage.from_(APP_ASSET_BUCKET).upload(
                COMPANY_LOGO_PATH,
                content,
                {"content-type": content_type}
            )
    load_company_logo_bytes.clear()
    return True

def company_logo_data_uri():
    data = load_company_logo_bytes()
    if not data:
        return ""
    encoded = base64.b64encode(data).decode("ascii")
    # 브라우저 표시용. PNG/JPG 모두 대부분 정상 렌더링됨
    return f"data:image/png;base64,{encoded}"


# 역할별 앱 권한
ROLE_PERMISSIONS = {
    "admin": {
        "대시보드",
        "employee_write", "attendance_write", "leave_write", "payroll_write",
        "payroll_confirm", "trip_write", "trip_approve", "trip_settle",
        "file_upload", "file_delete", "print_export"
    },
    "manager": {
        "대시보드",
        "attendance_write", "leave_write", "trip_write", "trip_approve",
        "trip_settle", "file_upload", "print_export"
    },
    "employee": {
        "대시보드",
        "attendance_write", "leave_write", "trip_write", "file_upload", "print_export"
    },
    "viewer": {
        "대시보드","print_export"},
}

def has_permission(permission):
    return permission in ROLE_PERMISSIONS.get(CURRENT_ROLE, ROLE_PERMISSIONS["viewer"])

def require_permission(permission, message="이 작업을 수행할 권한이 없습니다."):
    if not has_permission(permission):
        st.warning(f"🔒 {message}")
        return False
    return True


def truncate_ten(value):
    return int(value // 10) * 10

def safe_int(value):
    """pandas/numpy 숫자를 JSON 저장이 가능한 정수로 통일한다."""
    if pd.isna(value):
        return 0
    return int(value)

def build_payroll_snapshot(pay_month, pay_date, payroll_df, pay_run_no=1, pay_run_name="정기급여"):
    """편집 완료된 급여대장을 확정·회계연계용 스냅샷으로 만든다."""
    employees = []
    totals = {
        "gross_pay": 0, "employee_deductions": 0, "net_pay": 0,
        "employer_insurance": 0, "retirement_accrual": 0
    }

    for _, row in payroll_df.iterrows():
        gross_pay = sum(safe_int(row[c]) for c in [
            '기본급', '초과수당(승인)', '가족수당', '명절상여', '비과세', '기타수당'
        ])
        employee_deductions = sum(safe_int(row[c]) for c in [
            '국민연금(본인)', '건강보험(본인)', '장기요양(본인)',
            '고용보험(본인)', '소득세', '지방소득세', '기타공제'
        ])
        employer_insurance = sum(safe_int(row[c]) for c in [
            '국민연금(사업자)', '건강보험(사업자)', '장기요양(사업자)',
            '고용보험(사업자)', '산재보험(사업자)'
        ])
        retirement = safe_int(row['퇴직적립금'])
        net_pay = gross_pay - employee_deductions

        employees.append({
            "emp_id": str(row['사번']), "emp_name": str(row['이름']),
            "dept": str(row['부서']), "position": str(row['직위']),
            "base_salary": safe_int(row['기본급']),
            "overtime_pay": safe_int(row['초과수당(승인)']),
            "family_allowance": safe_int(row['가족수당']),
            "non_taxable": safe_int(row['비과세']),
            "other_allowance": safe_int(row['기타수당']),
            "holiday_bonus": safe_int(row['명절상여']),
            "gross_pay": gross_pay,
            "employee_deductions": employee_deductions,
            "net_pay": net_pay,
            "employer_insurance": employer_insurance,
            "retirement_accrual": retirement
        })
        totals["gross_pay"] += gross_pay
        totals["employee_deductions"] += employee_deductions
        totals["net_pay"] += net_pay
        totals["employer_insurance"] += employer_insurance
        totals["retirement_accrual"] += retirement

    snapshot = {
        "schema_version": "1.1", "pay_month": pay_month, "pay_run_no": pay_run_no,
        "pay_run_name": pay_run_name, "pay_date": pay_date.isoformat(), "employees": employees, "totals": totals
    }

    accounting_export = {
        "schema_version": "payroll-accounting-1.1",
        "source_type": "payroll", "source_key": f"payroll:{pay_month}:run-{pay_run_no}",
        "pay_month": pay_month, "pay_run_no": pay_run_no, "pay_run_name": pay_run_name,
        "pay_date": pay_date.isoformat(),
        "title": f"{pay_month} {pay_run_name} 및 사용자부담금",
        "items": [
            {"itemDate": pay_date.isoformat(), "accountItem": "급여 및 제수당",
             "botamCategory": "인건비", "payMethod": "계좌이체",
             "vendor": "임직원", "detailSummary": f"{pay_month} {pay_run_name} 총액",
             "amount": totals["gross_pay"]},
            {"itemDate": pay_date.isoformat(), "accountItem": "사회보험료",
             "botamCategory": "인건비", "payMethod": "계좌이체",
             "vendor": "사회보험 기관", "detailSummary": f"{pay_month} 사업주 부담 사회보험",
             "amount": totals["employer_insurance"]},
            {"itemDate": pay_date.isoformat(), "accountItem": "퇴직급여",
             "botamCategory": "인건비", "payMethod": "계좌이체",
             "vendor": "퇴직연금 기관", "detailSummary": f"{pay_month} 퇴직적립금",
             "amount": totals["retirement_accrual"]}
        ],
        "payroll_summary": totals,
        "accounting_total": totals["gross_pay"] + totals["employer_insurance"] + totals["retirement_accrual"],
        "notice": "실지급액과 근로자 공제액은 급여총액의 구성정보이며 지출액에 다시 더하지 않습니다."
    }
    return snapshot, accounting_export

def payload_hash(snapshot, accounting_export):
    import json
    canonical = json.dumps(
        {"payroll": snapshot, "accounting": accounting_export},
        ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()

# 세션 내 로고 이미지 관리
if 'logo_b64' not in st.session_state:
    st.session_state.logo_b64 = ""

st.title("🏢 장기요양지원센터 통합 업무관리 시스템 · v20.5")

# 사이드바: 회사 로고 업로드 기능
with st.sidebar:
    st.header("🖼️ 회사 로고 설정")
    uploaded_logo = st.file_uploader("모든 문서에 적용할 로고 이미지 (PNG, JPG)", type=['png', 'jpg', 'jpeg'])
    if uploaded_logo is not None:
        bytes_data = uploaded_logo.getvalue()
        st.session_state.logo_b64 = base64.b64encode(bytes_data).decode()
        st.image(uploaded_logo, caption="등록된 회사 로고", use_container_width=True)
    elif st.session_state.logo_b64:
        st.info("💡 기존에 등록된 로고가 적용 중이다.")



# -------------------------------------------------------------------
# v19.3 공통 표/Excel 표시 규칙
# -------------------------------------------------------------------
COLUMN_KR = {
    # 공통
    "id":"번호","emp_id":"사번","emp_name":"이름","name":"이름","birth_date":"생년월일",
    "dept":"부서","department":"부서","position":"직위","pay_grade":"호봉",
    "apply_dt":"신청일","created_at":"등록일시","updated_at":"수정일시","status":"상태",
    "note":"비고","reason":"사유",
    # 연차
    "start_date":"시작일","end_date":"종료일","leave_type":"휴가종류","used_days":"사용일수",
    "leave_days":"연차일수","day_count":"일수",
    # 초과근무
    "work_date":"근무일","ot_date":"초과근무일","work_type":"근무구분",
    "start_time":"시작시간","end_time":"종료시간","duration_hours":"신청시간",
    "estimated_pay":"예상수당","act_start_time":"실적시작시간","act_end_time":"실적종료시간",
    "actual_hours":"실적시간","work_hours":"근무시간","hours":"시간",
    "actual_pay":"실적수당","approved_pay":"승인수당",
    "actual_duration_hours":"실적시간","act_reason":"실적사유","requested_hours":"신청시간","approved_hours":"승인시간",
    # 출장
    "trip_type":"출장구분","purpose":"출장목적","start_at":"출장시작","end_at":"출장종료",
    "destination":"출장지","transport_type":"교통수단","distance_km":"출장거리(km)",
    "fuel_type":"유종","rule_base_amount":"기준금액","apply_status":"신청상태",
    "report_status":"복명상태","settlement_status":"정산상태","participants":"동행자",
    "report_content":"복명내용","transport_cost":"교통비","toll_cost":"통행료",
    "lodging_cost":"숙박비","other_cost":"기타비용","daily_cost":"일비",
    "meal_cost":"식비","total_cost":"출장비합계",
    # 급여
    "pay_month":"급여월","pay_run_no":"급여차수","pay_run_name":"급여명칭",
    "base_salary":"기본급","ot_pay":"초과수당","overtime_pay":"초과수당",
    "family_allowance":"가족수당","holiday_bonus":"명절상여","non_taxable":"비과세",
    "other_allowance":"기타수당","national_pension":"국민연금","health_insurance":"건강보험",
    "longterm_care":"장기요양보험","employment_insurance":"고용보험",
    "income_tax":"소득세","local_tax":"지방소득세","other_deduction":"기타공제",
    "ot_pay_overridden":"초과수당수정","gross_pay":"급여총액",
    "employee_deductions":"근로자공제합계","net_pay":"실지급액",
    "employer_insurance":"사업주부담보험","retirement_accrual":"퇴직적립금",
    "employer_national_pension":"사업주국민연금","employer_health_insurance":"사업주건강보험","employer_longterm_care":"사업주장기요양보험","employer_employment_insurance":"사업주고용보험","employer_industrial_insurance":"사업주산재보험","retirement_reserve":"퇴직적립금",
    # 파일/변경
    "file_type":"파일구분","file_name":"파일명","change_type":"변경구분",
    "change_reason":"변경사유","new_destination":"변경출장지","approval_type":"승인구분",
    "approval_status":"승인상태","approved_at":"승인일시"
}
MONEY_HINTS=("급여","수당","금액","비용","출장비","교통비","통행료","숙박비","식비","일비",
             "보험","연금","소득세","공제","지급액","적립금","기본급","상여","비과세")
DECIMAL2_HINTS=("사용일수","연차일수","일수","시간","거리")

def _kr_col(c):
    return COLUMN_KR.get(c, c)

def _normalize_table_view(df):
    """화면/Excel 공통: 한글 제목, 출장 일시 형식, 소수점 규칙."""
    if df is None: return pd.DataFrame()
    view=df.copy()
    # 출장 시작/종료: 초과근무처럼 날짜와 시간을 읽기 쉽게 표시
    for c in ["start_at","end_at"]:
        if c in view.columns:
            dt=pd.to_datetime(view[c], errors="coerce")
            view[c]=dt.dt.strftime("%Y-%m-%d %H:%M").where(dt.notna(), view[c])
    view=view.rename(columns={c:_kr_col(c) for c in view.columns})
    return view

def employee_filter_ui(df, key, label="직원 검색"):
    if df is None or df.empty:
        return df
    # 직원은 기존 본인자료 제한 유지
    if CURRENT_ROLE == "employee":
        return scope_dataframe_to_current_employee(df, "emp_id")
    name_col = "emp_name" if "emp_name" in df.columns else ("name" if "name" in df.columns else None)
    id_col = "emp_id" if "emp_id" in df.columns else None
    if not name_col and not id_col:
        return df
    choices = ["전체"]
    mapping = {}
    for _, row in df.iterrows():
        eid = str(row.get(id_col, "")) if id_col else ""
        nm = str(row.get(name_col, "")) if name_col else ""
        txt = f"{nm} ({eid})" if eid else nm
        if txt and txt not in mapping:
            mapping[txt] = (eid, nm)
            choices.append(txt)
    selected = st.selectbox(label, choices, key=key)
    if selected == "전체":
        return df
    eid, nm = mapping[selected]
    if id_col and eid:
        return df[df[id_col].astype(str) == eid].copy()
    return df[df[name_col].astype(str) == nm].copy()

def display_table_kr(df, use_container_width=True, hide_index=True, **kwargs):
    """모든 조회 표 공통 표시."""
    view=_normalize_table_view(df)
    fmt={}
    for c in view.columns:
        if pd.api.types.is_numeric_dtype(view[c]):
            if any(k in str(c) for k in DECIMAL2_HINTS):
                fmt[c]="{:,.2f}"
            elif any(k in str(c) for k in MONEY_HINTS) or pd.api.types.is_integer_dtype(view[c]):
                fmt[c]="{:,.0f}"
            else:
                fmt[c]="{:,.2f}"
    # 컬럼 제목은 중앙 정렬. 본문 숫자는 우측, 텍스트는 기본 표시.
    sty=view.style.format(fmt, na_rep="").set_table_styles([
        {"selector":"th","props":[("text-align","center")]},
    ])
    return st.dataframe(sty, use_container_width=use_container_width, hide_index=hide_index, **kwargs)

def excel_view_df(df):
    return _normalize_table_view(df)

def style_excel_sheet(ws):
    """화면과 같은 한글 제목/칸너비/전체 격자선/숫자 형식."""
    header_fill=PatternFill("solid", fgColor="F3F6FA")
    header_font=Font(name="맑은 고딕", size=10, bold=True, color="1F2937")
    body_font=Font(name="맑은 고딕", size=10)
    thin=Side(style="thin", color="D9DEE7")
    grid=Border(left=thin,right=thin,top=thin,bottom=thin)
    for cell in ws[1]:
        cell.fill=header_fill; cell.font=header_font; cell.border=grid
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font=body_font; cell.border=grid
            cell.alignment=Alignment(vertical="center",wrap_text=False)
    ws.freeze_panes="A2"
    if ws.max_row>=1 and ws.max_column>=1:
        ws.auto_filter.ref=ws.dimensions
    for col_cells in ws.columns:
        letter=get_column_letter(col_cells[0].column)
        header=str(col_cells[0].value or "")
        vals=[str(c.value) if c.value is not None else "" for c in col_cells[:150]]
        # 화면 가독성에 맞춘 폭. 긴 텍스트도 과도하게 넓어지지 않게 제한.
        width=min(max(max((len(v) for v in vals),default=0)+3, 11), 28)
        ws.column_dimensions[letter].width=width
        for cell in col_cells[1:]:
            if isinstance(cell.value,(int,float)) and not isinstance(cell.value,bool):
                if any(k in header for k in DECIMAL2_HINTS):
                    cell.number_format='#,##0.00'
                elif any(k in header for k in MONEY_HINTS) or isinstance(cell.value,int):
                    cell.number_format='#,##0'
                else:
                    cell.number_format='#,##0.00'
    ws.row_dimensions[1].height=26
    ws.sheet_view.showGridLines=False



def payroll_summary_values(df):
    """급여대장 화면/인쇄/Excel 공통 합계.
    편집용 한글 컬럼과 DB/스냅샷 영문 컬럼을 모두 지원한다.
    """
    zero = {
        "급여총액":0, "근로자 사회보험료":0, "소득세·지방소득세":0,
        "실지급액":0, "사업주 사회보험료":0, "퇴직적립금":0
    }
    if df is None or df.empty:
        return zero

    def s(*cols):
        for c in cols:
            if c in df.columns:
                return int(pd.to_numeric(df[c], errors="coerce").fillna(0).sum())
        return 0

    # 화면 편집 급여대장(한글)과 저장/DB 데이터(영문) 양쪽에서 같은 값을 읽는다.
    gross = s("급여총액", "gross_pay", "total_pay", "pay_total", "total_salary")
    if gross == 0:
        gross = (
            s("기본급", "base_salary") +
            s("초과수당(승인)", "초과수당", "overtime_pay", "ot_pay") +
            s("가족수당", "family_allowance") +
            s("명절상여", "holiday_bonus") +
            s("비과세", "non_taxable") +
            s("기타수당", "other_allowance")
        )

    worker_ins = (
        s("국민연금(본인)", "국민연금", "national_pension") +
        s("건강보험(본인)", "건강보험", "health_insurance") +
        s("장기요양(본인)", "장기요양보험", "longterm_care") +
        s("고용보험(본인)", "고용보험", "employment_insurance")
    )
    tax = (
        s("소득세", "income_tax") +
        s("지방소득세", "지방세", "local_tax")
    )
    net = s("실지급액", "net_pay", "real_pay")
    if net == 0 and gross:
        other_ded = s("기타공제", "other_deduction")
        net = gross - worker_ins - tax - other_ded

    employer_ins = s("사업주 부담보험", "사업주부담보험", "employer_insurance")
    if employer_ins == 0:
        employer_ins = (
            s("국민연금(사업자)", "사업주국민연금", "employer_national_pension") +
            s("건강보험(사업자)", "사업주건강보험", "employer_health_insurance") +
            s("장기요양(사업자)", "사업주장기요양보험", "employer_longterm_care") +
            s("고용보험(사업자)", "사업주고용보험", "employer_employment_insurance") +
            s("산재보험(사업자)", "사업주산재보험", "employer_industrial_insurance")
        )

    retirement = s("퇴직적립금", "retirement_accrual", "retirement_reserve")

    return {
        "급여총액": gross,
        "근로자 사회보험료": worker_ins,
        "소득세·지방소득세": tax,
        "실지급액": net,
        "사업주 사회보험료": employer_ins,
        "퇴직적립금": retirement,
    }

def show_payroll_summary(df, key_prefix="pay"):
    vals=payroll_summary_values(df)
    st.markdown("#### 💰 급여대장 주요 합계")
    c1,c2,c3=st.columns(3)
    c1.metric("급여총액",f"{vals['급여총액']:,}원")
    c2.metric("근로자 사회보험료 합계",f"{vals['근로자 사회보험료']:,}원")
    c3.metric("소득세·지방소득세 합계",f"{vals['소득세·지방소득세']:,}원")
    c4,c5,c6=st.columns(3)
    c4.metric("실지급액 합계",f"{vals['실지급액']:,}원")
    c5.metric("사업주 사회보험료 합계",f"{vals['사업주 사회보험료']:,}원")
    c6.metric("퇴직적립금 합계",f"{vals['퇴직적립금']:,}원")
    return vals


def current_emp_id():
    return str(CURRENT_META.get("emp_id", "") or "").strip()

def is_employee_scope():
    return CURRENT_ROLE == "employee"

def scope_dataframe_to_current_employee(df, emp_col="emp_id"):
    if not is_employee_scope() or df is None or df.empty:
        return df
    eid = current_emp_id()
    if not eid or emp_col not in df.columns:
        return df.iloc[0:0].copy()
    return df[df[emp_col].astype(str).str.strip() == eid].copy()


def scope_employee_master(df):
    """직원 계정에서는 employees 마스터도 본인 행만 노출."""
    return scope_dataframe_to_current_employee(df, "emp_id")

def can_manage_all_records():
    return CURRENT_ROLE in ("admin", "manager")

def write_audit_log(action, target_type="", target_id="", detail=""):
    """감사로그 실패가 본 업무를 막지 않도록 best-effort로 기록."""
    try:
        supabase.table("audit_logs").insert({
            "user_id": str(_user_value(CURRENT_USER, "id", "")),
            "user_email": CURRENT_EMAIL,
            "user_role": CURRENT_ROLE,
            "action": action,
            "target_type": target_type,
            "target_id": str(target_id),
            "detail": str(detail)[:2000],
        }).execute()
    except Exception:
        pass

# -------------------------------------------------------------------
# 통합 업무 메뉴 / UI
# -------------------------------------------------------------------
st.markdown("""
<style>
/* 전체 화면 폭과 기본 타이포 */
.block-container {max-width: 1500px; padding-top: 4.5rem; padding-bottom: 3rem;}
html, body, [class*="css"] {font-family: "Pretendard","Noto Sans KR","Malgun Gothic",sans-serif;}
h1,h2,h3 {letter-spacing:-0.035em; color:#172033;}
h1 {font-size:1.75rem!important; line-height:1.35!important; margin-top:0!important; padding-top:.15rem!important;} h2 {font-size:1.45rem!important;} h3 {font-size:1.15rem!important;}

/* 입력창 */
div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div,
div[data-baseweb="textarea"] > div {
    border-radius:10px!important; border-color:#d9e0e8!important; background:#fff!important;
}
div[data-testid="stDateInput"] input, div[data-testid="stTimeInput"] input {border-radius:10px!important;}

/* 버튼 */
.stButton > button, .stDownloadButton > button {
    border-radius:10px!important; min-height:40px; font-weight:650;
    border:1px solid #d8dee8; box-shadow:0 1px 2px rgba(16,24,40,.04);
}
.stButton > button[kind="primary"] {
    background:#315efb!important; border-color:#315efb!important;
}

/* 데이터 표 */
div[data-testid="stDataFrame"] {
    border:1px solid #e4e9f0; border-radius:12px; overflow:hidden;
    box-shadow:0 2px 8px rgba(16,24,40,.035);
}

/* Metric */
div[data-testid="stMetric"] {
    background:#fff; border:1px solid #e4e9f0; border-radius:12px;
    padding:14px 16px; box-shadow:0 2px 8px rgba(16,24,40,.035);
}
div[data-testid="stMetricLabel"] {font-weight:650; color:#667085;}
div[data-testid="stMetricValue"] {font-size:1.65rem; color:#172033;}

/* 안내창 */
div[data-testid="stAlert"] {border-radius:10px;}

/* 구분선 */
hr {margin:1.5rem 0; border-color:#edf0f4;}

/* 사이드바 */
section[data-testid="stSidebar"] {border-right:1px solid #e8ecf2;}
section[data-testid="stSidebar"] .stRadio label {padding:.18rem 0;}
section[data-testid="stSidebar"] div[role="radiogroup"] {gap:.15rem;}

/* 카드 */
.ui-card {
    background:#fff; border:1px solid #e4e9f0; border-radius:14px;
    padding:16px 18px; box-shadow:0 2px 10px rgba(16,24,40,.035); margin-bottom:12px;
}
.ui-eyebrow {font-size:.78rem; color:#667085; font-weight:700; margin-bottom:4px;}
.ui-title {font-size:1.12rem; font-weight:750; color:#172033;}
.ui-desc {font-size:.88rem; color:#667085; margin-top:5px; line-height:1.55;}
</style>
""", unsafe_allow_html=True)

MENU_GROUPS = {
    "🏠 홈": [
        ("대시보드", "나의 업무현황 또는 관리자 통합현황"),
    ],
    "👥 인사": [
        ("직원 관리", "직원 등록·정보 수정 및 기본 인사정보"),
        ("연차 관리", "연차 발생·사용·잔여 현황 및 전 직원 요약"),
        ("연차 신청서", "연차 신청서 조회·인쇄"),
        ("계정·권한 관리", "직원 계정 생성·권한 부여·비밀번호 초기화"),
    ],
    "⏱️ 근태": [
        ("초과근무 신청", "초과·휴일근무 사전 신청"),
        ("초과근무 실적", "실제 수행내역·승인·월별 급여 연계"),
    ],
    "💰 급여": [
        ("관리자 통계·보고서", "연·월별 급여·연차·초과근무·출장 통합 통계 및 보고서"),
        ("통합 급여대장", "월 급여 계산·수정·확정 및 Excel"),
        ("급여명세서", "직원별 급여명세서 조회·인쇄"),
        ("급여대장 인쇄", "월별 통합 급여대장 인쇄"),
        ("연간 급여총괄", "12개월 누적 급여대장 요약"),
    ],
    "🚗 출장": [
        ("출장 신청·관리", "출장신청·승인·현황 관리"),
        ("출장 복명·규정", "복명·여비정산·증빙·복무규정"),
    ],
}

MENU_TO_TAB = {
    "대시보드": 0,
    "직원 관리": 1, "초과근무 신청": 2, "초과근무 실적": 3,
    "연차 관리": 4, "연차 신청서": 5, "통합 급여대장": 6,
    "급여명세서": 7, "급여대장 인쇄": 8, "연간 급여총괄": 9,
    "출장 신청·관리": 10, "출장 복명·규정": 11, "계정·권한 관리": 12,
    "관리자 통계·보고서": 13
}

ROLE_ALLOWED_MENUS = {
    "admin": set(MENU_TO_TAB.keys()),
    "manager": {
        "대시보드", "관리자 통계·보고서",
        "초과근무 신청", "초과근무 실적", "연차 관리", "연차 신청서",
        "출장 신청·관리", "출장 복명·규정"
    },
    "employee": {
        "초과근무 신청", "초과근무 실적", "연차 관리", "연차 신청서",
        "급여명세서", "출장 신청·관리", "출장 복명·규정"
    },
    "viewer": {"연차 신청서", "급여명세서", "출장 복명·규정"},
}

with st.sidebar:
    st.markdown("## 🏢 업무관리")
    st.caption(f"화성시장기요양지원센터 · {APP_VERSION}")
    st.markdown(
        f"""<div style="padding:10px 12px;background:#f7f9fc;border:1px solid #e6eaf0;border-radius:10px;margin:10px 0 12px;">
        <div style="font-weight:700;">👤 {CURRENT_NAME}</div>
        <div style="font-size:.78rem;color:#667085;margin-top:3px;">{CURRENT_EMAIL}</div>
        <div style="font-size:.78rem;color:#315efb;margin-top:3px;">{_role_label(CURRENT_ROLE)}</div>
        </div>""",
        unsafe_allow_html=True
    )
    if st.button("🚪 로그아웃", use_container_width=True):
        try:
            supabase.auth.sign_out()
        except Exception:
            pass
        st.session_state.auth_user = None
        st.session_state.auth_role = "viewer"
        st.rerun()

    if CURRENT_ROLE == "admin":
        with st.expander("🎨 회사 로고 관리"):
            st.caption("저장한 로고는 Supabase Storage에 보관되어 앱 재시작·재배포 후에도 유지됩니다.")
            logo_file = st.file_uploader(
                "로고 파일",
                type=["png", "jpg", "jpeg"],
                key="company_logo_upload"
            )
            if st.button("💾 회사 로고 저장", use_container_width=True):
                if logo_file is None:
                    st.warning("저장할 로고 파일을 선택해 주세요.")
                else:
                    try:
                        save_company_logo(logo_file)
                        write_audit_log("회사 로고 변경", "app_assets", COMPANY_LOGO_PATH, logo_file.name)
                        st.success("회사 로고를 저장했습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"회사 로고 저장 실패: {e}")
    st.divider()
    allowed_menus = ROLE_ALLOWED_MENUS.get(CURRENT_ROLE, ROLE_ALLOWED_MENUS["viewer"])
    visible_groups = [g for g, items in MENU_GROUPS.items() if any(name in allowed_menus for name, _ in items)]
    menu_group = st.radio("업무 영역", visible_groups, label_visibility="collapsed")
    allowed_menus = ROLE_ALLOWED_MENUS.get(CURRENT_ROLE, ROLE_ALLOWED_MENUS["viewer"])
    choices = [x[0] for x in MENU_GROUPS[menu_group] if x[0] in allowed_menus]
    if not choices:
        st.info("현재 권한에서 사용할 수 있는 메뉴가 없습니다.")
        st.stop()
    menu = st.radio("세부 메뉴", choices, label_visibility="collapsed")
    desc = dict(MENU_GROUPS[menu_group])[menu]
    st.caption(desc)
    st.divider()
    st.caption("급여 · 근태 · 연차 · 출장 통합관리")

active_tab = MENU_TO_TAB[menu]

# 현재 위치 표시
st.markdown(f"### 🏢 화성시장기요양지원센터 통합 업무관리 시스템 · {APP_VERSION}")
if CURRENT_ROLE == "employee":
    if current_emp_id():
        st.caption(f"👤 직원 전용 화면 · 사번 {current_emp_id()} · 본인 자료만 조회/신청")
    else:
        st.error("이 계정에 직원 사번이 연결되어 있지 않습니다. 관리자에게 계정 연결을 요청해 주세요.")


st.markdown(
    f"""<div class="ui-card">
    <div class="ui-eyebrow">{menu_group.replace('👥 ','').replace('⏱️ ','').replace('💰 ','').replace('🚗 ','').replace('🏠 ','')} · 업무관리</div>
    <div class="ui-title">{menu}</div>
    <div class="ui-desc">{desc}</div>
    </div>""",
    unsafe_allow_html=True
)

# -------------------------------------------------------------------
# TAB 0: v19 직원 마이페이지 / 관리자 통합 대시보드
# -------------------------------------------------------------------
if active_tab == 0:
    if CURRENT_ROLE == "employee":
        st.header("👤 나의 업무 현황")
        _eid=current_emp_id()
        if not _eid:
            st.warning("로그인 계정과 직원 사번이 연결되어 있지 않습니다.")
        else:
            try:
                _me=(supabase.table("employees").select("*").eq("emp_id",_eid).limit(1).execute().data or [{}])[0]
            except Exception: _me={}
            st.subheader(f"{_me.get('emp_name',CURRENT_NAME)} 님")
            st.caption(f"{_me.get('dept','')} · {_me.get('position','')} · 사번 {_eid}")

            def _my_df(table):
                try: return pd.DataFrame(supabase.table(table).select("*").eq("emp_id",_eid).execute().data or [])
                except Exception: return pd.DataFrame()
            _lv=_my_df("leave_records"); _ot=_my_df("overtime_records"); _tr=_my_df("business_trips")
            _used=pd.to_numeric(_lv["used_days"],errors="coerce").fillna(0).sum() if (not _lv.empty and "used_days" in _lv) else 0
            _month=datetime.now().strftime("%Y-%m")
            _dc=next((c for c in ["work_date","ot_date","date"] if c in _ot.columns),None)
            _otm=int(_ot[_ot[_dc].astype(str).str.startswith(_month)].shape[0]) if _dc else 0
            c1,c2,c3=st.columns(3)
            c1.metric("누적 휴가 사용",f"{_used:g}일"); c2.metric("이번 달 초과근무",f"{_otm}건"); c3.metric("출장 기록",f"{len(_tr)}건")
            _oth=0.0
            if not _ot.empty:
                _hc=next((c for c in ["hours","actual_hours","work_hours"] if c in _ot.columns),None)
                if _hc: _oth=pd.to_numeric(_ot[_hc],errors="coerce").fillna(0).sum()
            _tc=int(pd.to_numeric(_tr["total_cost"],errors="coerce").fillna(0).sum()) if (not _tr.empty and "total_cost" in _tr.columns) else 0
            _pending=(int((~_ot["status"].astype(str).isin(["승인","반려"])).sum()) if (not _ot.empty and "status" in _ot.columns) else 0)
            _pending+=(int((~_tr["apply_status"].astype(str).isin(["승인","반려","취소"])).sum()) if (not _tr.empty and "apply_status" in _tr.columns) else 0)
            s1,s2,s3=st.columns(3)
            s1.metric("초과근무 누적시간",f"{_oth:g}시간")
            s2.metric("출장비 누계",f"{_tc:,}원")
            s3.metric("미처리 업무",f"{_pending}건")
            st.divider()
            a,b=st.columns(2)
            with a:
                st.subheader("🌴 최근 휴가")
                if _lv.empty: st.info("휴가 내역이 없습니다.")
                else:
                    cs=[c for c in ["start_date","end_date","leave_type","used_days"] if c in _lv.columns]
                    display_table_kr((_lv.sort_values("id",ascending=False).head(5) if "id" in _lv else _lv.tail(5))[cs],use_container_width=True,hide_index=True)
                st.subheader("⏱️ 최근 초과근무")
                if _ot.empty: st.info("초과근무 내역이 없습니다.")
                else:
                    cs=[c for c in ["work_date","ot_date","hours","status"] if c in _ot.columns]
                    display_table_kr((_ot.sort_values("id",ascending=False).head(5) if "id" in _ot else _ot.tail(5))[cs],use_container_width=True,hide_index=True)
            with b:
                st.subheader("🚗 최근 출장")
                if _tr.empty: st.info("출장 내역이 없습니다.")
                else:
                    cs=[c for c in ["start_at","end_at","destination","purpose","apply_status"] if c in _tr.columns]
                    display_table_kr((_tr.sort_values("id",ascending=False).head(5) if "id" in _tr else _tr.tail(5))[cs],use_container_width=True,hide_index=True)
                st.subheader("💰 급여")
                st.info("좌측 ‘급여명세서’ 메뉴에서 본인 급여명세서를 확인하세요.")
    else:
        st.header("📊 관리자 통합 대시보드")
        def _all_df(table):
            try: return pd.DataFrame(supabase.table(table).select("*").execute().data or [])
            except Exception: return pd.DataFrame()
        _em=_all_df("employees"); _lv=_all_df("leave_records"); _ot=_all_df("overtime_records"); _tr=_all_df("business_trips")
        _otp=int((~_ot["status"].astype(str).isin(["승인","반려"])).sum()) if (not _ot.empty and "status" in _ot) else 0
        _trp=int((~_tr["apply_status"].astype(str).isin(["승인","반려","취소"])).sum()) if (not _tr.empty and "apply_status" in _tr) else 0
        c1,c2,c3,c4=st.columns(4)
        c1.metric("직원",f"{len(_em)}명"); c2.metric("휴가 신청",f"{len(_lv)}건"); c3.metric("초과근무 미처리",f"{_otp}건"); c4.metric("출장 미처리",f"{_trp}건")
        _aoth=0.0
        if not _ot.empty:
            _hc=next((c for c in ["hours","actual_hours","work_hours"] if c in _ot.columns),None)
            if _hc: _aoth=pd.to_numeric(_ot[_hc],errors="coerce").fillna(0).sum()
        _atc=int(pd.to_numeric(_tr["total_cost"],errors="coerce").fillna(0).sum()) if (not _tr.empty and "total_cost" in _tr.columns) else 0
        x1,x2=st.columns(2)
        x1.metric("초과근무 누적시간",f"{_aoth:g}시간")
        x2.metric("출장비 누계",f"{_atc:,}원")
        st.divider()
        a,b=st.columns(2)
        with a:
            st.subheader("🌴 최근 휴가")
            if not _lv.empty:
                cs=[c for c in ["emp_name","start_date","end_date","leave_type","used_days"] if c in _lv.columns]
                display_table_kr((_lv.sort_values("id",ascending=False).head(8) if "id" in _lv else _lv.tail(8))[cs],use_container_width=True,hide_index=True)
            else: st.info("휴가 신청이 없습니다.")
            st.subheader("⏱️ 최근 초과근무")
            if not _ot.empty:
                cs=[c for c in ["emp_name","work_date","ot_date","hours","status"] if c in _ot.columns]
                display_table_kr((_ot.sort_values("id",ascending=False).head(8) if "id" in _ot else _ot.tail(8))[cs],use_container_width=True,hide_index=True)
            else: st.info("초과근무가 없습니다.")
        with b:
            st.subheader("🚗 최근 출장")
            if not _tr.empty:
                cs=[c for c in ["emp_name","destination","start_at","end_at","apply_status"] if c in _tr.columns]
                display_table_kr((_tr.sort_values("id",ascending=False).head(8) if "id" in _tr else _tr.tail(8))[cs],use_container_width=True,hide_index=True)
            else: st.info("출장 내역이 없습니다.")
            st.subheader("👥 부서별 인원")
            if not _em.empty and "dept" in _em.columns:
                display_table_kr(_em.groupby("dept").size().reset_index(name="인원"),use_container_width=True,hide_index=True)
            else: st.info("직원 데이터가 없습니다.")

# -------------------------------------------------------------------
# TAB 1: 직원 등록 및 정보 관리
# -------------------------------------------------------------------
if active_tab == 1:
    st.header("👥 직원 관리")
    if not has_permission("employee_write"):
        st.info("🔒 현재 계정은 직원정보 조회만 가능합니다. 등록·수정·삭제는 관리자 권한이 필요합니다.")
    try:
        emp_res = supabase.table("employees").select("*").execute()
        df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()
        df_emp = scope_employee_master(df_emp)
    except Exception as e:
        st.warning("⚠️ Supabase에서 직원 데이터를 불러오지 못했다. Secrets의 API Key 설정 또는 DB 상태를 확인해 주어야 한다.")
        df_emp = pd.DataFrame()
    
    if not df_emp.empty:
        employee_column_labels = {
            "emp_id": "사번", "emp_name": "이름", "birth_date": "생년월일",
            "dept": "부서", "position": "직위", "hobong": "호봉",
            "base_salary": "기본급", "hourly_wage": "통상시급",
            "family_allowance": "가족수당", "non_taxable": "비과세",
            "other_allowance": "기타수당", "other_deduction": "기타공제",
            "holiday_bonus": "명절상여",
            "national_pension": "국민연금(본인)", "health_insurance": "건강보험(본인)",
            "longterm_care": "장기요양(본인)", "employment_insurance": "고용보험(본인)",
            "income_tax": "소득세", "local_tax": "지방소득세",
            "employer_national_pension": "국민연금(회사)",
            "employer_health_insurance": "건강보험(회사)",
            "employer_longterm_care": "장기요양(회사)",
            "employer_employment_insurance": "고용보험(회사)",
            "employer_industrial_insurance": "산재보험(회사)",
            "retirement_accrual": "퇴직적립금",
            "is_national": "국민연금 가입", "is_health": "건강보험 가입",
            "is_employment": "고용보험 가입", "is_industrial": "산재보험 가입",
            "total_annual_leave": "연간 연차일수", "created_at": "등록일시", "updated_at": "수정일시",
            "id": "DB번호"
        }
        display_emp = df_emp.rename(columns=employee_column_labels)
        employee_money_columns = [
            "기본급", "통상시급", "가족수당", "명절상여", "비과세", "기타수당", "기타공제",
            "국민연금(본인)", "건강보험(본인)", "장기요양(본인)", "고용보험(본인)", "소득세", "지방소득세",
            "국민연금(회사)", "건강보험(회사)",
            "장기요양(회사)", "고용보험(회사)", "산재보험(회사)", "퇴직적립금"
        ]
        employee_config = {
            col: st.column_config.NumberColumn(col, min_value=0, step=10, format="localized", width="medium")
            for col in employee_money_columns if col in display_emp.columns
        }
        for col in ["국민연금 가입", "건강보험 가입", "고용보험 가입", "산재보험 가입"]:
            if col in display_emp.columns:
                display_emp[col] = display_emp[col].astype(bool)
                employee_config[col] = st.column_config.CheckboxColumn(col)
        edited_display_df = st.data_editor(
            display_emp, use_container_width=True, num_rows="dynamic", hide_index=True,
            column_config=employee_config
        )
        edited_df = edited_display_df.rename(columns={v: k for k, v in employee_column_labels.items()})
        if st.button("수정 데이터 DB 저장"):
            for _, row in edited_df.iterrows():
                save_row = row.to_dict()
                for col in ["is_national", "is_health", "is_employment", "is_industrial"]:
                    if col in save_row:
                        save_row[col] = 1 if bool(save_row[col]) else 0
                supabase.table("employees").upsert(save_row).execute()
            st.success("직원 데이터 수정사항이 Supabase DB에 반영되었다.")
            st.rerun()

    st.divider()

    st.header("2. 신규 직원 등록")
    with st.form("employee_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            emp_id = st.text_input("사번")
            emp_name = st.text_input("이름")
            birth_date = st.text_input("생년월일 (예: 1980-01-01)", value="1980-01-01")
            dept = st.text_input("부서")
            position = st.text_input("직위", value="주임")
            hobong = st.text_input("호봉", value="1호봉")
            total_leave = st.number_input("연간 총 연차 부여일수", min_value=0.0, value=15.0, step=0.1, format="%.1f")
        with col2:
            base_salary = st.number_input("기본급 (원)", min_value=0, value=2500000, step=100000)
            hourly_wage = st.number_input("통상시급 (원)", min_value=0, value=12000, step=500)
            family_allowance = st.number_input("가족수당 (원)", min_value=0, value=50000, step=10000)
            non_taxable = st.number_input("비과세 (원)", min_value=0, value=100000, step=10000)
            other_allowance = st.number_input("기타수당 (원)", min_value=0, value=0, step=10000)
            other_deduction = st.number_input("기타공제 (원)", min_value=0, value=0, step=10000)
            holiday_bonus = st.number_input("명절상여 기본액 (원)", min_value=0, value=0, step=10000)

            st.write("**보험 가입 여부**")
            is_national = st.checkbox("국민연금 가입", value=True)
            is_health = st.checkbox("건강/장기요양보험 가입", value=True)
            is_employment = st.checkbox("고용보험 가입", value=True)
            is_industrial = st.checkbox("산재보험 가입", value=True)

        st.markdown("**직원 부담 보험료·세금 및 회사 부담금 기본값**")
        d1, d2, d3 = st.columns(3)
        with d1:
            national_pension = st.number_input("국민연금 본인부담", min_value=0, value=0, step=10)
            health_insurance = st.number_input("건강보험 본인부담", min_value=0, value=0, step=10)
            longterm_care = st.number_input("장기요양 본인부담", min_value=0, value=0, step=10)
            employment_insurance = st.number_input("고용보험 본인부담", min_value=0, value=0, step=10)
        with d2:
            income_tax = st.number_input("소득세 본인부담", min_value=0, value=0, step=10)
            local_tax = st.number_input("지방소득세 본인부담", min_value=0, value=0, step=10)
            employer_national_pension = st.number_input("국민연금 회사부담", min_value=0, value=0, step=10)
            employer_health_insurance = st.number_input("건강보험 회사부담", min_value=0, value=0, step=10)
        with d3:
            employer_longterm_care = st.number_input("장기요양 회사부담", min_value=0, value=0, step=10)
            employer_employment_insurance = st.number_input("고용보험 회사부담", min_value=0, value=0, step=10)
            employer_industrial_insurance = st.number_input("산재보험 회사부담", min_value=0, value=0, step=10)
            retirement_accrual = st.number_input("퇴직적립금", min_value=0, value=0, step=10)
            
        submit_emp = st.form_submit_button("직원 DB 등록")
        
        if submit_emp:
            if emp_id and emp_name:
                emp_data = {
                    "emp_id": emp_id, "emp_name": emp_name, "birth_date": birth_date, "dept": dept,
                    "position": position, "hobong": hobong, "base_salary": base_salary,
                    "hourly_wage": hourly_wage, "family_allowance": family_allowance,
                    "non_taxable": non_taxable, "other_allowance": other_allowance, "holiday_bonus": holiday_bonus,
                    "other_deduction": other_deduction,
                    "national_pension": national_pension, "health_insurance": health_insurance,
                    "longterm_care": longterm_care, "employment_insurance": employment_insurance,
                    "income_tax": income_tax, "local_tax": local_tax,
                    "employer_national_pension": employer_national_pension,
                    "employer_health_insurance": employer_health_insurance,
                    "employer_longterm_care": employer_longterm_care,
                    "employer_employment_insurance": employer_employment_insurance,
                    "employer_industrial_insurance": employer_industrial_insurance,
                    "retirement_accrual": retirement_accrual,
                    "is_national": 1 if is_national else 0, "is_health": 1 if is_health else 0,
                    "is_employment": 1 if is_employment else 0, "is_industrial": 1 if is_industrial else 0,
                    "total_annual_leave": total_leave
                }
                res = supabase.table("employees").insert(emp_data).execute()
                if res.data:
                    st.success(f"{emp_name} ({position}) 직원이 Supabase DB에 정상 등록되었다.")
                else:
                    st.error("직원 등록 중 오류가 발생했거나 이미 존재하는 사번이다.")
            else:
                st.error("사번과 이름을 모두 입력해야 한다.")

# -------------------------------------------------------------------
# TAB 2: 초과근무 사전 신청
# -------------------------------------------------------------------
if active_tab == 2:
    st.header("1. 초과근무 / 휴일근무 사전 신청")
    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()
    df_emp = scope_employee_master(df_emp)

    df_emp = scope_employee_master(df_emp)
    if df_emp.empty:
        st.warning("먼저 '직원 등록 및 정보 관리' 탭에서 직원을 등록해야 한다.")
    else:
        emp_list = df_emp['emp_name'] + " (" + df_emp['position'] + " / " + df_emp['emp_id'] + ")"
        selected_emp_str = st.selectbox("직원 선택", emp_list, key="ot_emp")
        selected_emp_id = selected_emp_str.split("/")[-1].replace(")", "").strip()
        emp_info = df_emp[df_emp['emp_id'] == selected_emp_id].iloc[0]

        work_date = st.date_input("근무 예정 일자", datetime.now())
        work_type = st.radio("근무 구분", ["평일 초과근무 (18:00 이후)", "휴일근무"])
        
        col1, col2 = st.columns(2)
        with col1:
            start_time = st.time_input("예정 시작 시간", time(18, 0) if work_type == "평일 초과근무 (18:00 이후)" else time(9, 0))
        with col2:
            end_time = st.time_input("예정 종료 시간", time(20, 0))

        reason = st.text_area("신청 사유")

        start_dt = datetime.combine(work_date, start_time)
        end_dt = datetime.combine(work_date, end_time)
        duration_hours = (end_dt - start_dt).total_seconds() / 3600

        if duration_hours < 0:
            st.error("종료 시간은 시작 시간보다 빨라야 한다.")
        else:
            multiplier = 1.5
            raw_pay = duration_hours * emp_info['hourly_wage'] * multiplier
            estimated_pay = truncate_ten(raw_pay)

            st.info(f"💡 예상 근무시간: **{duration_hours:.1f}시간**")

            if st.button("사전 신청서 제출"):
                ot_data = {
                    "apply_dt": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "emp_id": emp_info['emp_id'], "emp_name": emp_info['emp_name'],
                    "dept": emp_info['dept'], "position": emp_info['position'],
                    "work_date": str(work_date), "work_type": work_type,
                    "start_time": str(start_time), "end_time": str(end_time),
                    "duration_hours": duration_hours, "estimated_pay": estimated_pay,
                    "act_start_time": str(start_time), "act_end_time": str(end_time),
                    "actual_duration_hours": duration_hours, "actual_pay": estimated_pay,
                    "status": "신청", "reason": reason, "act_reason": ""
                }
                supabase.table("overtime_records").insert(ot_data).execute()
                st.success("초과근무 신청 내역이 Supabase DB에 등록되었다.")

# -------------------------------------------------------------------
# TAB 3: 실제 수행 입력 & 근무일자별 전체 내역 & 급여 수동 연계 (전월 25일~당월 24일 기준)
# -------------------------------------------------------------------
if active_tab == 3:
    st.header("✅ 실제 초과/휴일근무 수행 내역 입력 & 급여 연계")
    
    ot_res = supabase.table("overtime_records").select("*").order("id", desc=True).execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()

    df_ot = scope_dataframe_to_current_employee(df_ot)
    if df_ot.empty:
        st.info("등록된 초과근무 신청 내역이 없다.")
    else:
        col_ot1, col_ot2 = st.columns([3, 1])
        with col_ot1:
            ot_options = [f"ID {r['id']} | [{r['status']}] [{r['work_date']}] {r['emp_name']} ({r['work_type']}) - 사전: {float(r['duration_hours']):.1f}h" for _, r in df_ot.iterrows()]
            selected_ot_idx = st.selectbox("처리 및 출력할 초과근무 내역 선택", range(len(ot_options)), format_func=lambda x: ot_options[x])
            target_ot = df_ot.iloc[selected_ot_idx]
        
        with col_ot2:
            st.write("**🗑️ 선택 내역 삭제**")
            if st.button("해당 초과근무 내역 삭제", type="primary", key="del_ot_btn"):
                supabase.table("overtime_records").delete().eq("id", int(target_ot['id'])).execute()
                st.success(f"ID {target_ot['id']} 초과근무 내역이 정상적으로 삭제되었다.")
                st.rerun()

        emp_s_res = supabase.table("employees").select("*").eq("emp_id", target_ot['emp_id']).execute()
        df_emp_single = pd.DataFrame(emp_s_res.data) if emp_s_res.data else pd.DataFrame()
        df_emp_single = scope_employee_master(df_emp_single)
        hourly_w = df_emp_single.iloc[0]['hourly_wage'] if not df_emp_single.empty else 12000

        act_s_val = target_ot.get('act_start_time', target_ot['start_time'])
        act_e_val = target_ot.get('act_end_time', target_ot['end_time'])
        if pd.isna(act_s_val) or not act_s_val: act_s_val = target_ot['start_time']
        if pd.isna(act_e_val) or not act_e_val: act_e_val = target_ot['end_time']

        st.subheader(f"✏️ 실제 수행 시간 및 승인 상태 변경: {target_ot['emp_name']} ({target_ot['work_date']})")
        
        with st.form("actual_ot_form"):
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                st.write(f"**사전 신청 정보**")
                st.write(f"- 근무 구분: {target_ot['work_type']}")
                st.write(f"- 신청 시간: {target_ot['start_time']} ~ {target_ot['end_time']} ({float(target_ot['duration_hours']):.1f}시간)")
                st.write(f"- 신청 사유: {target_ot['reason']}")
            
            with col_a2:
                st.write(f"**실제 수행 근무시간 & 업무 내용 입력**")
                try:
                    init_s_time = datetime.strptime(str(act_s_val)[:8], "%H:%M:%S").time()
                    init_e_time = datetime.strptime(str(act_e_val)[:8], "%H:%M:%S").time()
                except:
                    init_s_time = time(18, 0)
                    init_e_time = time(20, 0)

                act_s_time = st.time_input("실제 시작 시간", init_s_time)
                act_e_time = st.time_input("실제 종료 시간", init_e_time)

                dummy_date = datetime.now().date()
                s_dt = datetime.combine(dummy_date, act_s_time)
                e_dt = datetime.combine(dummy_date, act_e_time)
                calculated_act_hours = max(0.0, (e_dt - s_dt).total_seconds() / 3600)

                st.metric(label="실제 인정 시간", value=f"{calculated_act_hours:.1f} 시간")
                act_pay = truncate_ten(calculated_act_hours * hourly_w * 1.5)
                
                status_choice = st.selectbox("승인 상태", ["승인", "신청", "반려"], index=["승인", "신청", "반려"].index(target_ot['status']) if target_ot['status'] in ["승인", "신청", "반려"] else 0)
                act_reason_input = st.text_input("실제 업무 수행 내용 / 확인 메모", value=target_ot['act_reason'] if pd.notna(target_ot['act_reason']) else '')

            submit_act = st.form_submit_button("실제 수행 내역 저장 및 승인 반영")

            if submit_act:
                update_data = {
                    "act_start_time": str(act_s_time), "act_end_time": str(act_e_time),
                    "actual_duration_hours": calculated_act_hours, "actual_pay": act_pay,
                    "status": status_choice, "act_reason": act_reason_input
                }
                supabase.table("overtime_records").update(update_data).eq("id", int(target_ot['id'])).execute()
                st.success(f"[{target_ot['emp_name']}] 직원의 수행 내역 및 승인 상태({status_choice})가 반영되었다.")
                st.rerun()

        latest_res = supabase.table("overtime_records").select("*").eq("id", int(target_ot['id'])).execute()
        target_ot_latest = latest_res.data[0] if latest_res.data else target_ot

        st.divider()
        st.subheader("🖨️ 초과/휴일근무 신청 및 확인서 인쇄")

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 35px; float: left;">' if st.session_state.logo_b64 else ''
        act_reason_disp = target_ot_latest['act_reason'] if pd.notna(target_ot_latest['act_reason']) and target_ot_latest['act_reason'] != "" else "입력된 실제 수행 내용 없음"

        ot_confirm_template = f"""<style>
        * {{-webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;}}
        @media print {{ body, div, table, tr, th, td {{-webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;}} }}
        </style>
        
        <div style="text-align: right; margin-bottom: 10px;">
            <button onclick="window.print()" style="padding: 8px 16px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">🖨️ 해당 서식 인쇄하기</button>
        </div>
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 680px; margin: auto; background: #fff;">
            {logo_html}
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; clear: both;">
                <h2 style="margin: 0; padding-top: 15px; font-size: 22px; text-decoration: underline;">초 과 / 휴 일 근 무 신 청 및 확 인 서</h2>
                <table style="border-collapse: collapse; text-align: center; font-size: 12px; width: 210px;" border="1">
                    <tr style="height: 20px; background-color: #f2f2f2;">
                        <th rowspan="2" style="width: 25px; background-color: #e6e6e6;">결<br>재</th>
                        <th style="width: 60px;">담 당</th>
                        <th style="width: 60px;">대 리</th>
                        <th style="width: 65px;">센터장</th>
                    </tr>
                    <tr style="height: 50px;">
                        <td></td><td></td><td></td>
                    </tr>
                </table>
            </div>

            <table style="width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px;" border="1">
                <tr style="height: 38px;">
                    <th style="padding: 6px; background: #f9f9f9; width: 20%;">성 명</th>
                    <td style="padding: 6px; width: 30%;">{target_ot_latest['emp_name']} ({target_ot_latest['position']})</td>
                    <th style="padding: 6px; background: #f9f9f9; width: 20%;">소 속</th>
                    <td style="padding: 6px; width: 30%;">{target_ot_latest['dept']}</td>
                </tr>
                <tr style="height: 38px;">
                    <th style="padding: 6px; background: #f9f9f9;">근무구분</th>
                    <td style="padding: 6px;" colspan="3"><b>{target_ot_latest['work_type']}</b> (최종 상태: {target_ot_latest['status']})</td>
                </tr>
                <tr style="height: 38px;">
                    <th style="padding: 6px; background: #f9f9f9;">사전 신청일시</th>
                    <td style="padding: 6px;" colspan="3">{target_ot_latest['work_date']} ({target_ot_latest['start_time']} ~ {target_ot_latest['end_time']}) / {float(target_ot_latest['duration_hours']):.1f}시간</td>
                </tr>
                <tr style="height: 40px; background-color: #ffffcc;">
                    <th style="padding: 6px; background: #fff2cc;">실제 수행 인정</th>
                    <td style="padding: 6px;" colspan="3"><b>실제 근무시간: {target_ot_latest['act_start_time']} ~ {target_ot_latest['act_end_time']} ({float(target_ot_latest['actual_duration_hours']):.1f} 시간)</b></td>
                </tr>
                <tr>
                    <th style="padding: 6px; background: #f9f9f9;">사유 및 업무내용</th>
                    <td style="padding: 10px; height: 70px; vertical-align: top;" colspan="3">
                        <b>[신청 사유]</b> {target_ot_latest['reason']}<br>
                        <b>[실제 수행 내용]</b> {act_reason_disp}
                    </td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 35px; font-size: 14px;">위와 같이 근무를 신청하고 실제 수행 내역을 확인합니다.</p>
            <p style="text-align: center; margin-top: 10px; font-size: 13px;">{target_ot_latest['work_date'][:4]}년 {target_ot_latest['work_date'][5:7]}월 {target_ot_latest['work_date'][8:10]}일</p>
            
            <p style="text-align: right; margin-top: 30px; font-size: 14px; font-weight: bold; padding-right: 10px;">
                신청자 : {target_ot_latest['emp_name']} (인)
            </p>
        </div>
        """
        st.components.v1.html(ot_confirm_template, height=560, scrolling=True)

        st.divider()
        st.subheader("📊 월별 초과/휴일근무 일자별 내역 및 급여 연계")

        current_year = datetime.now().year
        c_y, c_m = st.columns(2)
        with c_y:
            sel_year = st.selectbox("조회 지급연도 선택", range(current_year - 2, current_year + 3), index=2, key="ot_year_sel")
        with c_m:
            sel_month = st.selectbox("조회 지급월 선택 (25일 지급 기준)", range(1, 13), index=datetime.now().month - 1, key="ot_month_sel")

        filter_month = f"{sel_year}-{sel_month:02d}"

        # 당월 25일 지급 기준 산정기간 계산 (전월 25일 ~ 당월 24일)
        pay_date_ref = datetime(sel_year, sel_month, 25).date()
        period_start = (pay_date_ref.replace(day=1) - timedelta(days=1)).replace(day=25)
        period_end = pay_date_ref - timedelta(days=1)
        
        st.caption(f"📅 **[{filter_month} 급여 산정 기간]**: {period_start} ~ {period_end} (전월 25일 ~ 당월 24일)")

        # 산정기간 내 전체 데이터 조회 (근무일자 오름차순)
        ot_m_res = supabase.table("overtime_records")\
            .select("*")\
            .gte("work_date", str(period_start))\
            .lte("work_date", str(period_end))\
            .order("work_date", desc=False)\
            .execute()
        df_ot_month = pd.DataFrame(ot_m_res.data) if ot_m_res.data else pd.DataFrame()

        if df_ot_month.empty:
            st.info(f"💡 해당 산정기간({period_start} ~ {period_end}) 내 등록된 초과/휴일근무 내역이 없습니다.")
        else:
            display_ot_df = df_ot_month[[
                'id', 'work_date', 'emp_id', 'emp_name', 'dept', 'position', 
                'work_type', 'act_start_time', 'act_end_time', 'actual_duration_hours', 'actual_pay', 'status'
            ]].copy()

            display_ot_df.columns = [
                'ID', '근무일자', '사번', '이름', '부서', '직위', 
                '근무구분', '시작시간', '종료시간', '인정시간(h)', '계산수당(원)', '승인상태'
            ]

            st.write(f"**[{filter_month} 지급분] 전체 초과/휴일근무 근무일자별 상세 목록 (총 {len(display_ot_df)}건)**")
            
            display_table_kr(
                display_ot_df, 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "계산수당(원)": st.column_config.NumberColumn("계산수당(원)", format="%d 원"),
                    "인정시간(h)": st.column_config.NumberColumn("인정시간(h)", format="%.1f 시간")
                }
            )

            approved_df = df_ot_month[df_ot_month['status'] == '승인']

            # 월별 초과/휴일근무 목록 엑셀 다운로드
            ot_export = display_ot_df.copy()
            ot_export["인정시간(h)"] = pd.to_numeric(ot_export["인정시간(h)"], errors="coerce").fillna(0.0)
            ot_export["계산수당(원)"] = pd.to_numeric(ot_export["계산수당(원)"], errors="coerce").fillna(0).astype(int)

            ot_excel = io.BytesIO()
            from openpyxl import Workbook
            wb_ot = Workbook()
            ws_ot = wb_ot.active
            ws_ot.title = f"{filter_month}_초과휴일근무"

            ot_thin = Side(style="thin", color="D9DDE3")
            ot_border = Border(left=ot_thin, right=ot_thin, top=ot_thin, bottom=ot_thin)
            ot_center = Alignment(horizontal="center", vertical="center")
            ot_left = Alignment(horizontal="left", vertical="center")
            ot_right = Alignment(horizontal="right", vertical="center")
            ot_title_fill = PatternFill("solid", fgColor="D9EAF7")
            ot_header_fill = PatternFill("solid", fgColor="F2F4F7")
            ot_approved_fill = PatternFill("solid", fgColor="E2F0D9")
            ot_metric_fill = PatternFill("solid", fgColor="FFF2CC")

            ws_ot.merge_cells("A1:L1")
            ws_ot["A1"] = "월별 초과/휴일근무 일자별 내역 및 급여 연계"
            ws_ot["A1"].font = Font(name="맑은 고딕", size=16, bold=True)
            ws_ot["A1"].fill = ot_title_fill
            ws_ot["A1"].alignment = ot_left
            ws_ot.row_dimensions[1].height = 30

            ws_ot.merge_cells("A2:L2")
            ws_ot["A2"] = f"[{filter_month} 급여 산정 기간] {period_start} ~ {period_end} (전월 25일 ~ 당월 24일)"
            ws_ot["A2"].font = Font(name="맑은 고딕", size=10)
            ws_ot["A2"].alignment = ot_left

            ws_ot.merge_cells("A4:L4")
            ws_ot["A4"] = f"[{filter_month} 지급분] 전체 초과/휴일근무 근무일자별 상세 목록 (총 {len(ot_export)}건)"
            ws_ot["A4"].font = Font(name="맑은 고딕", size=11, bold=True)

            for ci, name in enumerate(ot_export.columns, 1):
                c = ws_ot.cell(5, ci, name)
                c.font = Font(name="맑은 고딕", size=10, bold=True)
                c.fill = ot_header_fill
                c.alignment = ot_center
                c.border = ot_border

            for ri, (_, r) in enumerate(ot_export.iterrows(), 6):
                for ci, name in enumerate(ot_export.columns, 1):
                    value = r[name]
                    if name == "인정시간(h)": value = float(value)
                    if name == "계산수당(원)": value = int(value)
                    c = ws_ot.cell(ri, ci, value)
                    c.font = Font(name="맑은 고딕", size=10)
                    c.border = ot_border
                    if str(r["승인상태"]) == "승인":
                        c.fill = ot_approved_fill
                    c.alignment = ot_left if name in ["이름","부서","직위","근무구분"] else (ot_right if name in ["인정시간(h)","계산수당(원)"] else ot_center)
                    if name == "인정시간(h)": c.number_format = '0.0" 시간"'
                    if name == "계산수당(원)": c.number_format = '#,##0" 원"'

            mr = 7 + len(ot_export)
            metrics = [
                ("기간 내 전체 신청 건수", len(df_ot_month), '0" 건"', 1, 4),
                ("최종 승인 건수", len(approved_df), '0" 건"', 5, 8),
                ("승인 총 수당 합계", int(approved_df["actual_pay"].sum()), '#,##0" 원"', 9, 12)
            ]
            for label, value, fmt, sc, ec in metrics:
                ws_ot.merge_cells(start_row=mr, start_column=sc, end_row=mr, end_column=ec)
                ws_ot.merge_cells(start_row=mr+1, start_column=sc, end_row=mr+1, end_column=ec)
                ws_ot.cell(mr, sc, label)
                ws_ot.cell(mr+1, sc, value)
                for rr in [mr, mr+1]:
                    for cc in range(sc, ec+1):
                        ws_ot.cell(rr, cc).fill = ot_metric_fill
                        ws_ot.cell(rr, cc).border = ot_border
                        ws_ot.cell(rr, cc).alignment = ot_center
                ws_ot.cell(mr, sc).font = Font(name="맑은 고딕", size=10, bold=True)
                ws_ot.cell(mr+1, sc).font = Font(name="맑은 고딕", size=16, bold=True)
                ws_ot.cell(mr+1, sc).number_format = fmt

            for i, w in enumerate([8,14,10,12,27,12,29,13,13,14,16,12], 1):
                ws_ot.column_dimensions[get_column_letter(i)].width = w
            ws_ot.freeze_panes = "A6"
            ws_ot.sheet_view.showGridLines = False
            ws_ot.page_setup.orientation = "landscape"
            ws_ot.page_setup.paperSize = ws_ot.PAPERSIZE_A4
            ws_ot.page_setup.fitToWidth = 1
            ws_ot.page_setup.fitToHeight = 0
            ws_ot.sheet_properties.pageSetUpPr.fitToPage = True
            ws_ot.print_area = f"A1:L{mr+1}"
            wb_ot.save(ot_excel)

            ot_b1, ot_b2 = st.columns(2)
            with ot_b1:
                st.download_button(
                    "📥 해당월 목록 엑셀 다운로드 (.xlsx)",
                    ot_excel.getvalue(),
                    file_name=f"{filter_month}_초과휴일근무_일자별내역.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            ot_print_rows = ""
            for _, r in ot_export.iterrows():
                approved_style = "background:#f3faef !important;" if str(r["승인상태"]) == "승인" else ""
                ot_print_rows += f"""<tr style="{approved_style}">
                <td>{r['ID']}</td><td>{r['근무일자']}</td><td>{r['사번']}</td><td>{r['이름']}</td>
                <td style="text-align:left">{r['부서']}</td><td>{r['직위']}</td><td style="text-align:left">{r['근무구분']}</td>
                <td>{r['시작시간']}</td><td>{r['종료시간']}</td><td style="text-align:right">{float(r['인정시간(h)']):.1f} 시간</td>
                <td style="text-align:right">{int(r['계산수당(원)']):,} 원</td><td>{r['승인상태']}</td></tr>"""

            ot_print_html = f"""
            <style>
            * {{box-sizing:border-box;-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important;}}
            body {{font-family:'Malgun Gothic',sans-serif;color:#1f2937;margin:0;}}
            .pbtn {{width:100%;padding:9px;background:#fff;border:1px solid #d1d5db;border-radius:7px;cursor:pointer;font-weight:bold;}}
            .sheet {{padding:8px 2px;background:#fff;}}
            h2 {{font-size:19px;margin:0 0 14px;}}
            .period {{font-size:10px;color:#6b7280;margin-bottom:16px;}}
            .ttl {{font-size:11px;font-weight:bold;margin-bottom:8px;}}
            table {{width:100%;border-collapse:collapse;font-size:8px;}}
            th {{background:#f2f4f7 !important;color:#6b7280;padding:6px 3px;border:1px solid #dfe3e8;white-space:nowrap;}}
            td {{padding:6px 3px;border:1px solid #e5e7eb;text-align:center;white-space:nowrap;}}
            .metrics {{display:flex;gap:12px;margin-top:15px;}}
            .metric {{flex:1;background:#fff2cc !important;border:1px solid #eadf9c;padding:9px;border-radius:5px;}}
            .mv {{font-size:20px;margin-top:6px;}}
            @page {{size:A4 landscape;margin:7mm;}}
            @media print {{
              .no-print {{display:none !important;}}
              *,body,div,table,tr,th,td {{-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important;}}
              .sheet {{padding:0;}}
            }}
            </style>
            <div class="no-print"><button class="pbtn" onclick="window.print()">🖨️ 해당월 목록 인쇄</button></div>
            <div class="sheet">
            <h2>📊 월별 초과/휴일근무 일자별 내역 및 급여 연계</h2>
            <div class="period">📅 [{filter_month} 급여 산정 기간]: {period_start} ~ {period_end} (전월 25일 ~ 당월 24일)</div>
            <div class="ttl">[{filter_month} 지급분] 전체 초과/휴일근무 근무일자별 상세 목록 (총 {len(ot_export)}건)</div>
            <table><thead><tr>{''.join(f'<th>{c}</th>' for c in ot_export.columns)}</tr></thead><tbody>{ot_print_rows}</tbody></table>
            <div class="metrics">
              <div class="metric">기간 내 전체 신청 건수<div class="mv">{len(df_ot_month)} 건</div></div>
              <div class="metric">최종 승인 건수<div class="mv">{len(approved_df)} 건</div></div>
              <div class="metric">승인 총 수당 합계<div class="mv">{int(approved_df['actual_pay'].sum()):,} 원</div></div>
            </div></div>
            """
            with ot_b2:
                st.components.v1.html(ot_print_html, height=55, scrolling=False)

            
            col_stat1, col_stat2, col_stat3 = st.columns(3)
            col_stat1.metric("기간 내 전체 신청 건수", f"{len(df_ot_month)} 건")
            col_stat2.metric("최종 승인 건수", f"{len(approved_df)} 건")
            col_stat3.metric("승인 총 수당 합계", f"{int(approved_df['actual_pay'].sum()):,} 원")

            st.markdown("---")
            
            link_col1, link_col2 = st.columns([3, 1])
            with link_col1:
                st.write(f"💡 아래 버튼을 누르면 해당 산정기간({period_start} ~ {period_end}) 동안 **[승인]** 처리된 초과/휴일근무 수당이 **{filter_month} 통합 급여대장**에 일괄 반영됩니다.")
            with link_col2:
                if st.button("🔄 해당월 급여대장 수당 연계 반영", type="primary", use_container_width=True):
                    if approved_df.empty:
                        st.warning("승인된 근무 내역이 없어 연계할 데이터가 없습니다.")
                    else:
                        emp_totals = approved_df.groupby('emp_id')['actual_pay'].sum().to_dict()
                        pay_date_str = str(pay_date_ref)
                        
                        for emp_id, ot_sum_pay in emp_totals.items():
                            adj_check = supabase.table("monthly_payroll_adjust")\
                                .select("*")\
                                .eq("pay_month", filter_month)\
                                .eq("pay_run_no", 1)\
                                .eq("emp_id", emp_id)\
                                .execute()
                            
                            if adj_check.data:
                                supabase.table("monthly_payroll_adjust")\
                                    .update({"ot_pay": int(ot_sum_pay), "ot_pay_overridden": True})\
                                    .eq("pay_month", filter_month)\
                                    .eq("pay_run_no", 1)\
                                    .eq("emp_id", emp_id)\
                                    .execute()
                            else:
                                emp_info_res = supabase.table("employees").select("*").eq("emp_id", emp_id).execute()
                                if emp_info_res.data:
                                    e = emp_info_res.data[0]
                                    supabase.table("monthly_payroll_adjust").insert({
                                        "pay_month": filter_month,
                                        "pay_run_no": 1,
                                        "pay_run_name": "정기급여",
                                        "pay_date": pay_date_str,
                                        "emp_id": emp_id,
                                        "base_salary": safe_int(e.get("base_salary")),
                                        "ot_pay": int(ot_sum_pay),
                                        "ot_pay_overridden": True,
                                        "family_allowance": safe_int(e.get("family_allowance")),
                                        "non_taxable": safe_int(e.get("non_taxable")),
                                        "other_allowance": safe_int(e.get("other_allowance")),
                                        "national_pension": safe_int(e.get("national_pension")),
                                        "health_insurance": safe_int(e.get("health_insurance")),
                                        "longterm_care": safe_int(e.get("longterm_care")),
                                        "employment_insurance": safe_int(e.get("employment_insurance")),
                                        "income_tax": safe_int(e.get("income_tax")),
                                        "local_tax": safe_int(e.get("local_tax")),
                                        "other_deduction": safe_int(e.get("other_deduction"))
                                    }).execute()

                        st.success(f"✅ [{filter_month} 급여지급분] 산정기간({period_start} ~ {period_end}) 승인 내역({len(approved_df)}건)이 통합 급여대장에 성공적으로 연계·반영되었습니다!")
                        st.rerun()

# -------------------------------------------------------------------
# TAB 4: 개인별 연차 관리 & 전 직원 연차 요약표
# -------------------------------------------------------------------
if active_tab == 4:
    st.header("🌴 개인별 연차 관리 & 전 직원 연차 요약표")
    
    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()
    df_emp = scope_employee_master(df_emp)

    df_emp = scope_employee_master(df_emp)
    if df_emp.empty:
        st.warning("등록된 직원이 없다.")
    else:
        col_l1, col_l2 = st.columns([1, 1])
        
        with col_l1:
            st.subheader("1. 연차/휴가 신청서 작성")
            emp_leave_list = df_emp['emp_name'] + " (" + df_emp['position'] + " / " + df_emp['emp_id'] + ")"
            selected_l_emp = st.selectbox("직원 선택", emp_leave_list, key="leave_emp_select")
            selected_l_id = selected_l_emp.split("/")[-1].replace(")", "").strip()
            l_emp_info = df_emp[df_emp['emp_id'] == selected_l_id].iloc[0]

            leave_type = st.selectbox("휴가 종류", ["연차 (1일)", "오전반차 (0.5일)", "오후반차 (0.5일)", "병가", "경조휴가", "특별휴가"])
            
            l_start_date = st.date_input("휴가 시작일", datetime.now(), key="l_s_date")
            l_end_date = st.date_input("휴가 종료일", datetime.now(), key="l_e_date")
            
            if "반차" in leave_type:
                used_days = 0.5
            else:
                used_days = float((l_end_date - l_start_date).days + 1)

            leave_reason = st.text_area("휴가 사유", key="l_reason")

            if st.button("연차 신청서 제출 및 DB 저장"):
                leave_data = {
                    "apply_dt": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "emp_id": l_emp_info['emp_id'], "emp_name": l_emp_info['emp_name'],
                    "dept": l_emp_info['dept'], "position": l_emp_info['position'],
                    "leave_type": leave_type, "start_date": str(l_start_date),
                    "end_date": str(l_end_date), "used_days": used_days, "reason": leave_reason
                }
                supabase.table("leave_records").insert(leave_data).execute()
                st.success("연차 신청 내역이 Supabase DB에 저장되었다.")
                st.rerun()

        with col_l2:
            st.subheader("2. 개인별 연차 현황 요약 및 삭제")
            l_res = supabase.table("leave_records").select("*").eq("emp_id", l_emp_info['emp_id']).order("id", desc=True).execute()
            df_leave_all = pd.DataFrame(l_res.data) if l_res.data else pd.DataFrame()

            used_annual = df_leave_all[df_leave_all['leave_type'].str.contains("연차|반차", na=False)]['used_days'].sum() if not df_leave_all.empty else 0.0
            total_annual = l_emp_info['total_annual_leave']
            remaining_annual = total_annual - used_annual

            m1, m2, m3 = st.columns(3)
            m1.metric("총 부여 연차", f"{total_annual} 일")
            m2.metric("사용 연차", f"{used_annual} 일")
            m3.metric("잔여 연차", f"{remaining_annual} 일")

            st.write(f"**[{l_emp_info['emp_name']}] 개인 신청 이력**")
            if not df_leave_all.empty:
                display_table_kr(df_leave_all[['id', 'apply_dt', 'leave_type', 'start_date', 'end_date', 'used_days', 'reason']], use_container_width=True)

                st.divider()
                st.write("**🗑️ 연차 신청 내역 삭제**")
                del_leave_options = [f"ID {r['id']} | [{r['start_date']}] {r['leave_type']} ({r['used_days']}일)" for _, r in df_leave_all.iterrows()]
                selected_del_str = st.selectbox("삭제할 내역 선택", del_leave_options, key="del_leave_sel")
                selected_del_id = int(selected_del_str.split("|")[0].replace("ID", "").strip())

                if st.button("선택한 연차 내역 삭제", type="primary"):
                    supabase.table("leave_records").delete().eq("id", selected_del_id).execute()
                    st.success("해당 연차 내역이 정상적으로 삭제되었다.")
                    st.rerun()

        st.divider()
        st.header("📋 센터 등록 직원 전체 연차 내역 요약표")
        
        all_l_res = supabase.table("leave_records").select("*").execute()
        df_all_leaves = pd.DataFrame(all_l_res.data) if all_l_res.data else pd.DataFrame()

        summary_rows = []
        for idx, emp_row in df_emp.iterrows():
            emp_l_records = df_all_leaves[df_all_leaves['emp_id'] == emp_row['emp_id']] if not df_all_leaves.empty else pd.DataFrame()
            u_annual = emp_l_records[emp_l_records['leave_type'].str.contains("연차|반차", na=False)]['used_days'].sum() if not emp_l_records.empty else 0.0
            tot_annual = emp_row['total_annual_leave']
            rem_annual = tot_annual - u_annual
            
            summary_rows.append({
                "사번": emp_row['emp_id'], "이름": emp_row['emp_name'], "부서": emp_row['dept'], "직위": emp_row['position'],
                "총 부여 연차": tot_annual, "사용 연차": u_annual, "잔여 연차": rem_annual,
                "사용률 (%)": round((u_annual / tot_annual * 100), 1) if tot_annual > 0 else 0.0
            })

        df_summary_all = pd.DataFrame(summary_rows)

        output_leave = io.BytesIO()
        with pd.ExcelWriter(output_leave, engine='openpyxl') as writer:
            excel_view_df(df_summary_all).to_excel(writer, index=False, sheet_name="전직원_연차_요약")
            worksheet = writer.sheets["전직원_연차_요약"]
            style_excel_sheet(worksheet)

            header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            header_font = Font(name="맑은 고딕", size=11, bold=True)
            body_font = Font(name="맑은 고딕", size=10)

            thin_border = Border(
                left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
            )
            align_center = Alignment(horizontal='center', vertical='center')
            align_right = Alignment(horizontal='right', vertical='center')

            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.border = thin_border
                    if cell.row == 1:
                        cell.fill = header_fill; cell.font = header_font; cell.alignment = align_center
                    else:
                        cell.font = body_font
                        cell.alignment = align_right if cell.column in [5, 6, 7, 8] else align_center

            for col in worksheet.columns:
                max_len = max(sum(2 if ord(c) > 127 else 1 for c in str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

        excel_leave_data = output_leave.getvalue()

        st.download_button(
            label="📥 전 직원 연차 내역 요약표 엑셀 다운로드 (.xlsx)",
            data=excel_leave_data, file_name=f"전직원_연차요약_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        display_table_kr(df_summary_all, use_container_width=True)

# -------------------------------------------------------------------
# TAB 5: 연차 신청서 독립 출력 탭
# -------------------------------------------------------------------
if active_tab == 5:
    st.header("🖨️ 휴가 (연차) 신청서 인쇄")
    
    l_records_res = supabase.table("leave_records").select("*").order("id", desc=True).execute()
    df_leave_records = pd.DataFrame(l_records_res.data) if l_records_res.data else pd.DataFrame()

    df_leave_records = scope_dataframe_to_current_employee(df_leave_records)
    if df_leave_records.empty:
        st.info("등록된 연차/휴가 신청 내역이 없다.")
    else:
        leave_options = [f"[{r['start_date']}] {r['emp_name']} {r['position']} - {r['leave_type']}" for _, r in df_leave_records.iterrows()]
        selected_l_index = st.selectbox("출력할 연차 신청서 선택", range(len(leave_options)), format_func=lambda x: leave_options[x])
        target_l = df_leave_records.iloc[selected_l_index]

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 35px; float: left;">' if st.session_state.logo_b64 else ''

        leave_template = f"""<style>
        * {{-webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;}}
        @media print {{ body, div, table, tr, th, td {{-webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;}} }}
        </style>
        
        <div style="text-align: right; margin-bottom: 10px;">
            <button onclick="window.print()" style="padding: 8px 16px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">🖨️ 해당 서식 인쇄하기</button>
        </div>
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 680px; margin: auto; background: #fff;">
            {logo_html}
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; clear: both;">
                <h2 style="margin: 0; padding-top: 15px; font-size: 24px; text-decoration: underline;">휴 가 (연 차) 신 청 서</h2>
                <table style="border-collapse: collapse; text-align: center; font-size: 12px; width: 210px;" border="1">
                    <tr style="height: 20px; background-color: #f2f2f2;">
                        <th rowspan="2" style="width: 25px; background-color: #e6e6e6;">결<br>재</th>
                        <th style="width: 60px;">담 당</th>
                        <th style="width: 60px;">대 리</th>
                        <th style="width: 65px;">센터장</th>
                    </tr>
                    <tr style="height: 50px;">
                        <td></td><td></td><td></td>
                    </tr>
                </table>
            </div>

            <table style="width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 14px;" border="1">
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9; width: 20%;">성 명</th>
                    <td style="padding: 8px; width: 30%;">{target_l['emp_name']} ({target_l['position']})</td>
                    <th style="padding: 8px; background: #f9f9f9; width: 20%;">소 속</th>
                    <td style="padding: 8px; width: 30%;">{target_l['dept']}</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">휴가구분</th>
                    <td style="padding: 8px;" colspan="3">{target_l['leave_type']} (사용일수: {target_l['used_days']}일)</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">휴가기간</th>
                    <td style="padding: 8px;" colspan="3">{target_l['start_date']} ~ {target_l['end_date']}</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background: #f9f9f9;">휴가사유</th>
                    <td style="padding: 12px; height: 80px; vertical-align: top;" colspan="3">{target_l['reason']}</td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 50px; font-size: 15px;">위와 같이 휴가(연차)를 신청합니다.</p>
            <p style="text-align: center; margin-top: 15px; font-size: 13px;">{target_l['apply_dt'][:10]}</p>
            
            <p style="text-align: right; margin-top: 40px; font-size: 15px; font-weight: bold; padding-right: 10px;">
                신청인: {target_l['emp_name']} (인)
            </p>
        </div>
        """
        st.components.v1.html(leave_template, height=560, scrolling=True)

# -------------------------------------------------------------------
# TAB 6: 통합 급여대장 (수정 및 엑셀)
# -------------------------------------------------------------------
if active_tab == 6:
    st.header("📊 통합 급여대장 (수정 및 엑셀)")

    run_col1, run_col2, run_col3 = st.columns([1, 1, 2])
    with run_col1:
        pay_date = st.date_input("지급일 선택", datetime.now(), key="payroll_date")
    with run_col2:
        pay_run_no = st.selectbox("급여대장 차수", [1, 2], format_func=lambda x: f"{x}차 대장", key="pay_run_no")
    with run_col3:
        default_run_name = "정기급여" if pay_run_no == 1 else "명절상여금"
        pay_run_name = st.text_input("급여대장 명칭", value=default_run_name, key=f"pay_run_name_{pay_run_no}").strip() or default_run_name
    pay_month = pay_date.strftime("%Y-%m")
    st.caption(f"현재 편집 대상: {pay_month} · {pay_run_no}차 · {pay_run_name} · 지급일 {pay_date.isoformat()}")

    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()
    df_emp = scope_employee_master(df_emp)

    df_emp = scope_employee_master(df_emp)
    ot_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()
    df_ot = scope_employee_master(df_ot)

    multi_run_ready = True
    try:
        adj_res = supabase.table("monthly_payroll_adjust").select("*").eq("pay_month", pay_month).eq("pay_run_no", pay_run_no).execute()
        df_adjust = pd.DataFrame(adj_res.data) if adj_res.data else pd.DataFrame()
        df_adjust = scope_employee_master(df_adjust)
    except Exception:
        multi_run_ready = False
        df_adjust = pd.DataFrame()
        st.error("먼저 V1.3 다중 급여대장 SQL 설정 파일을 Supabase에서 실행해 주세요.")

    closing_table_ready = True
    current_closing = None
    try:
        closing_res = supabase.table("payroll_monthly_closings").select("*").eq("pay_month", pay_month).eq("pay_run_no", pay_run_no).execute()
        current_closing = closing_res.data[0] if closing_res.data else None
    except Exception:
        closing_table_ready = False

    if df_emp.empty:
        st.warning("등록된 직원이 없다.")
    else:
        calculated_rows = []
        no = 1
        
        ot_start_date = (pay_date.replace(day=1) - timedelta(days=1)).replace(day=25)
        ot_end_date = pay_date - timedelta(days=1)

        for idx, emp in df_emp.iterrows():
            adj_match = df_adjust[df_adjust['emp_id'] == emp['emp_id']] if not df_adjust.empty else pd.DataFrame()

            emp_ot = df_ot[
                (df_ot['emp_id'] == emp['emp_id']) & 
                (df_ot['work_date'] >= str(ot_start_date)) & 
                (df_ot['work_date'] <= str(ot_end_date))
            ] if not df_ot.empty else pd.DataFrame()

            calculated_ot_pay = int(emp_ot['actual_pay'].sum()) if not emp_ot.empty else 0

            if not adj_match.empty:
                adj = adj_match.iloc[0]
                base = adj['base_salary']
                ot_pay = adj['ot_pay'] if adj.get('ot_pay_overridden', False) else calculated_ot_pay
                family = adj['family_allowance']
                holiday_bonus = safe_int(adj.get('holiday_bonus'))
                non_tax = adj['non_taxable']
                other_allow = adj['other_allowance']
                emp_national = adj['national_pension']
                emp_health = adj['health_insurance']
                emp_longterm = adj['longterm_care']
                emp_employment = adj['employment_insurance']
                emp_income_tax = safe_int(adj.get('income_tax'))
                emp_local_tax = safe_int(adj.get('local_tax'))
                other_deduct = adj['other_deduction']
            else:
                ot_pay = calculated_ot_pay
                base = emp['base_salary']
                family = emp['family_allowance']
                holiday_bonus = 0
                non_tax = emp['non_taxable']
                other_allow = emp['other_allowance']
                other_deduct = emp['other_deduction']

                emp_national = safe_int(emp.get('national_pension'))
                emp_health = safe_int(emp.get('health_insurance'))
                emp_longterm = safe_int(emp.get('longterm_care'))
                emp_employment = safe_int(emp.get('employment_insurance'))
                emp_income_tax = safe_int(emp.get('income_tax'))
                emp_local_tax = safe_int(emp.get('local_tax'))

                if pay_run_no == 2:
                    base = ot_pay = family = non_tax = other_allow = other_deduct = 0
                    holiday_bonus = safe_int(emp.get('holiday_bonus'))
                    emp_national = emp_health = emp_longterm = emp_employment = 0
                    emp_income_tax = emp_local_tax = 0

            biz_national = safe_int(emp.get('employer_national_pension'))
            biz_health = safe_int(emp.get('employer_health_insurance'))
            biz_longterm = safe_int(emp.get('employer_longterm_care'))
            biz_employment = safe_int(emp.get('employer_employment_insurance'))
            biz_industrial = safe_int(emp.get('employer_industrial_insurance'))
            retirement_accrual = safe_int(emp.get('retirement_accrual'))
            if pay_run_no == 2 and adj_match.empty:
                biz_national = biz_health = biz_longterm = biz_employment = biz_industrial = retirement_accrual = 0

            if not adj_match.empty:
                biz_national = safe_int(adj.get('employer_national_pension')) if pd.notna(adj.get('employer_national_pension')) else biz_national
                biz_health = safe_int(adj.get('employer_health_insurance')) if pd.notna(adj.get('employer_health_insurance')) else biz_health
                biz_longterm = safe_int(adj.get('employer_longterm_care')) if pd.notna(adj.get('employer_longterm_care')) else biz_longterm
                biz_employment = safe_int(adj.get('employer_employment_insurance')) if pd.notna(adj.get('employer_employment_insurance')) else biz_employment
                biz_industrial = safe_int(adj.get('employer_industrial_insurance')) if pd.notna(adj.get('employer_industrial_insurance')) else biz_industrial
                retirement_accrual = safe_int(adj.get('retirement_accrual')) if pd.notna(adj.get('retirement_accrual')) else retirement_accrual

            calculated_rows.append({
                "No": no, "사번": emp['emp_id'], "이름": emp['emp_name'], "생년월일": emp['birth_date'], "부서": emp['dept'], "직위": emp['position'], "호봉": emp['hobong'],
                "기본급": base, "초과수당(승인)": ot_pay, "가족수당": family, "명절상여": holiday_bonus, "비과세": non_tax, "기타수당": other_allow,
                "국민연금(본인)": emp_national, "건강보험(본인)": emp_health, "장기요양(본인)": emp_longterm, "고용보험(본인)": emp_employment,
                "소득세": emp_income_tax, "지방소득세": emp_local_tax, "기타공제": other_deduct,
                "국민연금(사업자)": biz_national, "건강보험(사업자)": biz_health, "장기요양(사업자)": biz_longterm, "고용보험(사업자)": biz_employment, "산재보험(사업자)": biz_industrial,
                "퇴직적립금": retirement_accrual
            })
            no += 1

        df_calc = pd.DataFrame(calculated_rows)

        st.subheader(f"✏️ {pay_month} {pay_run_no}차 {pay_run_name} 엑셀형 편집기")
        st.info("셀을 클릭하여 직접 수정하거나 엑셀의 여러 셀을 붙여넣을 수 있습니다.")

        identity_columns = ["No", "사번", "이름", "생년월일", "부서", "직위", "호봉"]
        amount_columns = [
            "기본급", "초과수당(승인)", "가족수당", "명절상여", "비과세", "기타수당",
            "국민연금(본인)", "건강보험(본인)", "장기요양(본인)", "고용보험(본인)",
            "소득세", "지방소득세", "기타공제", "국민연금(사업자)",
            "건강보험(사업자)", "장기요양(사업자)", "고용보험(사업자)",
            "산재보험(사업자)", "퇴직적립금"
        ]
        column_config = {
            "No": st.column_config.NumberColumn("No", width="small"),
            "사번": st.column_config.TextColumn("사번", width="small"),
            "이름": st.column_config.TextColumn("이름", width="small"),
            "생년월일": st.column_config.TextColumn("생년월일", width="medium"),
            "부서": st.column_config.TextColumn("부서", width="small"),
            "직위": st.column_config.TextColumn("직위", width="small"),
            "호봉": st.column_config.TextColumn("호봉", width="small")
        }
        for col in amount_columns:
            column_config[col] = st.column_config.NumberColumn(
                col, min_value=0, step=10, format="localized", width="medium",
                help="10원 단위로 입력할 수 있습니다."
            )

        reset_version_key = f"payroll_editor_reset_version_{pay_month}_{pay_run_no}"
        if reset_version_key not in st.session_state:
            st.session_state[reset_version_key] = 0
        editor_key = f"payroll_excel_editor_{pay_month}_{pay_run_no}_{st.session_state[reset_version_key]}"
        edited_payroll = st.data_editor(
            df_calc,
            key=editor_key,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=identity_columns,
            column_config=column_config,
            height=min(650, max(220, 38 * (len(df_calc) + 2)))
        )

        invalid_cells = []
        for col in amount_columns:
            numeric_values = pd.to_numeric(edited_payroll[col], errors="coerce")
            bad_rows = edited_payroll[numeric_values.isna() | (numeric_values < 0)]
            invalid_cells.extend([f"{name} - {col}" for name in bad_rows["이름"].astype(str).tolist()])
            edited_payroll[col] = numeric_values.fillna(0).round().astype(int)

        if invalid_cells:
            st.error("금액은 0 이상의 숫자로 입력해 주세요: " + ", ".join(invalid_cells[:8]) + (" 외" if len(invalid_cells) > 8 else ""))

        gross_series = edited_payroll[["기본급", "초과수당(승인)", "가족수당", "명절상여", "비과세", "기타수당"]].sum(axis=1)
        deduction_series = edited_payroll[["국민연금(본인)", "건강보험(본인)", "장기요양(본인)", "고용보험(본인)", "소득세", "지방소득세", "기타공제"]].sum(axis=1)
        employer_series = edited_payroll[["국민연금(사업자)", "건강보험(사업자)", "장기요양(사업자)", "고용보험(사업자)", "산재보험(사업자)"]].sum(axis=1)
        net_series = gross_series - deduction_series
        if (net_series < 0).any():
            invalid_cells.append("공제합계가 급여총액을 초과한 직원")
            st.error("공제합계가 급여총액보다 큰 직원이 있습니다. 실지급액이 음수가 되지 않도록 확인해 주세요.")

        metric_cols = st.columns(5)
        metric_cols[0].metric("급여총액", f"{int(gross_series.sum()):,}원")
        metric_cols[1].metric("근로자 공제합계", f"{int(deduction_series.sum()):,}원")
        metric_cols[2].metric("실지급액", f"{int(net_series.sum()):,}원")
        metric_cols[3].metric("사업주 부담보험", f"{int(employer_series.sum()):,}원")
        metric_cols[4].metric("퇴직적립금", f"{int(edited_payroll['퇴직적립금'].sum()):,}원")

        action_col1, action_col2 = st.columns([3, 1])
        with action_col1:
            save_adjustments = st.button(
                "💾 수정사항 개별 급여명세서에 연동 저장",
                disabled=bool(invalid_cells),
                use_container_width=True
            )
        with action_col2:
            if st.button("↩ 저장값 다시 불러오기", use_container_width=True):
                st.session_state[reset_version_key] += 1
                st.rerun()

        if save_adjustments:
            for idx, r in edited_payroll.iterrows():
                pay_adj_data = {
                    "pay_month": pay_month, "pay_run_no": pay_run_no, "pay_run_name": pay_run_name,
                    "pay_date": pay_date.isoformat(), "emp_id": r['사번'], "base_salary": int(r['기본급']),
                    "ot_pay": int(r['초과수당(승인)']), "family_allowance": int(r['가족수당']),
                    "holiday_bonus": int(r['명절상여']),
                    "non_taxable": int(r['비과세']), "other_allowance": int(r['기타수당']),
                    "national_pension": int(r['국민연금(본인)']), "health_insurance": int(r['건강보험(본인)']),
                    "longterm_care": int(r['장기요양(본인)']), "employment_insurance": int(r['고용보험(본인)']),
                    "income_tax": int(r['소득세']), "local_tax": int(r['지방소득세']), "other_deduction": int(r['기타공제'])
                    , "ot_pay_overridden": True,
                    "employer_national_pension": int(r['국민연금(사업자)']),
                    "employer_health_insurance": int(r['건강보험(사업자)']),
                    "employer_longterm_care": int(r['장기요양(사업자)']),
                    "employer_employment_insurance": int(r['고용보험(사업자)']),
                    "employer_industrial_insurance": int(r['산재보험(사업자)']),
                    "retirement_accrual": int(r['퇴직적립금'])
                }
                supabase.table("monthly_payroll_adjust").upsert(
                    pay_adj_data, on_conflict="pay_month,pay_run_no,emp_id"
                ).execute()
            st.success(f"{pay_month} {pay_run_no}차 {pay_run_name} 수정 수구가 저장되었습니다.")

        payroll_html_rows = ""
        sum_base = sum_ot = sum_family = sum_holiday = sum_nontax = sum_gross = 0
        sum_nat = sum_hea = sum_long = sum_emp = sum_inc = sum_loc = sum_other_d = sum_deduct_tot = sum_net = 0
        sum_b_nat = sum_b_hea = sum_b_long = sum_b_emp = sum_b_ind = sum_b_tot = sum_retire = 0

        for idx, row in edited_payroll.iterrows():
            total_gross = row['기본급'] + row['초과수당(승인)'] + row['가족수당'] + row['명절상여'] + row['비과세'] + row['기타수당']
            emp_deduction_total = row['국민연금(본인)'] + row['건강보험(본인)'] + row['장기요양(본인)'] + row['고용보험(본인)'] + row['소득세'] + row['지방소득세'] + row['기타공제']
            net_pay = total_gross - emp_deduction_total
            biz_deduction_total = row['국민연금(사업자)'] + row['건강보험(사업자)'] + row['장기요양(사업자)'] + row['고용보험(사업자)'] + row['산재보험(사업자)']

            sum_base += row['기본급']; sum_ot += row['초과수당(승인)']; sum_family += row['가족수당']; sum_holiday += row['명절상여']; sum_nontax += row['비과세']; sum_gross += total_gross
            sum_nat += row['국민연금(본인)']; sum_hea += row['건강보험(본인)']; sum_long += row['장기요양(본인)']; sum_emp += row['고용보험(본인)']
            sum_inc += row['소득세']; sum_loc += row['지방소득세']; sum_other_d += row['기타공제']; sum_deduct_tot += emp_deduction_total; sum_net += net_pay
            sum_b_nat += row['국민연금(사업자)']; sum_b_hea += row['건강보험(사업자)']; sum_b_long += row['장기요양(사업자)']; sum_b_emp += row['고용보험(사업자)']; sum_b_ind += row['산재보험(사업자)']; sum_b_tot += biz_deduction_total; sum_retire += row['퇴직적립금']

            payroll_html_rows += f"""
            <tr>
                <td>{row['No']}</td><td>{row['이름']}</td><td>{row['생년월일']}</td><td>{row['호봉']}</td>
                <td style="text-align:right;">{row['기본급']:,}</td>
                <td style="text-align:right;">{row['초과수당(승인)']:,}</td>
                <td style="text-align:right;">{row['가족수당']:,}</td>
                <td style="text-align:right;">{row['명절상여']:,}</td>
                <td style="text-align:right;">{row['비과세']:,}</td>
                <td style="text-align:right; font-weight:bold;">{total_gross:,}</td>
                <td style="text-align:right;">{row['국민연금(본인)']:,}</td>
                <td style="text-align:right;">{row['건강보험(본인)']:,}</td>
                <td style="text-align:right;">{row['장기요양(본인)']:,}</td>
                <td style="text-align:right;">{row['고용보험(본인)']:,}</td>
                <td style="text-align:right;">{row['소득세']:,}</td>
                <td style="text-align:right;">{row['지방소득세']:,}</td>
                <td style="text-align:right; font-weight:bold;">{emp_deduction_total:,}</td>
                <td style="text-align:right; font-weight:bold; background-color:#fffae6;">{net_pay:,}</td>
                <td style="text-align:right;">{row['국민연금(사업자)']:,}</td>
                <td style="text-align:right;">{row['건강보험(사업자)']:,}</td>
                <td style="text-align:right;">{row['장기요양(사업자)']:,}</td>
                <td style="text-align:right;">{row['고용보험(사업자)']:,}</td>
                <td style="text-align:right;">{row['산재보험(사업자)']:,}</td>
                <td style="text-align:right; font-weight:bold;">{biz_deduction_total:,}</td>
                <td style="text-align:right;">{row['퇴직적립금']:,}</td>
            </tr>
            """

        st.divider()
        # v20.4 급여대장 주요 합계
        show_payroll_summary(edited_payroll, "ledger")
        
        st.subheader("🔒 월 급여 확정 및 회계자료")
        snapshot, accounting_export = build_payroll_snapshot(pay_month, pay_date, edited_payroll, pay_run_no, pay_run_name)
        current_hash = payload_hash(snapshot, accounting_export)

        if not closing_table_ready or not multi_run_ready:
            st.warning("먼저 제공된 V1.4 SQL 설정 파일을 Supabase에서 실행해야 급여 확정 기능을 사용할 수 있습니다.")
        elif current_closing and current_closing.get("status") == "finalized":
            saved_hash = current_closing.get("content_hash", "")
            if saved_hash == current_hash:
                st.success(f"{pay_month} {pay_run_no}차 {pay_run_name}이 확정되었습니다. (수정확정 {current_closing.get('revision', 1)}차)")
            else:
                st.warning("확정 후 급여대장 값이 변경되었습니다. 변경 내용을 저장한 뒤 재확정해 주세요.")
        else:
            st.info("현재 월은 아직 확정되지 않았습니다. 편집값을 확인한 뒤 확정해 주세요.")

        confirm_col, cancel_col = st.columns(2)
        with confirm_col:
            confirm_label = "🔒 월 급여 확정" if not current_closing else "🔁 변경 내용 재확정"
            if st.button(confirm_label, disabled=(not closing_table_ready or not multi_run_ready or bool(invalid_cells)), use_container_width=True):
                for _, r in edited_payroll.iterrows():
                    supabase.table("monthly_payroll_adjust").upsert({
                        "pay_month": pay_month, "pay_run_no": pay_run_no, "pay_run_name": pay_run_name,
                        "pay_date": pay_date.isoformat(), "emp_id": r['사번'],
                        "base_salary": safe_int(r['기본급']), "ot_pay": safe_int(r['초과수당(승인)']),
                        "ot_pay_overridden": True, "family_allowance": safe_int(r['가족수당']),
                        "holiday_bonus": safe_int(r['명절상여']),
                        "non_taxable": safe_int(r['비과세']), "other_allowance": safe_int(r['기타수당']),
                        "national_pension": safe_int(r['국민연금(본인)']), "health_insurance": safe_int(r['건강보험(본인)']),
                        "longterm_care": safe_int(r['장기요양(본인)']), "employment_insurance": safe_int(r['고용보험(본인)']),
                        "income_tax": safe_int(r['소득세']), "local_tax": safe_int(r['지방소득세']),
                        "other_deduction": safe_int(r['기타공제']),
                        "employer_national_pension": safe_int(r['국민연금(사업자)']),
                        "employer_health_insurance": safe_int(r['건강보험(사업자)']),
                        "employer_longterm_care": safe_int(r['장기요양(사업자)']),
                        "employer_employment_insurance": safe_int(r['고용보험(사업자)']),
                        "employer_industrial_insurance": safe_int(r['산재보험(사업자)']),
                        "retirement_accrual": safe_int(r['퇴직적립금'])
                    }, on_conflict="pay_month,pay_run_no,emp_id").execute()

                revision = safe_int(current_closing.get("revision", 0)) + 1 if current_closing else 1
                supabase.table("payroll_monthly_closings").upsert({
                    "pay_month": pay_month, "pay_run_no": pay_run_no, "pay_run_name": pay_run_name,
                    "pay_date": pay_date.isoformat(),
                    "status": "finalized", "revision": revision,
                    "payroll_data": snapshot, "accounting_export": accounting_export,
                    "content_hash": current_hash, "finalized_at": datetime.now().isoformat(),
                    "updated_at": datetime.now().isoformat()
                }, on_conflict="pay_month,pay_run_no").execute()
                st.success(f"{pay_month} {pay_run_no}차 {pay_run_name}을 수정확정 {revision}차로 저장했습니다.")
                st.rerun()

        with cancel_col:
            can_cancel = bool(current_closing and current_closing.get("status") == "finalized")
            if st.button("🔓 확정 취소", disabled=not can_cancel, use_container_width=True):
                supabase.table("payroll_monthly_closings").update({
                    "status": "cancelled", "updated_at": datetime.now().isoformat()
                }).eq("pay_month", pay_month).eq("pay_run_no", pay_run_no).execute()
                st.warning(f"{pay_month} {pay_run_no}차 급여 확정을 취소했습니다.")
                st.rerun()

        export_ready = bool(
            current_closing and current_closing.get("status") == "finalized"
            and current_closing.get("content_hash") == current_hash
        )
        if export_ready:
            import json
            export_bytes = json.dumps(
                current_closing["accounting_export"], ensure_ascii=False, indent=2
            ).encode("utf-8")
            st.download_button(
                "📤 확정 회계자료 JSON 다운로드", data=export_bytes,
                file_name=f"급여회계자료_{pay_month}_{pay_run_no}차_{pay_run_name}_확정{current_closing.get('revision', 1)}차.json",
                mime="application/json", use_container_width=True
            )
            st.caption("회계 지출 합계: " + f"{current_closing['accounting_export']['accounting_total']:,}원 · 근로자 공제액은 중복 합산하지 않습니다.")

        summary_html_row = f"""
        <tr style="background-color: #e6f2ff; font-weight: bold;">
            <td colspan="4">합 계</td>
            <td style="text-align:right;">{sum_base:,}</td>
            <td style="text-align:right;">{sum_ot:,}</td>
            <td style="text-align:right;">{sum_family:,}</td>
            <td style="text-align:right;">{sum_holiday:,}</td>
            <td style="text-align:right;">{sum_nontax:,}</td>
            <td style="text-align:right;">{sum_gross:,}</td>
            <td style="text-align:right;">{sum_nat:,}</td>
            <td style="text-align:right;">{sum_hea:,}</td>
            <td style="text-align:right;">{sum_long:,}</td>
            <td style="text-align:right;">{sum_emp:,}</td>
            <td style="text-align:right;">{sum_inc:,}</td>
            <td style="text-align:right;">{sum_loc:,}</td>
            <td style="text-align:right;">{sum_deduct_tot:,}</td>
            <td style="text-align:right; background-color:#ffe680;">{sum_net:,}</td>
            <td style="text-align:right;">{sum_b_nat:,}</td>
            <td style="text-align:right;">{sum_b_hea:,}</td>
            <td style="text-align:right;">{sum_b_long:,}</td>
            <td style="text-align:right;">{sum_b_emp:,}</td>
            <td style="text-align:right;">{sum_b_ind:,}</td>
            <td style="text-align:right;">{sum_b_tot:,}</td>
            <td style="text-align:right;">{sum_retire:,}</td>
        </tr>
        """

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            excel_view_df(edited_payroll).to_excel(writer, index=False, sheet_name=f"{pay_month}_{pay_run_no}차")
            _sv=payroll_summary_values(edited_payroll)
            pd.DataFrame([{"항목":k,"합계금액":v} for k,v in _sv.items()]).to_excel(
                writer,index=False,sheet_name="급여주요합계"
            )
            worksheet = writer.sheets[f"{pay_month}_{pay_run_no}차"]
            style_excel_sheet(worksheet)
            if "급여주요합계" in writer.sheets:
                style_excel_sheet(writer.sheets["급여주요합계"])
            for col_idx, col_name in enumerate(edited_payroll.columns, 1):
                if col_name in amount_columns:
                    for cell in worksheet.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                        cell[0].number_format = '#,##0'
        excel_data = output.getvalue()

        st.download_button(
            label="📥 통합 급여대장 엑셀 다운로드 (.xlsx)",
            data=excel_data, file_name=f"통합급여대장_{pay_month}_{pay_run_no}차_{pay_run_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 28px; float: left;">' if st.session_state.logo_b64 else ''

        payroll_template = f"""
        <div style="font-family: 'Malgun Gothic', sans-serif; font-size: 11px; width: 100%; overflow-x: auto;">
            {logo_html}
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px; clear: both;">
                <div style="font-size: 14px; font-weight: bold;">지급일 &nbsp;&nbsp;&nbsp; {pay_date}</div>
                <div>(단위: 원 / 원단위 절사)</div>
            </div>
            
            <table border="1" style="width: 100%; border-collapse: collapse; text-align: center;" cellpadding="3">
                <thead>
                    <tr style="background-color: #ffffcc;">
                        <th rowspan="3" style="width: 25px;">No</th>
                        <th rowspan="3" style="width: 50px;">이름</th>
                        <th rowspan="3" style="width: 65px;">생년월일</th>
                        <th rowspan="3" style="width: 40px;">호봉</th>
                        <th colspan="5">지급 내역</th>
                        <th rowspan="3">급여총액</th>
                        <th colspan="7">근로자 본인 부담금</th>
                        <th rowspan="3" style="background-color: #fff2cc;">실지급액</th>
                        <th colspan="6">사업자 부담 사회보험금</th>
                        <th rowspan="3">사업주부담<br>퇴직적립금</th>
                    </tr>
                    <tr style="background-color: #ffffcc;">
                        <th rowspan="2">기본급</th>
                        <th rowspan="2">초과수당</th>
                        <th rowspan="2">가족수당</th>
                        <th rowspan="2">명절상여</th>
                        <th rowspan="2">비과세</th>
                        <th>국민</th><th>건강</th><th>장기요양</th><th>고용</th><th>소득세</th><th>지방세</th>
                        <th rowspan="2">공제합계</th>
                        <th>국민</th><th>건강</th><th>장기요양</th><th>고용</th><th>산재</th>
                        <th rowspan="2">사업자합계</th>
                    </tr>
                    <tr style="background-color: #ffffcc;">
                        <th>4.75%</th><th>3.595%</th><th>12.95%</th><th>0.90%</th><th>간이세액</th><th>10%</th>
                        <th>4.75%</th><th>3.595%</th><th>12.95%</th><th>1.15%</th><th>7.26%</th>
                    </tr>
                </thead>
                <tbody>
                    {summary_html_row}
                    {payroll_html_rows}
                </tbody>
            </table>
        </div>
        """
        st.components.v1.html(payroll_template, height=520, scrolling=True)

# -------------------------------------------------------------------
# TAB 7: 개별 급여명세서 인쇄
# -------------------------------------------------------------------
if active_tab == 7:
    st.header("📄 개별 급여명세서 인쇄")
    
    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()
    df_emp = scope_employee_master(df_emp)

    df_emp = scope_employee_master(df_emp)
    ot_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()
    df_ot = scope_employee_master(df_ot)

    if df_emp.empty:
        st.warning("등록된 직원이 없다.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            pay_month_slip = st.date_input("명세서 지급 월 선택", datetime.now(), key="slip_month").strftime("%Y-%m")
            emp_slip_list = df_emp['emp_name'] + " (" + df_emp['position'] + " / " + df_emp['emp_id'] + ")"
            selected_slip_str = st.selectbox("직원 선택", emp_slip_list, key="slip_emp")
            selected_slip_id = selected_slip_str.split("/")[-1].replace(")", "").strip()
            emp = df_emp[df_emp['emp_id'] == selected_slip_id].iloc[0]
        with col2:
            slip_run_no = st.selectbox("출력할 급여대장", [1, 2], format_func=lambda x: f"{x}차 대장", key="slip_run_no")

        current_hourly_wage = int(emp['hourly_wage'])

        adj_single_res = supabase.table("monthly_payroll_adjust").select("*").eq("pay_month", pay_month_slip).eq("pay_run_no", slip_run_no).eq("emp_id", emp['emp_id']).execute()
        df_adj_single = pd.DataFrame(adj_single_res.data) if adj_single_res.data else pd.DataFrame()

        slip_pay_date = datetime.strptime(f"{pay_month_slip}-25", "%Y-%m-%d").date()
        ot_start_date_slip = (slip_pay_date.replace(day=1) - timedelta(days=1)).replace(day=25)
        ot_end_date_slip = slip_pay_date - timedelta(days=1)

        emp_ot = df_ot[
            (df_ot['emp_id'] == emp['emp_id']) & 
            (df_ot['work_date'] >= str(ot_start_date_slip)) & 
            (df_ot['work_date'] <= str(ot_end_date_slip))
        ] if not df_ot.empty else pd.DataFrame()

        weekday_ot_hours = emp_ot[emp_ot['work_type'].str.contains("평일", na=False)]['actual_duration_hours'].sum() if not emp_ot.empty else 0.0
        holiday_ot_hours = emp_ot[emp_ot['work_type'].str.contains("휴일", na=False)]['actual_duration_hours'].sum() if not emp_ot.empty else 0.0
        total_ot_hours = weekday_ot_hours + holiday_ot_hours
        calculated_ot_pay = int(emp_ot['actual_pay'].sum()) if not emp_ot.empty else 0

        if not df_adj_single.empty:
            adj = df_adj_single.iloc[0]
            base = adj['base_salary']
            ot_pay = adj['ot_pay'] if adj.get('ot_pay_overridden', False) else calculated_ot_pay
            family = adj['family_allowance']
            holiday_bonus = safe_int(adj.get('holiday_bonus'))
            non_tax = adj['non_taxable']
            other_allow = adj['other_allowance']
            emp_national = adj['national_pension']
            emp_health = adj['health_insurance']
            emp_longterm = adj['longterm_care']
            emp_employment = adj['employment_insurance']
            emp_income_tax = safe_int(adj.get('income_tax'))
            emp_local_tax = safe_int(adj.get('local_tax'))
            other_deduct = adj['other_deduction']
        else:
            ot_pay = calculated_ot_pay
            base = emp['base_salary']
            family = emp['family_allowance']
            holiday_bonus = 0
            non_tax = emp['non_taxable']
            other_allow = emp['other_allowance']
            other_deduct = emp['other_deduction']

            emp_national = safe_int(emp.get('national_pension'))
            emp_health = safe_int(emp.get('health_insurance'))
            emp_longterm = safe_int(emp.get('longterm_care'))
            emp_employment = safe_int(emp.get('employment_insurance'))
            emp_income_tax = safe_int(emp.get('income_tax'))
            emp_local_tax = safe_int(emp.get('local_tax'))
            if slip_run_no == 2:
                base = ot_pay = family = non_tax = other_allow = other_deduct = 0
                holiday_bonus = safe_int(emp.get('holiday_bonus'))
                emp_national = emp_health = emp_longterm = emp_employment = 0
                emp_income_tax = emp_local_tax = 0

        total_gross = base + ot_pay + family + holiday_bonus + non_tax + other_allow
        emp_deduction_total = emp_national + emp_health + emp_longterm + emp_employment + emp_income_tax + emp_local_tax + other_deduct
        net_pay = total_gross - emp_deduction_total

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 32px; float: left;">' if st.session_state.logo_b64 else ''

        payslip_template = f"""
        <div style="text-align: right; margin-bottom: 10px;">
            <button onclick="window.print()" style="padding: 8px 16px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">🖨️ 급여명세서 인쇄하기</button>
        </div>
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 680px; margin: auto; background: #fff;">
            {logo_html}
            <h2 style="text-align: center; margin-top: 0; margin-bottom: 25px; font-size: 24px; text-decoration: underline; clear: both;">
                {pay_month_slip}월 {slip_run_no}차 급 여 명 세 서
            </h2>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 13px;" border="1">
                <tr>
                    <th style="padding: 6px; background: #f2f2f2; width: 15%;">성 명</th>
                    <td style="padding: 6px; width: 35%;">{emp['emp_name']} ({emp['position']})</td>
                    <th style="padding: 6px; background: #f2f2f2; width: 15%;">소 속</th>
                    <td style="padding: 6px; width: 35%;">{emp['dept']}</td>
                </tr>
                <tr>
                    <th style="padding: 6px; background: #f2f2f2;">사 번</th>
                    <td style="padding: 6px;">{emp['emp_id']}</td>
                    <th style="padding: 6px; background: #f2f2f2;">생년월일</th>
                    <td style="padding: 6px;">{emp['birth_date']}</td>
                </tr>
            </table>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 13px;" border="1">
                <tr style="background: #e6f2ff;">
                    <th style="padding: 8px; width: 50%;" colspan="2">지급 항목</th>
                    <th style="padding: 8px; width: 50%;" colspan="2">공제 항목</th>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">기본급</td>
                    <td style="padding: 6px; text-align: right;">{base:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">국민연금</td>
                    <td style="padding: 6px; text-align: right;">{emp_national:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">시간외/휴일수당 ({total_ot_hours:.1f}h)</td>
                    <td style="padding: 6px; text-align: right;">{ot_pay:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">건강보험</td>
                    <td style="padding: 6px; text-align: right;">{emp_health:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">가족수당</td>
                    <td style="padding: 6px; text-align: right;">{family:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">장기요양보험</td>
                    <td style="padding: 6px; text-align: right;">{emp_longterm:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">비과세 식대/보육</td>
                    <td style="padding: 6px; text-align: right;">{non_tax:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">고용보험</td>
                    <td style="padding: 6px; text-align: right;">{emp_employment:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">명절상여</td>
                    <td style="padding: 6px; text-align: right;">{holiday_bonus:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">소득세 / 지방소득세</td>
                    <td style="padding: 6px; text-align: right;">{(emp_income_tax + emp_local_tax):,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">기타수당</td>
                    <td style="padding: 6px; text-align: right;">{other_allow:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">기타공제</td>
                    <td style="padding: 6px; text-align: right;">{other_deduct:,} 원</td>
                </tr>
                <tr style="font-weight: bold; background: #f2f2f2;">
                    <td style="padding: 8px;">지급액 계</td>
                    <td style="padding: 8px; text-align: right;">{total_gross:,} 원</td>
                    <td style="padding: 8px;">공제액 계</td>
                    <td style="padding: 8px; text-align: right;">{emp_deduction_total:,} 원</td>
                </tr>
            </table>

            <div style="border: 2px solid #333; padding: 10px; text-align: center; background: #fffde7; margin-bottom: 20px;">
                <span style="font-size: 16px; font-weight: bold;">실지급액 : {net_pay:,} 원</span>
            </div>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; text-align: center; font-size: 12px;" border="1">
                <tr style="background-color: #f2f2f2; height: 28px;">
                    <th style="width: 25%;">통상시급</th>
                    <th style="width: 25%;">시간외·연장근로시간</th>
                    <th style="width: 25%;">휴일근로시간</th>
                    <th style="width: 25%;">야간근로시간</th>
                </tr>
                <tr style="height: 32px;">
                    <td><b>{current_hourly_wage:,} 원</b></td>
                    <td>{f"{weekday_ot_hours:.1f}h" if weekday_ot_hours > 0 else '-'}</td>
                    <td>{f"{holiday_ot_hours:.1f}h" if holiday_ot_hours > 0 else '-'}</td>
                    <td>-</td>
                </tr>
            </table>

            <div style="font-size: 13px; font-weight: bold; text-align: center; margin-bottom: 8px; background-color: #f2f2f2; padding: 5px; border: 1px solid #000;">
                계 산 방 법
            </div>
            <table style="width: 100%; border-collapse: collapse; font-size: 11px;" border="1">
                <tr style="background-color: #f9f9f9; text-align: center; height: 25px;">
                    <th style="width: 35%;">구분</th>
                    <th style="width: 65%;">산출식 또는 산출방법</th>
                </tr>
                <tr style="height: 24px;">
                    <td style="text-align: center;">일할급여 계산기준<br>(중간 입사/퇴사 시)</td>
                    <td style="padding-left: 10px;">기본급 ÷ 해당 월일수 × 근무일수</td>
                </tr>
                <tr style="height: 24px;">
                    <td style="text-align: center;">연장근로수당</td>
                    <td style="padding-left: 10px;">연장근로시간수 × 통상시급({current_hourly_wage:,}원) × 1.5</td>
                </tr>
                <tr style="height: 24px;">
                    <td style="text-align: center;">휴일근로수당</td>
                    <td style="padding-left: 10px;">휴일근로시간수 × 통상시급({current_hourly_wage:,}원) × 1.5</td>
                </tr>
                <tr style="height: 24px;">
                    <td style="text-align: center;">야간근로수당</td>
                    <td style="padding-left: 10px;">야간근로시간수 × 통상시급({current_hourly_wage:,}원) × 1.5</td>
                </tr>
            </table>
            <p style="font-size: 10px; color: #333; margin-top: 5px; margin-bottom: 15px;">
                * 통상시급 : 정기적이고 일률적으로 지급하는 급여 ÷ 소정근로시간(209시간)
            </p>

            <p style="text-align: center; margin-top: 20px; font-size: 12px; color: #444;">귀하의 노고에 진심으로 감사드립니다.</p>
        </div>
        """
        st.components.v1.html(payslip_template, height=890, scrolling=True)

# -------------------------------------------------------------------
# TAB 8: 통합 급여대장 인쇄
# -------------------------------------------------------------------
if active_tab in (7, 8):
    if active_tab == 7:
        st.divider()
        st.header("🖨️ 급여대장 인쇄")
        st.caption("통합 급여대장과 같은 화면에서 지급일·대장 차수를 선택한 뒤 바로 인쇄할 수 있습니다.")
    else:
        st.header("🖨️ 통합 급여대장 인쇄")

    print_col1, print_col2 = st.columns(2)
    with print_col1:
        pay_date_print = st.date_input("출력할 급여 지급일 선택", datetime.now(), key="payroll_print_date")
    with print_col2:
        print_run_no = st.selectbox("출력할 급여대장", [1, 2], format_func=lambda x: f"{x}차 대장", key="print_run_no")
    pay_month_print = pay_date_print.strftime("%Y-%m")

    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()
    df_emp = scope_employee_master(df_emp)

    df_emp = scope_employee_master(df_emp)
    ot_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()
    df_ot = scope_employee_master(df_ot)

    adj_res = supabase.table("monthly_payroll_adjust").select("*").eq("pay_month", pay_month_print).eq("pay_run_no", print_run_no).execute()
    df_adjust = pd.DataFrame(adj_res.data) if adj_res.data else pd.DataFrame()
    df_adjust = scope_employee_master(df_adjust)

    if df_emp.empty:
        st.warning("등록된 직원이 없다.")
    else:
        payroll_print_rows = ""
        no = 1

        sum_base = sum_ot = sum_family = sum_holiday = sum_nontax = sum_gross = 0
        sum_nat = sum_hea = sum_long = sum_emp = sum_inc = sum_loc = sum_other_d = sum_deduct_tot = sum_net = 0
        sum_b_nat = sum_b_hea = sum_b_long = sum_b_emp = sum_b_ind = sum_b_tot = sum_retire = 0

        ot_start_date_print = (pay_date_print.replace(day=1) - timedelta(days=1)).replace(day=25)
        ot_end_date_print = pay_date_print - timedelta(days=1)

        for idx, emp in df_emp.iterrows():
            adj_match = df_adjust[df_adjust['emp_id'] == emp['emp_id']] if not df_adjust.empty else pd.DataFrame()
            emp_ot = df_ot[
                (df_ot['emp_id'] == emp['emp_id']) & 
                (df_ot['work_date'] >= str(ot_start_date_print)) & 
                (df_ot['work_date'] <= str(ot_end_date_print))
            ] if not df_ot.empty else pd.DataFrame()

            calculated_ot_pay = int(emp_ot['actual_pay'].sum()) if not emp_ot.empty else 0

            if not adj_match.empty:
                adj = adj_match.iloc[0]
                base = adj['base_salary']
                ot_pay = adj['ot_pay'] if adj.get('ot_pay_overridden', False) else calculated_ot_pay
                family = adj['family_allowance']
                holiday_bonus = safe_int(adj.get('holiday_bonus'))
                non_tax = adj['non_taxable']
                other_allow = adj['other_allowance']
                emp_national = adj['national_pension']
                emp_health = adj['health_insurance']
                emp_longterm = adj['longterm_care']
                emp_employment = adj['employment_insurance']
                emp_income_tax = safe_int(adj.get('income_tax'))
                emp_local_tax = safe_int(adj.get('local_tax'))
                other_deduct = adj['other_deduction']
            else:
                ot_pay = calculated_ot_pay
                base = emp['base_salary']
                family = emp['family_allowance']
                holiday_bonus = 0
                non_tax = emp['non_taxable']
                other_allow = emp['other_allowance']
                other_deduct = emp['other_deduction']

                emp_national = safe_int(emp.get('national_pension'))
                emp_health = safe_int(emp.get('health_insurance'))
                emp_longterm = safe_int(emp.get('longterm_care'))
                emp_employment = safe_int(emp.get('employment_insurance'))
                emp_income_tax = safe_int(emp.get('income_tax'))
                emp_local_tax = safe_int(emp.get('local_tax'))
                if print_run_no == 2:
                    base = ot_pay = family = non_tax = other_allow = other_deduct = 0
                    holiday_bonus = safe_int(emp.get('holiday_bonus'))
                    emp_national = emp_health = emp_longterm = emp_employment = 0
                    emp_income_tax = emp_local_tax = 0

            tot_g = base + ot_pay + family + holiday_bonus + non_tax + other_allow
            emp_deduction_total = emp_national + emp_health + emp_longterm + emp_employment + emp_income_tax + emp_local_tax + other_deduct
            net_pay = tot_g - emp_deduction_total

            biz_national = safe_int(adj.get('employer_national_pension')) if not adj_match.empty else safe_int(emp.get('employer_national_pension'))
            biz_health = safe_int(adj.get('employer_health_insurance')) if not adj_match.empty else safe_int(emp.get('employer_health_insurance'))
            biz_longterm = safe_int(adj.get('employer_longterm_care')) if not adj_match.empty else safe_int(emp.get('employer_longterm_care'))
            biz_employment = safe_int(adj.get('employer_employment_insurance')) if not adj_match.empty else safe_int(emp.get('employer_employment_insurance'))
            biz_industrial = safe_int(adj.get('employer_industrial_insurance')) if not adj_match.empty else safe_int(emp.get('employer_industrial_insurance'))
            biz_deduction_total = biz_national + biz_health + biz_longterm + biz_employment + biz_industrial
            retirement_accrual = safe_int(adj.get('retirement_accrual')) if not adj_match.empty else safe_int(emp.get('retirement_accrual'))

            sum_base += base; sum_ot += ot_pay; sum_family += family; sum_holiday += holiday_bonus; sum_nontax += non_tax; sum_gross += tot_g
            sum_nat += emp_national; sum_hea += emp_health; sum_long += emp_longterm; sum_emp += emp_employment
            sum_inc += emp_income_tax; sum_loc += emp_local_tax; sum_other_d += other_deduct; sum_deduct_tot += emp_deduction_total; sum_net += net_pay
            sum_b_nat += biz_national; sum_b_hea += biz_health; sum_b_long += biz_longterm; sum_b_emp += biz_employment; sum_b_ind += biz_industrial; sum_b_tot += biz_deduction_total; sum_retire += retirement_accrual

            payroll_print_rows += f"""
            <tr style="height: 26px;">
                <td>{no}</td><td>{emp['emp_name']}</td><td>{emp['birth_date']}</td><td>{emp['hobong']}</td>
                <td style="text-align:right;">{base:,}</td>
                <td style="text-align:right;">{ot_pay:,}</td>
                <td style="text-align:right;">{family:,}</td>
                <td style="text-align:right;">{holiday_bonus:,}</td>
                <td style="text-align:right;">{non_tax:,}</td>
                <td style="text-align:right; font-weight:bold;">{tot_g:,}</td>
                <td style="text-align:right;">{emp_national:,}</td>
                <td style="text-align:right;">{emp_health:,}</td>
                <td style="text-align:right;">{emp_longterm:,}</td>
                <td style="text-align:right;">{emp_employment:,}</td>
                <td style="text-align:right;">{emp_income_tax:,}</td>
                <td style="text-align:right;">{emp_local_tax:,}</td>
                <td style="text-align:right; font-weight:bold;">{emp_deduction_total:,}</td>
                <td style="text-align:right; font-weight:bold; background-color:#fffae6;">{net_pay:,}</td>
                <td style="text-align:right;">{biz_national:,}</td>
                <td style="text-align:right;">{biz_health:,}</td>
                <td style="text-align:right;">{biz_longterm:,}</td>
                <td style="text-align:right;">{biz_employment:,}</td>
                <td style="text-align:right;">{biz_industrial:,}</td>
                <td style="text-align:right; font-weight:bold;">{biz_deduction_total:,}</td>
                <td style="text-align:right;">{retirement_accrual:,}</td>
            </tr>
            """
            no += 1

        summary_print_row = f"""
        <tr style="background-color: #e6f2ff; font-weight: bold; height: 28px;">
            <td colspan="4">합 계</td>
            <td style="text-align:right;">{sum_base:,}</td>
            <td style="text-align:right;">{sum_ot:,}</td>
            <td style="text-align:right;">{sum_family:,}</td>
            <td style="text-align:right;">{sum_holiday:,}</td>
            <td style="text-align:right;">{sum_nontax:,}</td>
            <td style="text-align:right;">{sum_gross:,}</td>
            <td style="text-align:right;">{sum_nat:,}</td>
            <td style="text-align:right;">{sum_hea:,}</td>
            <td style="text-align:right;">{sum_long:,}</td>
            <td style="text-align:right;">{sum_emp:,}</td>
            <td style="text-align:right;">{sum_inc:,}</td>
            <td style="text-align:right;">{sum_loc:,}</td>
            <td style="text-align:right;">{sum_deduct_tot:,}</td>
            <td style="text-align:right; background-color:#ffe680;">{sum_net:,}</td>
            <td style="text-align:right;">{sum_b_nat:,}</td>
            <td style="text-align:right;">{sum_b_hea:,}</td>
            <td style="text-align:right;">{sum_b_long:,}</td>
            <td style="text-align:right;">{sum_b_emp:,}</td>
            <td style="text-align:right;">{sum_b_ind:,}</td>
            <td style="text-align:right;">{sum_b_tot:,}</td>
            <td style="text-align:right;">{sum_retire:,}</td>
        </tr>
        """

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 28px; float: left;">' if st.session_state.logo_b64 else ''

        payroll_print_template = f"""
        <div style="text-align: right; margin-bottom: 10px;">
            <button onclick="window.print()" style="padding: 8px 16px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">🖨️ 급여대장 인쇄하기</button>
        </div>
        <div style="border: 2px solid #000; padding: 20px; font-family: 'Malgun Gothic', sans-serif; background: #fff; width: 100%; box-sizing: border-box;">
            {logo_html}
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 15px; clear: both;">
                <h2 style="margin: 0; padding-top: 5px; font-size: 22px; text-decoration: underline;">{pay_month_print}월 {print_run_no}차 통합 급여대장</h2>
                <table style="border-collapse: collapse; text-align: center; font-size: 11px; width: 210px;" border="1">
                    <tr style="height: 18px; background-color: #f2f2f2;">
                        <th rowspan="2" style="width: 25px; background-color: #e6e6e6;">결<br>재</th>
                        <th style="width: 60px;">담 당</th>
                        <th style="width: 60px;">대 리</th>
                        <th style="width: 65px;">센터장</th>
                    </tr>
                    <tr style="height: 40px;">
                        <td></td><td></td><td></td>
                    </tr>
                </table>
            </div>

            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; font-size: 11px;">
                <div style="font-size: 12px; font-weight: bold;">지급일자: {pay_date_print}</div>
                <div>(단위: 원 / 원단위 절사)</div>
            </div>

            <table border="1" style="width: 100%; border-collapse: collapse; text-align: center; font-size: 10px;" cellpadding="2">
                <thead>
                    <tr style="background-color: #ffffcc;">
                        <th rowspan="3" style="width: 25px;">No</th>
                        <th rowspan="3" style="width: 50px;">이름</th>
                        <th rowspan="3" style="width: 65px;">생년월일</th>
                        <th rowspan="3" style="width: 40px;">호봉</th>
                        <th colspan="5">지급 내역</th>
                        <th rowspan="3">급여총액</th>
                        <th colspan="7">근로자 본인 부담금</th>
                        <th rowspan="3" style="background-color: #fff2cc;">실지급액</th>
                        <th colspan="6">사업자 부담 사회보험금</th>
                        <th rowspan="3">사업주부담<br>퇴직적립금</th>
                    </tr>
                    <tr style="background-color: #ffffcc;">
                        <th rowspan="2">기본급</th>
                        <th rowspan="2">초과수당</th>
                        <th rowspan="2">가족수당</th>
                        <th rowspan="2">명절상여</th>
                        <th rowspan="2">비과세</th>
                        <th>국민</th><th>건강</th><th>장기요양</th><th>고용</th><th>소득세</th><th>지방세</th>
                        <th rowspan="2">공제합계</th>
                        <th>국민</th><th>건강</th><th>장기요양</th><th>고용</th><th>산재</th>
                        <th rowspan="2">사업자합계</th>
                    </tr>
                    <tr style="background-color: #ffffcc;">
                        <th>4.75%</th><th>3.595%</th><th>12.95%</th><th>0.90%</th><th>간이세액</th><th>10%</th>
                        <th>4.75%</th><th>3.595%</th><th>12.95%</th><th>1.15%</th><th>7.26%</th>
                    </tr>
                </thead>
                <tbody>
                    {summary_print_row}
                    {payroll_print_rows}
                </tbody>
            </table>
        </div>
        """
        st.components.v1.html(payroll_print_template, height=600, scrolling=True)

# -------------------------------------------------------------------
# TAB 9: 월별 급여대장 총괄표
# -------------------------------------------------------------------
if active_tab == 9:
    st.header("📑 월별 급여대장 총괄표 (12개월 누적 요약)")
    
    c_y9 = st.selectbox("조회 연도 선택", range(datetime.now().year - 2, datetime.now().year + 3), index=2, key="annual_summary_year")
    
    emp_all_res = supabase.table("employees").select("*").execute()
    df_emp_all = pd.DataFrame(emp_all_res.data) if emp_all_res.data else pd.DataFrame()
    df_emp_all = scope_employee_master(df_emp_all)

    ot_all_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot_all = pd.DataFrame(ot_all_res.data) if ot_all_res.data else pd.DataFrame()
    df_ot_all = scope_employee_master(df_ot_all)

    adj_all_res = supabase.table("monthly_payroll_adjust").select("*").execute()
    df_adj_all = pd.DataFrame(adj_all_res.data) if adj_all_res.data else pd.DataFrame()
    df_adj_all = scope_employee_master(df_adj_all)

    if df_emp_all.empty:
        st.warning("등록된 직원 정보가 없다.")
    else:
        monthly_summary_rows = []
        tot_ann_count = tot_ann_base = tot_ann_ot_hours = tot_ann_ot = tot_ann_fam = tot_ann_holiday = tot_ann_nontax = tot_ann_other_a = 0
        tot_ann_gross = tot_ann_nat = tot_ann_hea = tot_ann_long = tot_ann_emp = tot_ann_inc = tot_ann_loc = 0
        tot_ann_other_d = tot_ann_deduct = tot_ann_net = tot_ann_retire = 0

        for m in range(1, 13):
            m_str = f"{c_y9}-{m:02d}"
            m_emp_count = len(df_emp_all)
            m_base = m_ot = m_fam = m_holiday = m_nontax = m_other_a = 0
            m_nat = m_hea = m_long = m_emp = m_inc = m_loc = m_other_d = 0
            m_retire = 0
            m_ot_hours = 0.0

            # 지급월 산정기간(전월 25일 ~ 당월 24일) 계산
            pay_ref = datetime(c_y9, m, 25).date()
            m_p_start = (pay_ref.replace(day=1) - timedelta(days=1)).replace(day=25)
            m_p_end = pay_ref - timedelta(days=1)

            m_ot_records = df_ot_all[
                (df_ot_all['work_date'] >= str(m_p_start)) & 
                (df_ot_all['work_date'] <= str(m_p_end))
            ] if not df_ot_all.empty else pd.DataFrame()

            if not m_ot_records.empty:
                m_ot_hours = m_ot_records['actual_duration_hours'].sum()

            for _, emp in df_emp_all.iterrows():
                adj_m = df_adj_all[(df_adj_all['pay_month'] == m_str) & (df_adj_all['emp_id'] == emp['emp_id'])] if not df_adj_all.empty else pd.DataFrame()
                
                emp_ot = df_ot_all[
                    (df_ot_all['emp_id'] == emp['emp_id']) & 
                    (df_ot_all['work_date'] >= str(m_p_start)) & 
                    (df_ot_all['work_date'] <= str(m_p_end))
                ] if not df_ot_all.empty else pd.DataFrame()
                
                calc_ot = int(emp_ot['actual_pay'].sum()) if not emp_ot.empty else 0

                if not adj_m.empty:
                    base = safe_int(adj_m['base_salary'].sum())
                    fam = safe_int(adj_m['family_allowance'].sum())
                    m_holiday += safe_int(adj_m['holiday_bonus'].sum()) if 'holiday_bonus' in adj_m.columns else 0
                    nontax = safe_int(adj_m['non_taxable'].sum())
                    other_a = safe_int(adj_m['other_allowance'].sum())
                    nat = safe_int(adj_m['national_pension'].sum())
                    hea = safe_int(adj_m['health_insurance'].sum())
                    lng = safe_int(adj_m['longterm_care'].sum())
                    e_emp = safe_int(adj_m['employment_insurance'].sum())
                    inc = safe_int(adj_m['income_tax'].sum()) if 'income_tax' in adj_m.columns else 0
                    loc = safe_int(adj_m['local_tax'].sum()) if 'local_tax' in adj_m.columns else 0
                    other_d = safe_int(adj_m['other_deduction'].sum())
                    retire = safe_int(adj_m['retirement_accrual'].fillna(0).sum()) if 'retirement_accrual' in adj_m.columns else 0
                    ot = 0
                    for _, adj in adj_m.iterrows():
                        if bool(adj.get('ot_pay_overridden', False)):
                            run_ot = safe_int(adj.get('ot_pay'))
                        elif safe_int(adj.get('pay_run_no', 1)) == 1:
                            run_ot = calc_ot
                        else:
                            run_ot = 0
                        ot += run_ot
                else:
                    ot = calc_ot
                    base = emp['base_salary']
                    fam = emp['family_allowance']
                    nontax = emp['non_taxable']
                    other_a = emp['other_allowance']
                    other_d = emp['other_deduction']

                    nat = safe_int(emp.get('national_pension'))
                    hea = safe_int(emp.get('health_insurance'))
                    lng = safe_int(emp.get('longterm_care'))
                    e_emp = safe_int(emp.get('employment_insurance'))
                    inc = safe_int(emp.get('income_tax'))
                    loc = safe_int(emp.get('local_tax'))
                    retire = safe_int(emp.get('retirement_accrual'))

                m_base += base; m_ot += ot; m_fam += fam; m_nontax += nontax; m_other_a += other_a
                m_nat += nat; m_hea += hea; m_long += lng; m_emp += e_emp
                m_inc += inc; m_loc += loc; m_other_d += other_d
                m_retire += retire

            m_gross = m_base + m_ot + m_fam + m_holiday + m_nontax + m_other_a
            m_deduct = m_nat + m_hea + m_long + m_emp + m_inc + m_loc + m_other_d
            m_net = m_gross - m_deduct
            tot_ann_base += m_base; tot_ann_ot_hours += m_ot_hours; tot_ann_ot += m_ot; tot_ann_fam += m_fam; tot_ann_holiday += m_holiday; tot_ann_nontax += m_nontax; tot_ann_other_a += m_other_a
            tot_ann_gross += m_gross; tot_ann_nat += m_nat; tot_ann_hea += m_hea; tot_ann_long += m_long; tot_ann_emp += m_emp
            tot_ann_inc += m_inc; tot_ann_loc += m_loc; tot_ann_other_d += m_other_d; tot_ann_deduct += m_deduct; tot_ann_net += m_net
            tot_ann_retire += m_retire

            monthly_summary_rows.append(f"""
            <tr style="height: 26px;">
                <td><b>{m}월</b> ({m_str})</td>
                <td>{m_emp_count} 명</td>
                <td style="text-align:right;">{m_base:,}</td>
                <td style="text-align:right; background-color:#f0f8ff;"><b>{m_ot_hours:.1f} 시간</b></td>
                <td style="text-align:right;">{m_ot:,}</td>
                <td style="text-align:right;">{m_fam:,}</td>
                <td style="text-align:right;">{m_holiday:,}</td>
                <td style="text-align:right;">{m_nontax:,}</td>
                <td style="text-align:right; font-weight:bold; background-color:#f9f9f9;">{m_gross:,}</td>
                <td style="text-align:right;">{m_nat:,}</td>
                <td style="text-align:right;">{m_hea:,}</td>
                <td style="text-align:right;">{m_long:,}</td>
                <td style="text-align:right;">{m_emp:,}</td>
                <td style="text-align:right;">{m_inc:,}</td>
                <td style="text-align:right;">{m_loc:,}</td>
                <td style="text-align:right; font-weight:bold; background-color:#f9f9f9;">{m_deduct:,}</td>
                <td style="text-align:right; font-weight:bold; background-color:#fffae6;">{m_net:,}</td>
                <td style="text-align:right;">{m_retire:,}</td>
            </tr>
            """)

        ann_sum_html_row = f"""
        <tr style="background-color: #e6f2ff; font-weight: bold; height: 30px;">
            <td colspan="2">연간 누적 합계</td>
            <td style="text-align:right;">{tot_ann_base:,}</td>
            <td style="text-align:right; background-color:#d0e8ff;">{tot_ann_ot_hours:.1f} 시간</td>
            <td style="text-align:right;">{tot_ann_ot:,}</td>
            <td style="text-align:right;">{tot_ann_fam:,}</td>
            <td style="text-align:right;">{tot_ann_holiday:,}</td>
            <td style="text-align:right;">{tot_ann_nontax:,}</td>
            <td style="text-align:right;">{tot_ann_gross:,}</td>
            <td style="text-align:right;">{tot_ann_nat:,}</td>
            <td style="text-align:right;">{tot_ann_hea:,}</td>
            <td style="text-align:right;">{tot_ann_long:,}</td>
            <td style="text-align:right;">{tot_ann_emp:,}</td>
            <td style="text-align:right;">{tot_ann_inc:,}</td>
            <td style="text-align:right;">{tot_ann_loc:,}</td>
            <td style="text-align:right;">{tot_ann_deduct:,}</td>
            <td style="text-align:right; background-color:#ffe680;">{tot_ann_net:,}</td>
            <td style="text-align:right;">{tot_ann_retire:,}</td>
        </tr>
        """

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 28px; float: left;">' if st.session_state.logo_b64 else ''

        annual_summary_template = f"""
        <div style="text-align: right; margin-bottom: 10px;">
            <button onclick="window.print()" style="padding: 8px 16px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">🖨️ 총괄표 인쇄하기</button>
        </div>
        <div style="border: 2px solid #000; padding: 25px; font-family: 'Malgun Gothic', sans-serif; background: #fff; width: 100%; box-sizing: border-box;">
            {logo_html}
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; clear: both;">
                <h2 style="margin: 0; padding-top: 5px; font-size: 22px; text-decoration: underline;">{c_y9}년도 월별 급여대장 총괄표</h2>
                <table style="border-collapse: collapse; text-align: center; font-size: 11px; width: 210px;" border="1">
                    <tr style="height: 18px; background-color: #f2f2f2;">
                        <th rowspan="2" style="width: 25px; background-color: #e6e6e6;">결<br>재</th>
                        <th style="width: 60px;">담 당</th>
                        <th style="width: 60px;">대 리</th>
                        <th style="width: 65px;">센터장</th>
                    </tr>
                    <tr style="height: 40px;">
                        <td></td><td></td><td></td>
                    </tr>
                </table>
            </div>

            <div style="text-align: right; margin-bottom: 8px; font-size: 11px;">(단위: 원 / 원단위 절사)</div>

            <table border="1" style="width: 100%; border-collapse: collapse; text-align: center; font-size: 11px;" cellpadding="3">
                <thead>
                    <tr style="background-color: #ffffcc; height: 32px;">
                        <th style="width: 80px;">지급월</th>
                        <th style="width: 45px;">인원</th>
                        <th>기본급</th>
                        <th style="background-color: #e6f2ff;">승인 초과시간</th>
                        <th>초과수당</th>
                        <th>가족수당</th>
                        <th>명절상여</th>
                        <th>비과세</th>
                        <th style="background-color: #fff2cc;">급여총액</th>
                        <th>국민연금</th>
                        <th>건강보험</th>
                        <th>장기요양</th>
                        <th>고용보험</th>
                        <th>소득세</th>
                        <th>지방세</th>
                        <th style="background-color: #fff2cc;">공제합계</th>
                        <th style="background-color: #ffe680;">실지급액</th>
                        <th>퇴직적립금</th>
                    </tr>
                </thead>
                <tbody>
                    {ann_sum_html_row}
                    {''.join(monthly_summary_rows)}
                    {ann_sum_html_row}
                </tbody>
            </table>
        </div>
        """
        st.components.v1.html(annual_summary_template, height=650, scrolling=True)


# -------------------------------------------------------------------
# TAB 10: 출장 신청·관리
# -------------------------------------------------------------------
if active_tab == 10:
    st.header("🚗 출장 신청·관리")
    st.caption("복무규정 제25조~제28조 및 여비관리 세칙 기준")

    with st.expander("📖 출장 관련 복무규정 바로보기", expanded=False):
        st.markdown("""
**제25조(출장명령)**  
직무수행을 위해 출장하는 직원은 출장신청서를 제출하고 출장명령을 받아야 합니다. 지정된 출장기일 안에 업무수행에 전력을 다하고, 업무를 완수하지 못할 사유가 생기면 즉시 상사에게 보고하여 지시를 받아야 합니다.

**제26조(출장 중의 사정변경)**  
출장 목적지 외 장소 방문 또는 출장기일 연장이 필요한 경우 사전에 상사에게 보고하고 승인을 받아야 합니다. 부득이하게 사전승인을 받지 못한 경우 귀원 즉시 사후승인을 받아야 합니다.

**제27조(출장보고)**  
출장용무를 마치고 귀원한 때에는 지체 없이 출장보고서를 제출해야 합니다. 다만 경미하거나 비밀에 속하는 사항은 구두보고가 가능합니다.

**제28조(출장여비)**  
출장여비는 「여비관리 세칙」에 따릅니다.
        """)

    # 테이블 존재 여부 안내
    try:
        trip_res = supabase.table("business_trips").select("*").order("id", desc=True).execute()
        df_trips = pd.DataFrame(trip_res.data) if trip_res.data else pd.DataFrame()
        df_trips = scope_dataframe_to_current_employee(df_trips)
        trip_table_ok = True
    except Exception:
        df_trips = pd.DataFrame()
        trip_table_ok = False

    if not trip_table_ok:
        st.warning("⚠️ 출장관리 DB 테이블이 아직 없습니다. 아래 'Supabase SQL'을 먼저 실행해 주세요.")

    emp_trip_res = supabase.table("employees").select("*").execute()
    df_trip_emp = pd.DataFrame(emp_trip_res.data) if emp_trip_res.data else pd.DataFrame()
    df_trip_emp = scope_employee_master(df_trip_emp)

    df_trip_emp = scope_employee_master(df_trip_emp)
    st.subheader("📝 출장 신청")
    if df_trip_emp.empty:
        st.info("직원 데이터가 없습니다.")
    else:
        # employees 실제 스키마(emp_name, dept)에 맞춰 출장자 연계
        emp_opts = {
            f"{str(r.get('emp_name','')).strip()} ({str(r.get('emp_id','')).strip()})": r
            for _, r in df_trip_emp.iterrows()
            if str(r.get("emp_name","")).strip()
        }
        if not emp_opts:
            st.warning("직원 DB의 emp_name 컬럼에서 출장자를 찾지 못했습니다.")
            st.stop()
        trip_emp_key = st.selectbox("출장자", list(emp_opts.keys()), key="trip_emp")
        trip_emp = emp_opts[trip_emp_key]

        tc1, tc2, tc3 = st.columns(3)
        with tc1:
            st.text_input("성명", value=str(trip_emp.get("emp_name","")), disabled=True)
        with tc2:
            st.text_input("직위", value=str(trip_emp.get("position","")), disabled=True)
        with tc3:
            st.text_input("부서", value=str(trip_emp.get("dept","")), disabled=True)

        trip_type = st.radio("출장 구분", ["일반출장", "교육·연수출장"], horizontal=True)
        purpose = st.text_input("출장목적", placeholder="예: 장기요양기관 회계 교육 참석")
        dc1, dc2 = st.columns(2)
        with dc1:
            start_date = st.date_input("출장 시작일", key="trip_start_date")
            start_time = st.time_input("시작시간", value=time(9,0), key="trip_start_time")
        with dc2:
            end_date = st.date_input("출장 종료일", key="trip_end_date")
            end_time = st.time_input("종료시간", value=time(18,0), key="trip_end_time")

        destination = st.text_input("출장지")
        mv1, mv2, mv3 = st.columns(3)
        with mv1:
            transport = st.selectbox("이동수단", ["대중교통", "자차", "센터차량", "동승", "도보", "기타"])
        with mv2:
            distance_km = st.number_input("총 거리(km)", min_value=0.0, step=1.0)
        with mv3:
            fuel_type = st.selectbox("자차 유종", ["해당없음", "가솔린", "디젤", "LPG"], disabled=(transport != "자차"))

        st.markdown("##### 💰 여비 기준 확인")
        start_dt = datetime.combine(start_date, start_time)
        end_dt = datetime.combine(end_date, end_time)
        trip_hours = max(0, (end_dt - start_dt).total_seconds() / 3600)
        if trip_hours >= 4:
            base_trip_pay = 20000
            rule_text = "출장여행시간 4시간 이상 → 20,000원"
        elif trip_hours > 2:
            base_trip_pay = 10000
            rule_text = "출장여행시간 4시간 미만 → 10,000원"
        else:
            base_trip_pay = 0
            rule_text = "2시간 이내/보행 가능 거리 → 실교통비 지급 가능"

        if transport == "센터차량":
            base_trip_pay = max(0, base_trip_pay - 10000)
            rule_text += " / 센터차량 사용 → 기준금액에서 10,000원 감액"

        rc1, rc2 = st.columns(2)
        with rc1:
            st.info(rule_text)
        with rc2:
            st.metric("규정 기준 여비(참고)", f"{base_trip_pay:,}원")

        st.caption("※ 자가차량 운임은 여비관리 세칙상 총거리×기준단가 방식입니다. 유류 기준가격 입력 없이 임의 계산하지 않고 정산 단계에서 증빙·기준단가를 확인하도록 설계했습니다.")
        note = st.text_area("비고 / 출장 사전 특이사항")

        if st.button("🚗 출장 신청 등록", type="primary", use_container_width=True, disabled=not has_permission("trip_write")):
            if not trip_table_ok:
                st.error("먼저 Supabase에 business_trips 테이블을 생성해 주세요.")
            elif not purpose or not destination:
                st.warning("출장목적과 출장지를 입력해 주세요.")
            elif end_dt <= start_dt:
                st.warning("출장 종료일시는 시작일시보다 늦어야 합니다.")
            else:
                payload = {
                    "emp_id": str(trip_emp.get("emp_id","")),
                    "emp_name": str(trip_emp.get("emp_name","")),
                    "department": str(trip_emp.get("dept","")),
                    "position": str(trip_emp.get("position","")),
                    "trip_type": trip_type,
                    "purpose": purpose,
                    "start_at": start_dt.isoformat(),
                    "end_at": end_dt.isoformat(),
                    "destination": destination,
                    "transport_type": transport,
                    "distance_km": float(distance_km),
                    "fuel_type": fuel_type,
                    "rule_base_amount": int(base_trip_pay),
                    "note": note,
                    "apply_status": "신청",
                    "report_status": "미작성",
                    "settlement_status": "미정산"
                }
                supabase.table("business_trips").insert(payload).execute()
                st.success("출장 신청이 등록되었습니다.")
                st.rerun()

    st.divider()
    # 과거 출장자료의 빈 이름/부서는 emp_id 기준으로 화면에서 자동 보정
    if trip_table_ok and not df_trips.empty and not df_trip_emp.empty and "emp_id" in df_trips.columns:
        emp_lookup = df_trip_emp.set_index(df_trip_emp["emp_id"].astype(str))
        for idx, tr in df_trips.iterrows():
            eid = str(tr.get("emp_id",""))
            if eid in emp_lookup.index:
                er = emp_lookup.loc[eid]
                if isinstance(er, pd.DataFrame):
                    er = er.iloc[0]
                if not str(tr.get("emp_name","") or "").strip():
                    df_trips.at[idx, "emp_name"] = str(er.get("emp_name",""))
                if not str(tr.get("department","") or "").strip():
                    df_trips.at[idx, "department"] = str(er.get("dept",""))
                if not str(tr.get("position","") or "").strip():
                    df_trips.at[idx, "position"] = str(er.get("position",""))

    st.subheader("📊 출장 현황 / 승인 관리")
    if trip_table_ok and not df_trips.empty:
        show_cols = [c for c in ["id","start_at","emp_name","position","purpose","destination","transport_type",
                                 "distance_km","rule_base_amount","apply_status","report_status","settlement_status"]
                     if c in df_trips.columns]
        display_table_kr(df_trips[show_cols], use_container_width=True, hide_index=True)

        trip_ids = df_trips["id"].tolist()
        sel_trip_id = st.selectbox("처리할 출장 ID", trip_ids, key="trip_manage_id")
        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            if st.button("✅ 출장 승인(명령)", use_container_width=True, disabled=not has_permission("trip_approve")):
                supabase.table("business_trips").update({"apply_status":"승인"}).eq("id", sel_trip_id).execute()
                write_audit_log("출장 승인", "business_trips", sel_trip_id)
                st.success("출장명령/승인 처리했습니다.")
                st.rerun()
        with mc2:
            if st.button("↩️ 반려", use_container_width=True, disabled=not has_permission("trip_approve")):
                supabase.table("business_trips").update({"apply_status":"반려"}).eq("id", sel_trip_id).execute()
                write_audit_log("출장 반려", "business_trips", sel_trip_id)
                st.rerun()
        with mc3:
            if st.button("❌ 출장 취소", use_container_width=True, disabled=not has_permission("trip_approve")):
                supabase.table("business_trips").update({"apply_status":"취소"}).eq("id", sel_trip_id).execute()
                write_audit_log("출장 취소", "business_trips", sel_trip_id)
                st.rerun()
    elif trip_table_ok:
        st.info("등록된 출장 내역이 없습니다.")

# -------------------------------------------------------------------
# TAB 11: 출장 복명·복무규정
# -------------------------------------------------------------------
if active_tab == 11:
    st.header("📋 출장 복명·여비정산")

    try:
        trip_res2 = supabase.table("business_trips").select("*").order("id", desc=True).execute()
        df_trip2 = pd.DataFrame(trip_res2.data) if trip_res2.data else pd.DataFrame()
        df_trip2 = scope_dataframe_to_current_employee(df_trip2)
    except Exception:
        df_trip2 = pd.DataFrame()

    if df_trip2.empty:
        st.info("출장 신청 내역이 없습니다. 먼저 출장 신청을 등록해 주세요.")
    else:
        approved_trip = df_trip2[df_trip2["apply_status"] == "승인"] if "apply_status" in df_trip2.columns else df_trip2
        if approved_trip.empty:
            st.info("승인된 출장 건이 없습니다.")
        else:
            labels = {
                int(r["id"]): f"#{int(r['id'])} | {r.get('emp_name','')} | {str(r.get('start_at',''))[:10]} | {r.get('destination','')}"
                for _, r in approved_trip.iterrows()
            }
            rid = st.selectbox("복명할 출장 선택", list(labels.keys()), format_func=lambda x: labels[x])
            rr = approved_trip[approved_trip["id"] == rid].iloc[0]

            st.markdown("##### 📌 신청 내용 자동연계")
            a,b,c,d = st.columns(4)
            a.metric("출장자", str(rr.get("emp_name","")))
            b.metric("출장지", str(rr.get("destination","")))
            c.metric("이동수단", str(rr.get("transport_type","")))
            d.metric("거리", f"{float(rr.get('distance_km',0) or 0):,.1f} km")
            st.write("**출장목적:**", rr.get("purpose",""))
            st.write("**출장시간:**", str(rr.get("start_at","")), "~", str(rr.get("end_at","")))

            participants = st.text_input("참석자", value=str(rr.get("participants","") or ""))
            default_report = """[참여 목적]

[참여 내용]

[주요 교육·업무 내용]

[업무 적용 및 결과]
"""
            report_content = st.text_area(
                "출장 보고 내용",
                value=str(rr.get("report_content","") or default_report),
                height=280
            )

            st.markdown("##### 💰 여비 정산")
            e1,e2,e3,e4 = st.columns(4)
            with e1:
                transport_cost = st.number_input("교통비/자가차량 운임", min_value=0, step=1000, value=int(rr.get("transport_cost",0) or 0))
            with e2:
                toll_cost = st.number_input("통행료", min_value=0, step=1000, value=int(rr.get("toll_cost",0) or 0))
            with e3:
                lodging_cost = st.number_input("숙박비", min_value=0, step=1000, value=int(rr.get("lodging_cost",0) or 0))
            with e4:
                other_cost = st.number_input("기타 증빙경비", min_value=0, step=1000, value=int(rr.get("other_cost",0) or 0))

            f1,f2 = st.columns(2)
            with f1:
                daily_cost = st.number_input("일비", min_value=0, step=1000, value=int(rr.get("daily_cost", rr.get("rule_base_amount",0)) or 0))
            with f2:
                meal_cost = st.number_input("식비", min_value=0, step=1000, value=int(rr.get("meal_cost",0) or 0))

            total_trip_cost = int(transport_cost+toll_cost+lodging_cost+other_cost+daily_cost+meal_cost)
            st.metric("정산 합계", f"{total_trip_cost:,}원")
            st.caption("국내여비 기준표: 모든 직원 일비 25,000원/일, 식비 20,000원/일, 숙박료 실비(광역시 상한 80,000원, 그 밖의 지역 70,000원). 별도 단시간 출장여비 기준도 함께 적용되므로 담당자 확인 후 확정하도록 구성했습니다.")

            if st.button("💾 복명 및 정산내용 저장", type="primary", use_container_width=True, disabled=not has_permission("trip_settle")):
                supabase.table("business_trips").update({
                    "participants": participants,
                    "report_content": report_content,
                    "transport_cost": int(transport_cost),
                    "toll_cost": int(toll_cost),
                    "lodging_cost": int(lodging_cost),
                    "other_cost": int(other_cost),
                    "daily_cost": int(daily_cost),
                    "meal_cost": int(meal_cost),
                    "total_cost": total_trip_cost,
                    "report_status": "완료",
                    "settlement_status": "정산대기"
                }).eq("id", rid).execute()
                st.success("출장 복명 및 정산내용을 저장했습니다.")
                st.rerun()

    st.divider()
    st.subheader("📚 시스템에서 확인하는 출장·여비 규정")
    reg_tab1, reg_tab2, reg_tab3 = st.tabs(["출장 복무규정", "여비관리 세칙", "교육·연수 결과보고"])
    with reg_tab1:
        st.markdown("""
- **제25조 출장명령**: 출장신청서 제출 및 출장명령 필요
- **제26조 출장 중 사정변경**: 목적지 변경·기간 연장 시 사전승인, 부득이한 경우 귀원 즉시 사후승인
- **제27조 출장보고**: 귀원 후 지체 없이 출장보고서 제출
- **제28조 출장여비**: 「여비관리 세칙」 적용
        """)
    with reg_tab2:
        st.markdown("""
**여비관리 세칙 주요 기준**
- 여비는 출장 계획의 경로와 방법에 따라 계산
- 다른 기관에서 여비가 지급되면 해당 금액만큼 감액
- 센터 교통수단 또는 요금이 들지 않는 교통수단 이용 시 자동차운임 미지급
- 출장여행시간 **4시간 이상 20,000원 / 4시간 미만 10,000원**
- 보행 가능 거리나 **2시간 이내는 실교통비 지급 가능**
- 센터차량 배정 시 위 단시간 출장여비 기준에서 **10,000원 감액**
- 국내여비 기준표: 일비 **25,000원/일**, 식비 **20,000원/일**
- 숙박료: 실비, 상한 **광역시 80,000원 / 그 밖의 지역 70,000원**
- 자가차량 운임: **총거리(km) × 기준단가**, 유종별 연비기준은 가솔린 12km/L, 디젤 10km/L, LPG 8km/L
- 도로통행료는 센터차량·자가차량 구분 없이 실비 지급
        """)
    with reg_tab3:
        st.markdown("""
**제90조 결과보고**
- 위탁교육자는 **교육이수 후 5일 이내** 보고서를 제출
- 교재·출석표·수료증 첨부
- 소속부서장을 경유하여 센터장에게 제출
- 정당한 사유 없이 기한 내 수료증·보고서를 제출하지 않으면 규정상 교육 미이수로 볼 수 있음
        """)


    st.divider()
    st.subheader("🖨️ 출장신청서 · 출장복명서 A4 출력")

    if not df_trip2.empty:
        print_ids = df_trip2["id"].tolist()
        print_id = st.selectbox(
            "출력할 출장 선택",
            print_ids,
            format_func=lambda x: f"#{x} | {df_trip2.loc[df_trip2['id']==x, 'emp_name'].iloc[0]} | {str(df_trip2.loc[df_trip2['id']==x, 'start_at'].iloc[0])[:10]}",
            key="trip_print_id"
        )
        pr = df_trip2[df_trip2["id"] == print_id].iloc[0]

        def _fmt_dt(v):
            try:
                d = pd.to_datetime(v)
                return d.strftime("%Y년 %m월 %d일 %H:%M")
            except:
                return str(v or "")

        # 현재 센터에서 사용하던 출장신청서 형식 기반
        trip_apply_html = f"""
        <style>
        * {{box-sizing:border-box;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}}
        body {{font-family:'Malgun Gothic',sans-serif;color:#000;margin:0;background:white;}}
        .toolbar {{margin-bottom:8px;}}
        .pbtn {{width:100%;padding:9px;border:1px solid #bbb;border-radius:6px;background:#fff;cursor:pointer;font-weight:bold;}}
        .paper {{width:190mm;min-height:260mm;margin:auto;padding:8mm 8mm;background:#fff;}}
        .form-table {{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px;}}
        .form-table td,.form-table th {{border:1px solid #111;padding:7px 6px;height:34px;vertical-align:middle;}}
        .title {{font-size:20px;font-weight:700;letter-spacing:8px;text-align:center;text-decoration:underline;}}
        .label {{font-weight:700;text-align:center;background:#f5f5f5!important;}}
        .center {{text-align:center;}} .left {{text-align:left;}}
        .notice {{margin-top:18px;padding-top:14px;border-top:1px dashed #222;font-size:12px;}}
        @page {{size:A4 portrait;margin:8mm;}}
        @media print {{
          .toolbar {{display:none!important;}}
          .paper {{width:auto;min-height:auto;padding:0;}}
          *,td,th {{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}}
        }}
        </style>
        <div class="toolbar"><button class="pbtn" onclick="window.print()">🖨️ 출장신청서 인쇄</button></div>
        <div class="paper">
          <div style="font-size:11px;margin-bottom:4px;">[서식 제1호]</div>
          <table class="form-table">
            <tr><td colspan="7" class="title">출장신청서</td></tr>
            <tr>
              <td colspan="3" rowspan="2"></td>
              <td rowspan="2" class="label" style="width:7%;">결<br>재</td>
              <td class="label">담당</td><td class="label">대리</td><td class="label">센터장</td>
            </tr>
            <tr><td style="height:42px;"></td><td></td><td></td></tr>
            <tr>
              <td class="label" style="width:16%;">성 명</td>
              <td class="center">{pr.get('emp_name','')}</td>
              <td class="label">직 위</td>
              <td colspan="4" class="center">{pr.get('position','')}</td>
            </tr>
            <tr><td class="label">출장목적</td><td colspan="6">{pr.get('purpose','')}</td></tr>
            <tr><td class="label">출장시간</td><td colspan="6" class="center">{_fmt_dt(pr.get('start_at'))} ~ {_fmt_dt(pr.get('end_at'))}</td></tr>
            <tr><td class="label">출 장 지</td><td colspan="2">{pr.get('destination','')}</td><td class="label">부 서</td><td colspan="3">{pr.get('department','')}</td></tr>
            <tr>
              <td class="label">이동사항</td>
              <td colspan="2" class="center">{pr.get('transport_type','')}</td>
              <td class="label">거리구분</td>
              <td colspan="3" class="center">{float(pr.get('distance_km',0) or 0):,.1f} km</td>
            </tr>
          <tr><td colspan="7" class="sign" style="text-align:center;line-height:1.8;padding:12px;">
              위와 같이 출장을 신청합니다.<br>
              {pd.to_datetime(pr.get('start_at')).strftime('%Y년 %m월 %d일') if pr.get('start_at') else ''}<br>
              신청자 : {pr.get('emp_name','')} &nbsp;&nbsp;(인)
            </td></tr>
          </table>
          <div class="notice">* 출장근거가 불충분하면 출장비는 지급하지 않습니다.</div>
        </div>
        """

        # 현재 센터에서 사용하던 출장복명서 형식 기반
        report_text = str(pr.get("report_content","") or "").replace("\n","<br>")
        trip_report_html = f"""
        <style>
        * {{box-sizing:border-box;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}}
        body {{font-family:'Malgun Gothic',sans-serif;color:#000;margin:0;background:#fff;}}
        .toolbar {{margin-bottom:8px;}} .pbtn {{width:100%;padding:9px;border:1px solid #bbb;border-radius:6px;background:#fff;cursor:pointer;font-weight:bold;}}
        .paper {{width:190mm;min-height:270mm;margin:auto;padding:6mm;background:#fff;}}
        table {{width:100%;border-collapse:collapse;table-layout:fixed;font-size:11px;}}
        td {{border:1px solid #111;padding:6px;vertical-align:middle;}}
        .title {{font-size:19px;font-weight:700;letter-spacing:8px;text-align:center;text-decoration:underline;}}
        .label {{font-weight:700;text-align:center;background:#f5f5f5!important;}}
        .report {{height:92mm;vertical-align:top!important;line-height:1.55;}}
        .photo {{height:58mm;text-align:center;color:#777;vertical-align:middle!important;background:#fafafa!important;}}
        .sign {{text-align:center;line-height:1.8;padding:12px;}}
        @page {{size:A4 portrait;margin:7mm;}}
        @media print {{.toolbar{{display:none!important}} .paper{{width:auto;min-height:auto;padding:0}} *,td{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}}}}
        </style>
        <div class="toolbar"><button class="pbtn" onclick="window.print()">🖨️ 출장복명서 인쇄</button></div>
        <div class="paper">
        <table>
          <tr><td colspan="7" class="title">출장복명서</td></tr>
          <tr><td colspan="3" rowspan="2"></td><td rowspan="2" class="label">결<br>재</td><td class="label">담당</td><td class="label">대리</td><td class="label">센터장</td></tr>
          <tr><td style="height:36px"></td><td></td><td></td></tr>
          <tr><td class="label">성 명</td><td>{pr.get('emp_name','')}</td><td class="label">직 위</td><td colspan="4">{pr.get('position','')}</td></tr>
          <tr><td class="label">출장목적</td><td colspan="6">{pr.get('purpose','')}</td></tr>
          <tr><td class="label">참 석 자</td><td colspan="6">{pr.get('participants','')}</td></tr>
          <tr><td class="label">출장시간</td><td colspan="6">{_fmt_dt(pr.get('start_at'))} ~ {_fmt_dt(pr.get('end_at'))}</td></tr>
          <tr><td class="label">출 장 지</td><td colspan="2">{pr.get('destination','')}</td><td class="label">이동사항</td><td colspan="3">{pr.get('transport_type','')} / {float(pr.get('distance_km',0) or 0):,.1f}km</td></tr>
          <tr><td class="label">여 비</td><td colspan="2">₩ {int(pr.get('total_cost',0) or 0):,}</td><td class="label">정산상태</td><td colspan="3">{pr.get('settlement_status','')}</td></tr>
          <tr><td class="label">출장 보고<br>내용</td><td colspan="6" class="report">{report_text}</td></tr>
          <tr><td class="label">출장 사진</td><td colspan="6" class="photo">첨부 사진 출력 영역</td></tr>
          <tr><td colspan="7" class="sign">금번 출장 결과를 위와 같이 복명합니다.<br><br>{datetime.now().strftime('%Y년 %m월 %d일')}<br><br>출장인 : {pr.get('emp_name','')} &nbsp;&nbsp;(인)</td></tr>
        </table>
        <div style="font-size:10px;margin-top:8px;">* 출장근거가 불충분하면 출장비는 지급하지 않습니다.</div>
        </div>
        """

        pc1, pc2 = st.columns(2)
        with pc1:
            st.components.v1.html(trip_apply_html, height=55, scrolling=False)
        with pc2:
            st.components.v1.html(trip_report_html, height=55, scrolling=False)

    st.divider()
    st.subheader("📎 출장 사진 · 영수증 · 수료증 첨부")
    st.caption("출장별 증빙파일을 Supabase Storage의 business-trip-files 버킷에 저장합니다.")
    if not df_trip2.empty:
        file_trip_id = st.selectbox("증빙을 연결할 출장 ID", df_trip2["id"].tolist(), key="trip_file_id")
        file_type = st.selectbox("증빙 종류", ["출장사진","교통영수증","주차/통행료 영수증","교육수료증","교육자료","기타"])
        uploaded_trip_files = st.file_uploader(
            "파일 선택",
            type=["jpg","jpeg","png","pdf"],
            accept_multiple_files=True,
            key="trip_files"
        )
        if st.button("📤 출장 증빙 업로드", use_container_width=True, disabled=not has_permission("file_upload")):
            if not uploaded_trip_files:
                st.warning("업로드할 파일을 선택해 주세요.")
            else:
                # 버킷 목록 조회는 anon 권한에서 실제 버킷이 있어도 빈 목록이 반환될 수 있으므로
                # 사전 존재검사를 하지 않고 실제 업로드 결과로 판단합니다.
                ok, fail = 0, 0
                for uf in uploaded_trip_files:
                    # Supabase Storage object key는 ASCII 기반으로 생성.
                    # 한글 원본 파일명은 DB file_name에 그대로 보존합니다.
                    original_name = str(uf.name)
                    ext = Path(original_name).suffix.lower()
                    if not re.fullmatch(r"\.[a-z0-9]{1,10}", ext):
                        mime_ext = {
                            "image/jpeg": ".jpg",
                            "image/png": ".png",
                            "application/pdf": ".pdf",
                        }
                        ext = mime_ext.get(str(uf.type).lower(), ".bin")

                    object_name = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{uuid.uuid4().hex[:10]}{ext}"
                    storage_path = f"trip-{int(file_trip_id)}/{object_name}"

                    try:
                        supabase.storage.from_(TRIP_STORAGE_BUCKET).upload(
                            storage_path,
                            uf.getvalue(),
                            {"content-type": uf.type or "application/octet-stream", "upsert": "false"}
                        )
                        supabase.table("business_trip_files").insert({
                            "trip_id": int(file_trip_id),
                            "file_type": file_type,
                            "file_name": original_name,
                            "storage_path": storage_path
                        }).execute()
                        ok += 1
                    except Exception as e:
                        fail += 1
                        err = str(e)
                        if "row-level security" in err.lower() or "unauthorized" in err.lower() or "403" in err:
                            st.error(
                                f"'{uf.name}' 업로드 실패: Storage 업로드 권한이 없습니다. "
                                "Supabase의 storage.objects 정책을 적용해 주세요."
                            )
                        elif "Bucket not found" in err or "404" in err:
                            st.error(
                                f"'{uf.name}' 업로드 실패: 앱이 연결된 Supabase 프로젝트에서 "
                                f"'{TRIP_STORAGE_BUCKET}' 버킷을 찾지 못했습니다. "
                                "현재 앱의 SUPABASE_URL과 버킷을 만든 프로젝트가 같은지 확인해 주세요."
                            )
                        elif "InvalidKey" in err or "Invalid key" in err:
                            st.error(
                                f"'{uf.name}' 업로드 실패: Storage 객체 경로가 유효하지 않습니다. "
                                "v9에서는 원본 한글 파일명 대신 안전한 영문 객체키를 자동 생성합니다."
                            )
                        else:
                            st.error(f"'{uf.name}' 업로드 실패: {err}")

                if ok:
                    st.success(f"✅ {ok}개 출장 증빙파일을 저장했습니다.")
                if fail:
                    st.warning(f"⚠️ {fail}개 파일 업로드에 실패했습니다.")
        try:
            f_res = supabase.table("business_trip_files").select("*").eq("trip_id", int(file_trip_id)).order("id", desc=True).execute()
            df_files = pd.DataFrame(f_res.data) if f_res.data else pd.DataFrame()
            if not df_files.empty:
                display_table_kr(df_files[[c for c in ["file_type","file_name","created_at"] if c in df_files.columns]], use_container_width=True, hide_index=True)
        except Exception:
            pass

    st.divider()
    st.subheader("🔄 출장 변경 · 사후승인 이력")
    if not df_trip2.empty:
        change_trip_id = st.selectbox("변경할 출장 ID", df_trip2["id"].tolist(), key="trip_change_id")
        change_type = st.selectbox("변경 구분", ["목적지 변경","출장기간 변경","이동수단 변경","기타"])
        change_reason = st.text_area("변경 사유 / 보고 내용", key="trip_change_reason")
        cc1, cc2 = st.columns(2)
        with cc1:
            new_destination = st.text_input("변경 출장지(해당 시)", key="trip_new_dest")
        with cc2:
            approval_type = st.radio("승인 구분", ["사전승인","사후승인"], horizontal=True, key="trip_approval_type")
        if st.button("📝 출장 변경 신청 기록", use_container_width=True):
            if not change_reason:
                st.warning("변경 사유를 입력해 주세요.")
            else:
                supabase.table("business_trip_changes").insert({
                    "trip_id": int(change_trip_id),
                    "change_type": change_type,
                    "change_reason": change_reason,
                    "new_destination": new_destination,
                    "approval_type": approval_type,
                    "approval_status": "신청"
                }).execute()
                st.success("출장 변경/승인 이력을 등록했습니다.")
                st.rerun()

        try:
            ch_res = supabase.table("business_trip_changes").select("*").eq("trip_id", int(change_trip_id)).order("id", desc=True).execute()
            df_ch = pd.DataFrame(ch_res.data) if ch_res.data else pd.DataFrame()
            if not df_ch.empty:
                display_table_kr(df_ch, use_container_width=True, hide_index=True)
                pending = df_ch[df_ch["approval_status"]=="신청"] if "approval_status" in df_ch.columns else pd.DataFrame()
                if not pending.empty:
                    ch_id = st.selectbox("승인 처리할 변경 ID", pending["id"].tolist(), key="trip_change_approve_id")
                    ca1,ca2 = st.columns(2)
                    with ca1:
                        if st.button("✅ 변경 승인", use_container_width=True, disabled=not has_permission("trip_approve")):
                            supabase.table("business_trip_changes").update({"approval_status":"승인","approved_at":datetime.now().isoformat()}).eq("id", ch_id).execute()
                            st.rerun()
                    with ca2:
                        if st.button("↩️ 변경 반려", use_container_width=True, disabled=not has_permission("trip_approve")):
                            supabase.table("business_trip_changes").update({"approval_status":"반려","approved_at":datetime.now().isoformat()}).eq("id", ch_id).execute()
                            st.rerun()
        except Exception:
            st.caption("출장 변경이력 테이블을 생성하면 이력이 표시됩니다.")

    st.divider()
    st.subheader("📊 월별 출장관리대장 · Excel")
    if not df_trip2.empty:
        tmp = df_trip2.copy()
        tmp["_start"] = pd.to_datetime(tmp["start_at"], errors="coerce")
        years = sorted(tmp["_start"].dropna().dt.year.unique().tolist(), reverse=True)
        if years:
            gy, gm = st.columns(2)
            with gy:
                ledger_year = st.selectbox("조회 연도", years, key="trip_ledger_year")
            with gm:
                ledger_month = st.selectbox("조회 월", list(range(1,13)), index=datetime.now().month-1, key="trip_ledger_month")
            ledger = tmp[(tmp["_start"].dt.year==ledger_year)&(tmp["_start"].dt.month==ledger_month)].copy()
            ledger_cols = [c for c in ["id","start_at","end_at","emp_name","position","purpose","destination","transport_type","distance_km","total_cost","apply_status","report_status","settlement_status"] if c in ledger.columns]
            display_table_kr(ledger[ledger_cols], use_container_width=True, hide_index=True)

            lm1,lm2,lm3,lm4 = st.columns(4)
            lm1.metric("출장 건수", f"{len(ledger)}건")
            lm2.metric("복명 완료", f"{int((ledger['report_status']=='완료').sum()) if 'report_status' in ledger else 0}건")
            lm3.metric("총 출장거리", f"{pd.to_numeric(ledger.get('distance_km',0),errors='coerce').fillna(0).sum():,.2f}km")
            lm4.metric("총 출장비", f"{pd.to_numeric(ledger.get('total_cost',0),errors='coerce').fillna(0).sum():,.0f}원")

            trip_xlsx = io.BytesIO()
            wb = Workbook()
            ws = wb.active
            ws.title = f"{ledger_year}-{ledger_month:02d}_출장대장"
            title_fill = PatternFill("solid", fgColor="D9EAF7")
            header_fill = PatternFill("solid", fgColor="FFF2CC")
            done_fill = PatternFill("solid", fgColor="E2F0D9")
            thin = Side(style="thin", color="808080")
            bd = Border(left=thin,right=thin,top=thin,bottom=thin)
            ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(ledger_cols))
            ws.cell(1,1,f"{ledger_year}년 {ledger_month}월 출장관리대장")
            ws.cell(1,1).font = Font(name="맑은 고딕",size=16,bold=True)
            ws.cell(1,1).fill = title_fill
            ws.cell(1,1).alignment = Alignment(horizontal="center")
            labels = {"id":"No","start_at":"출장시작","end_at":"출장종료","emp_name":"성명","position":"직위","purpose":"출장목적",
                      "destination":"출장지","transport_type":"이동수단","distance_km":"거리(km)","total_cost":"출장비",
                      "apply_status":"승인","report_status":"복명","settlement_status":"정산"}
            for ci,cname in enumerate(ledger_cols,1):
                c=ws.cell(3,ci,labels.get(cname,cname)); c.fill=header_fill;c.font=Font(name="맑은 고딕",bold=True);c.border=bd;c.alignment=Alignment(horizontal="center")
            for ri,(_,row) in enumerate(ledger.iterrows(),4):
                for ci,cname in enumerate(ledger_cols,1):
                    val=row[cname]
                    if cname in ["start_at","end_at"]:
                        try: val=pd.to_datetime(val).strftime("%Y-%m-%d %H:%M")
                        except: pass
                    c=ws.cell(ri,ci,val);c.border=bd;c.font=Font(name="맑은 고딕",size=9)
                    if str(row.get("report_status",""))=="완료": c.fill=done_fill
                    c.alignment=Alignment(horizontal="center" if cname not in ["purpose","destination"] else "left")
                    if cname=="total_cost": c.number_format='#,##0"원"'
                    if cname=="distance_km": c.number_format='0.0"km"'
            widths=[7,18,18,11,11,30,22,12,11,13,10,10,10]
            for i,w in enumerate(widths[:len(ledger_cols)],1): ws.column_dimensions[get_column_letter(i)].width=w
            ws.sheet_view.showGridLines=False
            ws.page_setup.orientation="landscape";ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0
            ws.sheet_properties.pageSetUpPr.fitToPage=True
            wb.save(trip_xlsx)
            st.download_button("📥 월별 출장관리대장 Excel 다운로드",trip_xlsx.getvalue(),file_name=f"{ledger_year}-{ledger_month:02d}_출장관리대장.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)






# -------------------------------------------------------------------
# TAB 12: 계정 · 권한 관리 (관리자 전용)
# -------------------------------------------------------------------
if active_tab == 12:
    st.header("🔐 계정·권한 관리")
    st.caption("관리자 전용 · 직원 계정 생성, 역할 변경, 비밀번호 초기화")

    if CURRENT_ROLE != "admin":
        st.error("관리자만 접근할 수 있습니다.")
        st.stop()

    if supabase_admin is None:
        st.error("관리자 기능을 사용하려면 Streamlit Secrets에 SUPABASE_SECRET_KEY를 등록해야 합니다.")
        st.code('SUPABASE_SECRET_KEY = "Supabase Secret/Service Role Key"', language="toml")
        st.stop()

    st.info("관리자 Secret Key는 계정 생성·비밀번호 변경 같은 Auth Admin 작업에만 사용하며 화면에는 노출하지 않습니다.")

    # 직원 DB
    try:
        _emp_res = supabase.table("employees").select("*").order("emp_name").execute()
        _emp_df = pd.DataFrame(_emp_res.data) if _emp_res.data else pd.DataFrame()
    except Exception:
        _emp_df = pd.DataFrame()

    # Auth 사용자
    try:
        _users_res = supabase_admin.auth.admin.list_users(page=1, per_page=1000)
        _users = getattr(_users_res, "users", _users_res)
        if not isinstance(_users, list):
            _users = []
    except Exception as e:
        _users = []
        st.error(f"Auth 사용자 목록 조회 실패: {e}")

    def _auth_field(u, name, default=""):
        if isinstance(u, dict):
            return u.get(name, default)
        return getattr(u, name, default)

    account_rows = []
    for u in _users:
        meta = _auth_field(u, "user_metadata", {}) or {}
        account_rows.append({
            "user_id": str(_auth_field(u, "id", "")),
            "이메일": str(_auth_field(u, "email", "")),
            "이름": str(meta.get("name", "")),
            "사번": str(meta.get("emp_id", "")),
            "권한": _role_label(str(meta.get("role", "employee"))),
            "role_code": str(meta.get("role", "employee")),
            "최근로그인": str(_auth_field(u, "last_sign_in_at", "") or ""),
            "계정상태": "중지" if str(_auth_field(u, "banned_until", "") or "") else "사용",
        })
    _account_df = pd.DataFrame(account_rows)

    m1, m2, m3 = st.columns(3)
    m1.metric("전체 계정", f"{len(_account_df)}명")
    m2.metric("관리자", f"{int((_account_df['role_code']=='admin').sum()) if not _account_df.empty else 0}명")
    m3.metric("일반 직원", f"{int((_account_df['role_code']=='employee').sum()) if not _account_df.empty else 0}명")

    if not _emp_df.empty:
        linked_ids = set(_account_df["사번"].astype(str)) if not _account_df.empty else set()
        _emp_link = _emp_df.copy()
        _emp_link["계정연결"] = _emp_link["emp_id"].astype(str).apply(lambda x: "연결됨" if x in linked_ids else "미연결")
        with st.expander("🔗 직원 ↔ 로그인 계정 연결 현황"):
            _cols = [c for c in ["emp_id","emp_name","dept","position","계정연결"] if c in _emp_link.columns]
            display_table_kr(_emp_link[_cols], use_container_width=True, hide_index=True)

    st.subheader("👤 직원 계정 현황")
    if _account_df.empty:
        st.info("등록된 로그인 계정이 없습니다.")
    else:
        display_table_kr(
            _account_df[["이메일","이름","사번","권한","계정상태","최근로그인"]],
            use_container_width=True, hide_index=True
        )

    st.divider()
    create_tab, role_tab, pw_tab, status_tab = st.tabs(["➕ 계정 생성", "🛡️ 권한 변경", "🔑 비밀번호 초기화", "⏯️ 계정 사용·중지"])

    with create_tab:
        if _emp_df.empty:
            st.warning("직원 DB가 없어 계정을 생성할 수 없습니다.")
        else:
            emp_map = {
                f"{str(r.get('emp_name','')).strip()} ({str(r.get('emp_id','')).strip()})": r
                for _, r in _emp_df.iterrows()
                if str(r.get("emp_name","")).strip()
            }
            selected_emp_label = st.selectbox("계정을 만들 직원", list(emp_map.keys()), key="admin_create_emp")
            selected_emp = emp_map[selected_emp_label]
            new_email = st.text_input("로그인 이메일", key="admin_create_email")
            new_pw = st.text_input("초기 비밀번호", type="password", key="admin_create_pw",
                                   help="직원에게 전달할 임시 비밀번호입니다. 8자 이상을 권장합니다.")
            new_role_label = st.selectbox("권한", ["직원", "담당자", "조회자", "관리자"], key="admin_create_role")
            role_map = {"직원":"employee", "담당자":"manager", "조회자":"viewer", "관리자":"admin"}

            if st.button("계정 생성", type="primary", use_container_width=True, key="admin_create_btn"):
                selected_emp_id = str(selected_emp.get("emp_id","")).strip()
                existing_emp_ids = set(_account_df["사번"].astype(str)) if not _account_df.empty else set()
                if selected_emp_id in existing_emp_ids:
                    st.warning("이 직원은 이미 로그인 계정과 연결되어 있습니다.")
                elif not new_email.strip() or len(new_pw) < 8:
                    st.warning("이메일을 입력하고 초기 비밀번호는 8자 이상으로 설정해 주세요.")
                else:
                    try:
                        result = supabase_admin.auth.admin.create_user({
                            "email": new_email.strip(),
                            "password": new_pw,
                            "email_confirm": True,
                            "user_metadata": {
                                "name": str(selected_emp.get("emp_name","")),
                                "emp_id": str(selected_emp.get("emp_id","")),
                                "dept": str(selected_emp.get("dept","")),
                                "position": str(selected_emp.get("position","")),
                                "role": role_map[new_role_label],
                            }
                        })
                        created = _auth_user_obj(result)
                        write_audit_log(
                            "계정 생성", "auth.users",
                            str(_user_value(created, "id", "")),
                            f"{new_email.strip()} / {new_role_label}"
                        )
                        st.success(f"{selected_emp.get('emp_name','')} 직원의 계정을 생성했습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"계정 생성 실패: {e}")

    with role_tab:
        if _account_df.empty:
            st.info("권한을 변경할 계정이 없습니다.")
        else:
            role_email = st.selectbox("계정 선택", _account_df["이메일"].tolist(), key="admin_role_email")
            row = _account_df[_account_df["이메일"] == role_email].iloc[0]
            target_uid = row["user_id"]
            role_label_map = {"employee":"직원","manager":"담당자","viewer":"조회자","admin":"관리자"}
            labels = ["직원","담당자","조회자","관리자"]
            current_label = role_label_map.get(row["role_code"], "직원")
            new_role2 = st.selectbox("변경할 권한", labels, index=labels.index(current_label), key="admin_role_new")
            role_code_map = {"직원":"employee","담당자":"manager","조회자":"viewer","관리자":"admin"}

            if st.button("권한 저장", type="primary", use_container_width=True, key="admin_role_save"):
                try:
                    # 기존 user_metadata를 유지하면서 role만 변경
                    target_user = next((u for u in _users if str(_auth_field(u,"id","")) == target_uid), None)
                    meta = dict(_auth_field(target_user, "user_metadata", {}) or {})
                    meta["role"] = role_code_map[new_role2]
                    supabase_admin.auth.admin.update_user_by_id(target_uid, {"user_metadata": meta})
                    write_audit_log("계정 권한 변경", "auth.users", target_uid, f"{role_email}: {new_role2}")
                    st.success("권한을 변경했습니다. 해당 사용자는 다음 로그인부터 변경된 권한이 적용됩니다.")
                    st.rerun()
                except Exception as e:
                    st.error(f"권한 변경 실패: {e}")

    with pw_tab:
        if _account_df.empty:
            st.info("비밀번호를 초기화할 계정이 없습니다.")
        else:
            pw_email = st.selectbox("계정 선택", _account_df["이메일"].tolist(), key="admin_pw_email")
            pw_row = _account_df[_account_df["이메일"] == pw_email].iloc[0]
            pw_uid = pw_row["user_id"]
            reset_mode = st.radio(
                "초기화 방식",
                ["관리자가 임시 비밀번호 설정", "비밀번호 재설정 이메일 발송"],
                horizontal=True
            )

            if reset_mode == "관리자가 임시 비밀번호 설정":
                temp_pw = st.text_input("새 임시 비밀번호", type="password", key="admin_temp_pw")
                confirm_pw = st.text_input("새 임시 비밀번호 확인", type="password", key="admin_temp_pw2")
                if st.button("비밀번호 초기화", type="primary", use_container_width=True, key="admin_pw_reset"):
                    if len(temp_pw) < 8:
                        st.warning("임시 비밀번호는 8자 이상으로 설정해 주세요.")
                    elif temp_pw != confirm_pw:
                        st.warning("비밀번호 확인이 일치하지 않습니다.")
                    else:
                        try:
                            supabase_admin.auth.admin.update_user_by_id(pw_uid, {"password": temp_pw})
                            write_audit_log("비밀번호 관리자 초기화", "auth.users", pw_uid, pw_email)
                            st.success("임시 비밀번호로 초기화했습니다.")
                        except Exception as e:
                            st.error(f"비밀번호 초기화 실패: {e}")
            else:
                st.caption("직원의 로그인 이메일로 Supabase 비밀번호 재설정 메일을 발송합니다.")
                if st.button("재설정 이메일 발송", type="primary", use_container_width=True, key="admin_pw_mail"):
                    try:
                        supabase.auth.reset_password_for_email(pw_email)
                        write_audit_log("비밀번호 재설정 메일", "auth.users", pw_uid, pw_email)
                        st.success("비밀번호 재설정 이메일을 발송했습니다.")
                    except Exception as e:
                        st.error(f"재설정 이메일 발송 실패: {e}")
    with status_tab:
        if _account_df.empty:
            st.info("관리할 계정이 없습니다.")
        else:
            status_email = st.selectbox("계정 선택", _account_df["이메일"].tolist(), key="admin_status_email")
            status_row = _account_df[_account_df["이메일"] == status_email].iloc[0]
            status_uid = status_row["user_id"]
            current_status = status_row["계정상태"]
            st.write(f"현재 상태: **{current_status}**")
            sc1, sc2 = st.columns(2)
            with sc1:
                if st.button("⛔ 계정 중지", use_container_width=True, disabled=(current_status=="중지" or status_email==CURRENT_EMAIL)):
                    try:
                        supabase_admin.auth.admin.update_user_by_id(status_uid, {"ban_duration":"876000h"})
                        write_audit_log("계정 중지", "auth.users", status_uid, status_email)
                        st.success("계정 사용을 중지했습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"계정 중지 실패: {e}")
            with sc2:
                if st.button("▶️ 계정 사용", use_container_width=True, disabled=(current_status=="사용")):
                    try:
                        supabase_admin.auth.admin.update_user_by_id(status_uid, {"ban_duration":"none"})
                        write_audit_log("계정 사용 재개", "auth.users", status_uid, status_email)
                        st.success("계정을 다시 사용할 수 있습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"계정 사용 재개 실패: {e}")
            if status_email == CURRENT_EMAIL:
                st.caption("현재 로그인 중인 관리자 본인 계정은 중지할 수 없도록 보호됩니다.")


# -------------------------------------------------------------------

# TAB 13: v20 관리자 통합 통계·보고서
if active_tab == 13:
    st.subheader("📊 관리자 통합 통계·보고서")
    if not can_manage_all_records():
        st.warning("관리자 또는 담당자만 이용할 수 있습니다.")
    else:
        _now=datetime.now()
        _years=list(range(_now.year-3,_now.year+2))
        _c1,_c2=st.columns(2)
        _year=_c1.selectbox("연도",_years,index=_years.index(_now.year),key="v20_year")
        _month=_c2.selectbox("월",range(1,13),index=_now.month-1,format_func=lambda x:f"{x}월",key="v20_month")
        _ym=f"{_year:04d}-{_month:02d}"

        def _load20(table):
            try: return pd.DataFrame(supabase.table(table).select("*").execute().data or [])
            except Exception: return pd.DataFrame()

        def _month20(df, ym, preferred):
            if df.empty: return df
            col=next((c for c in preferred if c in df.columns),None)
            if not col: return df.iloc[0:0].copy()
            return df[df[col].astype(str).str[:7]==ym].copy()

        def _sum20(df, cols):
            if df.empty:return 0.0
            c=next((x for x in cols if x in df.columns),None)
            return float(pd.to_numeric(df[c],errors="coerce").fillna(0).sum()) if c else 0.0

        _pay_all=_load20("monthly_payroll_adjust")
        _lv_all=_load20("leave_records")
        _ot_all=_load20("overtime_records")
        _tr_all=_load20("business_trips")
        _emp_all=_load20("employees")

        def _datasets20(ym):
            return (
                _month20(_pay_all,ym,["pay_month","created_at"]),
                _month20(_lv_all,ym,["start_date","leave_date","created_at"]),
                _month20(_ot_all,ym,["work_date","ot_date","created_at"]),
                _month20(_tr_all,ym,["start_at","start_date","created_at"])
            )

        _pay,_lv,_ot,_tr=_datasets20(_ym)
        _prev_dt=(pd.Timestamp(f"{_ym}-01")-pd.offsets.MonthBegin(1))
        _pym=_prev_dt.strftime("%Y-%m")
        _ppay,_plv,_pot,_ptr=_datasets20(_pym)

        _pay_sum=_sum20(_pay,["gross_pay","total_pay","pay_total","total_salary","net_pay"])
        _lv_sum=_sum20(_lv,["used_days","leave_days","days","day_count"])
        _ot_sum=_sum20(_ot,["actual_duration_hours","actual_hours","work_hours","duration_hours","hours"])
        _tr_sum=_sum20(_tr,["total_cost"])
        _prev=[_sum20(_ppay,["gross_pay","total_pay","pay_total","total_salary","net_pay"]),
               _sum20(_plv,["used_days","leave_days","days","day_count"]),
               _sum20(_pot,["actual_duration_hours","actual_hours","work_hours","duration_hours","hours"]),
               _sum20(_ptr,["total_cost"])]

        tab_month,tab_year,tab_check,tab_export=st.tabs(["월간 현황","연간 현황","데이터 점검","Excel 통합보고서"])

        with tab_month:
            st.markdown(f"#### {_year}년 {_month}월 통합 현황")
            st.caption("※ 초과근무 월간 통계는 실제 근무일(1일~말일) 기준입니다. 급여 연계의 전월 25일~당월 24일 기준과는 다릅니다.")
            a,b,c,d=st.columns(4)
            a.metric("급여",f"{int(_pay_sum):,}원" if _pay_sum else f"{len(_pay)}건",
                     delta=f"{int(_pay_sum-_prev[0]):+,}원 전월대비")
            b.metric("연차",f"{_lv_sum:,.2f}일" if _lv_sum else f"{len(_lv)}건",
                     delta=f"{_lv_sum-_prev[1]:+.2f}일 전월대비")
            c.metric("초과근무",f"{_ot_sum:,.2f}시간" if _ot_sum else f"{len(_ot)}건",
                     delta=f"{_ot_sum-_prev[2]:+.2f}시간 전월대비")
            d.metric("출장비",f"{int(_tr_sum):,}원" if _tr_sum else f"{len(_tr)}건",
                     delta=f"{int(_tr_sum-_prev[3]):+,}원 전월대비")
            st.markdown("#### 월간 상세자료")
            m1,m2,m3,m4=st.tabs(["급여","연차","초과근무","출장"])
            with m1:
                display_table_kr(_pay,use_container_width=True,hide_index=True)
            with m2:
                _lv_view=employee_filter_ui(_lv,"v201_report_leave_emp","연차 직원 검색")
                display_table_kr(_lv_view,use_container_width=True,hide_index=True)
            with m3:
                _ot_view=employee_filter_ui(_ot,"v201_report_ot_emp","초과근무 직원 검색")
                display_table_kr(_ot_view,use_container_width=True,hide_index=True)
            with m4:
                _tr_view=employee_filter_ui(_tr,"v201_report_trip_emp","출장 직원 검색")
                display_table_kr(_tr_view,use_container_width=True,hide_index=True)

            st.markdown("#### 🔎 연차·초과근무·출장 기간 검색")
            st.caption("예: 시작월 3월 / 종료월 5월을 선택하면 해당 연도 3월 1일부터 5월 말일까지 조회합니다.")
            _pc1,_pc2=st.columns(2)
            _sm=_pc1.selectbox("시작월",range(1,13),index=max(0,_month-1),format_func=lambda x:f"{x}월",key="v202_start_month")
            _em=_pc2.selectbox("종료월",range(1,13),index=max(0,_month-1),format_func=lambda x:f"{x}월",key="v202_end_month")
            if _sm > _em:
                st.warning("시작월은 종료월보다 클 수 없습니다.")
            else:
                _range_start=f"{_year:04d}-{_sm:02d}-01"
                _range_end=(pd.Timestamp(f"{_year:04d}-{_em:02d}-01")+pd.offsets.MonthEnd(1)).strftime("%Y-%m-%d")

                def _range20(df, preferred):
                    if df.empty:
                        return df
                    col=next((c for c in preferred if c in df.columns),None)
                    if not col:
                        return df.iloc[0:0].copy()
                    # Supabase timestamptz(+00:00)와 일반 날짜가 섞여 있어도
                    # 모두 UTC 기준으로 변환한 뒤 날짜만 비교하여 tz-aware/naive 충돌 방지
                    ds=pd.to_datetime(df[col],errors="coerce",utc=True)
                    start_ts=pd.to_datetime(_range_start,utc=True)
                    end_ts=pd.to_datetime(_range_end+" 23:59:59",utc=True)
                    return df[(ds>=start_ts) & (ds<=end_ts)].copy()

                _rlv=_range20(_lv_all,["start_date","leave_date","created_at"])
                _rot=_range20(_ot_all,["work_date","ot_date","created_at"])
                _rtr=_range20(_tr_all,["start_at","start_date","created_at"])

                st.caption(f"조회기간: {_range_start} ~ {_range_end}")
                _r1,_r2,_r3=st.tabs(["연차 기간조회","초과근무 기간조회","출장 기간조회"])
                with _r1:
                    _rlv_view=employee_filter_ui(_rlv,"v202_range_leave_emp","연차 직원 검색")
                    display_table_kr(_rlv_view,use_container_width=True,hide_index=True)
                with _r2:
                    _rot_view=employee_filter_ui(_rot,"v202_range_ot_emp","초과근무 직원 검색")
                    display_table_kr(_rot_view,use_container_width=True,hide_index=True)
                with _r3:
                    _rtr_view=employee_filter_ui(_rtr,"v202_range_trip_emp","출장 직원 검색")
                    display_table_kr(_rtr_view,use_container_width=True,hide_index=True)

        with tab_year:
            rows=[]
            for mm in range(1,13):
                ym=f"{_year:04d}-{mm:02d}"
                p,l,o,r=_datasets20(ym)
                rows.append({"월":f"{mm}월",
                    "급여총액":_sum20(p,["gross_pay","total_pay","pay_total","total_salary","net_pay"]),
                    "연차사용일수":_sum20(l,["used_days","leave_days","days","day_count"]),
                    "초과근무시간":_sum20(o,["actual_duration_hours","actual_hours","work_hours","duration_hours","hours"]),
                    "출장비":_sum20(r,["total_cost"])})
            _annual=pd.DataFrame(rows)
            display_table_kr(_annual,use_container_width=True,hide_index=True)
            st.markdown("##### 연간 월별 추이")
            _chart=_annual.copy()
            _chart["월번호"]=range(1,13)
            st.line_chart(_chart.set_index("월번호")[["연차사용일수","초과근무시간"]])
            st.bar_chart(_chart.set_index("월번호")[["급여총액","출장비"]])

        with tab_check:
            st.markdown("#### ⚠️ 데이터 점검")
            issues=[]
            if not _pay.empty:
                pc=next((c for c in ["gross_pay","total_pay","pay_total","total_salary","net_pay"] if c in _pay.columns),None)
                if pc:
                    n=int((pd.to_numeric(_pay[pc],errors="coerce").fillna(0)<=0).sum())
                    if n: issues.append({"점검항목":"급여 0원/누락","건수":n,"조치":"급여대장 확인"})
            if not _ot.empty:
                hc=next((c for c in ["actual_duration_hours","actual_hours","work_hours"] if c in _ot.columns),None)
                if hc:
                    n=int(pd.to_numeric(_ot[hc],errors="coerce").isna().sum())
                    if n: issues.append({"점검항목":"초과근무 실적시간 누락","건수":n,"조치":"초과근무 실적 확인"})
            if not _tr.empty:
                if "report_status" in _tr.columns:
                    n=int((~_tr["report_status"].astype(str).isin(["완료","승인"])).sum())
                    if n: issues.append({"점검항목":"출장 복명 미완료","건수":n,"조치":"출장 복명 확인"})
                if "settlement_status" in _tr.columns:
                    n=int((~_tr["settlement_status"].astype(str).isin(["완료","승인"])).sum())
                    if n: issues.append({"점검항목":"출장비 미정산","건수":n,"조치":"출장 정산 확인"})
            if not _emp_all.empty:
                required=[c for c in ["emp_id","emp_name","dept","position"] if c in _emp_all.columns]
                if required:
                    n=int(_emp_all[required].isna().any(axis=1).sum())
                    if n: issues.append({"점검항목":"직원 기본정보 누락","건수":n,"조치":"직원관리 확인"})
            if issues: display_table_kr(pd.DataFrame(issues),use_container_width=True,hide_index=True)
            else: st.success("선택한 월 기준 주요 데이터 점검에서 이상 항목이 발견되지 않았습니다.")

        with tab_export:
            st.markdown("#### 📥 월간·연간 통합보고서")
            try:
                from io import BytesIO
                _bio=BytesIO()
                with pd.ExcelWriter(_bio,engine="openpyxl") as w:
                    pd.DataFrame([{"기준월":_ym,"직원수":len(_emp_all),"급여합계":_pay_sum,
                        "연차사용일수":_lv_sum,"초과근무시간":_ot_sum,"출장비":_tr_sum}]).to_excel(w,sheet_name="업무현황요약",index=False)
                    excel_view_df(_pay).to_excel(w,sheet_name="급여",index=False)
                    excel_view_df(_lv).to_excel(w,sheet_name="연차",index=False)
                    excel_view_df(_ot).to_excel(w,sheet_name="초과근무",index=False)
                    excel_view_df(_tr).to_excel(w,sheet_name="출장",index=False)
                    excel_view_df(_annual).to_excel(w,sheet_name="연간현황",index=False)
                    for ws in w.book.worksheets: style_excel_sheet(ws)
                st.download_button("📥 월간·연간 통합 Excel 다운로드",_bio.getvalue(),
                    file_name=f"통합업무보고서_{_ym}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"보고서 생성 중 확인이 필요합니다: {e}")


# 모든 업무화면 공통 하단 회사 로고
# 기존 상단 표시 크기 대비 약 80% 수준(136px)으로 축소
# -------------------------------------------------------------------
_footer_logo = load_company_logo_bytes()
if _footer_logo:
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    _fc1, _fc2, _fc3 = st.columns([4, 1, 4])
    with _fc2:
        st.image(_footer_logo, width=136)
    st.markdown(
        f"<div style='text-align:center;color:#98a2b3;font-size:.72rem;margin-top:3px;'>System Version {APP_VERSION}</div>",
        unsafe_allow_html=True
    )
