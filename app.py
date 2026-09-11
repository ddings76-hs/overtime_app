import streamlit as st
import pandas as pd
from datetime import datetime, time
import io
import base64
import json
import hashlib
import zlib
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from supabase import create_client, Client

# 페이지 기본 설정
st.set_page_config(page_title="통합 급여·초과근무·연차 관리 시스템 V1.5.1", layout="wide")

# -------------------------------------------------------------------
# Supabase 클라우드 DB 연결 설정 (Secrets 참조)
# -------------------------------------------------------------------
try:
    url: str = st.secrets["SUPABASE_URL"]
    key: str = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(url, key)
except Exception as e:
    st.error("Supabase 연결 실패! Streamlit Secrets에 SUPABASE_URL과 SUPABASE_KEY가 정상 설정되었는지 확인이 필요하다.")
    st.stop()

def truncate_ten(value):
    return int(value // 10) * 10

def safe_int(value):
    """pandas/numpy 숫자를 JSON 저장이 가능한 정수로 통일한다."""
    if pd.isna(value):
        return 0
    return int(value)

# 국세청 근로소득 간이세액표(2026.03.01 시행) 중 공제대상가족 1명 열.
# 배포 시 별도 엑셀 파일 없이 동일 기준을 적용하도록 원본 표의 구간·세액을 압축 내장한다.
_TAX_TABLE_1_PERSON_B64 = "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"

def _load_tax_table_1_person():
    raw = zlib.decompress(base64.b64decode(_TAX_TABLE_1_PERSON_B64)).decode("utf-8")
    return json.loads(raw)

TAX_TABLE_1_PERSON = _load_tax_table_1_person()

def calculate_income_tax_1_person(taxable_monthly_pay):
    """비과세 제외 월 급여를 기준으로 부양가족 1명·자녀 0명·100% 소득세를 계산한다."""
    pay = max(0, safe_int(taxable_monthly_pay))
    if pay < 10_000_000:
        for lower, upper, tax in TAX_TABLE_1_PERSON:
            if lower <= pay < upper:
                return safe_int(tax)
        return 0

    base_tax = 1_507_400
    if pay == 10_000_000:
        return base_tax
    if pay <= 14_000_000:
        tax = base_tax + (pay - 10_000_000) * 0.98 * 0.35 + 25_000
    elif pay <= 28_000_000:
        tax = base_tax + 1_397_000 + (pay - 14_000_000) * 0.98 * 0.38
    elif pay <= 30_000_000:
        tax = base_tax + 6_610_600 + (pay - 28_000_000) * 0.98 * 0.40
    elif pay <= 45_000_000:
        tax = base_tax + 7_394_600 + (pay - 30_000_000) * 0.40
    elif pay <= 87_000_000:
        tax = base_tax + 13_394_600 + (pay - 45_000_000) * 0.42
    else:
        tax = base_tax + 31_034_600 + (pay - 87_000_000) * 0.45
    return truncate_ten(tax)

def calculate_local_income_tax(income_tax):
    """지방소득세는 산출 소득세의 10%를 10원 단위로 절사한다."""
    return truncate_ten(safe_int(income_tax) * 0.10)

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

    # 회계 지출에는 총급여·사업주 보험·퇴직적립만 반영한다.
    # 실지급액과 근로자 공제액은 payroll_summary에만 두어 중복 지출을 방지한다.
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
    canonical = json.dumps(
        {"payroll": snapshot, "accounting": accounting_export},
        ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()

# 세션 내 로고 이미지 관리
if 'logo_b64' not in st.session_state:
    st.session_state.logo_b64 = ""

st.title("🏢 통합 급여·초과근무·연차 관리 시스템")

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

# 탭 구성 (tab1 ~ tab9)
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
    "👥 직원 등록 및 정보 관리", 
    "📝 초과/휴일근무 신청", 
    "✅ 초과근무 수행 입력 & 요약표", 
    "🌴 연차 관리 & 전 직원 요약표",
    "🖨️ 연차 신청서 출력",
    "📊 통합 급여대장 (수정 및 엑셀)", 
    "📄 개별 급여명세서 인쇄",
    "🖨️ 통합 급여대장 인쇄",
    "📑 월별 급여대장 총괄표"
])

# -------------------------------------------------------------------
# TAB 1: 직원 등록 및 정보 관리 (순서 변경 및 소수점 1자리 적용)
# -------------------------------------------------------------------
with tab1:
    st.header("1. 직원 데이터 조회 및 편집")
    try:
        emp_res = supabase.table("employees").select("*").execute()
        df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()
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
            "국민연금(본인)", "건강보험(본인)", "장기요양(본인)", "고용보험(본인)",
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
            column_config=employee_config,
            disabled=[col for col in ["소득세", "지방소득세"] if col in display_emp.columns]
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
            # 소수점 1자리까지 입력 가능하도록 step 및 format 설정
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

        st.markdown("**직원 부담 보험료·세금 및 회사 부담금 기본값**  ")
        st.caption("입력한 금액은 새 월 급여대장의 최초값으로 사용되며, 급여대장에서 다시 수정할 수 있습니다.")
        d1, d2, d3 = st.columns(3)
        with d1:
            national_pension = st.number_input("국민연금 본인부담", min_value=0, value=0, step=10)
            health_insurance = st.number_input("건강보험 본인부담", min_value=0, value=0, step=10)
            longterm_care = st.number_input("장기요양 본인부담", min_value=0, value=0, step=10)
            employment_insurance = st.number_input("고용보험 본인부담", min_value=0, value=0, step=10)
        with d2:
            st.info("소득세는 부양가족 1명·자녀 0명·100% 기준으로 자동 계산되며, 지방소득세는 소득세의 10%로 자동 계산됩니다.")
            income_tax = 0
            local_tax = 0
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
with tab2:
    st.header("1. 초과근무 / 휴일근무 사전 신청")
    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()

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
# TAB 3: 실제 수행 입력 & 삭제 기능 & 월별 승인 요약표
# -------------------------------------------------------------------
with tab3:
    st.header("✅ 실제 초과근무 수행 내역 입력 & 월별 승인 요약표")
    
    ot_res = supabase.table("overtime_records").select("*").order("id", desc=True).execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()

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
        st.subheader("🖨️ 초과근무 신청 및 확인서 인쇄")

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 35px; float: left;">' if st.session_state.logo_b64 else ''
        act_reason_disp = target_ot_latest['act_reason'] if pd.notna(target_ot_latest['act_reason']) and target_ot_latest['act_reason'] != "" else "입력된 실제 수행 내용 없음"

        ot_confirm_template = f"""
        <div style="text-align: right; margin-bottom: 10px;">
            <button onclick="window.print()" style="padding: 8px 16px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">🖨️ 해당 서식 인쇄하기</button>
        </div>
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 680px; margin: auto; background: #fff;">
            {logo_html}
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; clear: both;">
                <h2 style="margin: 0; padding-top: 15px; font-size: 22px; text-decoration: underline;">초 과 근 무 신 청 및 확 인 서</h2>
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
                    <td style="padding: 6px;" colspan="3">{target_ot_latest['work_type']} (최종 상태: {target_ot_latest['status']})</td>
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

            <p style="text-align: center; margin-top: 35px; font-size: 14px;">위와 같이 초과근무를 신청하고 실제 수행 내역을 확인합니다.</p>
            <p style="text-align: center; margin-top: 10px; font-size: 13px;">{target_ot_latest['work_date'][:4]}년 {target_ot_latest['work_date'][5:7]}월 {target_ot_latest['work_date'][8:10]}일</p>
            
            <p style="text-align: right; margin-top: 30px; font-size: 14px; font-weight: bold; padding-right: 10px;">
                신청자 : {target_ot_latest['emp_name']} (인)
            </p>
        </div>
        """
        st.components.v1.html(ot_confirm_template, height=560, scrolling=True)

        st.divider()
        st.subheader("📊 월별 승인 초과근무 집계 요약표")
        
        current_year = datetime.now().year
        c_y, c_m = st.columns(2)
        with c_y:
            sel_year = st.selectbox("조회 연도 선택", range(current_year - 2, current_year + 3), index=2, key="ot_year_sel")
        with c_m:
            sel_month = st.selectbox("조회 월 선택", range(1, 13), index=datetime.now().month - 1, key="ot_month_sel")
        
        filter_month = f"{sel_year}-{sel_month:02d}"

        ot_m_res = supabase.table("overtime_records").select("*").like("work_date", f"{filter_month}%").eq("status", "승인").execute()
        df_ot_month = pd.DataFrame(ot_m_res.data) if ot_m_res.data else pd.DataFrame()

        if df_ot_month.empty:
            st.info(f"💡 [{filter_month}] 승인 완료된 초과근무 내역이 없다.")
        else:
            summary_ot = df_ot_month.groupby(['emp_id', 'emp_name', 'dept', 'position', 'status']).agg(
                승인_건수=('id', 'count'),
                총_인정시간_h=('actual_duration_hours', 'sum')
            ).reset_index()
            summary_ot['총_인정시간_h'] = summary_ot['총_인정시간_h'].round(1)
            
            st.write(f"**[{filter_month}] 최종 승인된 직원별 초과근무 인정 시간 현황**")
            st.dataframe(summary_ot, use_container_width=True)

# -------------------------------------------------------------------
# TAB 4: 개인별 연차 관리 & 전 직원 연차 요약표
# -------------------------------------------------------------------
with tab4:
    st.header("🌴 개인별 연차 관리 & 전 직원 연차 요약표")
    
    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()

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
                st.dataframe(df_leave_all[['id', 'apply_dt', 'leave_type', 'start_date', 'end_date', 'used_days', 'reason']], use_container_width=True)

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
            df_summary_all.to_excel(writer, index=False, sheet_name="전직원_연차_요약")
            worksheet = writer.sheets["전직원_연차_요약"]

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

        st.dataframe(df_summary_all, use_container_width=True)

# -------------------------------------------------------------------
# TAB 5: 연차 신청서 독립 출력 탭
# -------------------------------------------------------------------
with tab5:
    st.header("🖨️ 휴가 (연차) 신청서 인쇄")
    
    l_records_res = supabase.table("leave_records").select("*").order("id", desc=True).execute()
    df_leave_records = pd.DataFrame(l_records_res.data) if l_records_res.data else pd.DataFrame()

    if df_leave_records.empty:
        st.info("등록된 연차/휴가 신청 내역이 없다.")
    else:
        leave_options = [f"[{r['start_date']}] {r['emp_name']} {r['position']} - {r['leave_type']}" for _, r in df_leave_records.iterrows()]
        selected_l_index = st.selectbox("출력할 연차 신청서 선택", range(len(leave_options)), format_func=lambda x: leave_options[x])
        target_l = df_leave_records.iloc[selected_l_index]

        logo_html = f'<img src="data:image/png;base64,{st.session_state.logo_b64}" style="max-height: 35px; float: left;">' if st.session_state.logo_b64 else ''

        leave_template = f"""
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
with tab6:
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

    ot_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()

    multi_run_ready = True
    try:
        adj_res = supabase.table("monthly_payroll_adjust").select("*").eq("pay_month", pay_month).eq("pay_run_no", pay_run_no).execute()
        df_adjust = pd.DataFrame(adj_res.data) if adj_res.data else pd.DataFrame()
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
        
        for idx, emp in df_emp.iterrows():
            adj_match = df_adjust[df_adjust['emp_id'] == emp['emp_id']] if not df_adjust.empty else pd.DataFrame()

            emp_ot = df_ot[(df_ot['emp_id'] == emp['emp_id']) & (df_ot['work_date'].str.startswith(pay_month))] if not df_ot.empty else pd.DataFrame()
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
                emp_income_tax = adj['income_tax']
                emp_local_tax = adj['local_tax']
                other_deduct = adj['other_deduction']
            else:
                ot_pay = calculated_ot_pay
                base = emp['base_salary']
                family = emp['family_allowance']
                holiday_bonus = 0
                non_tax = emp['non_taxable']
                other_allow = emp['other_allowance']
                other_deduct = emp['other_deduction']

                total_gross_calc = truncate_ten(base + ot_pay + family + non_tax + other_allow)
                taxable_gross_calc = total_gross_calc - non_tax

                # V1.4부터 자동 요율 계산 대신 직원등록 시 입력한 실제 금액을 최초값으로 사용한다.
                emp_national = safe_int(emp.get('national_pension'))
                emp_health = safe_int(emp.get('health_insurance'))
                emp_longterm = safe_int(emp.get('longterm_care'))
                emp_employment = safe_int(emp.get('employment_insurance'))
                emp_income_tax = safe_int(emp.get('income_tax'))
                emp_local_tax = safe_int(emp.get('local_tax'))

                # 2차 대장은 정기급여가 자동 중복되지 않도록 최초 생성 시 금액을 0원으로 시작한다.
                # 필요한 상여금·공제액은 엑셀형 편집기에서 직접 입력한다.
                if pay_run_no == 2:
                    base = ot_pay = family = non_tax = other_allow = other_deduct = 0
                    holiday_bonus = safe_int(emp.get('holiday_bonus'))
                    emp_national = emp_health = emp_longterm = emp_employment = 0
                    emp_income_tax = emp_local_tax = 0

            tot_g = base + ot_pay + family + holiday_bonus + non_tax + other_allow
            taxable_gross = tot_g - non_tax
            emp_income_tax = calculate_income_tax_1_person(taxable_gross)
            emp_local_tax = calculate_local_income_tax(emp_income_tax)

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
        st.info("셀을 클릭하여 직접 수정하거나 엑셀의 여러 셀을 붙여넣을 수 있습니다. 소득세·지방소득세는 부양가족 1명·자녀 0명·100% 기준으로 자동 계산됩니다.")

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
            disabled=identity_columns + ["소득세", "지방소득세"],
            column_config=column_config,
            height=min(650, max(220, 38 * (len(df_calc) + 2)))
        )

        invalid_cells = []
        for col in amount_columns:
            numeric_values = pd.to_numeric(edited_payroll[col], errors="coerce")
            bad_rows = edited_payroll[numeric_values.isna() | (numeric_values < 0)]
            invalid_cells.extend([f"{name} - {col}" for name in bad_rows["이름"].astype(str).tolist()])
            edited_payroll[col] = numeric_values.fillna(0).round().astype(int)

        taxable_series = edited_payroll[["기본급", "초과수당(승인)", "가족수당", "명절상여", "기타수당"]].sum(axis=1)
        edited_payroll["소득세"] = taxable_series.apply(calculate_income_tax_1_person).astype(int)
        edited_payroll["지방소득세"] = edited_payroll["소득세"].apply(calculate_local_income_tax).astype(int)

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
            st.success(f"{pay_month} {pay_run_no}차 {pay_run_name} 수정 수치가 저장되었습니다.")

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
                # 확정 시 편집값도 함께 저장하여 명세서·인쇄 화면과 동일하게 유지한다.
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
            edited_payroll.to_excel(writer, index=False, sheet_name=f"{pay_month}_{pay_run_no}차")
            worksheet = writer.sheets[f"{pay_month}_{pay_run_no}차"]
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
with tab7:
    st.header("📄 개별 급여명세서 인쇄")
    
    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()

    ot_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()

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

        emp_ot = df_ot[(df_ot['emp_id'] == emp['emp_id']) & (df_ot['work_date'].str.startswith(pay_month_slip))] if not df_ot.empty else pd.DataFrame()
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
            emp_income_tax = adj['income_tax']
            emp_local_tax = adj['local_tax']
            other_deduct = adj['other_deduction']
        else:
            ot_pay = calculated_ot_pay
            base = emp['base_salary']
            family = emp['family_allowance']
            holiday_bonus = 0
            non_tax = emp['non_taxable']
            other_allow = emp['other_allowance']
            other_deduct = emp['other_deduction']

            total_gross_tmp = truncate_ten(base + ot_pay + family + non_tax + other_allow)
            taxable_gross_tmp = total_gross_tmp - non_tax

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
        taxable_gross = total_gross - non_tax
        emp_income_tax = calculate_income_tax_1_person(taxable_gross)
        emp_local_tax = calculate_local_income_tax(emp_income_tax)
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
                    <td style="padding: 6px; background: #f9f9f9;">시간외수당 ({total_ot_hours:.1f}h)</td>
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
                    <td>{weekday_ot_hours if weekday_ot_hours > 0 else '-'}</td>
                    <td>{holiday_ot_hours if holiday_ot_hours > 0 else '-'}</td>
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
with tab8:
    st.header("🖨️ 통합 급여대장 인쇄")

    print_col1, print_col2 = st.columns(2)
    with print_col1:
        pay_date_print = st.date_input("출력할 급여 지급일 선택", datetime.now(), key="payroll_print_date")
    with print_col2:
        print_run_no = st.selectbox("출력할 급여대장", [1, 2], format_func=lambda x: f"{x}차 대장", key="print_run_no")
    pay_month_print = pay_date_print.strftime("%Y-%m")

    emp_res = supabase.table("employees").select("*").execute()
    df_emp = pd.DataFrame(emp_res.data) if emp_res.data else pd.DataFrame()

    ot_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot = pd.DataFrame(ot_res.data) if ot_res.data else pd.DataFrame()

    adj_res = supabase.table("monthly_payroll_adjust").select("*").eq("pay_month", pay_month_print).eq("pay_run_no", print_run_no).execute()
    df_adjust = pd.DataFrame(adj_res.data) if adj_res.data else pd.DataFrame()

    if df_emp.empty:
        st.warning("등록된 직원이 없다.")
    else:
        payroll_print_rows = ""
        no = 1

        sum_base = sum_ot = sum_family = sum_holiday = sum_nontax = sum_gross = 0
        sum_nat = sum_hea = sum_long = sum_emp = sum_inc = sum_loc = sum_other_d = sum_deduct_tot = sum_net = 0
        sum_b_nat = sum_b_hea = sum_b_long = sum_b_emp = sum_b_ind = sum_b_tot = sum_retire = 0

        for idx, emp in df_emp.iterrows():
            adj_match = df_adjust[df_adjust['emp_id'] == emp['emp_id']] if not df_adjust.empty else pd.DataFrame()
            emp_ot = df_ot[(df_ot['emp_id'] == emp['emp_id']) & (df_ot['work_date'].str.startswith(pay_month_print))] if not df_ot.empty else pd.DataFrame()
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
                emp_income_tax = adj['income_tax']
                emp_local_tax = adj['local_tax']
                other_deduct = adj['other_deduction']
            else:
                ot_pay = calculated_ot_pay
                base = emp['base_salary']
                family = emp['family_allowance']
                holiday_bonus = 0
                non_tax = emp['non_taxable']
                other_allow = emp['other_allowance']
                other_deduct = emp['other_deduction']

                total_gross_calc = truncate_ten(base + ot_pay + family + non_tax + other_allow)
                taxable_gross_calc = total_gross_calc - non_tax

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
            taxable_gross = tot_g - non_tax
            emp_income_tax = calculate_income_tax_1_person(taxable_gross)
            emp_local_tax = calculate_local_income_tax(emp_income_tax)
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
with tab9:
    st.header("📑 월별 급여대장 총괄표 (12개월 누적 요약)")
    
    c_y9 = st.selectbox("조회 연도 선택", range(datetime.now().year - 2, datetime.now().year + 3), index=2, key="annual_summary_year")
    
    emp_all_res = supabase.table("employees").select("*").execute()
    df_emp_all = pd.DataFrame(emp_all_res.data) if emp_all_res.data else pd.DataFrame()

    ot_all_res = supabase.table("overtime_records").select("*").eq("status", "승인").execute()
    df_ot_all = pd.DataFrame(ot_all_res.data) if ot_all_res.data else pd.DataFrame()

    adj_all_res = supabase.table("monthly_payroll_adjust").select("*").execute()
    df_adj_all = pd.DataFrame(adj_all_res.data) if adj_all_res.data else pd.DataFrame()

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

            m_ot_records = df_ot_all[df_ot_all['work_date'].str.startswith(m_str)] if not df_ot_all.empty else pd.DataFrame()
            if not m_ot_records.empty:
                m_ot_hours = m_ot_records['actual_duration_hours'].sum()

            for _, emp in df_emp_all.iterrows():
                adj_m = df_adj_all[(df_adj_all['pay_month'] == m_str) & (df_adj_all['emp_id'] == emp['emp_id'])] if not df_adj_all.empty else pd.DataFrame()
                
                emp_ot = df_ot_all[(df_ot_all['emp_id'] == emp['emp_id']) & (df_ot_all['work_date'].str.startswith(m_str))] if not df_ot_all.empty else pd.DataFrame()
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
                    inc = loc = 0
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
                        run_taxable = (
                            safe_int(adj.get('base_salary')) + run_ot
                            + safe_int(adj.get('family_allowance'))
                            + safe_int(adj.get('holiday_bonus'))
                            + safe_int(adj.get('other_allowance'))
                        )
                        run_income_tax = calculate_income_tax_1_person(run_taxable)
                        inc += run_income_tax
                        loc += calculate_local_income_tax(run_income_tax)
                else:
                    ot = calc_ot
                    base = emp['base_salary']
                    fam = emp['family_allowance']
                    nontax = emp['non_taxable']
                    other_a = emp['other_allowance']
                    other_d = emp['other_deduction']

                    tot_g_tmp = truncate_ten(base + ot + fam + nontax + other_a)
                    taxable_tmp = tot_g_tmp - nontax

                    nat = safe_int(emp.get('national_pension'))
                    hea = safe_int(emp.get('health_insurance'))
                    lng = safe_int(emp.get('longterm_care'))
                    e_emp = safe_int(emp.get('employment_insurance'))
                    inc = calculate_income_tax_1_person(taxable_tmp)
                    loc = calculate_local_income_tax(inc)
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
