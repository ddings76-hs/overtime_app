import streamlit as st
import pandas as pd
from datetime import datetime, time
import sqlite3
import io

# 페이지 기본 설정
st.set_page_config(page_title="통합 급여·초과근무·휴가 관리 시스템", layout="wide")

# -------------------------------------------------------------------
# DB 연결 및 테이블 생성 (데이터 영구 보존)
# -------------------------------------------------------------------
DB_FILE = "office_management.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    # 직원 테이블
    c.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            emp_id TEXT PRIMARY KEY,
            emp_name TEXT,
            dept TEXT,
            position TEXT,
            hobong TEXT,
            base_salary INTEGER,
            hourly_wage INTEGER,
            family_allowance INTEGER,
            non_taxable INTEGER,
            other_allowance INTEGER,
            other_deduction INTEGER,
            total_annual_leave REAL DEFAULT 15.0
        )
    ''')
    
    # 초과근무 테이블
    c.execute('''
        CREATE TABLE IF NOT EXISTS overtime_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            apply_dt TEXT,
            emp_id TEXT,
            emp_name TEXT,
            dept TEXT,
            position TEXT,
            work_date TEXT,
            work_type TEXT,
            start_time TEXT,
            end_time TEXT,
            duration_hours REAL,
            estimated_pay INTEGER,
            reason TEXT
        )
    ''')

    # 휴가 관리 테이블
    c.execute('''
        CREATE TABLE IF NOT EXISTS leave_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            apply_dt TEXT,
            emp_id TEXT,
            emp_name TEXT,
            dept TEXT,
            position TEXT,
            leave_type TEXT,
            start_date TEXT,
            end_date TEXT,
            used_days REAL,
            reason TEXT
        )
    ''')
    
    conn.commit()
    conn.close()

init_db()

# DB 헬퍼 함수
def get_db_connection():
    return sqlite3.connect(DB_FILE)

def truncate_ten(value):
    return int(value // 10) * 10

st.title("🏢 통합 급여·초과근무·휴가 관리 시스템")

# 탭 구성
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "👥 직원 등록 및 정보 관리", 
    "📝 초과/휴일근무 신청", 
    "🖨️ 초과근무 신청서", 
    "🌴 개인별 휴가 관리",
    "📊 통합 급여대장 (엑셀)", 
    "📄 개별 급여명세서 인쇄"
])

# -------------------------------------------------------------------
# TAB 1: 직원 등록 및 누적 관리
# -------------------------------------------------------------------
with tab1:
    st.header("1. 신규 직원 등록")
    with st.form("employee_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            emp_id = st.text_input("사번")
            emp_name = st.text_input("이름")
            dept = st.text_input("부서")
            position = st.text_input("직위", value="주임")
            hobong = st.text_input("호봉", value="1호봉")
            total_leave = st.number_input("연간 총 연차 부여일수", min_value=0.0, value=15.0, step=0.5)
        with col2:
            base_salary = st.number_input("기본급 (원)", min_value=0, value=2500000, step=100000)
            hourly_wage = st.number_input("통상시급 (원)", min_value=0, value=12000, step=500)
            family_allowance = st.number_input("가족수당 (원)", min_value=0, value=50000, step=10000)
            non_taxable = st.number_input("비과세 (원)", min_value=0, value=100000, step=10000)
            other_allowance = st.number_input("기타수당 (원)", min_value=0, value=0, step=10000)
            other_deduction = st.number_input("기타공제 (원)", min_value=0, value=0, step=10000)
            
        submit_emp = st.form_submit_button("직원 DB 등록")
        
        if submit_emp:
            if emp_id and emp_name:
                conn = get_db_connection()
                c = conn.cursor()
                try:
                    c.execute('''
                        INSERT INTO employees VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (emp_id, emp_name, dept, position, hobong, base_salary, hourly_wage, family_allowance, non_taxable, other_allowance, other_deduction, total_leave))
                    conn.commit()
                    st.success(f"{emp_name} ({position}) 직원이 DB에 정상 등록되었다.")
                except sqlite3.IntegrityError:
                    st.error("이미 존재하는 사번이다.")
                finally:
                    conn.close()
            else:
                st.error("사번과 이름을 모두 입력해야 한다.")

    st.divider()
    st.header("2. 누적 직원 데이터 조회 및 편집")
    conn = get_db_connection()
    df_emp = pd.read_sql_query("SELECT * FROM employees", conn)
    conn.close()

    if not df_emp.empty:
        edited_df = st.data_editor(df_emp, use_container_width=True, num_rows="dynamic")
        if st.button("수정 데이터 DB 저장"):
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("DELETE FROM employees")
            for _, row in edited_df.iterrows():
                c.execute('''
                    INSERT INTO employees VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', tuple(row))
            conn.commit()
            conn.close()
            st.success("직원 데이터 수정사항이 DB에 반영되었다.")
            st.rerun()

# -------------------------------------------------------------------
# TAB 2: 초과근무 신청
# -------------------------------------------------------------------
with tab2:
    st.header("초과근무 / 휴일근무 신청")
    
    conn = get_db_connection()
    df_emp = pd.read_sql_query("SELECT * FROM employees", conn)
    conn.close()

    if df_emp.empty:
        st.warning("먼저 '직원 등록 및 정보 관리' 탭에서 직원을 등록해야 한다.")
    else:
        emp_list = df_emp['emp_name'] + " (" + df_emp['position'] + " / " + df_emp['emp_id'] + ")"
        selected_emp_str = st.selectbox("직원 선택", emp_list, key="ot_emp")
        selected_emp_id = selected_emp_str.split("/")[-1].replace(")", "").strip()
        emp_info = df_emp[df_emp['emp_id'] == selected_emp_id].iloc[0]

        work_date = st.date_input("근무 일자", datetime.now())
        work_type = st.radio("근무 구분", ["평일 초과근무 (18:00 이후)", "휴일근무"])
        
        col1, col2 = st.columns(2)
        with col1:
            start_time = st.time_input("시작 시간", time(18, 0) if work_type == "평일 초과근무 (18:00 이후)" else time(9, 0))
        with col2:
            end_time = st.time_input("종료 시간", time(20, 0))

        reason = st.text_area("근무 사유")

        start_dt = datetime.combine(work_date, start_time)
        end_dt = datetime.combine(work_date, end_time)
        duration_hours = (end_dt - start_dt).total_seconds() / 3600

        if duration_hours < 0:
            st.error("종료 시간은 시작 시간보다 빨라야 한다.")
        else:
            multiplier = 1.5
            raw_pay = duration_hours * emp_info['hourly_wage'] * multiplier
            estimated_pay = truncate_ten(raw_pay)

            st.info(f"💡 인정 근무시간: **{duration_hours:.1f}시간** / 예상 수당: **{estimated_pay:,}원** (원단위 절사)")

            if st.button("신청서 제출 및 DB 저장"):
                conn = get_db_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO overtime_records (apply_dt, emp_id, emp_name, dept, position, work_date, work_type, start_time, end_time, duration_hours, estimated_pay, reason)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    emp_info['emp_id'], emp_info['emp_name'], emp_info['dept'], emp_info['position'],
                    str(work_date), work_type, str(start_time), str(end_time), duration_hours, estimated_pay, reason
                ))
                conn.commit()
                conn.close()
                st.success("초과근무 신청 내역이 DB에 등록되었다.")

# -------------------------------------------------------------------
# TAB 3: 초과근무 신청서 출력
# -------------------------------------------------------------------
with tab3:
    st.header("🖨️ 초과근무 신청서 인쇄")
    conn = get_db_connection()
    df_ot = pd.read_sql_query("SELECT * FROM overtime_records ORDER BY id DESC", conn)
    conn.close()

    if df_ot.empty:
        st.info("제출된 초과근무 신청 내역이 없다.")
    else:
        record_options = [f"[{r['work_date']}] {r['emp_name']} {r['position']} - {r['work_type']}" for _, r in df_ot.iterrows()]
        selected_index = st.selectbox("출력할 내역 선택", range(len(record_options)), format_func=lambda x: record_options[x])
        target = df_ot.iloc[selected_index]

        ot_template = f"""
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 680px; margin: auto; background: #fff;">
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px;">
                <h2 style="margin: 0; padding-top: 15px; font-size: 24px; text-decoration: underline;">초 과 근 무 신 청 서</h2>
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
                    <td style="padding: 8px; width: 30%;">{target['emp_name']} ({target['position']})</td>
                    <th style="padding: 8px; background: #f9f9f9; width: 20%;">소 속</th>
                    <td style="padding: 8px; width: 30%;">{target['dept']}</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">근무구분</th>
                    <td style="padding: 8px;" colspan="3">{target['work_type']}</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">근무일시</th>
                    <td style="padding: 8px;" colspan="3">{target['work_date']} ({target['start_time']} ~ {target['end_time']}) / {target['duration_hours']}시간</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background: #f9f9f9;">근무사유</th>
                    <td style="padding: 12px; height: 80px; vertical-align: top;" colspan="3">{target['reason']}</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">예상수당</th>
                    <td style="padding: 8px; font-weight: bold;" colspan="3">{target['estimated_pay']:,}원</td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 50px; font-size: 15px;">위와 같이 초과근무를 신청합니다.</p>
            <p style="text-align: center; margin-top: 15px; font-size: 13px;">{target['work_date'][:4]}년 {target['work_date'][5:7]}월 {target['work_date'][8:10]}일</p>
            
            <p style="text-align: right; margin-top: 40px; font-size: 15px; font-weight: bold; padding-right: 10px;">
                신청인: {target['emp_name']} (인)
            </p>
        </div>
        """
        st.components.v1.html(ot_template, height=520, scrolling=True)

# -------------------------------------------------------------------
# TAB 4: 개인별 휴가(연차) 관리 (신규 고도화 기능)
# -------------------------------------------------------------------
with tab4:
    st.header("🌴 개인별 휴가 관리 및 현황")
    
    conn = get_db_connection()
    df_emp = pd.read_sql_query("SELECT * FROM employees", conn)
    conn.close()

    if df_emp.empty:
        st.warning("등록된 직원이 없다.")
    else:
        col_l1, col_l2 = st.columns([1, 2])
        
        with col_l1:
            st.subheader("1. 휴가 신청")
            emp_leave_list = df_emp['emp_name'] + " (" + df_emp['position'] + " / " + df_emp['emp_id'] + ")"
            selected_l_emp = st.selectbox("직원 선택", emp_leave_list, key="leave_emp_select")
            selected_l_id = selected_l_emp.split("/")[-1].replace(")", "").strip()
            l_emp_info = df_emp[df_emp['emp_id'] == selected_l_id].iloc[0]

            leave_type = st.selectbox("휴가 종류", ["연차 (1일)", "오전반차 (0.5일)", "오후반차 (0.5일)", "병가", "경조휴가", "특별휴가"])
            
            l_start_date = st.date_input("휴가 시작일", datetime.now())
            l_end_date = st.date_input("휴가 종료일", datetime.now())
            
            if "반차" in leave_type:
                used_days = 0.5
            else:
                used_days = float((l_end_date - l_start_date).days + 1)

            leave_reason = st.text_area("휴가 사유")

            if st.button("휴가 신청서 제출"):
                conn = get_db_connection()
                c = conn.cursor()
                c.execute('''
                    INSERT INTO leave_records (apply_dt, emp_id, emp_name, dept, position, leave_type, start_date, end_date, used_days, reason)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    l_emp_info['emp_id'], l_emp_info['emp_name'], l_emp_info['dept'], l_emp_info['position'],
                    leave_type, str(l_start_date), str(l_end_date), used_days, leave_reason
                ))
                conn.commit()
                conn.close()
                st.success("휴가 신청 내역이 DB에 저구되었다.")
                st.rerun()

        with col_l2:
            st.subheader("2. 개인별 연차 사용 현황")
            conn = get_db_connection()
            df_leave_all = pd.read_sql_query("SELECT * FROM leave_records WHERE emp_id = ?", conn, params=(l_emp_info['emp_id'],))
            conn.close()

            # 연차 차감 대상 집계 (연차, 반차 항목)
            used_annual = df_leave_all[df_leave_all['leave_type'].str.contains("연차|반차", na=False)]['used_days'].sum()
            total_annual = l_emp_info['total_annual_leave']
            remaining_annual = total_annual - used_annual

            m1, m2, m3 = st.columns(3)
            m1.metric("총 부여 연차", f"{total_annual} 일")
            m2.metric("사용 연차", f"{used_annual} 일")
            m3.metric("잔여 연차", f"{remaining_annual} 일")

            st.write(f"**[{l_emp_info['emp_name']}] 휴가 신청 이력**")
            st.dataframe(df_leave_all[['apply_dt', 'leave_type', 'start_date', 'end_date', 'used_days', 'reason']], use_container_width=True)

# -------------------------------------------------------------------
# TAB 5: 통합 급여대장 (DB 연동 및 엑셀 다운로드)
# -------------------------------------------------------------------
with tab5:
    st.header("📊 통합 급여대장")
    
    pay_date = st.date_input("지급일 선택", datetime.now())
    pay_month = pay_date.strftime("%Y-%m")

    conn = get_db_connection()
    df_emp = pd.read_sql_query("SELECT * FROM employees", conn)
    df_ot = pd.read_sql_query("SELECT * FROM overtime_records", conn)
    conn.close()

    if df_emp.empty:
        st.warning("등록된 직원이 없다.")
    else:
        payroll_data = []
        payroll_html_rows = ""
        no = 1
        
        for idx, emp in df_emp.iterrows():
            # DB 기반 해당월 초과근무 수당 합산
            emp_ot = df_ot[(df_ot['emp_id'] == emp['emp_id']) & (df_ot['work_date'].str.startswith(pay_month))]
            ot_pay = emp_ot['estimated_pay'].sum() if not emp_ot.empty else 0

            base = emp['base_salary']
            family = emp['family_allowance']
            non_tax = emp['non_taxable']
            other_allow = emp['other_allowance']
            other_deduct = emp['other_deduction']

            total_gross = truncate_ten(base + ot_pay + family + non_tax + other_allow)
            taxable_gross = total_gross - non_tax

            emp_national = truncate_ten(taxable_gross * 0.0475)
            emp_health = truncate_ten(taxable_gross * 0.03595)
            emp_longterm = truncate_ten(emp_health * 0.1295)
            emp_employment = truncate_ten(taxable_gross * 0.0090)
            emp_income_tax = truncate_ten(taxable_gross * 0.03)
            emp_local_tax = truncate_ten(emp_income_tax * 0.10)
            
            emp_deduction_total = emp_national + emp_health + emp_longterm + emp_employment + emp_income_tax + emp_local_tax + other_deduct
            net_pay = total_gross - emp_deduction_total

            biz_national = truncate_ten(taxable_gross * 0.0475)
            biz_health = truncate_ten(taxable_gross * 0.03595)
            biz_longterm = truncate_ten(biz_health * 0.1295)
            biz_employment = truncate_ten(taxable_gross * 0.0115)
            biz_industrial = truncate_ten(taxable_gross * 0.0726)
            
            biz_deduction_total = biz_national + biz_health + biz_longterm + biz_employment + biz_industrial
            retirement_accrual = truncate_ten(total_gross / 12)

            row_dict = {
                "No": no, "사번": emp['emp_id'], "이름": emp['emp_name'], "부서": emp['dept'], "직위": emp['position'], "호봉": emp['hobong'],
                "기본급": base, "초과근무수당": ot_pay, "가족수당": family, "비과세": non_tax, "기타수당": other_allow,
                "급여총액": total_gross,
                "국민연금(본인)": emp_national, "건강보험(본인)": emp_health, "장기요양(본인)": emp_longterm,
                "고용보험(본인)": emp_employment, "소득세": emp_income_tax, "지방소득세": emp_local_tax,
                "기타공제": other_deduct, "공제합계": emp_deduction_total, "실지급액": net_pay,
                "국민연금(사업자)": biz_national, "건강보험(사업자)": biz_health, "장기요양(사업자)": biz_longterm,
                "고용보험(사업자)": biz_employment, "산재보험(사업자)": biz_industrial, "사업자공제합계": biz_deduction_total,
                "퇴직적립금": retirement_accrual
            }
            payroll_data.append(row_dict)

            payroll_html_rows += f"""
            <tr>
                <td>{no}</td><td>{emp['emp_name']}</td><td>{emp['hobong']}</td>
                <td style="text-align:right;">{base:,}</td>
                <td style="text-align:right;">{ot_pay:,}</td>
                <td style="text-align:right;">{family:,}</td>
                <td style="text-align:right;">{non_tax:,}</td>
                <td style="text-align:right; font-weight:bold;">{total_gross:,}</td>
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

        df_payroll = pd.DataFrame(payroll_data)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_payroll.to_excel(writer, index=False, sheet_name=f"{pay_month}_급여대장")
        excel_data = output.getvalue()

        st.download_button(
            label="📥 통합 급여대장 엑셀 다운로드 (.xlsx)",
            data=excel_data,
            file_name=f"통합급여대장_{pay_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        payroll_template = f"""
        <div style="font-family: 'Malgun Gothic', sans-serif; font-size: 11px; width: 100%; overflow-x: auto;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
                <div style="font-size: 14px; font-weight: bold;">지급일 &nbsp;&nbsp;&nbsp; {pay_date}</div>
                <div>(단위: 원 / 원단위 절사 적용)</div>
            </div>
            
            <table border="1" style="width: 100%; border-collapse: collapse; text-align: center;" cellpadding="3">
                <thead>
                    <tr style="background-color: #ffffcc;">
                        <th rowspan="3" style="width: 25px;">No</th>
                        <th rowspan="3" style="width: 50px;">이름</th>
                        <th rowspan="3" style="width: 40px;">호봉</th>
                        <th colspan="4">지급 내역</th>
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
                        <th rowspan="2">비과세</th>
                        
                        <th>국민</th>
                        <th>건강</th>
                        <th>장기요양</th>
                        <th>고용</th>
                        <th>소득세</th>
                        <th>지방세</th>
                        <th rowspan="2">공제합계</th>
                        
                        <th>국민</th>
                        <th>건강</th>
                        <th>장기요양</th>
                        <th>고용</th>
                        <th>산재</th>
                        <th rowspan="2">사업자합계</th>
                    </tr>
                    <tr style="background-color: #ffffcc;">
                        <th>4.75%</th>
                        <th>3.595%</th>
                        <th>12.95%</th>
                        <th>0.90%</th>
                        <th>간이세액</th>
                        <th>10%</th>
                        
                        <th>4.75%</th>
                        <th>3.595%</th>
                        <th>12.95%</th>
                        <th>1.15%</th>
                        <th>7.26%</th>
                    </tr>
                </thead>
                <tbody>
                    {payroll_html_rows}
                </tbody>
            </table>
        </div>
        """
        st.components.v1.html(payroll_template, height=500, scrolling=True)

# -------------------------------------------------------------------
# TAB 6: 개별 급여명세서 인쇄
# -------------------------------------------------------------------
with tab6:
    st.header("📄 개별 급여명세서 인쇄")
    
    conn = get_db_connection()
    df_emp = pd.read_sql_query("SELECT * FROM employees", conn)
    df_ot = pd.read_sql_query("SELECT * FROM overtime_records", conn)
    conn.close()

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

        emp_ot = df_ot[(df_ot['emp_id'] == emp['emp_id']) & (df_ot['work_date'].str.startswith(pay_month_slip))]
        
        weekday_ot_hours = emp_ot[emp_ot['work_type'].str.contains("평일", na=False)]['duration_hours'].sum() if not emp_ot.empty else 0.0
        holiday_ot_hours = emp_ot[emp_ot['work_type'].str.contains("휴일", na=False)]['duration_hours'].sum() if not emp_ot.empty else 0.0
        ot_pay = emp_ot['estimated_pay'].sum() if not emp_ot.empty else 0

        total_ot_hours = weekday_ot_hours + holiday_ot_hours

        base = emp['base_salary']
        family = emp['family_allowance']
        non_tax = emp['non_taxable']
        other_allow = emp['other_allowance']
        other_deduct = emp['other_deduction']

        total_gross = truncate_ten(base + ot_pay + family + non_tax + other_allow)
        taxable_gross = total_gross - non_tax

        emp_national = truncate_ten(taxable_gross * 0.0475)
        emp_health = truncate_ten(taxable_gross * 0.03595)
        emp_longterm = truncate_ten(emp_health * 0.1295)
        emp_employment = truncate_ten(taxable_gross * 0.0090)
        emp_income_tax = truncate_ten(taxable_gross * 0.03)
        emp_local_tax = truncate_ten(emp_income_tax * 0.10)

        emp_deduction_total = emp_national + emp_health + emp_longterm + emp_employment + emp_income_tax + emp_local_tax + other_deduct
        net_pay = total_gross - emp_deduction_total

        payslip_template = f"""
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 680px; margin: auto; background: #fff;">
            
            <h2 style="text-align: center; margin-top: 0; margin-bottom: 25px; font-size: 24px; text-decoration: underline;">
                {pay_month_slip}월 급 여 명 세 서
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
                    <th style="padding: 6px; background: #f2f2f2;">지급일</th>
                    <td style="padding: 6px;">{pay_month_slip}-25</td>
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
                    <td style="padding: 6px; background: #f9f9f9;">국민연금 (4.75%)</td>
                    <td style="padding: 6px; text-align: right;">{emp_national:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">시간외수당 ({total_ot_hours:.1f}h)</td>
                    <td style="padding: 6px; text-align: right;">{ot_pay:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">건강보험 (3.595%)</td>
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
                    <td style="padding: 6px; background: #f9f9f9;">고용보험 (0.9%)</td>
                    <td style="padding: 6px; text-align: right;">{emp_employment:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">기타수당</td>
                    <td style="padding: 6px; text-align: right;">{other_allow:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">소득세 / 지방소득세</td>
                    <td style="padding: 6px; text-align: right;">{(emp_income_tax + emp_local_tax):,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">-</td>
                    <td style="padding: 6px; text-align: right;">-</td>
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
                    <td>{int(emp['hourly_wage']):,}</td>
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
                    <td style="padding-left: 10px;">연장근로시간수 × 통상시급 × 1.5</td>
                </tr>
                <tr style="height: 24px;">
                    <td style="text-align: center;">휴일근로수당</td>
                    <td style="padding-left: 10px;">휴일근로시간수 × 통상시급 × 1.5</td>
                </tr>
                <tr style="height: 24px;">
                    <td style="text-align: center;">야간근로수당</td>
                    <td style="padding-left: 10px;">야간근로시간수 × 통상시급 × 1.5</td>
                </tr>
            </table>
            <p style="font-size: 10px; color: #333; margin-top: 5px; margin-bottom: 15px;">
                * 통상시급 : 정기적이고 일률적으로 지급하는 급여 ÷ 소정근로시간(209시간)
            </p>

            <p style="text-align: center; margin-top: 20px; font-size: 12px; color: #444;">귀하의 노고에 진심으로 감사드립니다.</p>
        </div>
        """
        st.components.v1.html(payslip_template, height=850, scrolling=True)
