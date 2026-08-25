import streamlit as st
import pandas as pd
from datetime import datetime, time

# 페이지 기본 설정
st.set_page_config(page_title="통합 급여대장 및 초과근무 관리 시스템", layout="wide")

# 세션 상태 초기화
if 'employees' not in st.session_state:
    st.session_state.employees = pd.DataFrame(columns=['사번', '이름', '부서', '직위', '호봉', '기본급', '통상시급', '가족수당'])

if 'overtime_records' not in st.session_state:
    st.session_state.overtime_records = []

st.title("🏢 통합 급여대장 및 초과근무 관리 시스템")

# 탭 구성
tab1, tab2, tab3, tab4 = st.tabs(["👥 직원 등록 및 정보 수정", "📝 초과/휴일근무 신청", "🖨️ 초과근무 신청서 출력", "📊 통합 급여대장"])

# -------------------------------------------------------------------
# TAB 1: 직원 등록 및 수정
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
        with col2:
            base_salary = st.number_input("기본급 (원)", min_value=0, value=2500000, step=100000)
            hourly_wage = st.number_input("통상시급 (원)", min_value=0, value=12000, step=500)
            family_allowance = st.number_input("가족수당 (원)", min_value=0, value=50000, step=10000)
            
        submit_emp = st.form_submit_button("직원 등록")
        
        if submit_emp:
            if emp_id and emp_name:
                new_emp = pd.DataFrame({
                    '사번': [emp_id],
                    '이름': [emp_name],
                    '부서': [dept],
                    '직위': [position],
                    '호봉': [hobong],
                    '기본급': [base_salary],
                    '통상시급': [hourly_wage],
                    '가족수당': [family_allowance]
                })
                st.session_state.employees = pd.concat([st.session_state.employees, new_emp], ignore_index=True)
                st.success(f"{emp_name} ({position}) 직원의 정보가 등록되었다.")
            else:
                st.error("사번과 이름을 모두 입력해야 한다.")

    st.divider()
    st.header("2. 등록된 직원 정보 수정")
    
    if not st.session_state.employees.empty:
        st.subheader("등록된 직원 목록")
        st.dataframe(st.session_state.employees, use_container_width=True)
        
        # 수정할 직원 선택
        emp_edit_list = st.session_state.employees['이름'] + " (" + st.session_state.employees['사번'] + ")"
        selected_edit_str = st.selectbox("수정할 직원을 선택하세요", emp_edit_list)
        selected_edit_id = selected_edit_str.split("(")[1].replace(")", "")
        
        # 선택된 직원의 기존 데이터 추출
        idx = st.session_state.employees[st.session_state.employees['사번'] == selected_edit_id].index[0]
        emp_curr = st.session_state.employees.loc[idx]

        with st.form("edit_employee_form"):
            st.write(f"**[{emp_curr['이름']}] 직원 정보 수정**")
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                edit_dept = st.text_input("부서", value=emp_curr['부서'])
                edit_position = st.text_input("직위", value=emp_curr['직위'])
                edit_hobong = st.text_input("호봉", value=emp_curr['호봉'])
            with col_e2:
                edit_base = st.number_input("기본급 (원)", min_value=0, value=int(emp_curr['기본급']), step=100000)
                edit_hourly = st.number_input("통상시급 (원)", min_value=0, value=int(emp_curr['통상시급']), step=500)
                edit_family = st.number_input("가족수당 (원)", min_value=0, value=int(emp_curr['가족수당']), step=10000)

            submit_edit = st.form_submit_button("수정사항 저장")
            
            if submit_edit:
                st.session_state.employees.at[idx, '부서'] = edit_dept
                st.session_state.employees.at[idx, '직위'] = edit_position
                st.session_state.employees.at[idx, '호봉'] = edit_hobong
                st.session_state.employees.at[idx, '기본급'] = edit_base
                st.session_state.employees.at[idx, '통상시급'] = edit_hourly
                st.session_state.employees.at[idx, '가족수당'] = edit_family
                st.success(f"{emp_curr['이름']} 직원의 정보가 수정되었다.")
                st.rerun()

# -------------------------------------------------------------------
# TAB 2: 초과근무 신청
# -------------------------------------------------------------------
with tab2:
    st.header("초과근무 / 휴일근무 신청")
    
    if st.session_state.employees.empty:
        st.warning("먼저 '직원 등록 및 정보 수정' 탭에서 직원을 등록해야 한다.")
    else:
        emp_list = st.session_state.employees['이름'] + " (" + st.session_state.employees['직위'] + " / " + st.session_state.employees['사번'] + ")"
        selected_emp_str = st.selectbox("직원 선택", emp_list, key="ot_emp")
        selected_emp_id = selected_emp_str.split("/")[-1].replace(")", "").strip()
        emp_info = st.session_state.employees[st.session_state.employees['사번'] == selected_emp_id].iloc[0]

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
            estimated_pay = int(duration_hours * emp_info['통상시급'] * multiplier)

            st.info(f"💡 인정 근무시간: **{duration_hours:.1f}시간** / 예상 수당: **{estimated_pay:,}원**")

            if st.button("신청서 제출"):
                record = {
                    "신청일시": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "사번": emp_info['사번'],
                    "이름": emp_info['이름'],
                    "직위": emp_info['직위'],
                    "부서": emp_info['부서'],
                    "근무일자": str(work_date),
                    "근무구분": work_type,
                    "시작시간": str(start_time),
                    "종료시간": str(end_time),
                    "인정시간": duration_hours,
                    "수당": estimated_pay,
                    "사유": reason
                }
                st.session_state.overtime_records.append(record)
                st.success("초과근무 신청이 저장되었다.")

# -------------------------------------------------------------------
# TAB 3: 초과근무 신청서 출력 (결재란: 담당 / 대리 / 센터장)
# -------------------------------------------------------------------
with tab3:
    st.header("🖨️ 초과근무 신청서 인쇄")
    if not st.session_state.overtime_records:
        st.info("제출된 초과근무 신청 내역이 없다.")
    else:
        record_options = [f"[{r['근무일자']}] {r['이름']} {r['직위']} - {r['근무구분']}" for r in st.session_state.overtime_records]
        selected_index = st.selectbox("출력할 내역 선택", range(len(record_options)), format_func=lambda x: record_options[x])
        target = st.session_state.overtime_records[selected_index]

        # 결재란 (담당, 대리, 센터장) 포함 양식 HTML
        ot_template = f"""
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 680px; margin: auto; background: #fff;">
            
            <!-- 결재란 (우측 상단) -->
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
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                </table>
            </div>

            <table style="width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 14px;" border="1">
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9; width: 20%;">성 명</th>
                    <td style="padding: 8px; width: 30%;">{target['이름']} ({target['직위']})</td>
                    <th style="padding: 8px; background: #f9f9f9; width: 20%;">소 속</th>
                    <td style="padding: 8px; width: 30%;">{target['부서']}</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">근무구분</th>
                    <td style="padding: 8px;" colspan="3">{target['근무구분']}</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">근무일시</th>
                    <td style="padding: 8px;" colspan="3">{target['근무일자']} ({target['시작시간']} ~ {target['종료시간']}) / {target['인정시간']}시간</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background: #f9f9f9;">근무사유</th>
                    <td style="padding: 12px; height: 80px; vertical-align: top;" colspan="3">{target['사유']}</td>
                </tr>
                <tr style="height: 40px;">
                    <th style="padding: 8px; background: #f9f9f9;">예상수당</th>
                    <td style="padding: 8px; font-weight: bold;" colspan="3">{target['수당']:,}원</td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 50px; font-size: 15px;">위와 같이 초과근무를 신청합니다.</p>
            <p style="text-align: center; margin-top: 15px; font-size: 13px;">{target['근무일자'][:4]}년 {target['근무일자'][5:7]}월 {target['근무일자'][8:10]}일</p>
            
            <p style="text-align: right; margin-top: 40px; font-size: 15px; font-weight: bold; padding-right: 10px;">
                신청인: {target['이름']} (인)
            </p>
        </div>
        """
        st.components.v1.html(ot_template, height=520, scrolling=True)
        st.info("💡 인쇄 방법: 단축키 `Ctrl + P` (맥은 `Cmd + P`)를 누르면 결재란이 포함된 서식을 깔끔하게 출력할 수 있다.")

# -------------------------------------------------------------------
# TAB 4: 급여대장
# -------------------------------------------------------------------
with tab4:
    st.header("📊 통합 급여대장")
    
    pay_date = st.date_input("지급일 선택", datetime.now())
    pay_month = pay_date.strftime("%Y-%m")

    if st.session_state.employees.empty:
        st.warning("등록된 직원이 없다.")
    else:
        payroll_rows = ""
        no = 1
        
        for idx, emp in st.session_state.employees.iterrows():
            ot_pay = 0
            for r in st.session_state.overtime_records:
                if r['사번'] == emp['사번'] and r['근무일자'].startswith(pay_month):
                    ot_pay += r['수당']

            base = emp['기본급']
            family = emp['가족수당']
            total_gross = base + ot_pay + family

            emp_national = int(total_gross * 0.0475)
            emp_health = int(total_gross * 0.03595)
            emp_longterm = int(emp_health * 0.1295)
            emp_employment = int(total_gross * 0.0090)
            emp_income_tax = int(total_gross * 0.03)
            emp_local_tax = int(emp_income_tax * 0.10)
            
            emp_deduction_total = emp_national + emp_health + emp_longterm + emp_employment + emp_income_tax + emp_local_tax
            net_pay = total_gross - emp_deduction_total

            biz_national = int(total_gross * 0.0475)
            biz_health = int(total_gross * 0.03595)
            biz_longterm = int(biz_health * 0.1295)
            biz_employment = int(total_gross * 0.0115)
            biz_industrial = int(total_gross * 0.0726)
            
            biz_deduction_total = biz_national + biz_health + biz_longterm + biz_employment + biz_industrial
            retirement_accrual = int(total_gross / 12)

            payroll_rows += f"""
            <tr>
                <td>{no}</td>
                <td>{emp['이름']}</td>
                <td>{emp['호봉']}</td>
                <td style="text-align:right;">{base:,}</td>
                <td style="text-align:right;">{ot_pay:,}</td>
                <td style="text-align:right;">{family:,}</td>
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

        payroll_template = f"""
        <div style="font-family: 'Malgun Gothic', sans-serif; font-size: 11px; width: 100%; overflow-x: auto;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
                <div style="font-size: 14px; font-weight: bold;">지급일 &nbsp;&nbsp;&nbsp; {pay_date}</div>
                <div>(단위: 원)</div>
            </div>
            
            <table border="1" style="width: 100%; border-collapse: collapse; text-align: center;" cellpadding="3">
                <thead>
                    <tr style="background-color: #ffffcc;">
                        <th rowspan="3" style="width: 25px;">No</th>
                        <th rowspan="3" style="width: 50px;">이름</th>
                        <th rowspan="3" style="width: 40px;">호봉</th>
                        <th colspan="3">과세</th>
                        <th rowspan="3">급여총액</th>
                        <th colspan="7">근로자 본인 부담금</th>
                        <th rowspan="3" style="background-color: #fff2cc;">실지급액</th>
                        <th colspan="6">사업자 부담 사회보험금</th>
                        <th rowspan="3">사업주부담<br>퇴직적립금</th>
                    </tr>
                    <tr style="background-color: #ffffcc;">
                        <th rowspan="2">기본급</th>
                        <th rowspan="2">초과근무수당</th>
                        <th rowspan="2">가족수당</th>
                        
                        <th>국민</th>
                        <th>건강</th>
                        <th>장기요양</th>
                        <th>고용</th>
                        <th>소득세</th>
                        <th>지방소득세</th>
                        <th rowspan="2">부담금<br>공제합계</th>
                        
                        <th>국민</th>
                        <th>건강</th>
                        <th>장기요양</th>
                        <th>고용</th>
                        <th>산재</th>
                        <th rowspan="2">사업자<br>부담금<br>공제합계</th>
                    </tr>
                    <tr style="background-color: #ffffcc;">
                        <th>4.75%</th>
                        <th>3.595%</th>
                        <th>12.95%</th>
                        <th>0.90%</th>
                        <th>간이세액표</th>
                        <th>소득세의 10%</th>
                        
                        <th>4.75%</th>
                        <th>3.595%</th>
                        <th>12.95%</th>
                        <th>1.15%</th>
                        <th>7.26%</th>
                    </tr>
                </thead>
                <tbody>
                    {payroll_rows}
                </tbody>
            </table>
        </div>
        """
        st.components.v1.html(payroll_template, height=500, scrolling=True)
