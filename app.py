import streamlit as st
import pandas as pd
from datetime import datetime, time

# 페이지 기본 설정
st.set_page_config(page_title="통합 급여 및 초과근무 관리 시스템", layout="wide")

# 세션 상태 초기화
if 'employees' not in st.session_state:
    st.session_state.employees = pd.DataFrame(columns=['사번', '이름', '부서', '기본급', '통상시급'])

if 'overtime_records' not in st.session_state:
    st.session_state.overtime_records = []

st.title("🏢 통합 급여 및 초과근무 관리 시스템")

# 탭 구성
tab1, tab2, tab3, tab4 = st.tabs(["👥 직원 등록 및 관리", "📝 초과/휴일근무 신청", "🖨️ 초과근무 신청서", "💰 급여명세서 조회·출력"])

# -------------------------------------------------------------------
# TAB 1: 직원 등록
# -------------------------------------------------------------------
with tab1:
    st.header("직원 정보 및 기본 급여 등록")
    with st.form("employee_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            emp_id = st.text_input("사번")
            emp_name = st.text_input("이름")
            dept = st.text_input("부서")
        with col2:
            base_salary = st.number_input("기본급 (원)", min_value=0, value=2500000, step=100000)
            hourly_wage = st.number_input("통상시급 (원)", min_value=0, value=12000, step=500)
            
        submit_emp = st.form_submit_button("직원 등록")
        
        if submit_emp:
            if emp_id and emp_name:
                new_emp = pd.DataFrame({
                    '사번': [emp_id],
                    '이름': [emp_name],
                    '부서': [dept],
                    '기본급': [base_salary],
                    '통상시급': [hourly_wage]
                })
                st.session_state.employees = pd.concat([st.session_state.employees, new_emp], ignore_index=True)
                st.success(f"{emp_name} 직원의 정보가 등록되었다.")
            else:
                st.error("사번과 이름을 모두 입력해야 한다.")

    st.subheader("등록된 직원 목록")
    st.dataframe(st.session_state.employees, use_container_width=True)

# -------------------------------------------------------------------
# TAB 2: 초과근무 신청
# -------------------------------------------------------------------
with tab2:
    st.header("초과근무 / 휴일근무 신청")
    
    if st.session_state.employees.empty:
        st.warning("먼저 '직원 등록 및 관리' 탭에서 직원을 등록해야 한다.")
    else:
        emp_list = st.session_state.employees['이름'] + " (" + st.session_state.employees['사번'] + ")"
        selected_emp_str = st.selectbox("직원 선택", emp_list, key="ot_emp")
        selected_emp_id = selected_emp_str.split("(")[1].replace(")", "")
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
# TAB 3: 초과근무 신청서 출력
# -------------------------------------------------------------------
with tab3:
    st.header("초과근무 신청서 인쇄")
    if not st.session_state.overtime_records:
        st.info("제출된 초과근무 신청 내역이 없다.")
    else:
        record_options = [f"[{r['근무일자']}] {r['이름']} - {r['근무구분']}" for r in st.session_state.overtime_records]
        selected_index = st.selectbox("출력할 내역 선택", range(len(record_options)), format_func=lambda x: record_options[x])
        target = st.session_state.overtime_records[selected_index]

        ot_template = f"""
        <div style="border: 2px solid #000; padding: 25px; font-family: 'Malgun Gothic'; max-width: 650px; margin: auto;">
            <h2 style="text-align: center; text-decoration: underline;">초 과 근 무 신 청 서</h2>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px;" border="1">
                <tr><th style="padding: 8px; background: #f2f2f2; width: 20%;">성 명</th><td style="padding: 8px;">{target['이름']}</td><th style="padding: 8px; background: #f2f2f2; width: 20%;">소 속</th><td style="padding: 8px;">{target['부서']}</td></tr>
                <tr><th style="padding: 8px; background: #f2f2f2;">근무구분</th><td style="padding: 8px;" colspan="3">{target['근무구분']}</td></tr>
                <tr><th style="padding: 8px; background: #f2f2f2;">근무일시</th><td style="padding: 8px;" colspan="3">{target['근무일자']} ({target['시작시간']} ~ {target['종료시간']}) / {target['인정시간']}시간</td></tr>
                <tr><th style="padding: 8px; background: #f2f2f2;">근무사유</th><td style="padding: 8px; height: 60px;" colspan="3">{target['사유']}</td></tr>
                <tr><th style="padding: 8px; background: #f2f2f2;">예상수당</th><td style="padding: 8px;" colspan="3">{target['수당']:,}원</td></tr>
            </table>
            <p style="text-align: center; margin-top: 30px;">위와 같이 초과근무를 신청합니다.</p>
            <p style="text-align: right; margin-top: 20px;">신청인: {target['이름']} (인)</p>
        </div>
        """
        st.components.v1.html(ot_template, height=450, scrolling=True)

# -------------------------------------------------------------------
# TAB 4: 급여명세서 생성 및 출력
# -------------------------------------------------------------------
with tab4:
    st.header("월별 급여명세서 생성 및 인쇄")
    
    if st.session_state.employees.empty:
        st.warning("등록된 직원이 없다.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            pay_month = st.date_input("급여 지급 월", datetime.now()).strftime("%Y-%m")
            emp_pay_list = st.session_state.employees['이름'] + " (" + st.session_state.employees['사번'] + ")"
            selected_pay_emp = st.selectbox("직원 선택", emp_pay_list, key="pay_emp")
            selected_pay_id = selected_pay_emp.split("(")[1].replace(")", "")
            emp = st.session_state.employees[st.session_state.employees['사번'] == selected_pay_id].iloc[0]

        # 누적 초과근무 수당 자동 합산
        total_ot_pay = 0
        total_ot_hours = 0
        for r in st.session_state.overtime_records:
            if r['사번'] == emp['사번'] and r['근무일자'].startswith(pay_month):
                total_ot_pay += r['수당']
                total_ot_hours += r['인정시간']

        with col2:
            st.metric(label=f"{pay_month} 누적 초과근무 시간", value=f"{total_ot_hours:.1f} 시간")
            st.metric(label=f"{pay_month} 누적 초과근무 수당", value=f"{total_ot_pay:,} 원")

        # 급여 항목 산정
        base_pay = emp['기본급']
        gross_pay = base_pay + total_ot_pay # 총 지급액

        # 4대보험 및 공제액 산출 (4대보험 법정 요율 기준 대략 계산)
        national_pension = int(gross_pay * 0.045)  # 국민연금 4.5%
        health_insurance = int(gross_pay * 0.03545) # 건강보험 3.545%
        longterm_care = int(health_insurance * 0.1295) # 장기요양보험
        employment_insurance = int(gross_pay * 0.009) # 고용보험 0.9%
        income_tax = int(gross_pay * 0.03) # 근로소득세 (간이세액 추정 3%)
        local_tax = int(income_tax * 0.1) # 지방소득세 10%

        total_deduction = national_pension + health_insurance + longterm_care + employment_insurance + income_tax + local_tax
        net_pay = gross_pay - total_deduction # 실수령액

        st.divider()
        st.subheader("📄 급여명세서 미리보기")

        payslip_template = f"""
        <div style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic'; max-width: 700px; margin: auto;">
            <h2 style="text-align: center; text-decoration: underline; margin-bottom: 20px;">{pay_month}월 급여명세서</h2>
            
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;" border="1">
                <tr>
                    <th style="padding: 6px; background: #f2f2f2; width: 15%;">성명</th>
                    <td style="padding: 6px; width: 35%;">{emp['이름']}</td>
                    <th style="padding: 6px; background: #f2f2f2; width: 15%;">부서</th>
                    <td style="padding: 6px; width: 35%;">{emp['부서']}</td>
                </tr>
                <tr>
                    <th style="padding: 6px; background: #f2f2f2;">사번</th>
                    <td style="padding: 6px;">{emp['사번']}</td>
                    <th style="padding: 6px; background: #f2f2f2;">지급일</th>
                    <td style="padding: 6px;">{pay_month}-25</td>
                </tr>
            </table>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;" border="1">
                <tr style="background: #e6f2ff;">
                    <th style="padding: 8px; width: 50%;" colspan="2">지급 내역</th>
                    <th style="padding: 8px; width: 50%;" colspan="2">공제 내역</th>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">기본급</td>
                    <td style="padding: 6px; text-align: right;">{base_pay:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">국민연금</td>
                    <td style="padding: 6px; text-align: right;">{national_pension:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">시간외수당 ({total_ot_hours:.1f}h)</td>
                    <td style="padding: 6px; text-align: right;">{total_ot_pay:,} 원</td>
                    <td style="padding: 6px; background: #f9f9f9;">건강보험</td>
                    <td style="padding: 6px; text-align: right;">{health_insurance:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">-</td>
                    <td style="padding: 6px; text-align: right;">-</td>
                    <td style="padding: 6px; background: #f9f9f9;">장기요양보험</td>
                    <td style="padding: 6px; text-align: right;">{longterm_care:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">-</td>
                    <td style="padding: 6px; text-align: right;">-</td>
                    <td style="padding: 6px; background: #f9f9f9;">고용보험</td>
                    <td style="padding: 6px; text-align: right;">{employment_insurance:,} 원</td>
                </tr>
                <tr>
                    <td style="padding: 6px; background: #f9f9f9;">-</td>
                    <td style="padding: 6px; text-align: right;">-</td>
                    <td style="padding: 6px; background: #f9f9f9;">소득세 / 지방세</td>
                    <td style="padding: 6px; text-align: right;">{(income_tax + local_tax):,} 원</td>
                </tr>
                <tr style="font-weight: bold; background: #f2f2f2;">
                    <td style="padding: 8px;">지급액 계</td>
                    <td style="padding: 8px; text-align: right;">{gross_pay:,} 원</td>
                    <td style="padding: 8px;">공제액 계</td>
                    <td style="padding: 8px; text-align: right;">{total_deduction:,} 원</td>
                </tr>
            </table>

            <div style="border: 2px solid #333; padding: 12px; text-align: center; background: #fffde7; margin-top: 15px;">
                <span style="font-size: 16px; font-weight: bold;">실수령액 : {net_pay:,} 원</span>
            </div>
            
            <p style="text-align: center; margin-top: 25px; font-size: 12px; color: #666;">귀하의 노고에 감사드립니다.</p>
        </div>
        """
        st.components.v1.html(payslip_template, height=580, scrolling=True)
        st.info("💡 인쇄 방법: 브라우저 단축키 `Ctrl + P` (맥은 `Cmd + P`)를 눌러 급여명세서 서식을 바로 출력하거나 PDF로 저장할 수 있다.")
