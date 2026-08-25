import streamlit as st
import pandas as pd
from datetime import datetime, time
import io

# 페이지 기본 설정
st.set_page_config(page_title="통합 급여대장 및 초과근무 관리 시스템", layout="wide")

# 원단위 절사 함수 (10원 미만 버림)
def truncate_ten(value):
    return int(value // 10) * 10

# 세션 상태 초기화
if 'employees' not in st.session_state:
    st.session_state.employees = pd.DataFrame(columns=[
        '사번', '이름', '부서', '직위', '호봉', '기본급', '통상시급', '가족수당', '비과세', '기타수당', '기타공제'
    ])

if 'overtime_records' not in st.session_state:
    st.session_state.overtime_records = []

st.title("🏢 통합 급여대장 및 초과근무 관리 시스템")

# 탭 구성
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "👥 직원 등록 및 수정", 
    "📝 초과/휴일근무 신청", 
    "🖨️ 초과근무 신청서", 
    "📊 통합 급여대장 (엑셀 다운로드)", 
    "📄 개별 급여명세서 인쇄"
])

# -------------------------------------------------------------------
# TAB 1: 직원 등록 및 항목별 데이터 관리
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
            non_taxable = st.number_input("비과세 (원)", min_value=0, value=100000, step=10000)
            other_allowance = st.number_input("기타수당 (원)", min_value=0, value=0, step=10000)
            other_deduction = st.number_input("기타공제 (원)", min_value=0, value=0, step=10000)
            
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
                    '가족수당': [family_allowance],
                    '비과세': [non_taxable],
                    '기타수당': [other_allowance],
                    '기타공제': [other_deduction]
                })
                st.session_state.employees = pd.concat([st.session_state.employees, new_emp], ignore_index=True)
                st.success(f"{emp_name} ({position}) 직원의 정보가 등록되었다.")
            else:
                st.error("사번과 이름을 모두 입력해야 한다.")

    st.divider()
    st.header("2. 직원별 급여 항목 수정 (표에서 직접 수정 가능)")
    if not st.session_state.employees.empty:
        edited_df = st.data_editor(st.session_state.employees, use_container_width=True, num_rows="dynamic")
        if st.button("수정사항 대장에 저장"):
            st.session_state.employees = edited_df
            st.success("급여 항목 수정 내용이 저장되었다.")

# -------------------------------------------------------------------
# TAB 2: 초과근무 신청
# -------------------------------------------------------------------
with tab2:
    st.header("초과근무 / 휴일근무 신청")
    
    if st.session_state.employees.empty:
        st.warning("먼저 '직원 등록 및 수정' 탭에서 직원을 등록해야 한다.")
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
            # 원단위 절사 적용
            raw_pay = duration_hours * emp_info['통상시급'] * multiplier
            estimated_pay = truncate_ten(raw_pay)

            st.info(f"💡 인정 근무시간: **{duration_hours:.1f}시간** / 예상 수당: **{estimated_pay:,}원** (원단위 절사 적용)")

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
# TAB 3: 초과근무 신청서 출력
# -------------------------------------------------------------------
with tab3:
    st.header("🖨️ 초과근무 신청서 인쇄")
    if not st.session_state.overtime_records:
        st.info("제출된 초과근무 신청 내역이 없다.")
    else:
        record_options = [f"[{r['근무일자']}] {r['이름']} {r['직위']} - {r['근무구분']}" for r in st.session_state.overtime_records]
        selected_index = st.selectbox("출력할 내역 선택", range(len(record_options)), format_func=lambda x: record_options[x])
        target = st.session_state.overtime_records[selected_index]

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
        st.info("💡 인쇄 방법: 단축키 `Ctrl + P` (맥은 `Cmd + P`)를 누르면 결재란이 포함된 서식을 출력할 수 있다.")

# -------------------------------------------------------------------
# TAB 4: 통합 급여대장 (원단위 절사 및 엑셀 다운로드)
# -------------------------------------------------------------------
with tab4:
    st.header("📊 통합 급여대장")
    
    pay_date = st.date_input("지급일 선택", datetime.now())
    pay_month = pay_date.strftime("%Y-%m")

    if st.session_state.employees.empty:
        st.warning("등록된 직원이 없다.")
    else:
        payroll_data = []
        payroll_html_rows = ""
        no = 1
        
        for idx, emp in st.session_state.employees.iterrows():
            ot_pay = 0
            for r in st.session_state.overtime_records:
                if r['사번'] == emp['사번'] and r['근무일자'].startswith(pay_month):
                    ot_pay += r['수당']

            base = emp['기본급']
            family = emp['가족수당']
            non_tax = emp.get('비과세', 0)
            other_allow = emp.get('기타수당', 0)
            other_deduct = emp.get('기타공제', 0)

            total_gross = truncate_ten(base + ot_pay + family + non_tax + other_allow)
            taxable_gross = total_gross - non_tax  # 과세 대상 금액

            # 근로자 부담금 (원단위 절사)
            emp_national = truncate_ten(taxable_gross * 0.0475)
            emp_health = truncate_ten(taxable_gross * 0.03595)
            emp_longterm = truncate_ten(emp_health * 0.1295)
            emp_employment = truncate_ten(taxable_gross * 0.0090)
            emp_income_tax = truncate_ten(taxable_gross * 0.03)
            emp_local_tax = truncate_ten(emp_income_tax * 0.10)
            
            emp_deduction_total = emp_national + emp_health + emp_longterm + emp_employment + emp_income_tax + emp_local_tax + other_deduct
            net_pay = total_gross - emp_deduction_total

            # 사업자 부담금 (원단위 절사)
            biz_national = truncate_ten(taxable_gross * 0.0475)
            biz_health = truncate_ten(taxable_gross * 0.03595)
            biz_longterm = truncate_ten(biz_health * 0.1295)
            biz_employment = truncate_ten(taxable_gross * 0.0115)
            biz_industrial = truncate_ten(taxable_gross * 0.0726)
            
            biz_deduction_total = biz_national + biz_health + biz_longterm + biz_employment + biz_industrial
            retirement_accrual = truncate_ten(total_gross / 12)

            row_dict = {
                "No": no, "사번": emp['사번'], "이름": emp['이름'], "부서": emp['부서'], "직위": emp['직위'], "호봉": emp['호봉'],
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
                <td>{no}</td><td>{emp['이름']}</td><td>{emp['호봉']}</td>
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

        # 엑셀 파일 생성 다운로드 로직
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
                <div>(단위: 원 / 모든 금액 원단위 절사)</div>
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
# TAB 5: 개별 급여명세서 자동생성 및 인쇄
# -------------------------------------------------------------------
with tab5:
    st.header("📄 개별 급여명세서 자동생성 및 인쇄")
    
    if st.session_state.employees.empty:
        st.warning("등록된 직원이 없다.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            pay_month_slip = st.date_input("명세서 지급 월 선택", datetime.now(), key="slip_month").strftime("%Y-%m")
            emp_slip_list = st.session_state.employees['이름'] + " (" + st.session_state.employees['직위'] + " / " + st.session_state.employees['사번'] + ")"
            selected_slip_str = st.selectbox("직원 선택", emp_slip_list, key="slip_emp")
            selected_slip_id = selected_slip_str.split("/")[-1].replace(")", "").strip()
            emp = st.session_state.employees[st.session_state.employees['사번'] == selected_slip_id].iloc[0]

        # 계산
        ot_pay = 0
        ot_hours = 0
        for r in st.session_state.overtime_records:
            if r['사번'] == emp['사번'] and r['근무일자'].startswith(pay_month_slip):
                ot_pay += r['수당']
                ot_hours += r['인정시간']

        base = emp['기본급']
        family = emp['가족수당']
        non_tax = emp.get('비과세', 0)
        other_allow = emp.get('기타수당', 0)
        other_deduct = emp.get('기타공제', 0)

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
            
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px;">
                <h2 style="margin: 0; padding-top: 15px; font-size: 24px; text-decoration: underline;">{pay_month_slip}월 급 여 명 세 서</h2>
                
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

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 13px;" border="1">
                <tr>
                    <th style="padding: 6px; background: #f2f2f2; width: 15%;">성 명</th>
                    <td style="padding: 6px; width: 35%;">{emp['이름']} ({emp['직위']})</td>
                    <th style="padding: 6px; background: #f2f2f2; width: 15%;">소 속</th>
                    <td style="padding: 6px; width: 35%;">{emp['부서']}</td>
                </tr>
                <tr>
                    <th style="padding: 6px; background: #f2f2f2;">사 번</th>
                    <td style="padding: 6px;">{emp['사번']}</td>
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
                    <td style="padding: 6px; background: #f9f9f9;">시간외수당 ({ot_hours:.1f}h)</td>
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

            <div style="border: 2px solid #333; padding: 12px; text-align: center; background: #fffde7; margin-top: 15px;">
                <span style="font-size: 16px; font-weight: bold;">실지급액 : {net_pay:,} 원</span>
                <span style="font-size: 11px; color: #666; display: block; margin-top: 3px;">(모든 계산 금액 원단위 절사 적용)</span>
            </div>
            
            <p style="text-align: center; margin-top: 30px; font-size: 12px; color: #444;">귀하의 노고에 진심으로 감사드립니다.</p>
        </div>
        """
        st.components.v1.html(payslip_template, height=600, scrolling=True)
        st.info("💡 인쇄 방법: 브라우저 단축키 `Ctrl + P`를 눌러 해당 직원의 급여명세서를 종이로 출력하거나 PDF로 저장할 수 있다.")
