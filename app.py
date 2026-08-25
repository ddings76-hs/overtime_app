import streamlit as st
import pandas as pd
from datetime import datetime, time

# 페이지 기본 설정
st.set_page_config(page_title="사무실 초과근무 관리 시스템", layout="wide")

# 세션 상태 초기화 (데이터 저장용)
if 'employees' not in st.session_state:
    st.session_state.employees = pd.DataFrame(columns=['사번', '이름', '부서', '통상시급'])

if 'overtime_records' not in st.session_state:
    st.session_state.overtime_records = []

st.title("🏢 사무실 초과근무 및 휴일근무 관리 시스템")

# 탭 구성
tab1, tab2, tab3 = st.tabs(["👥 직원 등록 및 관리", "📝 초과/휴일근무 신청", "🖨️ 신청서 조회 및 출력"])

# -------------------------------------------------------------------
# TAB 1: 직원 등록
# -------------------------------------------------------------------
with tab1:
    st.header("직원 정보 등록")
    with st.form("employee_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            emp_id = st.text_input("사번")
            emp_name = st.text_input("이름")
        with col2:
            dept = st.text_input("부서")
            hourly_wage = st.number_input("통상시급 (원)", min_value=0, value=10000, step=500)
            
        submit_emp = st.form_submit_button("직원 등록")
        
        if submit_emp:
            if emp_id and emp_name:
                new_emp = pd.DataFrame({
                    '사번': [emp_id],
                    '이름': [emp_name],
                    '부서': [dept],
                    '통상시급': [hourly_wage]
                })
                st.session_state.employees = pd.concat([st.session_state.employees, new_emp], ignore_index=True)
                st.success(f"{emp_name} 직원이 등록되었다.")
            else:
                st.error("사번과 이름을 모두 입력해야 한다.")

    st.subheader("등록된 직원 목록")
    st.dataframe(st.session_state.employees, use_container_width=True)

# -------------------------------------------------------------------
# TAB 2: 초과근무 신청
# -------------------------------------------------------------------
with tab2:
    st.header("초과근무 / 휴일근무 신청서 작성")
    
    if st.session_state.employees.empty:
        st.warning("먼저 '직원 등록 및 관리' 탭에서 직원을 등록해야 한다.")
    else:
        emp_list = st.session_state.employees['이름'] + " (" + st.session_state.employees['사번'] + ")"
        selected_emp_str = st.selectbox("직원 선택", emp_list)
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

        # 수당 자동 계산 공식
        # 시작/종료 시간차 계산 (시간 단위)
        start_dt = datetime.combine(work_date, start_time)
        end_dt = datetime.combine(work_date, end_time)
        
        duration_hours = (end_dt - start_dt).total_seconds() / 3600

        if duration_hours < 0:
            st.error("종료 시간은 시작 시간보다 빨라야 한다.")
        else:
            multiplier = 1.5  # 통상시급 1.5배 적용 (휴일/초과 공통 기준)
            estimated_pay = int(duration_hours * emp_info['통상시급'] * multiplier)

            st.info(f"💡 예상 인정 근무시간: **{duration_hours:.1f}시간** / 예상 수당: **{estimated_pay:,}원** (시급의 {multiplier}배 적용)")

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
                    "총시간": f"{duration_hours:.1f}시간",
                    "예상수당": f"{estimated_pay:,}원",
                    "사유": reason
                }
                st.session_state.overtime_records.append(record)
                st.success("초과근무 신청이 완료되었다.")

# -------------------------------------------------------------------
# TAB 3: 신청서 조회 및 인쇄용 양식
# -------------------------------------------------------------------
with tab3:
    st.header("초과근무 신청서 출력")
    
    if not st.session_state.overtime_records:
        st.info("제출된 초과근무 신청 내역이 없다.")
    else:
        df_records = pd.DataFrame(st.session_state.overtime_records)
        st.subheader("전체 신청 내역")
        st.dataframe(df_records, use_container_width=True)

        st.divider()
        st.subheader("📄 인쇄용 양식 생성")
        
        record_options = [f"[{r['근무일자']}] {r['이름']} - {r['근무구분']}" for r in st.session_state.overtime_records]
        selected_index = st.selectbox("출력할 신청 내역 선택", range(len(record_options)), format_func=lambda x: record_options[x])

        target = st.session_state.overtime_records[selected_index]

        # 인쇄 전용 HTML/CSS 서식
        print_template = f"""
        <div id="print-area" style="border: 2px solid #000; padding: 30px; font-family: 'Malgun Gothic', sans-serif; max-width: 700px; margin: auto;">
            <h2 style="text-align: center; text-decoration: underline; margin-bottom: 30px;">초 과 근 무 신 청 서</h2>
            
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;" border="1">
                <tr>
                    <th style="padding: 8px; background-color: #f2f2f2; width: 20%;">성 명</th>
                    <td style="padding: 8px; width: 30%;">{target['이름']}</td>
                    <th style="padding: 8px; background-color: #f2f2f2; width: 20%;">소 속</th>
                    <td style="padding: 8px; width: 30%;">{target['부서']}</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background-color: #f2f2f2;">사 번</th>
                    <td style="padding: 8px;">{target['사번']}</td>
                    <th style="padding: 8px; background-color: #f2f2f2;">신청일자</th>
                    <td style="padding: 8px;">{target['신청일시'][:10]}</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background-color: #f2f2f2;">근무 구분</th>
                    <td style="padding: 8px;" colspan="3">{target['근무구분']}</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background-color: #f2f2f2;">근무 일시</th>
                    <td style="padding: 8px;" colspan="3">{target['근무일자']} ({target['시작시간']} ~ {target['종료시간']}) / 총 {target['총시간']}</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background-color: #f2f2f2;">근무 사유</th>
                    <td style="padding: 8px; height: 80px; vertical-align: top;" colspan="3">{target['사유']}</td>
                </tr>
                <tr>
                    <th style="padding: 8px; background-color: #f2f2f2;">예상 수당</th>
                    <td style="padding: 8px;" colspan="3">{target['예상수당']}</td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 40px;">위와 같이 초과근무를 신청합니다.</p>
            <p style="text-align: center; margin-top: 20px;">{target['근무일자'][:4]}년 {target['근무일자'][5:7]}월 {target['근무일자'][8:10]}일</p>
            
            <p style="text-align: right; margin-top: 30px; margin-right: 20px;">신청인: {target['이름']} (인)</p>
            
            <div style="margin-top: 50px; border-top: 1px dashed #ccc; padding-top: 20px;">
                <table style="width: 100%; text-align: center; border-collapse: collapse;" border="1">
                    <tr>
                        <th rowspan="2" style="width: 10%; background-color: #f2f2f2;">결<br>재</th>
                        <th style="width: 30%;">담당</th>
                        <th style="width: 30%;">팀장</th>
                        <th style="width: 30%;">대표</th>
                    </tr>
                    <tr style="height: 60px;">
                        <td></td>
                        <td></td>
                        <td></td>
                    </tr>
                </table>
            </div>
        </div>
        """
        
        # HTML 렌더링
        st.components.v1.html(print_template, height=600, scrolling=True)
        st.info("💡 종이로 출력하려면 브라우저의 인쇄 기능(`Ctrl + P` 또는 `Cmd + P`)을 이용하면 된다.")