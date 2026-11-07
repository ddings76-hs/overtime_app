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
                _range_view=workflow_button_nav(
                    "기간조회 업무 선택",
                    ["🌴 연차 기간조회","⏱️ 초과근무 기간조회","🚗 출장 기간조회"],
                    key="v326_range_view",
                    caption="조회할 업무를 선택해 주세요.",
                    max_per_row=3
                )

                if _range_view=="🌴 연차 기간조회":
                    _rlv_view=employee_filter_ui(_rlv,"v202_range_leave_emp","연차 직원 검색")
                    display_table_kr(_rlv_view,use_container_width=True,hide_index=True)

                if _range_view=="⏱️ 초과근무 기간조회":
                    _rot_view=employee_filter_ui(_rot,"v202_range_ot_emp","초과근무 직원 검색")
                    display_table_kr(_rot_view,use_container_width=True,hide_index=True)

                if _range_view=="🚗 출장 기간조회":
                    _rtr_view=employee_filter_ui(_rtr,"v202_range_trip_emp","출장 직원 검색")
                    display_table_kr(_rtr_view,use_container_width=True,hide_index=True)

        if _report_view=="📈 연간 현황":
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

        if _report_view=="🔎 데이터 점검":
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

        if _report_view=="📥 Excel 통합보고서":
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
