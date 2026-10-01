import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="통합 거래처별 손익분석 시스템", layout="wide")

st.title("📊 [통합] 거래처별 손익분석 시스템")

# 기본 거래처 목록 (배부 대상)
DEFAULT_CUSTOMERS = ["중국법인", "롯데케미칼", "지에스큐", "린코스", "케이스티파이", "마곡상인"]

# 기본 대분류 옵션 목록
DEFAULT_MAJORS = [
    "항공 운임", "도급사", "밀크런", "사내매출 조정", "기타", "카드 수수료", "센터 비용",
    "건물 감가상각비", "지게차 리스", "출입/보안 (에스원/캡스)", "파렛트 렌탈",
    "부서운영비/야특근식비", "수선비", "토지임차료", "전산 (EDI)", "소모품비 外",
    "통관수수료", "직접비", "공통비", "일반관리비"
]

# ----------------------------------------------------
# 컬럼 규격 자동 판별 함수 (실제 내용 길이에 맞춤 설정)
# ----------------------------------------------------
def auto_col_width(header_text, sample_values=None):
    def get_len(s):
        s = str(s) if pd.notna(s) else ""
        return sum(2 if ord(c) > 128 else 1 for c in s)

    max_len = get_len(header_text)
    if sample_values is not None:
        for val in sample_values:
            max_len = max(max_len, get_len(val))

    if max_len <= 10:
        return "small"      # 약 80~100px (코드, 상태, 짧은 수치)
    elif max_len <= 22:
        return "medium"     # 약 160~200px (금액, 거래처명, 일반 항목명)
    else:
        return "large"      # 약 280~320px (긴 계정명, 상세 적요)

def find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        for cand in candidates:
            if cand in str(c):
                return c
    return None

# Session State 초기화
if "customers" not in st.session_state:
    st.session_state.customers = list(DEFAULT_CUSTOMERS)

if "major_categories" not in st.session_state:
    st.session_state.major_categories = list(DEFAULT_MAJORS)

if "editor_version" not in st.session_state:
    st.session_state.editor_version = 0

if "revenue_dict" not in st.session_state:
    st.session_state.revenue_dict = {c: 0.0 for c in st.session_state.customers}

# ----------------------------------------------------
# 업로드 가이드 및 샘플 양식 제공 영역
# ----------------------------------------------------
with st.expander("ℹ️ **업로드 엑셀 파일 표준 양식 안내 및 샘플 다운로드**", expanded=False):
    st.markdown("""
    **원장 파일 필수 컬럼 구성 안내**  
    업로드할 엑셀/CSV 파일은 아래와 같은 헤더 컬럼 구조를 갖추어야 시스템이 자동으로 계정과 비용을 인식합니다.  
    - **`연도/월`**: 귀속 연월 (예: `2026/07`)
    - **`계정`**: 원장 계정과목명 (예: `글로벌(국내로컬비용)_운송(일반)`)
    - **`적요`**: 전표 세부 적요 (예: `.UPS KOREA 7월 운송비`)
    - **`금액(현지 통화)`**: 원화 발생 금액 (숫자)
    - **`거래처명`**: 매입 거래처명 (예: `유나이티드파슬서비스코리아(주)`)
    - **`구분`**: 수출, 수입, 공통 등
    - **`상세`**: 세부 품목 또는 라우팅 (예: `UPS`, `케미칼(운송)`)
    """)

    sample_df = pd.DataFrame([
        {
            "연도/월": "2026/07",
            "계정": "글로벌(국내로컬비용)_운송(일반)",
            "적요": ".UPS KOREA 7월 운송비",
            "금액(현지 통화)": 2024303418,
            "거래처명": "유나이티드파슬서비스코리아(주)",
            "구분": "수출",
            "상세": "UPS"
        },
        {
            "연도/월": "2026/07",
            "계정": "글로벌(운송비)_항공운송비",
            "적요": "2026년 7월 DHL-롯데케미칼 수출입 내역서",
            "금액(현지 통화)": 244142275,
            "거래처명": "(주)디에이치엘코리아",
            "구분": "수출",
            "상세": "케미칼(운송)"
        },
        {
            "연도/월": "2026/07",
            "계정": "국내(위탁용역비)_기타위탁용역",
            "적요": "[특송수출ㅣ엠제이플렉스] 26년 7월 물류서비스",
            "금액(현지 통화)": 35700000,
            "거래처명": "(주)엠제이플렉스",
            "구분": "공통",
            "상세": "현장도급비용"
        }
    ])

    st.markdown("##### 📋 [미리보기] 엑셀 표준 서식 예시")
    st.dataframe(
        sample_df,
        column_config={
            "금액(현지 통화)": st.column_config.NumberColumn(format="%,d")
        },
        width="stretch",
        hide_index=True
    )

    sample_buffer = io.BytesIO()
    with pd.ExcelWriter(sample_buffer, engine="openpyxl") as writer:
        sample_df.to_excel(writer, sheet_name="원장_샘플양식", index=False)
        ws = writer.sheets["원장_샘플양식"]
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=4).number_format = '#,##0'
    sample_buffer.seek(0)

    st.download_button(
        label="📥 표준 원장 엑셀 샘플 파일 다운로드 (.xlsx)",
        data=sample_buffer,
        file_name="원장_업로드_표준양식_샘플.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

uploaded_file = st.file_uploader("📂 엑셀 파일 업로드 (.xlsx, .csv)", type=["xlsx", "csv"])

# 공통 거래처 컬럼 관리 함수
def render_customer_manager(key_suffix):
    with st.expander("🏢 **거래처 컬럼 관리 (추가 / 삭제)**", expanded=False):
        c_add, c_del = st.columns(2)
        with c_add:
            st.markdown("##### ➕ 새 거래처 추가")
            new_cust_name = st.text_input("새 거래처명 입력", placeholder="예: FEDEX, 에스원 등", key=f"input_new_cust_{key_suffix}")
            if st.button("거래처 추가", type="primary", key=f"btn_add_cust_{key_suffix}"):
                clean_name = new_cust_name.strip()
                if not clean_name:
                    st.warning("거래처명을 입력해주세요.")
                elif clean_name in st.session_state.customers:
                    st.warning("이미 존재하는 거래처입니다.")
                else:
                    st.session_state.customers.append(clean_name)
                    st.session_state.dash_df[clean_name] = 0.0
                    st.session_state.revenue_dict[clean_name] = 0.0
                    st.session_state.editor_version += 1
                    st.success(f"'{clean_name}' 거래처 열이 추가되었습니다.")
                    st.rerun()

        with c_del:
            st.markdown("##### 🗑️ 기존 거래처 삭제")
            if len(st.session_state.customers) > 0:
                del_target = st.selectbox("삭제할 거래처 선택", options=["(선택안함)"] + st.session_state.customers, key=f"select_del_cust_{key_suffix}")
                if st.button("선택한 거래처 삭제", type="secondary", key=f"btn_del_cust_{key_suffix}"):
                    if del_target == "(선택안함)":
                        st.warning("삭제할 거래처를 선택해주세요.")
                    else:
                        st.session_state.customers.remove(del_target)
                        if del_target in st.session_state.dash_df.columns:
                            st.session_state.dash_df.drop(columns=[del_target], inplace=True)
                        if del_target in st.session_state.revenue_dict:
                            del st.session_state.revenue_dict[del_target]
                        st.session_state.editor_version += 1
                        st.success(f"'{del_target}' 거래처가 삭제되었습니다. (배부 금액은 미분배 잔액으로 복원)")
                        st.rerun()
            else:
                st.info("등록된 거래처가 없습니다.")

# 대분류 항목 관리 함수 (추가 / 삭제)
def render_major_category_manager(key_suffix):
    with st.expander("📂 **대분류 항목 관리 (추가 / 삭제)**", expanded=False):
        m_add, m_del = st.columns(2)
        with m_add:
            st.markdown("##### ➕ 새 대분류 추가")
            new_major_name = st.text_input("새 대분류명 입력", placeholder="예: 포장재비, 차량유지비 등", key=f"input_new_major_{key_suffix}")
            if st.button("대분류 추가", type="primary", key=f"btn_add_major_{key_suffix}"):
                clean_major = new_major_name.strip()
                if not clean_major:
                    st.warning("대분류명을 입력해주세요.")
                elif clean_major in st.session_state.major_categories:
                    st.warning("이미 존재하는 대분류입니다.")
                else:
                    st.session_state.major_categories.append(clean_major)
                    st.session_state.editor_version += 1
                    st.success(f"'{clean_major}' 대분류가 추가되었습니다. (드롭다운 목록에 즉시 반영)")
                    st.rerun()

        with m_del:
            st.markdown("##### 🗑️ 기존 대분류 삭제")
            if len(st.session_state.major_categories) > 0:
                del_major_target = st.selectbox(
                    "삭제할 대분류 선택", 
                    options=["(선택안함)"] + st.session_state.major_categories, 
                    key=f"select_del_major_{key_suffix}"
                )
                if st.button("선택한 대분류 삭제", type="secondary", key=f"btn_del_major_{key_suffix}"):
                    if del_major_target == "(선택안함)":
                        st.warning("삭제할 대분류를 선택해주세요.")
                    else:
                        st.session_state.major_categories.remove(del_major_target)
                        if "dash_df" in st.session_state and "대분류" in st.session_state.dash_df.columns:
                            mask = st.session_state.dash_df["대분류"] == del_major_target
                            if mask.any():
                                st.session_state.dash_df.loc[mask, "대분류"] = "기타"
                        st.session_state.editor_version += 1
                        st.success(f"'{del_major_target}' 대분류가 삭제되었습니다. (기존 지정 행은 '기타'로 변경)")
                        st.rerun()
            else:
                st.info("등록된 대분류 항목이 없습니다.")

@st.cache_data
def build_initial_dashboard(file, customers):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    
    # 1. 컬럼 매핑 유연화 (원장의 실제 컬럼명 탐색)
    col_acct = find_col(df, ["계정과목", "계정", "원장 계정", "원장계정", "계정명"])
    col_desc = find_col(df, ["적요", "원장 적요", "원장적요", "적요 (내용)", "내용"])
    col_amt = find_col(df, ["금액(현지 통화)", "금액", "원장 금액", "원장금액", "발생금액"])
    col_vendor = find_col(df, ["거래처명", "매입 거래처명", "매입거래처명", "거래처"])
    col_detail = find_col(df, ["상세", "원장 상세", "원장상세", "구분", "비고"])

    # 2. 데이터 정제 (계정과목 원본 그대로 보존)
    amt_series = pd.to_numeric(df[col_amt] if col_amt else 0, errors="coerce").fillna(0)
    acct_series = df[col_acct].fillna("미분류").astype(str).str.strip() if col_acct else pd.Series(["미분류"] * len(df))
    desc_series = df[col_desc].fillna("-").astype(str).str.strip() if col_desc else pd.Series(["-"] * len(df))
    vendor_series = df[col_vendor].fillna("-").astype(str).str.strip() if col_vendor else pd.Series(["-"] * len(df))
    detail_series = df[col_detail].fillna("-").astype(str).str.strip() if col_detail else pd.Series(["-"] * len(df))

    # 3. 비용구분 및 대분류 초기 자동 판정
    categories = []
    majors = []
    sga_keywords = [
        "판관", "운송경비", "관리비", "임차료", "토지", "감가상각", "AFB", "EDI", "전산",
        "보안", "에스원", "캡스", "지게차", "리스", "파렛트", "수선비", "부서운영", "식대",
        "광열", "전기", "수도", "재산세", "세금", "보험료", "소모품"
    ]
    for acct, desc, detail in zip(acct_series, desc_series, detail_series):
        combined = f"{acct} {desc} {detail}"
        
        # 비용구분
        if any(kw in combined for kw in sga_keywords):
            categories.append("판관비")
        else:
            categories.append("매출원가")
            
        # 대분류 초기값
        if "밀크런" in desc or "브이투브이" in combined:
            majors.append("밀크런")
        elif "사내매출" in combined or "CBE" in combined:
            majors.append("사내매출 조정")
        elif "항공" in combined or "UPS" in combined or "DHL" in combined or "EMS" in combined or "린코스" in combined or "SF" in combined:
            majors.append("항공 운임")
        elif "도급" in combined or "엠제이" in combined or "인터비즈" in combined or "교원" in combined:
            majors.append("도급사")
        elif "토지" in combined or "임차" in combined:
            majors.append("토지임차료")
        elif "전산" in combined or "EDI" in combined or "케이씨넷" in combined:
            majors.append("전산 (EDI)")
        elif "감가상각" in combined or "AFB" in combined:
            majors.append("건물 감가상각비")
        elif "지게차" in combined or "리스" in combined:
            majors.append("지게차 리스")
        elif "보안" in combined or "에스원" in combined or "캡스" in combined:
            majors.append("출입/보안 (에스원/캡스)")
        elif "파렛트" in combined:
            majors.append("파렛트 렌탈")
        elif "수선" in combined:
            majors.append("수선비")
        elif "부서운영" in combined or "식대" in combined:
            majors.append("부서운영비/야특근식비")
        elif "카드" in combined:
            majors.append("카드 수수료")
        elif "전기" in combined or "수도" in combined or "맥서브" in combined:
            majors.append("센터 비용")
        else:
            majors.append("기타")

    # 4. 개별 전표 보존 (계정과목 원본 그대로, 대분류 포함)
    processed_df = pd.DataFrame({
        "비용구분": categories,
        "계정과목": acct_series,
        "대분류": majors,
        "적요": desc_series,
        "거래처명": vendor_series,
        "상세": detail_series,
        "전체(Sheet3 총액)": amt_series
    })

    # 화주별 배부액은 임의 배부 없이 0원 시작
    for c in customers:
        processed_df[c] = 0.0

    processed_df["row_id"] = processed_df.index
    return df, processed_df

if uploaded_file is not None:
    if "dash_df" not in st.session_state:
        raw_df, init_dash = build_initial_dashboard(uploaded_file, st.session_state.customers)
        st.session_state.raw_df = raw_df
        st.session_state.dash_df = init_dash
        for m in init_dash["대분류"].dropna().unique():
            m_clean = str(m).strip()
            if m_clean and m_clean not in st.session_state.major_categories:
                st.session_state.major_categories.append(m_clean)

    dash_df = st.session_state.dash_df

    for c in st.session_state.customers:
        if c not in st.session_state.revenue_dict:
            st.session_state.revenue_dict[c] = 0.0

    tab1, tab2, tab3 = st.tabs([
        "⚙️ [탭 1] 원가/판관비 배부", 
        "💵 [탭 2] 거래처별 매출 입력", 
        "📈 [탭 3] 최종 손익계산서 & 대시보드"
    ])

    # =========================================================================
    # [탭 1] 원가 및 판관비 배부 작업
    # =========================================================================
    with tab1:
        render_customer_manager("tab1")
        render_major_category_manager("tab1")

        st.write("---")

        for c in st.session_state.customers:
            if c not in dash_df.columns:
                dash_df[c] = 0.0

        valid_customers = [c for c in st.session_state.customers if c in dash_df.columns]
        if len(valid_customers) > 0:
            dash_df["배부 완료액"] = dash_df[valid_customers].sum(axis=1)
        else:
            dash_df["배부 완료액"] = 0.0

        dash_df["미분배 잔액"] = dash_df["전체(Sheet3 총액)"] - dash_df["배부 완료액"]

        cogs_mask = dash_df["비용구분"] == "매출원가"
        sga_mask = dash_df["비용구분"] == "판관비"

        total_amount = dash_df["전체(Sheet3 총액)"].sum()
        cogs_total = dash_df.loc[cogs_mask, "전체(Sheet3 총액)"].sum()
        sga_total = dash_df.loc[sga_mask, "전체(Sheet3 총액)"].sum()
        total_unallocated = dash_df["미분배 잔액"].sum()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("총 발생 비용 (원장 100%)", f"{total_amount:,.0f}")
        k2.metric("1. 매출원가 합계", f"{cogs_total:,.0f}")
        k3.metric("2. 판관비 합계", f"{sga_total:,.0f}")
        k4.metric(
            "전체 미분배 잔여 금액", 
            f"{total_unallocated:,.0f}", 
            delta="전체 귀속 완료 (0원) ✅" if total_unallocated == 0 else f"미귀속 잔여: {total_unallocated:,.0f}",
            delta_color="normal" if total_unallocated == 0 else "inverse"
        )

        all_major_options = sorted(list(set(st.session_state.major_categories + [str(x) for x in dash_df["대분류"].dropna().unique() if str(x).strip()])))

        base_column_config = {
            "비용구분": st.column_config.SelectboxColumn(
                label="🔄 비용구분 (이동)",
                options=["매출원가", "판관비"],
                required=True,
                width=auto_col_width("🔄 비용구분 (이동)", ["매출원가", "판관비"]),
                help="'매출원가'와 '판관비'를 변경하면 1번과 2번 표 사이로 즉시 이동합니다."
            ),
            "계정과목": st.column_config.TextColumn(
                label="📑 계정과목(원장)",
                disabled=True,
                width=auto_col_width("📑 계정과목(원장)", dash_df["계정과목"]),
                help="원장에 기록된 계정과목명 그대로입니다."
            ),
            "대분류": st.column_config.SelectboxColumn(
                label="👉 [선택] 대분류",
                options=all_major_options,
                required=True,
                disabled=False,
                width=auto_col_width("👉 [선택] 대분류", dash_df["대분류"]),
                help="클릭하여 대분류를 직접 선택하거나 수정할 수 있습니다."
            ),
            "적요": st.column_config.TextColumn(
                label="📝 적요",
                disabled=False,
                width=auto_col_width("📝 적요", dash_df["적요"]),
                help="전표 적요 내용을 직접 수정할 수 있습니다."
            ),
            "거래처명": st.column_config.TextColumn(
                label="🏢 거래처명(원장)",
                disabled=True,
                width=auto_col_width("🏢 거래처명(원장)", dash_df["거래처명"]),
                help="원장에 기록된 매입 거래처명입니다."
            ),
            "전체(Sheet3 총액)": st.column_config.NumberColumn(
                label="전체 발생금액", 
                format="%,d", 
                disabled=True,
                width=auto_col_width("전체 발생금액", ["3,504,010,871"])
            ),
            "미분배 잔액": st.column_config.NumberColumn(
                label="⚠️ 미분배 잔액", 
                format="%,d", 
                disabled=True,
                width=auto_col_width("⚠️ 미분배 잔액", ["3,504,010,871"]),
                help="거래처별로 배부하고 남은 잔액입니다. 0원이 되어야 완전히 귀속된 상태입니다."
            )
        }
        for c in st.session_state.customers:
            base_column_config[c] = st.column_config.NumberColumn(
                label=f"🏢 {c}", 
                format="%,d", 
                disabled=False,
                min_value=0.0,
                width=auto_col_width(f"🏢 {c}", ["1,000,000,000"])
            )

        # 계정과목(원장) 우측에 [선택] 대분류 배치
        display_cols = ["row_id", "비용구분", "계정과목", "대분류", "적요", "거래처명", "전체(Sheet3 총액)", "미분배 잔액"] + st.session_state.customers

        st.subheader("1. 📦 [매출원가] 거래처별 손익분석")
        cogs_display_df = dash_df[cogs_mask][display_cols].copy()
        cogs_config = dict(base_column_config)
        cogs_config["row_id"] = None

        edited_cogs = st.data_editor(
            cogs_display_df,
            column_config=cogs_config,
            width="stretch",
            hide_index=True,
            key=f"cogs_editor_{st.session_state.editor_version}"
        )

        st.subheader("2. 🏢 [판관비] 거래처별 손익분석")
        sga_display_df = dash_df[sga_mask][display_cols].copy()
        sga_config = dict(base_column_config)
        sga_config["row_id"] = None

        edited_sga = st.data_editor(
            sga_display_df,
            column_config=sga_config,
            width="stretch",
            hide_index=True,
            key=f"sga_editor_{st.session_state.editor_version}"
        )

        # 동기화 (대분류 및 적요 직접 수정 반영)
        modified = False
        for sub_edited in [edited_cogs, edited_sga]:
            for idx, row in sub_edited.iterrows():
                r_id = row["row_id"]
                if row["비용구분"] != st.session_state.dash_df.loc[r_id, "비용구분"]:
                    st.session_state.dash_df.loc[r_id, "비용구분"] = row["비용구분"]
                    modified = True
                if "대분류" in row and row["대분류"] != st.session_state.dash_df.loc[r_id, "대분류"]:
                    st.session_state.dash_df.loc[r_id, "대분류"] = row["대분류"]
                    modified = True
                if "적요" in row and row["적요"] != st.session_state.dash_df.loc[r_id, "적요"]:
                    st.session_state.dash_df.loc[r_id, "적요"] = row["적요"]
                    modified = True
                for c in st.session_state.customers:
                    new_val = float(row[c]) if pd.notna(row[c]) else 0.0
                    old_val = float(st.session_state.dash_df.loc[r_id, c]) if c in st.session_state.dash_df.columns else 0.0
                    if new_val != old_val:
                        st.session_state.dash_df.loc[r_id, c] = new_val
                        modified = True

        if modified:
            st.session_state.editor_version += 1
            st.rerun()

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("🔄 동일한 [비용구분/대분류/계정과목/거래처] 합산 정리"):
                agg_dict = {"전체(Sheet3 총액)": "sum", "적요": lambda s: ", ".join(dict.fromkeys(s))[:50]}
                for c in st.session_state.customers:
                    agg_dict[c] = "sum"
                merged_df = st.session_state.dash_df.groupby(["비용구분", "대분류", "계정과목", "거래처명"], as_index=False).agg(agg_dict)
                merged_df["row_id"] = merged_df.index
                st.session_state.dash_df = merged_df
                st.session_state.editor_version += 1
                st.success("동일 항목 기준 합산 정리가 완료되었습니다.")
                st.rerun()

    # =========================================================================
    # [탭 2] 거래처별 매출 입력 (내용 규격에 맞춘 단정한 너비 적용)
    # =========================================================================
    with tab2:
        render_customer_manager("tab2")

        st.write("---")
        st.subheader("💵 거래처별 매출액(Revenue) 입력")
        st.caption("각 거래처(행)별 당월 매출액을 입력해 주세요. 전체 총 매출액은 자동으로 합산되며 [탭 3] 손익계산서로 연동됩니다.")

        rev_rows = [{"거래처명": c, "당월 매출액": float(st.session_state.revenue_dict.get(c, 0.0))} for c in st.session_state.customers]
        rev_df = pd.DataFrame(rev_rows)

        cust_width = auto_col_width("🏢 거래처명", rev_df["거래처명"])
        rev_width = auto_col_width("💰 당월 매출액", [f"{v:,.0f}" for v in st.session_state.revenue_dict.values()])

        col_t2_left, col_t2_right = st.columns(2)
        with col_t2_left:
            edited_rev = st.data_editor(
                rev_df,
                column_config={
                    "거래처명": st.column_config.TextColumn(
                        label="🏢 거래처명", 
                        disabled=True,
                        width=cust_width
                    ),
                    "당월 매출액": st.column_config.NumberColumn(
                        label="💰 당월 매출액",
                        format="%,d",
                        min_value=0.0,
                        step=1000000.0,
                        width=rev_width,
                        help="해당 거래처의 발생 매출액을 입력하세요."
                    )
                },
                width="stretch",
                hide_index=True,
                key=f"rev_row_editor_{st.session_state.editor_version}"
            )

        # 매출액 변경사항 동기화
        rev_modified = False
        for _, r in edited_rev.iterrows():
            c_name = r["거래처명"]
            c_amt = float(r["당월 매출액"]) if pd.notna(r["당월 매출액"]) else 0.0
            if st.session_state.revenue_dict.get(c_name, 0.0) != c_amt:
                st.session_state.revenue_dict[c_name] = c_amt
                rev_modified = True

        if rev_modified:
            st.rerun()

        st.write("---")
        tot_rev_sum = sum(st.session_state.revenue_dict.values())
        active_cust_cnt = sum(1 for v in st.session_state.revenue_dict.values() if v > 0)
        rm1, rm2 = st.columns(2)
        rm1.metric("총 매출액 합계", f"{tot_rev_sum:,.0f}")
        rm2.metric("매출 발생 거래처 수", f"{active_cust_cnt} / {len(st.session_state.customers)}개 사")

    # =========================================================================
    # [탭 3] 최종 손익계산서 & 대시보드
    # =========================================================================
    with tab3:
        st.subheader("📈 [통합] 거래처별 최종 손익계산서")
        st.caption("💡 판관비와 매출원가를 각각 **합계금액**으로 집계하여 전체 손익 요약을 한눈에 확인합니다.")

        cogs_sub_df = dash_df[dash_df["비용구분"] == "매출원가"]
        sga_sub_df = dash_df[dash_df["비용구분"] == "판관비"]

        tot_rev = sum(st.session_state.revenue_dict.values())
        tot_cogs = cogs_sub_df["전체(Sheet3 총액)"].sum()
        tot_gp = tot_rev - tot_cogs
        gp_ratio = (tot_gp / tot_rev * 100) if tot_rev > 0 else 0.0

        tot_sga = sga_sub_df["전체(Sheet3 총액)"].sum()
        tot_op = tot_gp - tot_sga
        op_ratio = (tot_op / tot_rev * 100) if tot_rev > 0 else 0.0

        # 상단 핵심 KPI
        kp1, kp2, kp3, kp4, kp5 = st.columns(5)
        kp1.metric("1. 총 매출액", f"{tot_rev:,.0f}")
        kp2.metric("2. 매출원가 합계", f"{tot_cogs:,.0f}")
        kp3.metric("3. 매출총이익", f"{tot_gp:,.0f}", delta=f"이익률: {gp_ratio:.1f}%")
        kp4.metric("4. 판관비 합계", f"{tot_sga:,.0f}")
        kp5.metric(
            "5. 영업이익", 
            f"{tot_op:,.0f}", 
            delta=f"영업이익률: {op_ratio:.1f}% ({'흑자' if tot_op >= 0 else '적자'})",
            delta_color="normal" if tot_op >= 0 else "inverse"
        )

        st.write("---")

        summary_pnl_rows = []

        row_rev = {"항목": "1. 매출액", "전체(합계)": tot_rev}
        for c in st.session_state.customers:
            row_rev[c] = st.session_state.revenue_dict.get(c, 0.0)
        summary_pnl_rows.append(row_rev)

        row_cogs = {"항목": "2. 매출원가 (합계)", "전체(합계)": tot_cogs}
        for c in st.session_state.customers:
            row_cogs[c] = cogs_sub_df[c].sum() if c in cogs_sub_df.columns else 0.0
        summary_pnl_rows.append(row_cogs)

        row_gp = {"항목": "3. 매출총이익 (1 - 2)", "전체(합계)": tot_gp}
        for c in st.session_state.customers:
            row_gp[c] = row_rev[c] - row_cogs[c]
        summary_pnl_rows.append(row_gp)

        row_sga = {"항목": "4. 판관비 (합계)", "전체(합계)": tot_sga}
        for c in st.session_state.customers:
            row_sga[c] = sga_sub_df[c].sum() if c in sga_sub_df.columns else 0.0
        summary_pnl_rows.append(row_sga)

        row_op = {"항목": "5. 영업이익 (3 - 4)", "전체(합계)": tot_op}
        for c in st.session_state.customers:
            row_op[c] = row_gp[c] - row_sga[c]
        summary_pnl_rows.append(row_op)

        summary_pnl_df = pd.DataFrame(summary_pnl_rows)

        summary_pnl_config = {
            "항목": st.column_config.TextColumn(
                label="손익 구분 항목", 
                width=auto_col_width("손익 구분 항목", summary_pnl_df["항목"])
            ),
            "전체(합계)": st.column_config.NumberColumn(
                label="전체(합계)", 
                format="%,d", 
                width=auto_col_width("전체(합계)", ["3,504,010,871"])
            )
        }
        for c in st.session_state.customers:
            summary_pnl_config[c] = st.column_config.NumberColumn(
                label=f"🏢 {c}", 
                format="%,d", 
                width=auto_col_width(f"🏢 {c}", ["1,000,000,000"])
            )

        st.dataframe(
            summary_pnl_df,
            column_config=summary_pnl_config,
            width="stretch",
            hide_index=True
        )

        # 대분류/계정과목별 세부 내역
        with st.expander("🔍 **매출원가 및 판관비 세부 내역 펼쳐보기**", expanded=False):
            cogs_detail = cogs_sub_df.groupby(["대분류", "계정과목"], as_index=False)[["전체(Sheet3 총액)"] + st.session_state.customers].sum()
            sga_detail = sga_sub_df.groupby(["대분류", "계정과목"], as_index=False)[["전체(Sheet3 총액)"] + st.session_state.customers].sum()

            detail_config = {
                "대분류": st.column_config.TextColumn(
                    label="대분류", 
                    width=auto_col_width("대분류", list(cogs_detail["대분류"]) + list(sga_detail["대분류"]))
                ),
                "계정과목": st.column_config.TextColumn(
                    label="계정과목", 
                    width=auto_col_width("계정과목", list(cogs_detail["계정과목"]) + list(sga_detail["계정과목"]))
                ),
                "전체(Sheet3 총액)": st.column_config.NumberColumn(
                    label="발생 총액", 
                    format="%,d", 
                    width="medium"
                )
            }
            for c in st.session_state.customers:
                detail_config[c] = st.column_config.NumberColumn(
                    label=c, 
                    format="%,d", 
                    width="medium"
                )

            st.markdown("##### 📦 매출원가 대분류/계정과목별 내역")
            st.dataframe(cogs_detail, column_config=detail_config, width="stretch", hide_index=True)

            st.markdown("##### 🏢 판관비 대분류/계정과목별 내역")
            st.dataframe(sga_detail, column_config=detail_config, width="stretch", hide_index=True)

        # 거래처별 손익 기여 요약
        st.write("---")
        st.subheader("📊 거래처별 손익 기여 요약")
        sum_rows = []
        for c in st.session_state.customers:
            c_rev = row_rev[c]
            c_cogs = row_cogs[c]
            c_gp = row_gp[c]
            c_sga = row_sga[c]
            c_op = row_op[c]
            c_opm = (c_op / c_rev * 100) if c_rev > 0 else 0.0
            status = "🟢 흑자" if c_op >= 0 else "🔴 적자"
            sum_rows.append({
                "거래처": c,
                "매출액": c_rev,
                "매출원가": c_cogs,
                "매출총이익": c_gp,
                "판관비": c_sga,
                "영업이익": c_op,
                "영업이익률": f"{c_opm:.1f}%",
                "수익성": status
            })
        
        sum_df = pd.DataFrame(sum_rows)
        st.dataframe(
            sum_df,
            column_config={
                "거래처": st.column_config.TextColumn(width=auto_col_width("거래처", sum_df["거래처"])),
                "매출액": st.column_config.NumberColumn(format="%,d", width="medium"),
                "매출원가": st.column_config.NumberColumn(format="%,d", width="medium"),
                "매출총이익": st.column_config.NumberColumn(format="%,d", width="medium"),
                "판관비": st.column_config.NumberColumn(format="%,d", width="medium"),
                "영업이익": st.column_config.NumberColumn(format="%,d", width="medium"),
                "영업이익률": st.column_config.TextColumn(width="small"),
                "수익성": st.column_config.TextColumn(width="small"),
            },
            width="stretch",
            hide_index=True
        )

        # 엑셀 다운로드 (최종 손익요약 시트 + 세부시트)
        st.write("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_pnl_df.to_excel(writer, sheet_name="최종_손익요약", index=False)
            sum_df.to_excel(writer, sheet_name="거래처별_손익요약", index=False)
            dash_df[dash_df["비용구분"] == "매출원가"].to_excel(writer, sheet_name="1.매출원가_상세배부", index=False)
            dash_df[dash_df["비용구분"] == "판관비"].to_excel(writer, sheet_name="2.판관비_상세배부", index=False)
            st.session_state.raw_df.to_excel(writer, sheet_name="원장전체_원본내역", index=False)

            ws = writer.sheets["최종_손익요약"]
            for col_idx in range(2, len(summary_pnl_df.columns) + 1):
                for row_idx in range(2, ws.max_row + 1):
                    ws.cell(row=row_idx, column=col_idx).number_format = '#,##0'

        output.seek(0)

        st.download_button(
            label="💾 완성된 최종 손익계산서 엑셀 다운로드",
            data=output,
            file_name="거래처별_최종_손익계산서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )