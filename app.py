import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="통합 손익분석 시스템 (매출원가/판관비)", layout="wide")

st.title("📊 [통합] 거래처별 손익분석 시스템")

# 기본 거래처 목록 (배부 대상)
DEFAULT_CUSTOMERS = ["중국법인", "롯데케미칼", "지에스큐", "린코스", "케이스티파이", "마곡상인"]

# 표준 대분류 및 세부 항목 옵션
STD_MAJORS = [
    "항공 운임", "도급사", "밀크런", "기타", "카드 수수료", "센터 비용",
    "건물 감가상각비", "지게차 리스", "출입/보안 (에스원/캡스)", "파렛트 렌탈",
    "부서운영비/야특근식비", "수선비", "토지임차료", "전산 (EDI)", "소모품비 外"
]

STD_SUBS = [
    "UPS", "EMS", "ICS", "RINCOS", "마곡상인", "FEDEX",
    "(주)엠제이플렉스", "(주)인터비즈시스템", "(주)교원스타트원",
    "밀크런", "기타(항공/기타)", "카드 수수료", "센터 비용",
    "건물 감가상각비", "지게차 리스", "보안", "파렛트 렌탈",
    "부서운영비", "수선비", "토지임차료", "전산 (EDI)", "소모품비 外",
    "감가상각비", "센터운영비용", "DHL", "밀크런 픽업", "케미칼(운송)", "통관수수료", "전기료", "통신비"
]

MAPPING_RULES = [
    {"category": "매출원가", "sub_cat": "항공 운임", "item": "UPS", "condition": lambda r: "UPS" in str(r["상세"]) or "유나이티드파슬" in str(r["거래처명"])},
    {"category": "매출원가", "sub_cat": "항공 운임", "item": "EMS", "condition": lambda r: "EMS" in str(r["상세"]) and "마곡상인" in str(r["거래처명"]) and "K-packet" not in str(r["적요"])},
    {"category": "매출원가", "sub_cat": "항공 운임", "item": "ICS", "condition": lambda r: "아이씨에스" in str(r["거래처명"]) and "미국" in str(r["상세"])},
    {"category": "매출원가", "sub_cat": "항공 운임", "item": "RINCOS", "condition": lambda r: "린코스" in str(r["거래처명"])},
    {"category": "매출원가", "sub_cat": "항공 운임", "item": "마곡상인", "condition": lambda r: "K-packet" in str(r["적요"]) or "K-PACKET" in str(r["상세"])},
    {"category": "매출원가", "sub_cat": "항공 운임", "item": "FEDEX", "condition": lambda r: "FEDEX" in str(r["적요"]).upper()},
    {"category": "매출원가", "sub_cat": "도급사", "item": "(주)엠제이플렉스", "condition": lambda r: "엠제이플렉스" in str(r["거래처명"])},
    {"category": "매출원가", "sub_cat": "도급사", "item": "(주)인터비즈시스템", "condition": lambda r: "인터비즈" in str(r["거래처명"])},
    {"category": "매출원가", "sub_cat": "도급사", "item": "(주)교원스타트원", "condition": lambda r: "교원" in str(r["거래처명"])},
    {"category": "매출원가", "sub_cat": "밀크런", "item": "밀크런", "condition": lambda r: "사내매출 조정" in str(r["계정"]) or "롯데택배" in str(r["상세"])},
    {"category": "매출원가", "sub_cat": "기타", "item": "기타(항공/기타)", "condition": lambda r: "DHL" in str(r["거래처명"]) or "SF" in str(r["상세"])},
    {"category": "판관비", "sub_cat": "토지임차료", "item": "토지임차료", "condition": lambda r: "토지사용료" in str(r["상세"]) or "토지임차료" in str(r["적요"])},
    {"category": "판관비", "sub_cat": "전산 (EDI)", "item": "전산 (EDI)", "condition": lambda r: "EDI" in str(r["상세"]) or "케이씨넷" in str(r["거래처명"])},
    {"category": "판관비", "sub_cat": "지게차 리스", "item": "지게차 리스", "condition": lambda r: "지게차" in str(r["적요"])},
    {"category": "판관비", "sub_cat": "출입/보안 (에스원/캡스)", "item": "보안", "condition": lambda r: "에스원" in str(r["거래처명"])},
    {"category": "판관비", "sub_cat": "파렛트 렌탈", "item": "파렛트 렌탈", "condition": lambda r: "한국파렛트풀" in str(r["거래처명"])},
    {"category": "판관비", "sub_cat": "수선비", "item": "수선비", "condition": lambda r: "수선" in str(r["계정"]) or "블루원" in str(r["거래처명"])},
    {"category": "판관비", "sub_cat": "부서운영비/야특근식비", "item": "부서운영비", "condition": lambda r: "부서운영비" in str(r["계정"]) or "야특근식대" in str(r["계정"])},
    {"category": "판관비", "sub_cat": "센터 비용", "item": "센터 비용", "condition": lambda r: "맥서브" in str(r["거래처명"]) or "효성전기" in str(r["거래처명"])},
    {"category": "판관비", "sub_cat": "기타", "item": "소모품비 外", "condition": lambda r: "소모품비" in str(r["계정"]) or "보험료" in str(r["상세"]) or "재산세" in str(r["상세"])},
    {"category": "판관비", "sub_cat": "카드 수수료", "item": "카드 수수료", "condition": lambda r: "카드수수료" in str(r["적요"])},
    {"category": "판관비", "sub_cat": "건물 감가상각비", "item": "건물 감가상각비", "condition": lambda r: "리스자산 상각" in str(r["적요"])},
]

if "customers" not in st.session_state:
    st.session_state.customers = list(DEFAULT_CUSTOMERS)

if "editor_version" not in st.session_state:
    st.session_state.editor_version = 0

uploaded_file = st.file_uploader("📂 7월 원장 엑셀 파일 업로드 (.xlsx, .csv)", type=["xlsx", "csv"])

@st.cache_data
def build_initial_dashboard(file, customers):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    
    df["금액(현지 통화)"] = pd.to_numeric(df["금액(현지 통화)"], errors="coerce").fillna(0)
    df["계정"] = df["계정"].fillna("미분류")
    df["거래처명"] = df["거래처명"].fillna("-").astype(str).str.strip()
    df["적요"] = df["적요"].fillna("-")
    df["상세"] = df["상세"].fillna("-")

    categories, sub_cats, items, vendors = [], [], [], []
    for _, row in df.iterrows():
        matched = False
        if "브이투브이" in str(row["거래처명"]) or "한국통상" in str(row["거래처명"]) or "AFB/" in str(row["적요"]):
            matched = False
        else:
            for rule in MAPPING_RULES:
                if rule["condition"](row):
                    categories.append(rule["category"])
                    sub_cats.append(rule["sub_cat"])
                    items.append(rule["item"])
                    vendors.append(str(row["거래처명"]))
                    matched = True
                    break
        if not matched:
            acct_val = str(row["계정"]).strip() if str(row["계정"]).strip() not in ["", "nan"] else "미분류"
            # 계정명 또는 성격으로 매출원가 vs 판관비 1차 판별
            if "운송경비" in acct_val or "감가상각" in str(row["상세"]) or "AFB" in str(row["적요"]) or "재산세" in str(row["상세"]):
                categories.append("판관비")
            else:
                categories.append("매출원가")
            sub_cats.append(acct_val)
            detail_val = str(row["상세"]).strip() if str(row["상세"]).strip() not in ["", "-", "nan"] else str(row["적요"])[:20]
            items.append(detail_val)
            vendors.append(str(row["거래처명"]))

    df["비용구분"] = categories
    df["대분류"] = sub_cats
    df["세부 항목"] = items
    df["거래처명"] = vendors

    # [비용구분, 대분류, 세부 항목, 거래처명] 그룹핑
    grouped = df.groupby(["비용구분", "대분류", "세부 항목", "거래처명"], as_index=False)["금액(현지 통화)"].sum()
    grouped.rename(columns={"금액(현지 통화)": "전체(Sheet3 총액)"}, inplace=True)

    for c in customers:
        grouped[c] = 0.0

    # 명확한 항목 기본 귀속
    for idx, row in grouped.iterrows():
        it = row["세부 항목"]
        tot = row["전체(Sheet3 총액)"]
        if it == "(주)인터비즈시스템" and "롯데케미칼" in customers:
            grouped.loc[idx, "롯데케미칼"] = tot
        elif it == "RINCOS" and "린코스" in customers:
            grouped.loc[idx, "린코스"] = tot
        elif it == "마곡상인" and "마곡상인" in customers:
            grouped.loc[idx, "마곡상인"] = tot

    grouped["row_id"] = grouped.index
    return df, grouped

if uploaded_file is not None:
    if "dash_df" not in st.session_state:
        raw_df, init_dash = build_initial_dashboard(uploaded_file, st.session_state.customers)
        st.session_state.raw_df = raw_df
        st.session_state.dash_df = init_dash

    dash_df = st.session_state.dash_df

    # 1. 거래처 컬럼 관리 (추가 / 삭제)
    with st.expander("🏢 **거래처 컬럼 관리 (추가 / 삭제)**", expanded=False):
        c_add, c_del = st.columns(2)
        with c_add:
            st.markdown("##### ➕ 새 거래처 추가")
            new_cust_name = st.text_input("새 거래처명 입력", placeholder="예: FEDEX, 에스원 등", key="input_new_cust")
            if st.button("거래처 추가", type="primary"):
                clean_name = new_cust_name.strip()
                if not clean_name:
                    st.warning("거래처명을 입력해주세요.")
                elif clean_name in st.session_state.customers:
                    st.warning("이미 존재하는 거래처입니다.")
                else:
                    st.session_state.customers.append(clean_name)
                    st.session_state.dash_df[clean_name] = 0.0
                    st.session_state.editor_version += 1
                    st.success(f"'{clean_name}' 거래처 열이 추가되었습니다.")
                    st.rerun()

        with c_del:
            st.markdown("##### 🗑️ 기존 거래처 삭제")
            if len(st.session_state.customers) > 0:
                del_target = st.selectbox("삭제할 거래처 선택", options=["(선택안함)"] + st.session_state.customers, key="select_del_cust")
                if st.button("선택한 거래처 삭제", type="secondary"):
                    if del_target == "(선택안함)":
                        st.warning("삭제할 거래처를 선택해주세요.")
                    else:
                        st.session_state.customers.remove(del_target)
                        if del_target in st.session_state.dash_df.columns:
                            st.session_state.dash_df.drop(columns=[del_target], inplace=True)
                        st.session_state.editor_version += 1
                        st.success(f"'{del_target}' 거래처가 삭제되었습니다. (배부 금액은 미분배 잔액으로 복원)")
                        st.rerun()
            else:
                st.info("등록된 거래처가 없습니다.")

    st.write("---")

    # 2. 실시간 계산 및 지표 대시보드
    if len(st.session_state.customers) > 0:
        dash_df["배부 완료액"] = dash_df[st.session_state.customers].sum(axis=1)
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
    k1.metric("총 발생 비용 (원장 100%)", f"₩{total_amount:,.0f}")
    k2.metric("1. 매출원가 합계", f"₩{cogs_total:,.0f}")
    k3.metric("2. 판관비 합계", f"₩{sga_total:,.0f}")
    k4.metric(
        "전체 미분배 잔여 금액", 
        f"₩{total_unallocated:,.0f}", 
        delta="전체 귀속 완료 (0원) ✅" if total_unallocated == 0 else f"미귀속 잔여: ₩{total_unallocated:,.0f}",
        delta_color="normal" if total_unallocated == 0 else "inverse"
    )

    st.write("---")

    # 드롭다운 옵션 준비
    all_major_options = sorted(list(set(STD_MAJORS + dash_df["대분류"].dropna().unique().tolist())))
    all_sub_options = sorted(list(set(STD_SUBS + dash_df["세부 항목"].dropna().unique().tolist())))

    base_column_config = {
        "비용구분": st.column_config.SelectboxColumn(
            label="🔄 비용구분 (이동)",
            options=["매출원가", "판관비"],
            required=True,
            help="'매출원가'와 '판관비'를 변경하면 1번과 2번 표 사이로 즉시 이동합니다."
        ),
        "대분류": st.column_config.SelectboxColumn(
            label="👉 [선택] 대분류",
            options=all_major_options,
            required=True,
            help="클릭하여 대분류를 변경할 수 있습니다."
        ),
        "세부 항목": st.column_config.SelectboxColumn(
            label="👉 [선택] 세부 항목",
            options=all_sub_options,
            required=True,
            help="클릭하여 세부 항목을 변경할 수 있습니다."
        ),
        "거래처명": st.column_config.TextColumn(
            label="🏢 거래처명(원장)",
            disabled=True,
            help="원장에 기록된 매입 거래처명입니다."
        ),
        "전체(Sheet3 총액)": st.column_config.NumberColumn(label="전체(Sheet3 총액)", format="₩%,d", disabled=True),
        "미분배 잔액": st.column_config.NumberColumn(
            label="⚠️ 미분배 잔액", 
            format="₩%,d", 
            disabled=True,
            help="거래처별로 배부하고 남은 잔액입니다. 0원이 되어야 완전히 귀속된 상태입니다."
        )
    }
    for c in st.session_state.customers:
        base_column_config[c] = st.column_config.NumberColumn(
            label=f"🏢 {c}", 
            format="₩%,d", 
            disabled=False,
            min_value=0.0
        )

    display_cols = ["row_id", "비용구분", "대분류", "세부 항목", "거래처명", "전체(Sheet3 총액)", "미분배 잔액"] + st.session_state.customers

    # 3. [1번 테이블] 매출원가 손익분석
    st.subheader("1. 📦 [매출원가] 거래처별 손익분석")
    st.caption("💡 '비용구분' 셀을 클릭하여 **'판관비'**로 변경하면 2번 판관비 표로 즉시 이동합니다.")

    cogs_display_df = dash_df[cogs_mask][display_cols].copy()
    cogs_editor_key = f"cogs_editor_{st.session_state.editor_version}"

    cogs_config = dict(base_column_config)
    cogs_config["row_id"] = None  # 식별자 컬럼 숨김

    edited_cogs = st.data_editor(
        cogs_display_df,
        column_config=cogs_config,
        use_container_width=True,
        hide_index=True,
        key=cogs_editor_key
    )

    # 4. [2번 테이블] 판관비 손익분석
    st.subheader("2. 🏢 [판관비] 거래처별 손익분석")
    st.caption("💡 '비용구분' 셀을 클릭하여 **'매출원가'**로 변경하면 1번 매출원가 표로 즉시 이동합니다.")

    sga_display_df = dash_df[sga_mask][display_cols].copy()
    sga_editor_key = f"sga_editor_{st.session_state.editor_version}"

    sga_config = dict(base_column_config)
    sga_config["row_id"] = None  # 식별자 컬럼 숨김

    edited_sga = st.data_editor(
        sga_display_df,
        column_config=sga_config,
        use_container_width=True,
        hide_index=True,
        key=sga_editor_key
    )

    # 두 테이블의 변경사항 동기화 (이동, 대분류, 세부항목, 거래처 금액)
    modified = False
    for sub_edited in [edited_cogs, edited_sga]:
        for idx, row in sub_edited.iterrows():
            r_id = row["row_id"]
            # 비용구분 이동 감지
            if row["비용구분"] != st.session_state.dash_df.loc[r_id, "비용구분"]:
                st.session_state.dash_df.loc[r_id, "비용구분"] = row["비용구분"]
                modified = True
            # 대분류 변경 감지
            if row["대분류"] != st.session_state.dash_df.loc[r_id, "대분류"]:
                st.session_state.dash_df.loc[r_id, "대분류"] = row["대분류"]
                modified = True
            # 세부 항목 변경 감지
            if row["세부 항목"] != st.session_state.dash_df.loc[r_id, "세부 항목"]:
                st.session_state.dash_df.loc[r_id, "세부 항목"] = row["세부 항목"]
                modified = True
            # 거래처 금액 변경 감지
            for c in st.session_state.customers:
                new_val = float(row[c]) if pd.notna(row[c]) else 0.0
                old_val = float(st.session_state.dash_df.loc[r_id, c]) if c in st.session_state.dash_df.columns else 0.0
                if new_val != old_val:
                    st.session_state.dash_df.loc[r_id, c] = new_val
                    modified = True

    if modified:
        st.session_state.editor_version += 1
        st.rerun()

    # 5. 편의 기능 및 엑셀 다운로드
    st.write("---")
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("🔄 동일한 [비용구분/대분류/세부항목/거래처] 합산 정리", help="이동 후 같아진 행들을 하나로 합산합니다."):
            agg_dict = {"전체(Sheet3 총액)": "sum"}
            for c in st.session_state.customers:
                agg_dict[c] = "sum"
            merged_df = st.session_state.dash_df.groupby(["비용구분", "대분류", "세부 항목", "거래처명"], as_index=False).agg(agg_dict)
            merged_df["row_id"] = merged_df.index
            st.session_state.dash_df = merged_df
            st.session_state.editor_version += 1
            st.success("합산 정리가 완료되었습니다.")
            st.rerun()

    # 엑셀 다운로드 (통합, 1.매출원가, 2.판관비, 원장전체 시트)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cogs_out = dash_df[dash_df["비용구분"] == "매출원가"][display_cols[1:]]
        sga_out = dash_df[dash_df["비용구분"] == "판관비"][display_cols[1:]]
        total_out = dash_df[display_cols[1:]]

        total_out.to_excel(writer, sheet_name="통합대시보드", index=False)
        cogs_out.to_excel(writer, sheet_name="1.매출원가", index=False)
        sga_out.to_excel(writer, sheet_name="2.판관비", index=False)
        st.session_state.raw_df.to_excel(writer, sheet_name="원장전체_매핑내역", index=False)

        for sname in ["통합대시보드", "1.매출원가", "2.판관비"]:
            ws = writer.sheets[sname]
            for col_idx in range(5, len(display_cols)):
                for row_idx in range(2, ws.max_row + 1):
                    ws.cell(row=row_idx, column=col_idx).number_format = '#,##0'

    output.seek(0)

    st.download_button(
        label="💾 완성된 [매출원가/판관비] 손익분석 엑셀 다운로드",
        data=output,
        file_name="거래처별_손익분석_매출원가_판관비.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )