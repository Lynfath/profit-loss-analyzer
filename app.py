import streamlit as st
import pandas as pd
import io
import openpyxl
import math
from decimal import Decimal, ROUND_CEILING
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="통합 거래처별 손익분석 시스템", layout="wide")
st.title("📊 [통합] 거래처별 손익분석 및 견적 산정 시스템")

# 조작 버튼 20% 축소 및 텍스트 맞춤 CSS 주입
st.markdown("""
<style>
/* 모든 조작 버튼 일괄 20% 축소 및 텍스트 맞춤 (행추가 크기에 100% 일치) */
div[data-testid="stButton"] button, 
button[kind="secondary"], 
button[kind="primary"] {
    padding: 0.22rem 0.6rem !important;
    font-size: 0.85rem !important;
    min-height: 1.85rem !important;
    height: auto !important;
    width: auto !important;
    line-height: 1.2 !important;
    white-space: nowrap !important;
}
</style>
""", unsafe_allow_html=True)

DEFAULT_CUSTOMERS = ["중국법인", "롯데케미칼", "지에스큐", "린코스", "케이스티파이", "마곡상인"]
STD_MAJORS = [
    "항공 운임", "도급사", "밀크런", "사내매출 조정", "기타", "카드 수수료", "센터 비용",
    "건물 감가상각비", "지게차 리스", "출입/보안 (에스원/캡스)", "파렛트 렌탈",
    "부서운영비/야특근식비", "수선비", "토지임차료", "전산 (EDI)", "소모품비 外",
    "통관수수료", "직접비", "공통비", "일반관리비"
]

def auto_col_width(header_text, sample_values=None):
    def get_len(s): return sum(2 if ord(c) > 128 else 1 for c in str(s or ""))
    max_len = get_len(header_text)
    if sample_values is not None:
        items = list(sample_values) if not isinstance(sample_values, list) else sample_values
        if len(items) > 50: items = items[:50]
        for val in items: max_len = max(max_len, get_len(val))
    return "small" if max_len <= 10 else ("medium" if max_len <= 22 else "large")

def find_col(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    for cand in candidates:
        for c in df.columns:
            if cand in str(c): return c
    return None

def clean_numeric_series(series, index=None):
    if series is None:
        return pd.Series([0.0] * (len(index) if index is not None else 0), index=index)
    s_str = series.astype(str).str.strip()
    s_str = s_str.str.replace(",", "", regex=False).str.replace("₩", "", regex=False).str.replace("원", "", regex=False).str.replace("$", "", regex=False)
    mask_paren = s_str.str.startswith("(") & s_str.str.endswith(")")
    s_str = s_str.str.replace("(", "", regex=False).str.replace(")", "", regex=False)
    vals = pd.to_numeric(s_str, errors="coerce").fillna(0.0)
    if mask_paren.any():
        vals.loc[mask_paren] = -vals.loc[mask_paren]
    return vals

if "customers" not in st.session_state: st.session_state.customers = list(DEFAULT_CUSTOMERS)
if "major_categories" not in st.session_state: st.session_state.major_categories = list(STD_MAJORS)
if "editor_version" not in st.session_state: st.session_state.editor_version = 0
if "quote_editor_version" not in st.session_state: st.session_state.quote_editor_version = 0
if "revenue_dict" not in st.session_state: st.session_state.revenue_dict = {c: 0.0 for c in st.session_state.customers}

# 업로드 가이드 및 표준 서식
with st.expander("ℹ️ **업로드 엑셀 파일 표준 양식 안내 및 샘플 다운로드**", expanded=False):
    st.markdown("""
    **원장 파일 필수 컬럼 구성 안내**  
    업로드할 엑셀/CSV 파일은 아래와 같은 헤더 컬럼 구조를 갖추어야 시스템이 자동으로 계정과 비용을 인식합니다.  
    - **`연도/월`**: 귀속 연월 (예: `2026/07`) | **`계정`**: 원장 계정과목명 | **`적요`**: 전표 세부 적요  
    - **`금액(현지 통화)`**: 발생 금액 (숫자) | **`거래처명`**: 매입 거래처명 | **`구분`**: 수출/수입/공통 | **`상세`**: 품목/라우팅
    """)
    sample_df = pd.DataFrame([
        {"연도/월": "2026/07", "계정": "글로벌(국내로컬비용)_운송(일반)", "적요": ".UPS KOREA 7월 운송비", "금액(현지 통화)": 2024303418, "거래처명": "유나이티드파슬서비스코리아(주)", "구분": "수출", "상세": "UPS"},
        {"연도/월": "2026/07", "계정": "글로벌(운송비)_항공운송비", "적요": "2026년 7월 DHL-롯데케미칼 수출입 내역서", "금액(현지 통화)": 244142275, "거래처명": "(주)디에이치엘코리아", "구분": "수출", "상세": "케미칼(운송)"},
        {"연도/월": "2026/07", "계정": "국내(위탁용역비)_기타위탁용역", "적요": "[특송수출ㅣ엠제이플렉스] 26년 7월 물류서비스", "금액(현지 통화)": 35700000, "거래처명": "(주)엠제이플렉스", "구분": "공통", "상세": "현장도급비용"}
    ])
    st.dataframe(sample_df, column_config={"금액(현지 통화)": st.column_config.NumberColumn(format="%,d")}, width="stretch", hide_index=True)
    sample_buf = io.BytesIO()
    with pd.ExcelWriter(sample_buf, engine="openpyxl") as writer:
        sample_df.to_excel(writer, sheet_name="원장_샘플양식", index=False)
        for r in range(2, writer.sheets["원장_샘플양식"].max_row + 1):
            writer.sheets["원장_샘플양식"].cell(row=r, column=4).number_format = '#,##0'
    sample_buf.seek(0)
    st.download_button(label="📥 표준 원장 엑셀 샘플 다운로드 (.xlsx)", data=sample_buf, file_name="원장_업로드_표준양식_샘플.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dl_sample_template")

uploaded_file = st.file_uploader("📂 원장 엑셀 파일 업로드 (.xlsx, .csv)", type=["xlsx", "csv"], key="main_file_uploader")

def render_major_category_manager(key_suffix):
    with st.expander("📂 **대분류 항목 관리 (추가 / 삭제)**", expanded=False):
        m_add, m_del = st.columns(2)
        with m_add:
            st.markdown("##### ➕ 새 대분류 추가")
            new_major = st.text_input("새 대분류명 입력", placeholder="예: 포장재비, 차량유지비 등", key=f"input_new_major_{key_suffix}")
            if st.button("대분류 추가", type="primary", key=f"btn_add_major_{key_suffix}"):
                c_major = new_major.strip()
                if not c_major: st.warning("대분류명을 입력해주세요.")
                elif c_major in st.session_state.major_categories: st.warning("이미 존재하는 대분류입니다.")
                else:
                    st.session_state.major_categories.append(c_major)
                    st.session_state.editor_version += 1
                    st.success(f"'{c_major}' 대분류가 추가되었습니다.")
                    st.rerun()
        with m_del:
            st.markdown("##### 🗑️ 기존 대분류 삭제")
            if st.session_state.major_categories:
                del_target = st.selectbox("삭제할 대분류 선택", options=["(선택안함)"] + st.session_state.major_categories, key=f"select_del_major_{key_suffix}")
                if st.button("선택한 대분류 삭제", type="secondary", key=f"btn_del_major_{key_suffix}"):
                    if del_target == "(선택안함)": st.warning("삭제할 대분류를 선택해주세요.")
                    else:
                        st.session_state.major_categories.remove(del_target)
                        if "dash_df" in st.session_state and "대분류" in st.session_state.dash_df.columns:
                            mask = st.session_state.dash_df["대분류"] == del_target
                            if mask.any(): st.session_state.dash_df.loc[mask, "대분류"] = "기타"
                        st.session_state.editor_version += 1
                        st.success(f"'{del_target}' 대분류가 삭제되었습니다.")
                        st.rerun()

def render_customer_manager(key_suffix):
    with st.expander("🏢 **거래처 컬럼 관리 (추가 / 삭제)**", expanded=False):
        c_add, c_del = st.columns(2)
        with c_add:
            st.markdown("##### ➕ 새 거래처 추가")
            new_c = st.text_input("새 거래처명 입력", placeholder="예: FEDEX, 에스원 등", key=f"input_new_cust_{key_suffix}")
            if st.button("거래처 추가", type="primary", key=f"btn_add_cust_{key_suffix}"):
                clean_c = new_c.strip()
                if not clean_c: st.warning("거래처명을 입력해주세요.")
                elif clean_c in st.session_state.customers: st.warning("이미 존재하는 거래처입니다.")
                else:
                    st.session_state.customers.append(clean_c)
                    if "dash_df" in st.session_state: st.session_state.dash_df[clean_c] = 0.0
                    st.session_state.revenue_dict[clean_c] = 0.0
                    st.session_state.editor_version += 1
                    st.success(f"'{clean_c}' 거래처 열이 추가되었습니다.")
                    st.rerun()
        with c_del:
            st.markdown("##### 🗑️ 기존 거래처 삭제")
            if st.session_state.customers:
                del_c = st.selectbox("삭제할 거래처 선택", options=["(선택안함)"] + st.session_state.customers, key=f"select_del_cust_{key_suffix}")
                if st.button("선택한 거래처 삭제", type="secondary", key=f"btn_del_cust_{key_suffix}"):
                    if del_c == "(선택안함)": st.warning("삭제할 거래처를 선택해주세요.")
                    else:
                        st.session_state.customers.remove(del_c)
                        if "dash_df" in st.session_state and del_c in st.session_state.dash_df.columns:
                            st.session_state.dash_df.drop(columns=[del_c], inplace=True)
                        if del_c in st.session_state.revenue_dict: del st.session_state.revenue_dict[del_c]
                        st.session_state.editor_version += 1
                        st.success(f"'{del_c}' 거래처가 삭제되었습니다.")
                        st.rerun()

@st.cache_data
def build_initial_dashboard(file, customers):
    df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
    col_acct = find_col(df, ["계정과목", "계정", "원장 계정", "원장계정", "계정명"])
    col_desc = find_col(df, ["적요", "원장 적요", "원장적요", "적요 (내용)", "내용"])
    col_amt = find_col(df, ["금액(현지 통화)", "금액", "원장 금액", "원장금액", "발생금액"])
    col_vendor = find_col(df, ["거래처명", "매입 거래처명", "매입거래처명", "거래처"])
    col_detail = find_col(df, ["상세", "원장 상세", "원장상세", "구분", "비고"])

    amt_s = clean_numeric_series(df[col_amt] if col_amt else None, index=df.index)
    acct_s = df[col_acct].fillna("미분류").astype(str).str.strip() if col_acct else pd.Series(["미분류"] * len(df), index=df.index)
    desc_s = df[col_desc].fillna("-").astype(str).str.strip() if col_desc else pd.Series(["-"] * len(df), index=df.index)
    vend_s = df[col_vendor].fillna("-").astype(str).str.strip() if col_vendor else pd.Series(["-"] * len(df), index=df.index)
    det_s = df[col_detail].fillna("-").astype(str).str.strip() if col_detail else pd.Series(["-"] * len(df), index=df.index)

    categories, majors = [], []
    sga_kws = ["판관", "운송경비", "관리비", "임차료", "토지", "감가상각", "AFB", "EDI", "전산", "보안", "에스원", "캡스", "지게차", "리스", "파렛트", "수선비", "부서운영", "식대", "광열", "전기", "수도", "재산세", "세금", "보험료", "소모품"]
    major_rules = [
        ("밀크런", ["밀크런", "브이투브이"]), ("사내매출 조정", ["사내매출", "CBE"]),
        ("항공 운임", ["항공", "UPS", "DHL", "EMS", "린코스", "SF"]), ("도급사", ["도급", "엠제이", "인터비즈", "교원"]),
        ("토지임차료", ["토지", "임차"]), ("전산 (EDI)", ["전산", "EDI", "케이씨넷"]),
        ("건물 감가상각비", ["감가상각", "AFB"]), ("지게차 리스", ["지게차", "리스"]),
        ("출입/보안 (에스원/캡스)", ["보안", "에스원", "캡스"]), ("파렛트 렌탈", ["파렛트"]),
        ("수선비", ["수선"]), ("부서운영비/야특근식비", ["부서운영", "식대"]),
        ("카드 수수료", ["카드"]), ("센터 비용", ["전기", "수도", "맥서브"])
    ]
    for a, d, dt in zip(acct_s, desc_s, det_s):
        comb = f"{a} {d} {dt}"
        categories.append("판관비" if any(kw in comb for kw in sga_kws) else "매출원가")
        m_assigned = "기타"
        for m_name, kws in major_rules:
            if any(k in comb for k in kws):
                m_assigned = m_name; break
        majors.append(m_assigned)

    res_df = pd.DataFrame({
        "비용구분": categories, "계정과목": acct_s, "대분류": majors,
        "적요": desc_s, "거래처명": vend_s, "상세": det_s, "전체(Sheet3 총액)": amt_s
    })
    res_df["선택"] = False
    for c in customers: res_df[c] = 0.0
    res_df["row_id"] = list(range(len(res_df)))
    return df, res_df

# =============================================================================
# [견적 산정 & 매출원가 산출 & 예상손익 분석] 헬퍼 및 엑셀 빌더 (Single Source of Truth)
# =============================================================================
def init_quote_state():
    # Existing sessions and prices keep the legacy cost-markup interpretation.
    if "q_pricing_basis" not in st.session_state:
        st.session_state.q_pricing_basis = "markup"
    if "quote_initialized" not in st.session_state:
        st.session_state.quote_initialized = True
        st.session_state.q_company = ""
        st.session_state.q_doc_no = "제 2026-1001호"
        st.session_state.q_receiver = "(주)에이치앤비글로벌"
        st.session_state.q_contact = "물류/SCM 담당자 귀하"
        st.session_state.q_title = "글로벌 특송 및 풀필먼트 서비스 제안 단가"
        st.session_state.q_category = "화장품 / 뷰티 류"
        st.session_state.q_date = "2026-10-01"
        st.session_state.q_sender = ""
        st.session_state.q_terms = "[샘플] 초기 50평 필수 셋팅(고객사 50평 운임 제공, 200만원/월). 물동량 증가 시 증가 평수는 양사 협의 결정."
        st.session_state.q_validity = "[샘플] 세금계산서 발행월 기준 익월 말일내 현금 지급"
        st.session_state.q_insurance = "[샘플] 보증보험 증권 제출 또는 예치금 예치 협의"
        st.session_state.q_indemnity = "[샘플] 택배 - BOX당 50만원 한도 배상 / 운영 - 인정감모율 초과분 (공급가액 기준)"
        st.session_state.q_center = "[샘플] 인천공항 자유무역지역 물류센터"
        st.session_state.q_notes = "[샘플] 부자재 화주사 제공 조건 / 견적 외 추가업무는 발생 시 양사 간 협의"

        st.session_state.q_vol_b2c = 2000
        st.session_state.q_multi_pcs = 500
        st.session_state.q_storage_pyeong = 50
        st.session_state.q_inbound_qty = 2000
        st.session_state.q_dim_qty = 2000
        st.session_state.q_customs_qty = 2000
        st.session_state.q_express_qty = 2000
        st.session_state.q_include_express = False
        st.session_state.q_include_infra = False
        st.session_state.q_express_margin = 8.5
        st.session_state.q_wh_target_margin = 15.0
        st.session_state.q_apply_sga = True
        st.session_state.q_sga_rate = 7.0
        st.session_state.q_sga_amt = 321650

        st.session_state.q_warehouse_df = pd.DataFrame([
            {"구분": "입고", "작업 내용": "검수(기본)", "단위": "건(기본)", "매입단가": 250.0, "제안단가": 300.0, "비고": "입하확인/수량 및 육안불량 확인/라벨부착"},
            {"구분": "입고", "작업 내용": "검수(상세)", "단위": "건(상세)", "매입단가": 650.0, "제안단가": 700.0, "비고": "내품 상세검수(Tag, 라벨 등) 필요시"},
            {"구분": "측정", "작업 내용": "체적측정", "단위": "PCS", "매입단가": 100.0, "제안단가": 110.0, "비고": "SKU별 최초 측정 및 유지관리"},
            {"구분": "출고", "작업 내용": "B2C 기본출고", "단위": "건", "매입단가": 450.0, "제안단가": 450.0, "비고": "피킹, 패킹, 출고(운송장/서류 부착 포함)"},
            {"구분": "출고", "작업 내용": "합포장", "단위": "PCS", "매입단가": 0.0, "제안단가": 150.0, "비고": "2PC 이상 포장 시 2PCS부터 추가 정산"},
            {"구분": "출고", "작업 내용": "B2B 기본출고", "단위": "건", "매입단가": 1000.0, "제안단가": 1500.0, "비고": "출고지시서 기반 피킹/패킹/중량검수"},
            {"구분": "보관", "작업 내용": "작업스테이션(*)", "단위": "평", "매입단가": 0.0, "제안단가": 40000.0, "비고": "초기 50평 필수(고객사 50평 운임 제공, 200만원/월)"},
            {"구분": "VAS", "작업 내용": "양품화", "단위": "PCS", "매입단가": 1000.0, "제안단가": 1200.0, "비고": "재포장, 라벨 재부착 등"},
            {"구분": "VAS", "작업 내용": "동봉작업", "단위": "PCS", "매입단가": 20.0, "제안단가": 30.0, "비고": "별도 요청 첨부 인쇄물 동봉"},
            {"구분": "수출 신고", "작업 내용": "간이수출신고", "단위": "건", "매입단가": 200.0, "제안단가": 400.0, "비고": "관세청 전산 간이수출신고 대행"},
            {"구분": "수출 신고", "작업 내용": "일반수출신고", "단위": "건", "매입단가": 10000.0, "제안단가": 15000.0, "비고": "관세사 연계 일반수출신고 건당"}
        ])

        st.session_state.q_express_df = pd.DataFrame([
            {"국가": "미국 (US)", "기준구간": "≤ 2.0kg", "매입원가": 18055.0, "판매 제안단가": 19590.0, "발송비중(%)": 45.0, "비고": "UPS/FedEx/DHL 2kg 이하 특송"},
            {"국가": "이태리 (IT)", "기준구간": "≤ 2.0kg", "매입원가": 18786.0, "판매 제안단가": 20380.0, "발송비중(%)": 20.0, "비고": "유럽 주요국 특송"},
            {"국가": "프랑스 (FR)", "기준구간": "≤ 2.0kg", "매입원가": 17695.0, "판매 제안단가": 19200.0, "발송비중(%)": 15.0, "비고": "유럽 주요국 특송"},
            {"국가": "일본 (JP)", "기준구간": "≤ 2.0kg", "매입원가": 12850.0, "판매 제안단가": 13940.0, "발송비중(%)": 10.0, "비고": "아시아 근거리 특송"},
            {"국가": "싱가포르 (SG)", "기준구간": "≤ 2.0kg", "매입원가": 15700.0, "판매 제안단가": 17030.0, "발송비중(%)": 5.0, "비고": "동남아 허브 특송"},
            {"국가": "기타 국가 평균", "기준구간": "≤ 2.0kg", "매입원가": 18179.0, "판매 제안단가": 19720.0, "발송비중(%)": 5.0, "비고": "전세계 기타 권역 평균"}
        ])

        st.session_state.q_infra_df = pd.DataFrame([
            {"구분": "전산/EDI", "작업 내용": "전산(EDI) 비용", "단위": "건", "매입단가": 447.0, "비고": "수입:수출 9:1 배부 후 수출 물량 1/N 배부"},
            {"구분": "임차료", "작업 내용": "C3 토지임차료", "단위": "평/월", "매입단가": 9000.0, "비고": "월 2,700만원 / 3,000평 중 화주 50평 기준"},
            {"구분": "감가상각", "작업 내용": "건물/설비 감가상각비", "단위": "평/월", "매입단가": 5230.0, "비고": "설비/건물 상각비 1,569만원 / 3,000평 중 50평 기준"},
            {"구분": "보험료", "작업 내용": "화재보험료", "단위": "평/월", "매입단가": 1311.0, "비고": "센터 화재보험료 평당 월 배부 단가"},
            {"구분": "현장운영", "작업 내용": "소모품비 外", "단위": "식/월", "매입단가": 30000.0, "비고": "센터 현장 운영 기본 소모품 월 고정 배부액"}
        ])

PRICING_BASES = {"markup": "원가 가산율 (기존 방식)", "gross_margin": "목표 매출총이익률"}

def proposal_price(cost, rate, basis="markup"):
    """Return a scalar or Series of VAT-exclusive proposal prices in KRW.

    Legacy markup keeps its original nearest-10-won rounding. Gross margin
    rounds UP to 10 won so rounding cannot reduce a positive-cost target.
    This is a unit-price target before separately allocated infra and SG&A.
    """
    rate = float(rate)
    if basis not in PRICING_BASES or not math.isfinite(rate) or rate < 0:
        raise ValueError("단가 산정 기준과 비율을 확인해주세요.")
    if basis == "gross_margin" and rate >= 100:
        raise ValueError("목표 매출총이익률은 100% 미만이어야 합니다.")
    if basis == "markup":
        price = cost * (1.0 + rate / 100.0)
        return price.round(-1) if isinstance(cost, pd.Series) else round(price, -1)
    if isinstance(cost, pd.Series):
        return cost.map(lambda value: proposal_price(value, rate, basis))
    if not math.isfinite(float(cost)) or cost < 0:
        raise ValueError("매입원가는 0 이상의 유한한 값이어야 합니다.")
    price = Decimal(str(cost)) / (Decimal(1) - Decimal(str(rate)) / 100)
    return float((price / 10).to_integral_value(rounding=ROUND_CEILING) * 10)

def sync_quote_input(widget_key, state_key):
    # Callbacks run before all tabs, so earlier download buttons use fresh inputs.
    st.session_state[state_key] = st.session_state[widget_key]

def render_pricing_basis(tab_key):
    key = f"pricing_basis_{tab_key}"
    st.session_state[key] = st.session_state.q_pricing_basis
    st.radio("자동 단가 산정 기준 (탭 4·5·6 공통)", list(PRICING_BASES),
             format_func=PRICING_BASES.get, horizontal=True, key=key,
             on_change=sync_quote_input, args=(key, "q_pricing_basis"))
    st.caption("원가 가산율 = (판매가−원가)÷원가. 매출총이익률 = (판매가−원가)÷판매가. "
               "기준 선택만으로 기존 단가는 변경되지 않습니다. 재계산 버튼 또는 탭 6 특송 비율 변경 시 적용됩니다.")
    if st.session_state.q_pricing_basis == "gross_margin":
        st.caption("판매가 = 매입원가 ÷ (1−목표율/100), 10원 단위 올림. "
                   "개별 매입단가 기준이며 별도 인프라·일반관리비 차감 후 영업이익률은 달라집니다. "
                   "창고 매입단가 0원 항목은 기존 제안단가를 유지합니다.")
    else:
        st.caption("판매가 = 매입원가 × (1+가산율/100), 기존과 동일하게 10원 단위 반올림.")
    return "매출총이익률" if st.session_state.q_pricing_basis == "gross_margin" else "원가 가산율"

def get_quote_vol_dict():
    return {
        "검수(기본)": st.session_state.get("q_inbound_qty", 2000), "검수(상세)": 0,
        "체적측정": st.session_state.get("q_dim_qty", 2000),
        "B2C 기본출고": st.session_state.get("q_vol_b2c", 2000), "합포장": st.session_state.get("q_multi_pcs", 500),
        "B2B 기본출고": 0, "작업스테이션(*)": st.session_state.get("q_storage_pyeong", 50),
        "양품화": 0, "동봉작업": 0, "간이수출신고": st.session_state.get("q_customs_qty", 2000), "일반수출신고": 0
    }

def get_infra_vol_for_row(unit_str, item_name=""):
    u = str(unit_str).strip()
    if "평" in u:
        return st.session_state.get("q_storage_pyeong", 50)
    elif "건" in u:
        return st.session_state.get("q_vol_b2c", 2000)
    elif "PCS" in u.upper():
        return st.session_state.get("q_dim_qty", 2000)
    return 1

def calc_infra_total_cost(infra_df):
    tot = 0.0
    for _, r in infra_df.iterrows():
        q = get_infra_vol_for_row(r.get("단위", ""), r.get("작업 내용", ""))
        tot += float(r.get("매입단가", 0.0)) * q
    return tot

def build_quote_workbook(meta, wh_df, exp_df, exp_margin, vol_dict, include_exp=False, infra_df=None, include_infra=False):
    wb = openpyxl.Workbook()

    f_company = Font("맑은 고딕", size=18, bold=True, color="DC2626")
    f_title = Font("맑은 고딕", size=14, bold=True)
    f_subtitle = Font("맑은 고딕", size=11, bold=True, color="1E3A8A")
    f_sub = Font("맑은 고딕", size=11, bold=True)
    f_bold = Font("맑은 고딕", size=10, bold=True)
    f_body = Font("맑은 고딕", size=10)
    f_note = Font("맑은 고딕", size=9, color="64748B", italic=True)
    f_jeongyeol = Font("맑은 고딕", size=10, bold=True, color="DC2626")

    fill_h = PatternFill("solid", fgColor="F1F5F9")
    fill_appr = PatternFill("solid", fgColor="F8FAFC")
    fill_tot = PatternFill("solid", fgColor="E2E8F0")

    b_thin = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
                    top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))
    b_appr = Border(left=Side(style='thin', color='94A3B8'), right=Side(style='thin', color='94A3B8'),
                    top=Side(style='thin', color='94A3B8'), bottom=Side(style='thin', color='94A3B8'))
    b_bottom_thick = Border(bottom=Side(style='medium', color='0F172A'))

    def set_c(ws, coord, val, font=f_body, align=Alignment("center", "center"), num_fmt=None, fill=None, border=b_thin):
        c = ws[coord]
        if type(c).__name__ != 'MergedCell':
            c.value, c.font, c.alignment = val, font, align
            if num_fmt: c.number_format = num_fmt
        if border is not None: c.border = border
        if fill is not None: c.fill = fill
        return c

    def merge_and_format(ws, range_str, val=None, font=f_body, align=Alignment("center", "center"), fill=None, border=b_thin, num_fmt=None):
        ws.merge_cells(range_str)
        first = True
        for row in ws[range_str]:
            for cell in row:
                if fill is not None: cell.fill = fill
                if border is not None: cell.border = border
                if first and type(cell).__name__ != 'MergedCell':
                    if val is not None: cell.value = val
                    cell.font, cell.alignment = font, align
                    if num_fmt: cell.number_format = num_fmt
                    first = False

    ws1 = wb.active
    ws1.title = "제안서"
    ws1.views.sheetView[0].showGridLines = True

    ws1.column_dimensions['A'].width = 3
    ws1.column_dimensions['B'].width = 12
    ws1.column_dimensions['C'].width = 24
    ws1.column_dimensions['D'].width = 13
    ws1.column_dimensions['E'].width = 17
    ws1.column_dimensions['F'].width = 22
    ws1.column_dimensions['G'].width = 24

    r = 2
    if meta.get("company"):
        merge_and_format(ws1, f"B{r}:G{r}", str(meta["company"]), font=f_company, align=Alignment("center", "center"), border=None)
        ws1.row_dimensions[r].height = 32
        r += 2

    set_c(ws1, f"B{r}", f"문서번호 : {meta.get('doc_no', '')}", font=f_note, align=Alignment("left", "center"), border=None)
    set_c(ws1, f"G{r}", f"제안일자 : {meta.get('date', '')}", font=f_note, align=Alignment("right", "center"), border=None)
    for col_l in ["B", "C", "D", "E", "F", "G"]:
        ws1[f"{col_l}{r}"].border = b_bottom_thick
    r += 2

    merge_and_format(ws1, f"B{r}:C{r}", f"수   신 : {meta.get('receiver', '')}", font=f_body, align=Alignment("left", "center"), border=None)
    merge_and_format(ws1, f"B{r+1}:C{r+1}", f"참   조 : {meta.get('contact', '')}", font=f_body, align=Alignment("left", "center"), border=None)
    merge_and_format(ws1, f"B{r+2}:C{r+2}", f"제   목 : {meta.get('title', '')}", font=f_subtitle, align=Alignment("left", "center"), border=None)

    merge_and_format(ws1, f"D{r}:D{r+1}", "결\n재", font=f_bold, align=Alignment("center", "center", wrap_text=True), fill=fill_appr, border=b_appr)
    for col_idx, h_txt in [("E", "담 당"), ("F", "팀 장"), ("G", "부문장")]:
        set_c(ws1, f"{col_idx}{r}", h_txt, font=f_body, align=Alignment("center", "center"), fill=fill_appr, border=b_appr)
    set_c(ws1, f"E{r+1}", "", border=b_appr)
    set_c(ws1, f"F{r+1}", "", border=b_appr)
    set_c(ws1, f"G{r+1}", "[전결]", font=f_jeongyeol, align=Alignment("center", "center"), border=b_appr)
    ws1.row_dimensions[r+1].height = 30
    r += 3

    set_c(ws1, f"B{r}", "1. 귀사의 무궁한 발전을 기원합니다.", font=f_body, align=Alignment("left", "center"), border=None)
    r += 1
    set_c(ws1, f"B{r}", "2. 글로벌 특송 제안 단가를 아래와 같이 송부 드립니다.", font=f_body, align=Alignment("left", "center"), border=None)
    r += 2

    merge_and_format(ws1, f"B{r}:G{r}", "-   아       래   -", font=f_sub, align=Alignment("center", "center"), border=None)
    r += 2

    set_c(ws1, f"B{r}", f"1) 취급품목 : {meta.get('category', '')}", font=f_sub, align=Alignment("left", "center"), border=None)
    r += 2

    set_c(ws1, f"B{r}", "2) 요율표 (단위 : 원, VAT 별도)", font=f_sub, align=Alignment("left", "center"), border=None)
    r += 1
    set_c(ws1, f"B{r}", "(1) 창고 풀필먼트 운영", font=f_bold, align=Alignment("left", "center"), border=None)
    r += 1

    set_c(ws1, f"B{r}", "구 분", f_bold, fill=fill_h)
    set_c(ws1, f"C{r}", "작업 내용", f_bold, fill=fill_h)
    set_c(ws1, f"D{r}", "기본 단위", f_bold, fill=fill_h)
    set_c(ws1, f"E{r}", "제안 단가", f_bold, fill=fill_h)
    merge_and_format(ws1, f"F{r}:G{r}", "비 고", f_bold, fill=fill_h)
    r += 1

    for _, row in wh_df.iterrows():
        set_c(ws1, f"B{r}", str(row["구분"]))
        set_c(ws1, f"C{r}", str(row["작업 내용"]), align=Alignment("left", "center"))
        set_c(ws1, f"D{r}", str(row["단위"]))
        set_c(ws1, f"E{r}", float(row["제안단가"]), align=Alignment("right", "center"), num_fmt="#,##0")
        merge_and_format(ws1, f"F{r}:G{r}", str(row.get("비고", "")), align=Alignment("left", "center"))
        r += 1

    merge_and_format(ws1, f"B{r}:G{r}", str(meta.get("terms", "")), font=f_note, align=Alignment("left", "center"), border=None)
    r += 2

    set_c(ws1, f"B{r}", "(2) 해외 발송 특송 제안 단가표", font=f_bold, align=Alignment("left", "center"), border=None)
    r += 1

    set_c(ws1, f"B{r}", "국 가", f_bold, fill=fill_h)
    merge_and_format(ws1, f"C{r}:D{r}", "기준 구간", f_bold, fill=fill_h)
    set_c(ws1, f"E{r}", "판매 제안단가", f_bold, fill=fill_h)
    merge_and_format(ws1, f"F{r}:G{r}", "비 고", f_bold, fill=fill_h)
    r += 1

    for _, row in exp_df.iterrows():
        set_c(ws1, f"B{r}", str(row["국가"]))
        merge_and_format(ws1, f"C{r}:D{r}", str(row["기준구간"]))
        set_c(ws1, f"E{r}", float(row.get("판매 제안단가", 0)), align=Alignment("right", "center"), num_fmt="#,##0")
        merge_and_format(ws1, f"F{r}:G{r}", str(row.get("비고", "")), align=Alignment("left", "center"))
        r += 1

    r += 1
    set_c(ws1, f"B{r}", "3) 기본 계약 및 정산 조건", font=f_sub, align=Alignment("left", "center"), border=None)
    r += 1
    for cond_txt in [
        f"• (1) 대금 정산 : {meta.get('validity', '')}",
        f"• (2) 보증보험 또는 예치금 : {meta.get('insurance', '')}",
        f"• (3) 손해 배상 : {meta.get('indemnity', '')}",
        f"• (4) 기타 사항 : {meta.get('notes', '')}"
    ]:
        merge_and_format(ws1, f"B{r}:G{r}", cond_txt, font=f_body, align=Alignment("left", "center"), border=None)
        r += 1

    r += 2
    if meta.get("sender"):
        merge_and_format(ws1, f"B{r}:G{r}", str(meta["sender"]), font=Font("맑은 고딕", size=13, bold=True), align=Alignment("center", "center"), border=None)
        r += 1
        merge_and_format(ws1, f"B{r}:G{r}", "( 전자발송에 의한 직인생략 )", font=f_note, align=Alignment("center", "center"), border=None)

    ws1.page_setup.orientation = ws1.ORIENTATION_PORTRAIT
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0

    # Sheet 2: 매출원가산출표 (신규 탭 5 대응 시트)
    ws_cogs = wb.create_sheet(title="매출원가산출표")
    ws_cogs.views.sheetView[0].showGridLines = True
    r_name = meta.get("receiver", "화주사")
    ws_cogs['B2'], ws_cogs['B2'].font = f'▣ "{r_name}" 매출원가 산출 내역서', f_title
    ws_cogs['F2'], ws_cogs['F2'].font = "단위 : 원, VAT별도", f_note

    rc = 4
    set_c(ws_cogs, f"B{rc}", "1) 창고 풀필먼트 운영 매입 단가표", font=f_sub, align=Alignment("left", "center"), border=None)
    rc += 1
    for col_l, h in [("B", "구 분"), ("C", "작업 내용"), ("D", "기본 단위"), ("E", "매입 단가"), ("F", "작업 상세 및 비고")]:
        set_c(ws_cogs, f"{col_l}{rc}", h, f_bold, fill=fill_h)
    rc += 1
    for _, row in wh_df.iterrows():
        set_c(ws_cogs, f"B{rc}", str(row["구분"]))
        set_c(ws_cogs, f"C{rc}", str(row["작업 내용"]), align=Alignment("left", "center"))
        set_c(ws_cogs, f"D{rc}", str(row["단위"]))
        set_c(ws_cogs, f"E{rc}", float(row["매입단가"]), align=Alignment("right", "center"), num_fmt="#,##0")
        set_c(ws_cogs, f"F{rc}", str(row.get("비고", "")), align=Alignment("left", "center"))
        rc += 1

    rc += 1
    set_c(ws_cogs, f"B{rc}", "2) 해외 발송 특송 매입 단가표", font=f_sub, align=Alignment("left", "center"), border=None)
    rc += 1
    for col_l, h in [("B", "국 가"), ("C", "기준 구간"), ("D", "발송비중(%)"), ("E", "매입 단가"), ("F", "비 고")]:
        set_c(ws_cogs, f"{col_l}{rc}", h, f_bold, fill=fill_h)
    rc += 1
    for _, row in exp_df.iterrows():
        set_c(ws_cogs, f"B{rc}", str(row["국가"]))
        set_c(ws_cogs, f"C{rc}", str(row["기준구간"]))
        set_c(ws_cogs, f"D{rc}", float(row.get("발송비중(%)", 0.0)), align=Alignment("right", "center"), num_fmt="0.0")
        set_c(ws_cogs, f"E{rc}", float(row.get("매입원가", 0.0)), align=Alignment("right", "center"), num_fmt="#,##0")
        set_c(ws_cogs, f"F{rc}", str(row.get("비고", "")), align=Alignment("left", "center"))
        rc += 1

    if infra_df is not None and len(infra_df) > 0:
        rc += 1
        set_c(ws_cogs, f"B{rc}", "3) 센터 인프라 고정비 단가 산출 기준표", font=f_sub, align=Alignment("left", "center"), border=None)
        rc += 1
        for col_l, h in [("B", "구 분"), ("C", "작업 내용"), ("D", "기본 단위"), ("E", "매입 단가"), ("F", "작업 상세 및 비고")]:
            set_c(ws_cogs, f"{col_l}{rc}", h, f_bold, fill=fill_h)
        rc += 1
        for _, row in infra_df.iterrows():
            set_c(ws_cogs, f"B{rc}", str(row["구분"]))
            set_c(ws_cogs, f"C{rc}", str(row["작업 내용"]), align=Alignment("left", "center"))
            set_c(ws_cogs, f"D{rc}", str(row["단위"]))
            set_c(ws_cogs, f"E{rc}", float(row["매입단가"]), align=Alignment("right", "center"), num_fmt="#,##0")
            set_c(ws_cogs, f"F{rc}", str(row.get("비고", "")), align=Alignment("left", "center"))
            rc += 1

    for col in ws_cogs.columns:
        col_letter = get_column_letter(col[0].column)
        cells_to_measure = [c for c in col if not (col_letter == "B" and c.row in [2, 4, rc])]
        m_len = max((sum(2 if ord(ch) > 128 else 1 for ch in str(c.value or '')) for c in cells_to_measure), default=0)
        ws_cogs.column_dimensions[col_letter].width = max(m_len + 4, 12)

    # Sheet 3: 예상손익분석표
    ws2 = wb.create_sheet(title="예상손익분석표")
    ws2.views.sheetView[0].showGridLines = True
    ws2['B2'], ws2['B2'].font = f'▣ "{r_name}" 예상손익', f_title
    ws2['N2'], ws2['N2'].font = "단위 : 원/월, VAT별도", f_note

    for col_l, h in [("B","매출 구분"), ("C","작업 내용"), ("D","단위"), ("E","물량"), ("F","판매단가"), ("G","매출 금액"),
                     ("I","매입 구분"), ("J","단위"), ("K","물량"), ("L","매입단가"), ("M","매입 금액"),
                     ("N","월 영업이익"), ("O","이익률"), ("P","비고")]:
        set_c(ws2, f"{col_l}4", h, f_bold, fill=fill_h)

    r_s2, tot_rev, tot_cost = 5, 0.0, 0.0
    if include_exp:
        eqty = meta.get("express_qty", 2000)
        if len(exp_df) > 0 and "발송비중(%)" in exp_df.columns:
            tot_w = exp_df["발송비중(%)"].sum()
            abuy = (exp_df["매입원가"] * exp_df["발송비중(%)"]).sum() / tot_w if tot_w > 0 else float(exp_df["매입원가"].mean())
            asell = (exp_df["판매 제안단가"] * exp_df["발송비중(%)"]).sum() / tot_w if tot_w > 0 else float(exp_df["판매 제안단가"].mean())
        else:
            abuy, asell = 18179.0, 19720.0
        erev, ecost = eqty * asell, eqty * abuy
        eop = erev - ecost
        tot_rev += erev; tot_cost += ecost
        for col_l, v, a, nf in [("B","해외특송","center",None), ("C","저중량 구간 (≤2kg) 평균","left",None), ("D","건","center",None), ("E",eqty,"right","#,##0"), ("F",asell,"right","#,##0"), ("G",erev,"right","#,##0"), ("I","해외특송 매입운임","left",None), ("J","건","center",None), ("K",eqty,"right","#,##0"), ("L",abuy,"right","#,##0"), ("M",ecost,"right","#,##0"), ("N",eop,"right","#,##0"), ("O",(f"{(eop/erev*100):.1f}%" if erev > 0 else "0.0%"),"center",None), ("P","해외발송 특송 건당 정산","left",None)]:
            set_c(ws2, f"{col_l}{r_s2}", v, align=Alignment(a, "center"), num_fmt=nf)
        r_s2 += 1

    for _, row in wh_df.iterrows():
        iname = str(row["작업 내용"])
        vol = vol_dict.get(iname, 0)
        sp, bp = float(row["제안단가"]), float(row["매입단가"])
        rev, cost = vol * sp, vol * bp
        op = rev - cost
        tot_rev += rev; tot_cost += cost
        pct = f"{(op/rev*100):.1f}%" if rev > 0 else "0.0%"
        for col_l, v, a, nf in [("B",str(row["구분"]),"center",None), ("C",iname,"left",None), ("D",str(row["단위"]),"center",None), ("E",vol,"right","#,##0"), ("F",sp,"right","#,##0"), ("G",rev,"right","#,##0"), ("I",iname,"left",None), ("J",str(row["단위"]),"center",None), ("K",vol,"right","#,##0"), ("L",bp,"right","#,##0"), ("M",cost,"right","#,##0"), ("N",op,"right","#,##0"), ("O",pct,"center",None), ("P",str(row.get("비고","")),"left",None)]:
            set_c(ws2, f"{col_l}{r_s2}", v, align=Alignment(a, "center"), num_fmt=nf)
        r_s2 += 1

    if include_infra and infra_df is not None:
        for _, row in infra_df.iterrows():
            iname = str(row["작업 내용"])
            u_str = str(row["단위"])
            ivol = get_infra_vol_for_row(u_str, iname)
            ibp = float(row["매입단가"])
            icost = ivol * ibp
            tot_cost += icost
            for col_l, v, a, nf in [("B","인프라 고정비","center",None), ("C",iname,"left",None), ("D",u_str,"center",None), ("E",0,"right","#,##0"), ("F",0,"right","#,##0"), ("G",0,"right","#,##0"), ("I",iname,"left",None), ("J",u_str,"center",None), ("K",ivol,"right","#,##0"), ("L",ibp,"right","#,##0"), ("M",icost,"right","#,##0"), ("N",-icost,"right","#,##0"), ("O","-","center",None), ("P",str(row.get("비고","")),"left",None)]:
                set_c(ws2, f"{col_l}{r_s2}", v, align=Alignment(a, "center"), num_fmt=nf)
            r_s2 += 1

    sga = float(meta.get("sga_amt", 321650))
    sga_rate = float(meta.get("sga_rate", 7.0))
    if sga > 0:
        tot_cost += sga
        sga_label = f"SG&A 배부 ({sga_rate:.1f}%)"
        for col_l, v, a, nf in [("B","일반관리비","center",None), ("C",sga_label,"left",None), ("G",0,"right","#,##0"), ("I","일반관리비","left",None), ("M",sga,"right","#,##0"), ("N",-sga,"right","#,##0"), ("O",f"-{sga_rate:.1f}%","center",None), ("P",f"운영매출액 대비 {sga_rate:.1f}% 배부 기준","left",None)]:
            set_c(ws2, f"{col_l}{r_s2}", v, align=Alignment(a, "center"), num_fmt=nf)
        r_s2 += 1

    f_op = tot_rev - tot_cost
    f_margin = f"{(f_op/tot_rev*100):.1f}%" if tot_rev > 0 else "0.0%"
    
    cols_all = ["B","C","D","E","F","G","H","I","J","K","L","M","N","O","P"]
    for c_letter in cols_all:
        set_c(ws2, f"{c_letter}{r_s2}", "", font=f_bold, fill=fill_tot, border=b_thin)
    set_c(ws2, f"B{r_s2}", "매출액 합계", font=f_bold, align=Alignment("center", "center"), fill=fill_tot)
    set_c(ws2, f"G{r_s2}", tot_rev, font=f_bold, align=Alignment("right", "center"), num_fmt="#,##0", fill=fill_tot)
    set_c(ws2, f"I{r_s2}", "매입액 합계", font=f_bold, align=Alignment("center", "center"), fill=fill_tot)
    set_c(ws2, f"M{r_s2}", tot_cost, font=f_bold, align=Alignment("right", "center"), num_fmt="#,##0", fill=fill_tot)
    set_c(ws2, f"N{r_s2}", f_op, font=f_bold, align=Alignment("right", "center"), num_fmt="#,##0", fill=fill_tot)
    set_c(ws2, f"O{r_s2}", f_margin, font=f_bold, align=Alignment("center", "center"), fill=fill_tot)

    for col in ws2.columns:
        col_letter = get_column_letter(col[0].column)
        cells_to_measure = [c for c in col if not (col_letter == "B" and c.row == 2)]
        m_len = max((sum(2 if ord(ch) > 128 else 1 for ch in str(c.value or '')) for c in cells_to_measure), default=0)
        ws2.column_dimensions[col_letter].width = max(m_len + 3, 10)

    return wb

# =============================================================================
# [탭 4] 화주사 제출용 견적서
# =============================================================================
def render_tab4_proposal():
    init_quote_state()
    st.subheader("📄 화주사 제출용 견적서")
    rate_label = render_pricing_basis("tab4")
    st.caption("💡 화주사 제출용 정식 견적서 서식입니다. 제안 단가 및 작업 상세/비고는 **[탭 5] 매출원가 산출 및 [탭 6] 예상손익 분석표와 100% 자동 연동(Single Source of Truth)**됩니다.")

    with st.expander("📝 **제안서 기본 정보 (회사명 / 제목 / 수신처 / 특약 / 정산조건) 설정**", expanded=False):
        c_comp, c_title = st.columns([1, 2])
        st.session_state.q_company = c_comp.text_input("회사명 (당사명)", value=st.session_state.q_company, placeholder="예: (주)대한물류 (기재 시 A4 최상단 붉은색 대형 표기)", key="input_doc_company")
        st.session_state.q_title = c_title.text_input("제안서 제목", value=st.session_state.q_title, key="input_doc_title")

        c1, c2, c3 = st.columns(3)
        st.session_state.q_doc_no = c1.text_input("문서번호", value=st.session_state.q_doc_no, key="input_doc_no")
        st.session_state.q_date = c2.text_input("제안일자", value=st.session_state.q_date, key="input_doc_date")
        st.session_state.q_category = c3.text_input("취급품목", value=st.session_state.q_category, key="input_doc_category")

        c4, c5 = st.columns(2)
        st.session_state.q_receiver = c4.text_input("수신 (화주사명)", value=st.session_state.q_receiver, key="input_doc_receiver")
        st.session_state.q_contact = c5.text_input("참조 (담당자)", value=st.session_state.q_contact, key="input_doc_contact")

        st.session_state.q_sender = st.text_input("발신인 명의", value=st.session_state.q_sender, placeholder="예: 부서명 / 성명", key="input_doc_sender")

        st.markdown("##### 📌 계약 및 특약 사항 세부 설정")
        st.session_state.q_terms = st.text_input("운영 특약 사항", value=st.session_state.q_terms, key="input_doc_terms")
        tc1, tc2 = st.columns(2)
        st.session_state.q_validity = tc1.text_input("(1) 대금 정산", value=st.session_state.q_validity, key="input_doc_validity")
        st.session_state.q_insurance = tc2.text_input("(2) 보증보험 또는 예치금", value=st.session_state.q_insurance, key="input_doc_insurance")
        tc3, tc4 = st.columns(2)
        st.session_state.q_indemnity = tc3.text_input("(3) 손해 배상", value=st.session_state.q_indemnity, key="input_doc_indemnity")
        st.session_state.q_notes = tc4.text_input("(4) 기타 사항", value=st.session_state.q_notes, key="input_doc_notes")

    st.write("---")
    st.markdown("##### 📦 1) 창고 풀필먼트 운영 요율표 (단위: 원, VAT 별도)")
    st.caption("※ 아래 표에서 '제안 단가' 또는 '작업 상세 및 비고'를 수정하면 [탭 5] 매출원가 산출 및 [탭 6] 예상손익 분석표에 즉시 동기화됩니다.")
    edited_wh = st.data_editor(
        st.session_state.q_warehouse_df[["구분", "작업 내용", "단위", "제안단가", "비고"]],
        column_config={
            "구분": st.column_config.TextColumn(label="구분", width="small"),
            "작업 내용": st.column_config.TextColumn(label="작업 내용", width="medium", disabled=True),
            "단위": st.column_config.TextColumn(label="기본 단위", width="small"),
            "제안단가": st.column_config.NumberColumn(label="제안 단가 (판매)", format="%,d", width="medium"),
            "비고": st.column_config.TextColumn(label="작업 상세 및 비고", width="large")
        },
        width="stretch", hide_index=True, key=f"editor_proposal_wh_{st.session_state.quote_editor_version}"
    )
    wh_changed = False
    for idx, row in edited_wh.iterrows():
        for col_k in ["구분", "단위", "비고"]:
            new_v = str(row[col_k]) if pd.notna(row.get(col_k)) else ""
            old_v = str(st.session_state.q_warehouse_df.loc[idx, col_k])
            if new_v != old_v:
                st.session_state.q_warehouse_df.loc[idx, col_k] = new_v
                wh_changed = True
        new_p = float(row["제안단가"]) if pd.notna(row.get("제안단가")) else 0.0
        if abs(new_p - float(st.session_state.q_warehouse_df.loc[idx, "제안단가"])) > 1e-4:
            st.session_state.q_warehouse_df.loc[idx, "제안단가"] = new_p
            wh_changed = True
    if wh_changed:
        st.session_state.quote_editor_version += 1
        st.rerun()

    st.info(f"💡 **운영 특약 사항** : {st.session_state.q_terms}")
    st.write("---")

    st.markdown("##### ✈️ 2) 해외 발송 특송 제안 단가표")
    st.caption("※ '판매 제안단가'를 직접 수정하거나, 선택한 단가 산정 기준으로 일괄 재계산할 수 있습니다.")

    exp_tool1, exp_tool2 = st.columns([3, 1])
    with exp_tool1:
        pass
    with exp_tool2:
        with st.expander(f"🎯 {rate_label} 일괄 계산 (선택사항)", expanded=False):
            new_margin = st.slider(f"적용 {rate_label} (%)", min_value=1.0, max_value=30.0, value=float(st.session_state.q_express_margin), step=0.5, key=f"pop_margin_slider_{st.session_state.quote_editor_version}")
            if st.button(f"{rate_label}로 일괄 재계산", key="btn_recalc_express_margin"):
                st.session_state.q_express_margin = new_margin
                st.session_state.q_express_df["판매 제안단가"] = proposal_price(st.session_state.q_express_df["매입원가"], new_margin, st.session_state.q_pricing_basis)
                st.session_state.quote_editor_version += 1
                st.rerun()

    edited_exp = st.data_editor(
        st.session_state.q_express_df[["국가", "기준구간", "판매 제안단가", "비고"]],
        column_config={
            "국가": st.column_config.TextColumn(label="국가", disabled=True),
            "기준구간": st.column_config.TextColumn(label="기준 구간", disabled=True),
            "판매 제안단가": st.column_config.NumberColumn(label="판매 제안단가 (원)", format="%,d", min_value=0.0),
            "비고": st.column_config.TextColumn(label="비고")
        },
        width="stretch", hide_index=True, key=f"editor_proposal_exp_{st.session_state.quote_editor_version}"
    )
    exp_changed = False
    for idx, row in edited_exp.iterrows():
        new_sp = float(row["판매 제안단가"]) if pd.notna(row.get("판매 제안단가")) else 0.0
        if abs(new_sp - float(st.session_state.q_express_df.loc[idx, "판매 제안단가"])) > 1e-4:
            st.session_state.q_express_df.loc[idx, "판매 제안단가"] = new_sp
            exp_changed = True
        new_note = str(row["비고"]) if pd.notna(row.get("비고")) else ""
        if new_note != str(st.session_state.q_express_df.loc[idx, "비고"]):
            st.session_state.q_express_df.loc[idx, "비고"] = new_note
            exp_changed = True
    if exp_changed:
        st.session_state.quote_editor_version += 1
        st.rerun()

    st.write("---")

    st.markdown("##### 📋 3) 기본 계약 및 정산 조건")
    st.markdown(f"""
    - **• (1) 대금 정산** : {st.session_state.q_validity}  
    - **• (2) 보증보험 또는 예치금** : {st.session_state.q_insurance}  
    - **• (3) 손해 배상** : {st.session_state.q_indemnity}  
    - **• (4) 기타 사항** : {st.session_state.q_notes}
    """)
    st.write("---")

    with st.expander("🖨️ **화주사 전달용 A4 1장 인쇄 / PDF 출력 미리보기 (Print Preview)**", expanded=True):
        wh_rows = "".join(f"<tr><td style='border:1px solid #CBD5E1;padding:5px;text-align:center;'>{r['구분']}</td><td style='border:1px solid #CBD5E1;padding:5px;'>{r['작업 내용']}</td><td style='border:1px solid #CBD5E1;padding:5px;text-align:center;'>{r['단위']}</td><td style='border:1px solid #CBD5E1;padding:5px;text-align:right;font-weight:bold;'>{r['제안단가']:,.0f}원</td><td style='border:1px solid #CBD5E1;padding:5px;font-size:11px;'>{r['비고']}</td></tr>" for _, r in st.session_state.q_warehouse_df.iterrows())
        exp_rows = "".join(f"<tr><td style='border:1px solid #CBD5E1;padding:5px;text-align:center;'>{r['국가']}</td><td style='border:1px solid #CBD5E1;padding:5px;text-align:center;'>{r['기준구간']}</td><td style='border:1px solid #CBD5E1;padding:5px;text-align:right;font-weight:bold;'>{r['판매 제안단가']:,.0f}원</td><td style='border:1px solid #CBD5E1;padding:5px;font-size:11px;'>{r['비고']}</td></tr>" for _, r in st.session_state.q_express_df.iterrows())
        sender_html = f"<div style='font-size:15px; font-weight:bold; letter-spacing:2px;'>{st.session_state.q_sender}</div><div style='font-size:10px; color:#94A3B8; margin-top:2px;'>( 전자발송에 의한 직인생략 )</div>" if st.session_state.q_sender else "<div style='font-size:13px; color:#94A3B8;'>( 발신인 명의를 입력하세요 )</div>"
        company_header = f"<div style='color:#DC2626; font-size:24px; font-weight:bold; letter-spacing:3px; text-align:center; margin-bottom:16px;'>{st.session_state.q_company}</div>" if st.session_state.q_company else ""

        a4_preview = (
            f"<div style='background:#fff; color:#0f172a; padding:25px; font-family: Malgun Gothic, sans-serif; border:1px solid #94A3B8; max-width:800px; margin:auto;'>"
            f"{company_header}"
            f"<div style='display:flex; justify-content:space-between; border-bottom:2px solid #0f172a; padding-bottom:6px; margin-bottom:12px; font-size:12px; color:#64748B;'>"
            f"<div>문서번호 : {st.session_state.q_doc_no}</div><div>제안일자 : {st.session_state.q_date}</div>"
            f"</div>"
            f"<div style='display:flex; justify-content:space-between; margin-bottom:16px;'>"
            f"<div style='font-size:13px; line-height:1.7;'>"
            f"<div><strong>수 &nbsp; 신 :</strong> {st.session_state.q_receiver}</div>"
            f"<div><strong>참 &nbsp; 조 :</strong> {st.session_state.q_contact}</div>"
            f"<div><strong>제 &nbsp; 목 :</strong> <span style='font-size:15px; color:#1E3A8A; font-weight:bold;'>{st.session_state.q_title}</span></div>"
            f"</div>"
            f"<table style='border-collapse:collapse; width:170px; font-size:11px; text-align:center;'>"
            f"<tr><td rowspan='2' style='border:1px solid #94A3B8; width:25px; background:#F8FAFC; font-weight:bold;'>결<br>재</td><td style='border:1px solid #94A3B8; padding:3px; background:#F8FAFC;'>담 당</td><td style='border:1px solid #94A3B8; padding:3px; background:#F8FAFC;'>팀 장</td><td style='border:1px solid #94A3B8; padding:3px; background:#F8FAFC;'>부문장</td></tr>"
            f"<tr style='height:40px;'><td style='border:1px solid #94A3B8;'></td><td style='border:1px solid #94A3B8;'></td><td style='border:1px solid #94A3B8; color:#DC2626; font-weight:bold;'>[전결]</td></tr>"
            f"</table>"
            f"</div>"
            f"<div style='font-size:12px; line-height:1.5; margin-bottom:12px;'>1. 귀사의 무궁한 발전을 기원합니다.<br>2. 글로벌 특송 제안 단가를 아래와 같이 송부 드립니다.</div>"
            f"<div style='text-align:center; font-weight:bold; font-size:13px; margin:12px 0;'>- &nbsp; 아 &nbsp; &nbsp; &nbsp; 래 &nbsp; -</div>"
            f"<div style='font-size:12px; font-weight:bold; margin-bottom:6px;'>1) 취급품목 : {st.session_state.q_category}</div>"
            f"<div style='font-size:12px; font-weight:bold; margin-bottom:6px;'>2) 요율표 (단위 : 원, VAT 별도)</div>"
            f"<div style='font-size:11px; font-weight:bold; margin-bottom:4px;'>(1) 창고 풀필먼트 운영</div>"
            f"<table style='width:100%; border-collapse:collapse; font-size:11px; margin-bottom:6px;'>"
            f"<thead><tr style='background:#F1F5F9; font-weight:bold;'><th style='border:1px solid #CBD5E1; padding:5px; width:70px;'>구 분</th><th style='border:1px solid #CBD5E1; padding:5px; width:120px;'>작업 내용</th><th style='border:1px solid #CBD5E1; padding:5px; width:70px;'>기본 단위</th><th style='border:1px solid #CBD5E1; padding:5px; width:90px;'>제안 단가</th><th style='border:1px solid #CBD5E1; padding:5px;'>비 고</th></tr></thead>"
            f"<tbody>{wh_rows}</tbody>"
            f"</table>"
            f"<div style='font-size:10px; color:#475569; margin-bottom:12px; font-style:italic;'>{st.session_state.q_terms}</div>"
            f"<div style='font-size:11px; font-weight:bold; margin-bottom:4px;'>(2) 해외 발송 특송 제안 단가표</div>"
            f"<table style='width:100%; border-collapse:collapse; font-size:11px; margin-bottom:12px;'>"
            f"<thead><tr style='background:#F1F5F9; font-weight:bold;'><th style='border:1px solid #CBD5E1; padding:5px; width:110px;'>국 가</th><th style='border:1px solid #CBD5E1; padding:5px; width:90px;'>기준 구간</th><th style='border:1px solid #CBD5E1; padding:5px; width:110px;'>판매 제안단가</th><th style='border:1px solid #CBD5E1; padding:5px;'>비 고</th></tr></thead>"
            f"<tbody>{exp_rows}</tbody>"
            f"</table>"
            f"<div style='font-size:11px; font-weight:bold; margin-bottom:4px;'>3) 기본 계약 및 정산 조건</div>"
            f"<div style='font-size:11px; line-height:1.7; color:#334155; margin-bottom:18px;'>"
            f"• (1) 대금 정산 : {st.session_state.q_validity}<br>"
            f"• (2) 보증보험 또는 예치금 : {st.session_state.q_insurance}<br>"
            f"• (3) 손해 배상 : {st.session_state.q_indemnity}<br>"
            f"• (4) 기타 사항 : {st.session_state.q_notes}"
            f"</div>"
            f"<div style='text-align:center; padding-top:12px; border-top:1px dashed #CBD5E1;'>"
            f"{sender_html}"
            f"</div>"
            f"</div>"
        )
        st.markdown(a4_preview, unsafe_allow_html=True)

    vol_dict_cur = get_quote_vol_dict()
    wh_rev_current = sum(float(r["제안단가"]) * vol_dict_cur.get(str(r["작업 내용"]), 0) for _, r in st.session_state.q_warehouse_df.iterrows())
    st.session_state.q_sga_amt = round(wh_rev_current * (float(st.session_state.get("q_sga_rate", 7.0)) / 100.0))

    meta_d = {
        "company": st.session_state.q_company, "doc_no": st.session_state.q_doc_no, "date": st.session_state.q_date,
        "receiver": st.session_state.q_receiver, "contact": st.session_state.q_contact, "title": st.session_state.q_title,
        "category": st.session_state.q_category, "sender": st.session_state.q_sender, "terms": st.session_state.q_terms,
        "validity": st.session_state.q_validity, "insurance": st.session_state.q_insurance, "indemnity": st.session_state.q_indemnity,
        "center": st.session_state.q_center, "notes": st.session_state.q_notes, "express_qty": st.session_state.q_express_qty,
        "sga_rate": st.session_state.get("q_sga_rate", 7.0), "sga_amt": st.session_state.q_sga_amt
    }
    excel_wb = build_quote_workbook(
        meta_d, st.session_state.q_warehouse_df, st.session_state.q_express_df,
        st.session_state.q_express_margin, vol_dict_cur,
        include_exp=st.session_state.q_include_express,
        infra_df=st.session_state.q_infra_df,
        include_infra=st.session_state.get("q_include_infra", False)
    )
    buf = io.BytesIO()
    excel_wb.save(buf)
    buf.seek(0)
    st.download_button(label="📥 [화주사 제출용 견적서 · 매출원가 · 예상손익] 통합 엑셀 다운로드 (.xlsx)", data=buf, file_name=f"견적서_및_예상손익_{st.session_state.q_receiver}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dl_quote_tab4")

# =============================================================================
# [탭 5] 매출원가 산출 (신규 추가 탭)
# =============================================================================
def render_tab5_cogs_calc():
    init_quote_state()
    wh_df = st.session_state.q_warehouse_df
    exp_df = st.session_state.q_express_df
    infra_df = st.session_state.q_infra_df
    vol_dict = get_quote_vol_dict()

    st.subheader("🛠️ 매출원가 산출")
    rate_label = render_pricing_basis("tab5")
    st.caption("💡 **[탭 4] 견적서**의 작업 상세/비고 데이터와 100% 연동되며, 여기서 설정한 **매입 단가** 및 **센터 인프라 고정비**는 **[탭 6] 예상손익 분석표**의 원가로 실시간 반영됩니다.")

    # 상단 원가 요약 KPI 카드 (자체 추가 기능 1)
    wh_std_unit_cost = sum(float(r["매입단가"]) for _, r in wh_df[wh_df["작업 내용"].isin(["검수(기본)", "체적측정", "B2C 기본출고", "간이수출신고"])].iterrows())
    wh_total_cost = sum(float(r["매입단가"]) * vol_dict.get(str(r["작업 내용"]), 0) for _, r in wh_df.iterrows())
    tot_w = exp_df["발송비중(%)"].sum() if "발송비중(%)" in exp_df.columns else 0.0
    avg_exp_cost = ((exp_df["매입원가"] * exp_df["발송비중(%)"]).sum() / tot_w) if tot_w > 0 else float(exp_df["매입원가"].mean())
    infra_total_cost = calc_infra_total_cost(infra_df)

    ck1, ck2, ck3, ck4 = st.columns(4)
    ck1.metric("1. 풀필먼트 표준 건당 원가", f"{wh_std_unit_cost:,.0f}원/건", help="검수(기본) + 체적측정 + B2C 기본출고 + 간이수출신고 1건당 매입단가 합계")
    ck2.metric("2. 풀필먼트 월 예상 매입원가", f"{wh_total_cost:,.0f}원/월", help="현재 설정된 월 예상 물동량 기준 창고 운영 매입원가 총액")
    ck3.metric("3. 해외특송 가중평균 매입단가", f"{avg_exp_cost:,.0f}원/건", delta=f"비중 합계: {tot_w:.1f}%" if abs(tot_w - 100.0) < 0.1 else f"비중 점검 필요 ({tot_w:.1f}%)", delta_color="normal" if abs(tot_w - 100.0) < 0.1 else "inverse")
    ck4.metric("4. 센터 인프라 고정비 합계", f"{infra_total_cost:,.0f}원/월", help=f"보관 {st.session_state.q_storage_pyeong}평 및 B2C {st.session_state.q_vol_b2c:,}건 기준 월 고정 원가")

    st.write("---")

    # 1) 창고 풀필먼트 운영 요율표 (매입 단가)
    wh_head_l, wh_head_r = st.columns([3, 1])
    with wh_head_l:
        st.markdown("##### 📦 1) 창고 풀필먼트 운영 요율표 (매입 단가 / 단위: 원, VAT 별도)")
        st.caption("※ '작업 상세 및 비고'는 [탭 4] 견적서에서 자동 연동(수정 불가)되며, 이 표에서는 '매입 단가'만 수정하여 [탭 6] 예상손익에 반영합니다.")
    with wh_head_r:
        with st.expander(f"🎯 {rate_label}로 제안단가 계산", expanded=False):
            target_m = st.slider(f"창고 {rate_label} (%)", min_value=0.0, max_value=50.0, value=float(st.session_state.get("q_wh_target_margin", 15.0)), step=1.0, key="slider_wh_target_margin")
            if st.button("매입단가 기준 제안단가 일괄 반영", key="btn_apply_wh_markup"):
                st.session_state.q_wh_target_margin = target_m
                for idx_r, r_wh in st.session_state.q_warehouse_df.iterrows():
                    bp_v = float(r_wh["매입단가"])
                    if bp_v > 0:
                        st.session_state.q_warehouse_df.loc[idx_r, "제안단가"] = proposal_price(bp_v, target_m, st.session_state.q_pricing_basis)
                st.session_state.quote_editor_version += 1
                st.success(f"{rate_label} {target_m:.0f}% 기준 단가가 [탭 4]에 반영되었습니다.")
                st.rerun()

    edited_cogs_wh = st.data_editor(
        st.session_state.q_warehouse_df[["구분", "작업 내용", "단위", "매입단가", "비고"]],
        column_config={
            "구분": st.column_config.TextColumn(label="구분", width="small", disabled=True),
            "작업 내용": st.column_config.TextColumn(label="작업 내용", width="medium", disabled=True),
            "단위": st.column_config.TextColumn(label="기본 단위", width="small", disabled=True),
            "매입단가": st.column_config.NumberColumn(label="매입 단가 (원가)", format="%,d", min_value=0.0, width="medium", disabled=False),
            "비고": st.column_config.TextColumn(label="작업 상세 및 비고 ([탭 4] 견적서 연동)", width="large", disabled=True)
        },
        width="stretch", hide_index=True, key=f"editor_cogs_wh_{st.session_state.quote_editor_version}"
    )
    cogs_wh_changed = False
    for idx, row in edited_cogs_wh.iterrows():
        new_bp = float(row["매입단가"]) if pd.notna(row.get("매입단가")) else 0.0
        if abs(new_bp - float(st.session_state.q_warehouse_df.loc[idx, "매입단가"])) > 1e-4:
            st.session_state.q_warehouse_df.loc[idx, "매입단가"] = new_bp
            cogs_wh_changed = True
    if cogs_wh_changed:
        st.session_state.quote_editor_version += 1
        st.rerun()

    st.write("---")

    # 2) 해외 발송 특송 매입 단가표
    exp_head_l, exp_head_r = st.columns([3, 1])
    with exp_head_l:
        st.markdown("##### ✈️ 2) 해외 발송 특송 매입 단가표 (단위: 원, VAT 별도)")
        st.caption("※ '작업 상세 및 비고'는 [탭 4] 견적서에서 자동 연동(수정 불가)되며, 국가별 '매입 단가' 및 '발송 비중(%)'만 수정합니다.")
    with exp_head_r:
        with st.expander("🎯 매입원가 기준 제안단가 동기화", expanded=False):
            sync_m = st.slider(f"특송 {rate_label} (%)", min_value=1.0, max_value=30.0, value=float(st.session_state.q_express_margin), step=0.5, key=f"slider_cogs_exp_margin_{st.session_state.quote_editor_version}")
            if st.button("특송 제안단가 일괄 재계산", key="btn_cogs_recalc_exp"):
                st.session_state.q_express_margin = sync_m
                st.session_state.q_express_df["판매 제안단가"] = proposal_price(st.session_state.q_express_df["매입원가"], sync_m, st.session_state.q_pricing_basis)
                st.session_state.quote_editor_version += 1
                st.success(f"특송 {rate_label} {sync_m:.1f}% 기준 단가가 [탭 4]에 반영되었습니다.")
                st.rerun()

    edited_cogs_exp = st.data_editor(
        st.session_state.q_express_df[["국가", "기준구간", "매입원가", "발송비중(%)", "비고"]],
        column_config={
            "국가": st.column_config.TextColumn(label="국가", width="small", disabled=True),
            "기준구간": st.column_config.TextColumn(label="기준 구간", width="small", disabled=True),
            "매입원가": st.column_config.NumberColumn(label="매입 단가 (원)", format="%,d", min_value=0.0, width="medium", disabled=False),
            "발송비중(%)": st.column_config.NumberColumn(label="발송 비중 (%)", format="%.1f%%", min_value=0.0, max_value=100.0, width="small", disabled=False),
            "비고": st.column_config.TextColumn(label="작업 상세 및 비고 ([탭 4] 견적서 연동)", width="large", disabled=True)
        },
        width="stretch", hide_index=True, key=f"editor_cogs_exp_{st.session_state.quote_editor_version}"
    )
    cogs_exp_changed = False
    for idx, row in edited_cogs_exp.iterrows():
        new_bp = float(row["매입원가"]) if pd.notna(row.get("매입원가")) else 0.0
        if abs(new_bp - float(st.session_state.q_express_df.loc[idx, "매입원가"])) > 1e-4:
            st.session_state.q_express_df.loc[idx, "매입원가"] = new_bp
            cogs_exp_changed = True
        new_w = float(row["발송비중(%)"]) if pd.notna(row.get("발송비중(%)")) else 0.0
        if abs(new_w - float(st.session_state.q_express_df.loc[idx, "발송비중(%)"])) > 1e-4:
            st.session_state.q_express_df.loc[idx, "발송비중(%)"] = new_w
            cogs_exp_changed = True
    if cogs_exp_changed:
        st.session_state.quote_editor_version += 1
        st.rerun()

    st.write("---")

    # 3) 센터 인프라 고정비 단가 산출 기준 안내 (1) 창고 풀필먼트 운영 요율표와 100% 동일한 형식)
    inf_head_l, inf_head_r = st.columns([2.2, 1.8])
    with inf_head_l:
        st.markdown("##### 🏢 3) 센터 인프라 고정비 단가 산출 기준표 (단위: 원, VAT 별도)")
        st.caption("※ 1) 창고 풀필먼트 운영 요율표와 동일한 규격으로 센터 고정비 원가 단가 및 산출 기준을 관리합니다.")
    with inf_head_r:
        inf_b1, inf_b2, inf_b3 = st.columns([1.1, 1.1, 1.8])
        with inf_b1:
            if st.button("+ 항목 추가", key="btn_add_infra_row"):
                new_inf = pd.DataFrame([{"구분": "기타고정비", "작업 내용": "신규 인프라 항목", "단위": "식/월", "매입단가": 0.0, "비고": "산출 기준 입력"}])
                st.session_state.q_infra_df = pd.concat([st.session_state.q_infra_df, new_inf], ignore_index=True)
                st.session_state.quote_editor_version += 1
                st.rerun()
        with inf_b2:
            if st.button("- 마지막 행 삭제", key="btn_del_infra_row"):
                if len(st.session_state.q_infra_df) > 1:
                    st.session_state.q_infra_df = st.session_state.q_infra_df.iloc[:-1].reset_index(drop=True)
                    st.session_state.quote_editor_version += 1
                    st.rerun()
                else:
                    st.warning("최소 1개의 항목은 유지해야 합니다.")
        with inf_b3:
            chk_inf_5 = st.checkbox("📊 [탭 6] 예상손익에 고정비 합산", value=st.session_state.get("q_include_infra", False), key=f"chk_tab5_include_infra_{st.session_state.quote_editor_version}")
            if chk_inf_5 != st.session_state.get("q_include_infra", False):
                st.session_state.q_include_infra = chk_inf_5
                st.session_state.quote_editor_version += 1
                st.rerun()

    edited_infra = st.data_editor(
        st.session_state.q_infra_df[["구분", "작업 내용", "단위", "매입단가", "비고"]],
        column_config={
            "구분": st.column_config.TextColumn(label="구분", width="small"),
            "작업 내용": st.column_config.TextColumn(label="작업 내용", width="medium"),
            "단위": st.column_config.TextColumn(label="기본 단위", width="small"),
            "매입단가": st.column_config.NumberColumn(label="매입 단가 (원가)", format="%,d", min_value=0.0, width="medium"),
            "비고": st.column_config.TextColumn(label="작업 상세 및 비고", width="large")
        },
        width="stretch", hide_index=True, key=f"editor_cogs_infra_{st.session_state.quote_editor_version}"
    )
    infra_changed = False
    for idx, row in edited_infra.iterrows():
        for col_k in ["구분", "작업 내용", "단위", "비고"]:
            new_v = str(row[col_k]) if pd.notna(row.get(col_k)) else ""
            old_v = str(st.session_state.q_infra_df.loc[idx, col_k])
            if new_v != old_v:
                st.session_state.q_infra_df.loc[idx, col_k] = new_v
                infra_changed = True
        new_bp = float(row["매입단가"]) if pd.notna(row.get("매입단가")) else 0.0
        if abs(new_bp - float(st.session_state.q_infra_df.loc[idx, "매입단가"])) > 1e-4:
            st.session_state.q_infra_df.loc[idx, "매입단가"] = new_bp
            infra_changed = True
    if infra_changed:
        st.session_state.quote_editor_version += 1
        st.rerun()

    st.write("---")

    # 4) 자체 추가 기능: [매입단가 vs 제안단가] 단가 마진 브릿지 및 역마진 점검표
    with st.expander("🔍 **4) [단가 검증] 매입단가 vs 제안단가 마진 브릿지 및 수익성 점검표 (자동 분석)**", expanded=True):
        st.caption("💡 [탭 5]의 매입 단가와 [탭 4]의 제안 단가를 실시간 비교하여 단위당 마진액과 역마진(적자) 위험 항목을 자동 진단합니다.")
        bridge_rows = []
        for _, r in st.session_state.q_warehouse_df.iterrows():
            bp = float(r["매입단가"])
            sp = float(r["제안단가"])
            unit_m = sp - bp
            m_rate = (unit_m / sp * 100.0) if sp > 0 else 0.0
            q_vol = vol_dict.get(str(r["작업 내용"]), 0)
            status_txt = "🟢 정상 마진" if unit_m > 0 else ("⚪ 손익분기(BEP)" if unit_m == 0 else "🔴 역마진 주의")
            bridge_rows.append({
                "구분": str(r["구분"]),
                "작업 내용": str(r["작업 내용"]),
                "기본 단위": str(r["단위"]),
                "매입 단가": bp,
                "제안 단가": sp,
                "단위당 마진": unit_m,
                "단가 마진율": f"{m_rate:.1f}%",
                "월 적용물량": q_vol,
                "월 예상 마진액": unit_m * q_vol,
                "수익성 진단": status_txt,
                "작업 상세 및 비고": str(r["비고"])
            })
        bridge_df = pd.DataFrame(bridge_rows)
        st.dataframe(
            bridge_df,
            column_config={
                "매입 단가": st.column_config.NumberColumn(format="%,d"),
                "제안 단가": st.column_config.NumberColumn(format="%,d"),
                "단위당 마진": st.column_config.NumberColumn(format="%,d"),
                "월 적용물량": st.column_config.NumberColumn(format="%,d"),
                "월 예상 마진액": st.column_config.NumberColumn(format="%,d")
            },
            width="stretch", hide_index=True
        )

    wh_rev_current = sum(float(r["제안단가"]) * vol_dict.get(str(r["작업 내용"]), 0) for _, r in st.session_state.q_warehouse_df.iterrows())
    st.session_state.q_sga_amt = round(wh_rev_current * (float(st.session_state.get("q_sga_rate", 7.0)) / 100.0))
    meta_d = {
        "company": st.session_state.q_company, "doc_no": st.session_state.q_doc_no, "date": st.session_state.q_date,
        "receiver": st.session_state.q_receiver, "contact": st.session_state.q_contact, "title": st.session_state.q_title,
        "category": st.session_state.q_category, "sender": st.session_state.q_sender, "terms": st.session_state.q_terms,
        "validity": st.session_state.q_validity, "insurance": st.session_state.q_insurance, "indemnity": st.session_state.q_indemnity,
        "center": st.session_state.q_center, "notes": st.session_state.q_notes, "express_qty": st.session_state.q_express_qty,
        "sga_rate": st.session_state.get("q_sga_rate", 7.0), "sga_amt": st.session_state.q_sga_amt
    }
    excel_wb = build_quote_workbook(
        meta_d, st.session_state.q_warehouse_df, st.session_state.q_express_df,
        st.session_state.q_express_margin, vol_dict,
        include_exp=st.session_state.q_include_express,
        infra_df=st.session_state.q_infra_df,
        include_infra=st.session_state.get("q_include_infra", False)
    )
    buf = io.BytesIO()
    excel_wb.save(buf)
    buf.seek(0)
    st.download_button(label="📥 [매출원가 산출 내역서 포함] 통합 견적·손익 엑셀 다운로드 (.xlsx)", data=buf, file_name=f"매출원가_견적_예상손익_{st.session_state.q_receiver}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dl_cogs_tab5")

# =============================================================================
# [탭 6] 예상손익 분석표 및 시뮬레이션 (기존 탭 5에서 이동)
# =============================================================================
def render_tab6_pnl_simulation():
    init_quote_state()
    wh_df = st.session_state.q_warehouse_df
    exp_df = st.session_state.q_express_df
    infra_df = st.session_state.q_infra_df
    q_rec = st.session_state.q_receiver
    st.subheader(f'📊 [예상손익] "{q_rec}" 예상손익 분석표 및 시뮬레이터')
    rate_label = render_pricing_basis("tab6")
    st.caption("단가 산정 비율은 개별 매입단가 기준입니다. 최종 이익률은 인프라 원가와 일반관리비를 차감한 영업이익률이며, 일반관리비는 창고 운영매출에만 배부됩니다.")
    st.caption("💡 **[탭 4] 제안 단가**와 **[탭 5] 매입 단가(매출원가)** 및 월 예상 물동량에 따른 [매출액 → 매입원가 → 영업이익 및 이익률]을 실시간 산출합니다.")

    with st.expander("⚙️ **1. 조건별 물동량 및 시뮬레이션 입력창 (Input Dashboard)**", expanded=True):
        sc1, sc2, sc3, sc4 = st.columns(4)
        st.session_state.q_vol_b2c = sc1.number_input("📦 B2C 출고 건수 (건/월)", min_value=0, max_value=1000000, value=st.session_state.q_vol_b2c, step=100, key="in_q_vol_b2c", on_change=sync_quote_input, args=("in_q_vol_b2c", "q_vol_b2c"))
        st.session_state.q_multi_pcs = sc2.number_input("🧩 합포장 물량 (PCS/월)", min_value=0, max_value=500000, value=st.session_state.q_multi_pcs, step=50, key="in_q_multi_pcs", on_change=sync_quote_input, args=("in_q_multi_pcs", "q_multi_pcs"))
        st.session_state.q_storage_pyeong = sc3.number_input("🏬 보관 평수 (평)", min_value=0, max_value=10000, value=st.session_state.q_storage_pyeong, step=5, key="in_q_storage_pyeong", on_change=sync_quote_input, args=("in_q_storage_pyeong", "q_storage_pyeong"))
        st.session_state.q_express_qty = sc4.number_input("✈️ 해외특송 발송 건수 (건/월)", min_value=0, max_value=500000, value=st.session_state.q_express_qty, step=100, key="in_q_express_qty", on_change=sync_quote_input, args=("in_q_express_qty", "q_express_qty"))
        sc5, sc6, sc7, sc8 = st.columns(4)
        st.session_state.q_inbound_qty = sc5.number_input("📥 입고/검수 물량 (건/월)", min_value=0, max_value=500000, value=st.session_state.q_inbound_qty, step=100, key="in_q_inbound_qty", on_change=sync_quote_input, args=("in_q_inbound_qty", "q_inbound_qty"))
        st.session_state.q_dim_qty = sc6.number_input("📏 체적측정 물량 (PCS/월)", min_value=0, max_value=500000, value=st.session_state.q_dim_qty, step=100, key="in_q_dim_qty", on_change=sync_quote_input, args=("in_q_dim_qty", "q_dim_qty"))
        st.session_state.q_customs_qty = sc7.number_input("📑 간이수출신고 (건/월)", min_value=0, max_value=500000, value=st.session_state.q_customs_qty, step=100, key="in_q_customs_qty", on_change=sync_quote_input, args=("in_q_customs_qty", "q_customs_qty"))
        
        cur_margin = sc8.slider(f"🎯 해외특송 {rate_label} (%)", min_value=1.0, max_value=30.0, value=float(st.session_state.q_express_margin), step=0.5, key=f"slider_tab6_margin_{st.session_state.quote_editor_version}")
        if cur_margin != st.session_state.q_express_margin:
            st.session_state.q_express_margin = cur_margin
            st.session_state.q_express_df["판매 제안단가"] = proposal_price(st.session_state.q_express_df["매입원가"], cur_margin, st.session_state.q_pricing_basis)
            st.session_state.quote_editor_version += 1
            st.rerun()

        if "q_sga_rate" not in st.session_state:
            st.session_state.q_sga_rate = 7.0

        c_opt1, c_opt2, c_opt3 = st.columns([1.4, 1.4, 1.2])
        st.session_state.q_include_express = c_opt1.checkbox("✈️ 해외특송(매출/원가)을 전체 손익표에 정식 합산", value=st.session_state.q_include_express, key="in_q_include_express", on_change=sync_quote_input, args=("in_q_include_express", "q_include_express"))
        chk_inf_6 = c_opt2.checkbox("🏢 센터 인프라 고정비를 손익표 매입원가에 합산", value=st.session_state.get("q_include_infra", False), key=f"in_q_include_infra_{st.session_state.quote_editor_version}")
        if chk_inf_6 != st.session_state.get("q_include_infra", False):
            st.session_state.q_include_infra = chk_inf_6
            st.session_state.quote_editor_version += 1
            st.rerun()
        st.session_state.q_sga_rate = c_opt3.number_input("🏢 일반관리비 (%)", min_value=0.0, max_value=50.0, value=float(st.session_state.q_sga_rate), step=0.5, format="%.1f", key="in_q_sga_rate", on_change=sync_quote_input, args=("in_q_sga_rate", "q_sga_rate"))

    st.write("---")
    vol_dict = get_quote_vol_dict()
    pnl_rows, tot_rev, tot_cost = [], 0.0, 0.0

    if st.session_state.q_include_express:
        exp_qty = st.session_state.q_express_qty
        if len(exp_df) > 0 and "발송비중(%)" in exp_df.columns:
            tot_w = exp_df["발송비중(%)"].sum()
            avg_buy = (exp_df["매입원가"] * exp_df["발송비중(%)"]).sum() / tot_w if tot_w > 0 else float(exp_df["매입원가"].mean())
            avg_sell = (exp_df["판매 제안단가"] * exp_df["발송비중(%)"]).sum() / tot_w if tot_w > 0 else float(exp_df["판매 제안단가"].mean())
        else:
            avg_buy, avg_sell = 18179.0, 19720.0
        exp_rev, exp_cost = exp_qty * avg_sell, exp_qty * avg_buy
        exp_op = exp_rev - exp_cost
        tot_rev += exp_rev; tot_cost += exp_cost
        pnl_rows.append({"구분": "해외특송", "작업 내용": "저중량 구간 (≤2kg) 평균", "단위": "건", "물량": exp_qty, "판매단가": avg_sell, "매출액": exp_rev, "매입단가": avg_buy, "매입원가": exp_cost, "영업이익": exp_op, "이익률": (f"{(exp_op/exp_rev*100):.1f}%" if exp_rev > 0 else "0.0%"), "비고": "해외발송 특송 건당 정산"})

    for idx, row in wh_df.iterrows():
        iname = str(row["작업 내용"])
        vol = vol_dict.get(iname, 0)
        sp, bp = float(row["제안단가"]), float(row["매입단가"])
        r_amt, c_amt = vol * sp, vol * bp
        op_amt = r_amt - c_amt
        tot_rev += r_amt; tot_cost += c_amt
        pct = f"{(op_amt/r_amt*100):.1f}%" if r_amt > 0 else "0.0%"
        pnl_rows.append({"구분": str(row["구분"]), "작업 내용": iname, "단위": str(row["단위"]), "물량": vol, "판매단가": sp, "매출액": r_amt, "매입단가": bp, "매입원가": c_amt, "영업이익": op_amt, "이익률": pct, "비고": str(row.get("비고", ""))})

    if st.session_state.get("q_include_infra", False):
        for _, row in infra_df.iterrows():
            iname = str(row["작업 내용"])
            u_str = str(row["단위"])
            ivol = get_infra_vol_for_row(u_str, iname)
            ibp = float(row["매입단가"])
            icost = ivol * ibp
            tot_cost += icost
            pnl_rows.append({
                "구분": f"인프라({row['구분']})",
                "작업 내용": iname,
                "단위": u_str,
                "물량": ivol,
                "판매단가": 0.0,
                "매출액": 0.0,
                "매입단가": ibp,
                "매입원가": icost,
                "영업이익": -icost,
                "이익률": "-",
                "비고": str(row.get("비고", ""))
            })

    wh_rev_tot = sum(float(r["제안단가"]) * vol_dict.get(str(r["작업 내용"]), 0) for _, r in wh_df.iterrows())
    sga_amt = round(wh_rev_tot * (float(st.session_state.q_sga_rate) / 100.0))
    st.session_state.q_sga_amt = sga_amt
    tot_cost += sga_amt

    if sga_amt > 0:
        pnl_rows.append({
            "구분": "일반관리비",
            "작업 내용": f"SG&A 배부 ({st.session_state.q_sga_rate:.1f}%)",
            "단위": "식",
            "물량": 1,
            "판매단가": 0,
            "매출액": 0,
            "매입단가": sga_amt,
            "매입원가": sga_amt,
            "영업이익": -sga_amt,
            "이익률": f"-{st.session_state.q_sga_rate:.1f}%",
            "비고": f"창고 운영매출({wh_rev_tot:,.0f}원) 대비 {st.session_state.q_sga_rate:.1f}% 배부 기준"
        })

    final_op = tot_rev - tot_cost
    final_margin = (final_op / tot_rev * 100) if tot_rev > 0 else 0.0
    avg_order_rev = (tot_rev / st.session_state.q_vol_b2c) if st.session_state.q_vol_b2c > 0 else 0.0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("1. 예상 월 매출액", f"{tot_rev:,.0f}원")
    k2.metric("2. 예상 월 매입원가", f"{tot_cost:,.0f}원")
    k3.metric("3. 월 영업이익", f"{final_op:,.0f}원", delta=f"이익률: {final_margin:.1f}% ({'흑자 🟢' if final_op >= 0 else '적자 🔴'})", delta_color="normal" if final_op >= 0 else "inverse")
    k4.metric("4. 일반관리비(SG&A)", f"{sga_amt:,.0f}원 ({st.session_state.q_sga_rate:.1f}%)")
    k5.metric("5. 건당 평균단가", f"{avg_order_rev:,.1f}원/건")
    st.write("---")

    st.markdown("##### 📑 2. 예상손익 세부 항목 분석표 (Single Source of Truth)")
    st.dataframe(pd.DataFrame(pnl_rows), column_config={"물량": st.column_config.NumberColumn(format="%,d"), "판매단가": st.column_config.NumberColumn(format="%,d"), "매출액": st.column_config.NumberColumn(format="%,d"), "매입단가": st.column_config.NumberColumn(format="%,d"), "매입원가": st.column_config.NumberColumn(format="%,d"), "영업이익": st.column_config.NumberColumn(format="%,d")}, width="stretch", hide_index=True)
    st.write("---")

    with st.expander("✈️ **3. 해외특송 국가별 세부 마진 분석표 펼쳐보기**", expanded=False):
        exp_calc = exp_df.copy()
        exp_calc["건당 마진"] = exp_calc["판매 제안단가"] - exp_calc["매입원가"]
        exp_calc["예상 발송건수"] = (st.session_state.q_express_qty * exp_calc["발송비중(%)"] / 100.0).round(0)
        exp_calc["예상 매출액"] = exp_calc["판매 제안단가"] * exp_calc["예상 발송건수"]
        exp_calc["예상 매입원가"] = exp_calc["매입원가"] * exp_calc["예상 발송건수"]
        exp_calc["예상 영업이익"] = exp_calc["예상 매출액"] - exp_calc["예상 매입원가"]
        st.dataframe(exp_calc[["국가", "기준구간", "매입원가", "판매 제안단가", "건당 마진", "발송비중(%)", "예상 발송건수", "예상 매출액", "예상 매입원가", "예상 영업이익"]], column_config={"매입원가": st.column_config.NumberColumn(format="%,d"), "판매 제안단가": st.column_config.NumberColumn(format="%,d"), "건당 마진": st.column_config.NumberColumn(format="%,d"), "발송비중(%)": st.column_config.NumberColumn(format="%.1f%%"), "예상 발송건수": st.column_config.NumberColumn(format="%,d"), "예상 매출액": st.column_config.NumberColumn(format="%,d"), "예상 매입원가": st.column_config.NumberColumn(format="%,d"), "예상 영업이익": st.column_config.NumberColumn(format="%,d")}, width="stretch", hide_index=True)

    with st.expander("🏢 **4. 센터 인프라 고정비 단가 산출 기준표 ([탭 5] 매출원가 산출 연동)**", expanded=False):
        infra_sim = infra_df.copy()
        infra_sim["적용 물량"] = [get_infra_vol_for_row(r["단위"], r["작업 내용"]) for _, r in infra_sim.iterrows()]
        infra_sim["월 예상 원가"] = infra_sim["매입단가"] * infra_sim["적용 물량"]
        st.dataframe(
            infra_sim[["구분", "작업 내용", "단위", "매입단가", "적용 물량", "월 예상 원가", "비고"]],
            column_config={
                "구분": st.column_config.TextColumn(label="구분", width="small"),
                "작업 내용": st.column_config.TextColumn(label="작업 내용", width="medium"),
                "단위": st.column_config.TextColumn(label="기본 단위", width="small"),
                "매입단가": st.column_config.NumberColumn(label="매입 단가 (원가)", format="%,d", width="medium"),
                "적용 물량": st.column_config.NumberColumn(label="적용 물량", format="%,d", width="small"),
                "월 예상 원가": st.column_config.NumberColumn(label="월 예상 원가", format="%,d", width="medium"),
                "비고": st.column_config.TextColumn(label="작업 상세 및 비고", width="large")
            },
            width="stretch", hide_index=True
        )
        st.caption(f"💡 일반관리비(SG&A)는 창고 운영매출액({wh_rev_tot:,.0f}원) 대비 **{st.session_state.q_sga_rate:.1f}% ({sga_amt:,.0f}원)**로 실시간 배부됩니다. (단가 수정은 [탭 5] 매출원가 산출에서 가능)")
    st.write("---")

    meta_d = {
        "company": st.session_state.q_company, "doc_no": st.session_state.q_doc_no, "date": st.session_state.q_date,
        "receiver": st.session_state.q_receiver, "contact": st.session_state.q_contact, "title": st.session_state.q_title,
        "category": st.session_state.q_category, "sender": st.session_state.q_sender, "terms": st.session_state.q_terms,
        "validity": st.session_state.q_validity, "insurance": st.session_state.q_insurance, "indemnity": st.session_state.q_indemnity,
        "center": st.session_state.q_center, "notes": st.session_state.q_notes, "express_qty": st.session_state.q_express_qty,
        "sga_rate": st.session_state.get("q_sga_rate", 7.0), "sga_amt": st.session_state.q_sga_amt
    }
    excel_wb = build_quote_workbook(
        meta_d, st.session_state.q_warehouse_df, st.session_state.q_express_df,
        st.session_state.q_express_margin, get_quote_vol_dict(),
        include_exp=st.session_state.q_include_express,
        infra_df=st.session_state.q_infra_df,
        include_infra=st.session_state.get("q_include_infra", False)
    )
    buf = io.BytesIO()
    excel_wb.save(buf)
    buf.seek(0)
    st.download_button(label="💾 [완성본] 화주사 제출용 견적서 · 매출원가 · 예상손익 분석표 엑셀 다운로드 (.xlsx)", data=buf, file_name=f"견적서_및_예상손익_{st.session_state.q_receiver}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dl_pnl_tab6")

# =============================================================================
# 메인 탭 라우팅 (원장 업로드 여부에 따른 동적 탭 구성)
# =============================================================================
if uploaded_file is not None:
    file_id = f"{uploaded_file.name}_{uploaded_file.size}"
    if "current_file_id" not in st.session_state or st.session_state.current_file_id != file_id:
        raw_df, init_dash = build_initial_dashboard(uploaded_file, st.session_state.customers)
        st.session_state.raw_df = raw_df
        st.session_state.dash_df = init_dash
        st.session_state.current_file_id = file_id
        st.session_state.editor_version += 1
        for m in init_dash["대분류"].dropna().unique():
            m_clean = str(m).strip()
            if m_clean and m_clean not in st.session_state.major_categories:
                st.session_state.major_categories.append(m_clean)

    dash_df = st.session_state.dash_df
    for c in st.session_state.customers:
        if c not in st.session_state.revenue_dict: st.session_state.revenue_dict[c] = 0.0

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "⚙️ [탭 1] 원가/판관비 배부", 
        "💵 [탭 2] 거래처별 매출 입력", 
        "📈 [탭 3] 최종 손익계산서 & 대시보드",
        "📄 [탭 4] 화주사 제출용 견적서",
        "🛠️ [탭 5] 매출원가 산출",
        "📊 [탭 6] 예상손익 분석표 및 시뮬레이션"
    ])

    with tab1:
        render_customer_manager("tab1")
        render_major_category_manager("tab1")
        st.write("---")

        for c in st.session_state.customers:
            if c not in dash_df.columns: dash_df[c] = 0.0

        valid_customers = [c for c in st.session_state.customers if c in dash_df.columns]
        dash_df["배부 완료액"] = dash_df[valid_customers].sum(axis=1) if len(valid_customers) > 0 else 0.0
        dash_df["미분배 잔액"] = dash_df["전체(Sheet3 총액)"] - dash_df["배부 완료액"]

        cogs_mask = dash_df["비용구분"] == "매출원가"
        sga_mask = dash_df["비용구분"] == "판관비"
        total_amount = dash_df["전체(Sheet3 총액)"].sum()
        cogs_total = dash_df.loc[cogs_mask, "전체(Sheet3 총액)"].sum()
        sga_total = dash_df.loc[sga_mask, "전체(Sheet3 총액)"].sum()
        total_unallocated = dash_df["미분배 잔액"].sum()
        is_allocated_clean = abs(total_unallocated) < 1.0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("총 발생 비용 (원장 100%)", f"{total_amount:,.0f}")
        k2.metric("1. 매출원가 합계", f"{cogs_total:,.0f}")
        k3.metric("2. 판관비 합계", f"{sga_total:,.0f}")
        k4.metric("전체 미분배 잔여 금액", f"{total_unallocated:,.0f}", delta="전체 귀속 완료 (0원) ✅" if is_allocated_clean else f"미귀속 잔여: {total_unallocated:,.0f}", delta_color="normal" if is_allocated_clean else "inverse")

        st.write("---")

        all_major_options = sorted(list(set(st.session_state.major_categories + [str(x) for x in dash_df["대분류"].dropna().unique() if str(x).strip()])))

        if "선택" not in dash_df.columns:
            dash_df["선택"] = False

        base_column_config = {
            "선택": st.column_config.CheckboxColumn(label="선택", help="삭제할 행을 선택하세요", default=False, width="small"),
            "비용구분": st.column_config.SelectboxColumn(label="🔄 비용구분 (이동)", options=["매출원가", "판관비"], required=True, width=auto_col_width("🔄 비용구분 (이동)", ["매출원가", "판관비"])),
            "계정과목": st.column_config.TextColumn(label="📑 계정과목(원장)", disabled=False, width=auto_col_width("📑 계정과목(원장)", dash_df["계정과목"])),
            "대분류": st.column_config.SelectboxColumn(label="[선택] 대분류", options=all_major_options, required=True, disabled=False, width=auto_col_width("[선택] 대분류", dash_df["대분류"])),
            "적요": st.column_config.TextColumn(label="📝 적요", disabled=False, width=auto_col_width("📝 적요", dash_df["적요"])),
            "거래처명": st.column_config.TextColumn(label="🏢 거래처명(원장)", disabled=False, width=auto_col_width("🏢 거래처명(원장)", dash_df["거래처명"])),
            "전체(Sheet3 총액)": st.column_config.NumberColumn(label="전체 발생금액", format="%,d", disabled=False, width=auto_col_width("전체 발생금액", ["3,504,010,871"])),
            "미분배 잔액": st.column_config.NumberColumn(label="⚠️ 미분배 잔액", format="%,d", disabled=True, width=auto_col_width("⚠️ 미분배 잔액", ["3,504,010,871"]))
        }
        for c in st.session_state.customers:
            base_column_config[c] = st.column_config.NumberColumn(label=f"🏢 {c}", format="%,d", disabled=False, min_value=0.0, width=auto_col_width(f"🏢 {c}", ["1,000,000,000"]))

        display_cols = ["row_id", "선택", "비용구분", "계정과목", "대분류", "적요", "거래처명", "전체(Sheet3 총액)", "미분배 잔액"] + st.session_state.customers

        # =====================================================================
        # 1. 매출원가 섹션
        # =====================================================================
        st.subheader("1. 📦 [매출원가] 거래처별 손익분석")

        cg_b1, cg_b2, cg_b3, _ = st.columns([1.2, 1.2, 4.0, 10.0], gap="small")
        with cg_b1:
            btn_add_cogs = st.button("+ 행 추가", key="btn_add_cogs")
        with cg_b2:
            btn_del_cogs = st.button("- 행 삭제", key="btn_del_cogs")
        with cg_b3:
            btn_cogs_to_sga = st.button("📦 매출원가 전체 ➡️ 판관비 일괄 이동", key="btn_cogs_to_sga")

        if btn_cogs_to_sga:
            c_mask = st.session_state.dash_df["비용구분"] == "매출원가"
            if c_mask.any():
                st.session_state.dash_df.loc[c_mask, "비용구분"] = "판관비"
                st.session_state.editor_version += 1
                st.success("매출원가의 모든 항목이 판관비로 일괄 이동되었습니다.")
                st.rerun()
            else:
                st.warning("이동할 매출원가 항목이 없습니다.")

        if btn_add_cogs:
            new_id = int(st.session_state.dash_df["row_id"].max() + 1) if len(st.session_state.dash_df) > 0 else 0
            new_row = {
                "row_id": new_id, "선택": False, "비용구분": "매출원가", "계정과목": "(직접입력)",
                "대분류": st.session_state.major_categories[0] if st.session_state.major_categories else "기타",
                "적요": "(신규 항목)", "거래처명": "-", "상세": "-", "전체(Sheet3 총액)": 0.0, "미분배 잔액": 0.0
            }
            for c in st.session_state.customers: new_row[c] = 0.0
            st.session_state.dash_df = pd.concat([st.session_state.dash_df, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state.editor_version += 1
            st.success("매출원가에 신규 행이 추가되었습니다.")
            st.rerun()

        # 매출원가 테이블
        cogs_cfg = dict(base_column_config); cogs_cfg["row_id"] = None
        edited_cogs = st.data_editor(dash_df[cogs_mask][display_cols].copy(), column_config=cogs_cfg, width="stretch", hide_index=True, key=f"cogs_editor_{st.session_state.editor_version}")

        if btn_del_cogs:
            checked_cogs_ids = edited_cogs[edited_cogs["선택"] == True]["row_id"].tolist()
            if len(checked_cogs_ids) > 0:
                st.session_state.dash_df = st.session_state.dash_df[~st.session_state.dash_df["row_id"].isin(checked_cogs_ids)].reset_index(drop=True)
                st.session_state.editor_version += 1
                st.success(f"매출원가에서 선택한 {len(checked_cogs_ids)}개 행이 삭제되었습니다.")
                st.rerun()
            else:
                st.warning("삭제할 행의 '선택' 체크박스를 먼저 체크해 주세요.")

        # =====================================================================
        # 2. 판관비 섹션
        # =====================================================================
        st.subheader("2. 🏢 [판관비] 거래처별 손익분석")

        sg_b1, sg_b2, sg_b3, _ = st.columns([1.2, 1.2, 4.0, 10.0], gap="small")
        with sg_b1:
            btn_add_sga = st.button("+ 행 추가", key="btn_add_sga")
        with sg_b2:
            btn_del_sga = st.button("- 행 삭제", key="btn_del_sga")
        with sg_b3:
            btn_sga_to_cogs = st.button("🏢 판관비 전체 ➡️ 매출원가 일괄 이동", key="btn_sga_to_cogs")

        if btn_sga_to_cogs:
            s_mask = st.session_state.dash_df["비용구분"] == "판관비"
            if s_mask.any():
                st.session_state.dash_df.loc[s_mask, "비용구분"] = "매출원가"
                st.session_state.editor_version += 1
                st.success("판관비의 모든 항목이 매출원가로 일괄 이동되었습니다.")
                st.rerun()
            else:
                st.warning("이동할 판관비 항목이 없습니다.")

        if btn_add_sga:
            new_id = int(st.session_state.dash_df["row_id"].max() + 1) if len(st.session_state.dash_df) > 0 else 0
            new_row = {
                "row_id": new_id, "선택": False, "비용구분": "판관비", "계정과목": "(직접입력)",
                "대분류": st.session_state.major_categories[0] if st.session_state.major_categories else "기타",
                "적요": "(신규 항목)", "거래처명": "-", "상세": "-", "전체(Sheet3 총액)": 0.0, "미분배 잔액": 0.0
            }
            for c in st.session_state.customers: new_row[c] = 0.0
            st.session_state.dash_df = pd.concat([st.session_state.dash_df, pd.DataFrame([new_row])], ignore_index=True)
            st.session_state.editor_version += 1
            st.success("판관비에 신규 행이 추가되었습니다.")
            st.rerun()

        # 판관비 테이블
        sga_cfg = dict(base_column_config); sga_cfg["row_id"] = None
        edited_sga = st.data_editor(dash_df[sga_mask][display_cols].copy(), column_config=sga_cfg, width="stretch", hide_index=True, key=f"sga_editor_{st.session_state.editor_version}")

        if btn_del_sga:
            checked_sga_ids = edited_sga[edited_sga["선택"] == True]["row_id"].tolist()
            if len(checked_sga_ids) > 0:
                st.session_state.dash_df = st.session_state.dash_df[~st.session_state.dash_df["row_id"].isin(checked_sga_ids)].reset_index(drop=True)
                st.session_state.editor_version += 1
                st.success(f"판관비에서 선택한 {len(checked_sga_ids)}개 행이 삭제되었습니다.")
                st.rerun()
            else:
                st.warning("삭제할 행의 '선택' 체크박스를 먼저 체크해 주세요.")

        # O(N) 최적화 사전 매핑 인메모리 동기화
        row_id_to_idx = {rid: idx for idx, rid in enumerate(st.session_state.dash_df["row_id"])}
        for sub_edited in [edited_cogs, edited_sga]:
            for _, row in sub_edited.iterrows():
                r_id = row["row_id"]
                if r_id in row_id_to_idx:
                    target_idx = row_id_to_idx[r_id]
                    if "선택" in row and pd.notna(row["선택"]):
                        st.session_state.dash_df.loc[target_idx, "선택"] = bool(row["선택"])
                    for col_name in ["비용구분", "대분류", "적요", "계정과목", "거래처명", "전체(Sheet3 총액)"]:
                        if col_name in row and pd.notna(row[col_name]) and row[col_name] != st.session_state.dash_df.loc[target_idx, col_name]:
                            st.session_state.dash_df.loc[target_idx, col_name] = row[col_name]
                    for c in st.session_state.customers:
                        nv = float(row[c]) if pd.notna(row[c]) else 0.0
                        ov = float(st.session_state.dash_df.loc[target_idx, c]) if c in st.session_state.dash_df.columns else 0.0
                        if nv != ov:
                            st.session_state.dash_df.loc[target_idx, c] = nv

        st.write("---")
        bot_b1, bot_b2, _ = st.columns([3.8, 4.0, 10.0], gap="small")
        with bot_b1:
            btn_apply = st.button("💾 배부 변경사항 적용 (미분배 잔액 및 손익 갱신)", type="primary", key="btn_apply")
        with bot_b2:
            btn_agg_duplicates = st.button("🔄 동일한 [비용구분/대분류/계정과목/거래처] 합산 정리", key="btn_agg_duplicates_tab1")

        if btn_apply:
            st.session_state.editor_version += 1
            st.success("✅ 배부 변경사항이 성공적으로 반영되어 미분배 잔액 및 대시보드가 갱신되었습니다.")
            st.rerun()

        if btn_agg_duplicates:
            agg_d = {
                "전체(Sheet3 총액)": "sum",
                "적요": lambda s: ", ".join(dict.fromkeys(str(x) for x in s if pd.notna(x)))[:50],
                "상세": lambda s: ", ".join(dict.fromkeys(str(x) for x in s if pd.notna(x) and str(x) != "-"))[:30] or "-"
            }
            for c in st.session_state.customers: agg_d[c] = "sum"
            m_df = st.session_state.dash_df.groupby(["비용구분", "대분류", "계정과목", "거래처명"], as_index=False).agg(agg_d)
            m_df["row_id"] = list(range(len(m_df)))
            m_df["선택"] = False
            for c in st.session_state.customers:
                if c not in m_df.columns: m_df[c] = 0.0
            st.session_state.dash_df = m_df
            st.session_state.editor_version += 1
            st.success("동일 항목 기준 합산 정리가 완료되었습니다.")
            st.rerun()

    with tab2:
        render_customer_manager("tab2")
        st.write("---")
        st.subheader("💵 거래처별 매출액(Revenue) 입력")
        st.caption("각 거래처(행)별 당월 매출액을 입력해 주세요. 전체 총 매출액은 자동으로 합산되며 [탭 3] 손익계산서로 연동됩니다.")

        rev_rows = [{"거래처명": c, "당월 매출액": float(st.session_state.revenue_dict.get(c, 0.0))} for c in st.session_state.customers]
        rev_df = pd.DataFrame(rev_rows)

        col_t2_left, col_t2_right = st.columns(2)
        with col_t2_left:
            edited_rev = st.data_editor(
                rev_df,
                column_config={
                    "거래처명": st.column_config.TextColumn(label="🏢 거래처명", disabled=True, width=auto_col_width("🏢 거래처명", rev_df["거래처명"])),
                    "당월 매출액": st.column_config.NumberColumn(label="💰 당월 매출액", format="%,d", min_value=0.0, step=1000000.0, width=auto_col_width("💰 당월 매출액", [f"{v:,.0f}" for v in st.session_state.revenue_dict.values()]))
                },
                width="stretch", hide_index=True, key=f"rev_row_editor_{st.session_state.editor_version}"
            )

        rev_mod = False
        for _, r in edited_rev.iterrows():
            c_name, c_amt = r["거래처명"], float(r["당월 매출액"]) if pd.notna(r["당월 매출액"]) else 0.0
            if st.session_state.revenue_dict.get(c_name, 0.0) != c_amt:
                st.session_state.revenue_dict[c_name] = c_amt; rev_mod = True
        if rev_mod: st.rerun()

        st.write("---")
        rm1, rm2 = st.columns(2)
        rm1.metric("총 매출액 합계", f"{sum(st.session_state.revenue_dict.values()):,.0f}")
        rm2.metric("매출 발생 거래처 수", f"{sum(1 for v in st.session_state.revenue_dict.values() if v > 0)} / {len(st.session_state.customers)}개 사")

    with tab3:
        st.subheader("📈 [통합] 거래처별 최종 손익계산서")
        st.caption("💡 판관비와 매출원가를 각각 **합계금액**으로 집계하여 전체 손익 요약을 한눈에 확인합니다.")

        cogs_sub = dash_df[dash_df["비용구분"] == "매출원가"]
        sga_sub = dash_df[dash_df["비용구분"] == "판관비"]

        t_rev = sum(st.session_state.revenue_dict.values())
        t_cogs = cogs_sub["전체(Sheet3 총액)"].sum()
        t_gp = t_rev - t_cogs
        gp_r = (t_gp / t_rev * 100) if t_rev > 0 else 0.0
        t_sga = sga_sub["전체(Sheet3 총액)"].sum()
        t_op = t_gp - t_sga
        op_r = (t_op / t_rev * 100) if t_rev > 0 else 0.0

        kp1, kp2, kp3, kp4, kp5 = st.columns(5)
        kp1.metric("1. 총 매출액", f"{t_rev:,.0f}")
        kp2.metric("2. 매출원가 합계", f"{t_cogs:,.0f}")
        kp3.metric("3. 매출총이익", f"{t_gp:,.0f}", delta=f"이익률: {gp_r:.1f}%")
        kp4.metric("4. 판관비 합계", f"{t_sga:,.0f}")
        kp5.metric("5. 영업이익", f"{t_op:,.0f}", delta=f"영업이익률: {op_r:.1f}% ({'흑자' if t_op >= 0 else '적자'})", delta_color="normal" if t_op >= 0 else "inverse")
        st.write("---")

        pnl_data = [
            ("1. 매출액", t_rev, {c: st.session_state.revenue_dict.get(c, 0.0) for c in st.session_state.customers}),
            ("2. 매출원가 (합계)", t_cogs, {c: cogs_sub[c].sum() if c in cogs_sub.columns else 0.0 for c in st.session_state.customers}),
            ("3. 매출총이익 (1 - 2)", t_gp, {c: st.session_state.revenue_dict.get(c, 0.0) - (cogs_sub[c].sum() if c in cogs_sub.columns else 0.0) for c in st.session_state.customers}),
            ("4. 판관비 (합계)", t_sga, {c: sga_sub[c].sum() if c in sga_sub.columns else 0.0 for c in st.session_state.customers}),
            ("5. 영업이익 (3 - 4)", t_op, {c: (st.session_state.revenue_dict.get(c, 0.0) - (cogs_sub[c].sum() if c in cogs_sub.columns else 0.0)) - (sga_sub[c].sum() if c in sga_sub.columns else 0.0) for c in st.session_state.customers})
        ]
        sum_pnl_rows = []
        for label, tot_val, cust_vals in pnl_data:
            r_d = {"항목": label, "전체(합계)": tot_val}
            r_d.update(cust_vals)
            sum_pnl_rows.append(r_d)
        sum_pnl_df = pd.DataFrame(sum_pnl_rows)

        sum_cfg = {"항목": st.column_config.TextColumn(label="손익 구분 항목", width=auto_col_width("손익 구분 항목", sum_pnl_df["항목"])), "전체(합계)": st.column_config.NumberColumn(label="전체(합계)", format="%,d", width=auto_col_width("전체(합계)", ["3,504,010,871"]))}
        for c in st.session_state.customers: sum_cfg[c] = st.column_config.NumberColumn(label=f"🏢 {c}", format="%,d", width=auto_col_width(f"🏢 {c}", ["1,000,000,000"]))
        st.dataframe(sum_pnl_df, column_config=sum_cfg, width="stretch", hide_index=True)

        with st.expander("🔍 **매출원가 및 판관비 세부 내역 펼쳐보기**", expanded=False):
            c_det = cogs_sub.groupby(["대분류", "계정과목"], as_index=False)[["전체(Sheet3 총액)"] + st.session_state.customers].sum()
            s_det = sga_sub.groupby(["대분류", "계정과목"], as_index=False)[["전체(Sheet3 총액)"] + st.session_state.customers].sum()
            det_cfg = {"대분류": st.column_config.TextColumn(label="대분류"), "계정과목": st.column_config.TextColumn(label="계정과목"), "전체(Sheet3 총액)": st.column_config.NumberColumn(label="발생 총액", format="%,d")}
            for c in st.session_state.customers: det_cfg[c] = st.column_config.NumberColumn(label=c, format="%,d")
            st.markdown("##### 📦 매출원가 대분류/계정과목별 내역")
            st.dataframe(c_det, column_config=det_cfg, width="stretch", hide_index=True)
            st.markdown("##### 🏢 판관비 대분류/계정과목별 내역")
            st.dataframe(s_det, column_config=det_cfg, width="stretch", hide_index=True)

        st.write("---")
        st.subheader("📊 거래처별 손익 기여 요약")
        kpi_rows = []
        for c in st.session_state.customers:
            c_r = sum_pnl_rows[0][c]; c_cg = sum_pnl_rows[1][c]; c_g = sum_pnl_rows[2][c]; c_sg = sum_pnl_rows[3][c]; c_o = sum_pnl_rows[4][c]
            kpi_rows.append({"거래처": c, "매출액": c_r, "매출원가": c_cg, "매출총이익": c_g, "판관비": c_sg, "영업이익": c_o, "영업이익률": f"{(c_o/c_r*100):.1f}%" if c_r > 0 else "0.0%", "수익성": "🟢 흑자" if c_o >= 0 else "🔴 적자"})
        kpi_df = pd.DataFrame(kpi_rows)
        st.dataframe(kpi_df, column_config={"거래처": st.column_config.TextColumn(width=auto_col_width("거래처", kpi_df["거래처"])), "매출액": st.column_config.NumberColumn(format="%,d"), "매출원가": st.column_config.NumberColumn(format="%,d"), "매출총이익": st.column_config.NumberColumn(format="%,d"), "판관비": st.column_config.NumberColumn(format="%,d"), "영업이익": st.column_config.NumberColumn(format="%,d"), "영업이익률": st.column_config.TextColumn(width="small"), "수익성": st.column_config.TextColumn(width="small")}, width="stretch", hide_index=True)

        st.write("---")
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine="openpyxl") as wr:
            sum_pnl_df.to_excel(wr, sheet_name="최종_손익요약", index=False)
            kpi_df.to_excel(wr, sheet_name="거래처별_손익요약", index=False)
            dash_df[dash_df["비용구분"] == "매출원가"].drop(columns=["선택", "row_id"], errors="ignore").to_excel(wr, sheet_name="1.매출원가_상세배부", index=False)
            dash_df[dash_df["비용구분"] == "판관비"].drop(columns=["선택", "row_id"], errors="ignore").to_excel(wr, sheet_name="2.판관비_상세배부", index=False)
            if "raw_df" in st.session_state and st.session_state.raw_df is not None:
                st.session_state.raw_df.to_excel(wr, sheet_name="원장전체_원본내역", index=False)
            for sheetname in wr.sheets:
                ws_cur = wr.sheets[sheetname]
                for col in ws_cur.columns:
                    c_letter = get_column_letter(col[0].column)
                    max_l = max((sum(2 if ord(ch) > 128 else 1 for ch in str(c.value or '')) for c in col), default=0)
                    ws_cur.column_dimensions[c_letter].width = max(max_l + 3, 11)
                for row in ws_cur.iter_rows(min_row=2):
                    for cell in row:
                        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool) and cell.column != 1:
                            cell.number_format = '#,##0'
        out_buf.seek(0)
        st.download_button(label="💾 완성된 최종 손익계산서 엑셀 다운로드", data=out_buf, file_name="거래처별_최종_손익계산서.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dl_final_pnl_tab3")

    with tab4:
        render_tab4_proposal()

    with tab5:
        render_tab5_cogs_calc()

    with tab6:
        render_tab6_pnl_simulation()
else:
    tab_upload, tab4_solo, tab5_solo, tab6_solo = st.tabs([
        "📂 원장 업로드 및 분석",
        "📄 [탭 4] 화주사 제출용 견적서",
        "🛠️ [탭 5] 매출원가 산출",
        "📊 [탭 6] 예상손익 분석표 및 시뮬레이션"
    ])
    with tab_upload:
        st.info("💡 상단에서 원장 엑셀 파일(.xlsx, .csv)을 업로드하면 기존 원장 배부 및 정밀 손익분석(탭 1~3)이 활성화됩니다.")
    with tab4_solo:
        render_tab4_proposal()
    with tab5_solo:
        render_tab5_cogs_calc()
    with tab6_solo:
        render_tab6_pnl_simulation()
