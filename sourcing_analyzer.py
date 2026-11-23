"""
초보 오픈마켓 셀러를 위한 4대 마켓 동시 도매 소싱 & 마진 자동 분석기 (Streamlit)
-------------------------------------------------------------------------
주요 기능:
1. 도매처 원클릭 키워드 검색 (도매매, 오너클랜, 도매토피아, 도매창고, 스페셜오퍼, 네이버쇼핑)
2. 도매처 CSV/엑셀 파일 업로드 및 한국어 인코딩 자동 감지 + 파일 내 키워드 검색
3. 초보 셀러 회피 품목(유아·의료·지재권 위험) 및 KC 인증 주의 키워드 자동 필터링
4. 4대 마켓(스마트스토어, 11번가, G마켓, 옥션) 권장 판매가 & 예상 순이익 동시 계산
5. 일반과세자 기본값 적용, 최대 1,000개 표시, 엑셀 호환 CSV 다운로드
"""

import io
import math
from urllib.parse import quote
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. 기본 설정 및 상수 정의
# ==============================================================================
st.set_page_config(
    page_title="4대 마켓 동시 소싱·마진 분석기",
    page_icon="📦",
    layout="wide",
)

# 초보 셀러 절대 주의(지재권·어린이·의료·건강기능식품 등) 기본 금지 키워드
DEFAULT_RISK_KEYWORDS = [
    "어린이", "유아", "아기", "키즈", "베이비", "신생아", "젖병", "치발기",
    "의료", "치료", "혈압", "혈당", "다이어트", "체지방", "질병", "의약",
    "나이키", "아디다스", "삼성", "애플", "다이슨", "디즈니", "산리오",
    "포켓몬", "카카오프렌즈", "레고", "샤넬", "구찌", "루이비통",
]

# 전기/전파법 KC 인증 확인이 필요한 주의 키워드
DEFAULT_KC_KEYWORDS = [
    "전기", "충전", "USB", "블루투스", "무선", "열선", "히터", "온열",
    "어댑터", "LED", "보조배터리", "가습기", "선풍기", "면도기", "드라이기",
]


# ==============================================================================
# 2. 유틸리티 함수
# ==============================================================================
def create_sample_dataframe() -> pd.DataFrame:
    """테스트용 샘플 데이터 생성"""
    sample_data = [
        {"도매처": "오너클랜", "상품코드": "W1001", "상품명": "싱크대 하부장 정리대 2단 슬라이딩 선반", "공급가": 12500, "택배비": 3000, "시중가": 23900},
        {"도매처": "도매매", "상품코드": "D2045", "상품명": "프라이팬 정리대 길이조절 냄비 뚜껑 거치대", "공급가": 11000, "택배비": 3000, "시중가": 21900},
        {"도매처": "온채널", "상품코드": "O3312", "상품명": "냉장고 정리 트레이 투명 손잡이 바구니 6P 세트", "공급가": 13500, "택배비": 3000, "시중가": 25900},
        {"도매처": "오너클랜", "상품코드": "W1088", "상품명": "USB 충전식 차량용 열선 방석 시트", "공급가": 16000, "택배비": 3000, "시중가": 29900},
        {"도매처": "도매토피아", "상품코드": "T5521", "상품명": "유아용 실리콘 캐릭터 빨대컵 (디즈니)", "공급가": 7000, "택배비": 3000, "시중가": 15900},
        {"도매처": "도매매", "상품코드": "D2099", "상품명": "주방 틈새 선반 이동식 3단 트롤리", "공급가": 14900, "택배비": 3000, "시중가": 28900},
        {"도매처": "스페셜오퍼", "상품코드": "S7710", "상품명": "다용도 욕실 미끄럼방지 매트 6P 세트", "공급가": 12000, "택배비": 3000, "시중가": 22500},
        {"도매처": "도매매", "상품코드": "D2110", "상품명": "체지방 감소 다이어트 건강 보조 밴드", "공급가": 14000, "택배비": 3000, "시중가": 28000},
    ]
    return pd.DataFrame(sample_data)


def load_uploaded_file(uploaded_file) -> pd.DataFrame:
    """CSV 또는 Excel 파일 안전 로드 (한국어 인코딩 자동 감지)"""
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith((".xlsx", ".xls")):
            return pd.read_excel(uploaded_file)
        elif file_name.endswith(".csv"):
            raw_bytes = uploaded_file.read()
            for encoding in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
                try:
                    return pd.read_csv(io.BytesIO(raw_bytes), encoding=encoding)
                except UnicodeDecodeError:
                    continue
            raise ValueError("지원하지 않는 파일 인코딩입니다. 엑셀에서 'CSV UTF-8'로 다시 저장해 주세요.")
        else:
            raise ValueError("CSV 또는 Excel(.xlsx, .xls) 파일만 업로드할 수 있습니다.")
    except Exception as e:
        raise RuntimeError(f"파일을 읽는 중 오류가 발생했습니다: {e}")


def clean_numeric(series: pd.Series) -> pd.Series:
    """가격 컬럼을 안전하게 숫자로 변환"""
    cleaned = (
        series.astype(str)
        .str.replace(r"[^0-9.]", "", regex=True)
        .replace("", "0")
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def check_product_risk(name: str, risk_keywords: list, kc_keywords: list) -> tuple[str, str]:
    """금지어 및 KC 인증 주의 키워드 탐지"""
    name_str = str(name)
    found_risks = [kw for kw in risk_keywords if kw and kw in name_str]
    found_kc = [kw for kw in kc_keywords if kw and kw in name_str]

    if found_risks:
        return "🔴 제외 권장", f"위험 키워드: {', '.join(found_risks)}"
    if found_kc:
        return "🟡 KC 확인 필요", f"인증 확인: {', '.join(found_kc)}"
    return "🟢 안전 통과", "특이사항 없음"


def calculate_recommended_price(
    supply_price: float,
    shipping_cost: float,
    fee_rate_pct: float,
    target_margin_pct: float,
    shipping_mode: str,
    customer_shipping: float,
    tax_type: str,
) -> tuple[int, int]:
    """목표 순마진율 기준 권장 판매가(100원 단위 올림)와 예상 순이익 계산"""
    fee_rate = fee_rate_pct / 100.0
    target_margin = target_margin_pct / 100.0

    if shipping_mode == "무료배송 (판매가에 택배비 포함)":
        net_shipping_cost = shipping_cost
        extra_shipping_revenue = 0.0
        shipping_fee = 0.0
    else:
        net_shipping_cost = shipping_cost
        extra_shipping_revenue = customer_shipping
        shipping_fee = customer_shipping * 0.0363

    base_cost = supply_price + net_shipping_cost + shipping_fee - extra_shipping_revenue

    if tax_type == "일반과세자 (부가세 10% 반영)":
        denominator = 1.0 - fee_rate - (1.1 * target_margin)
    elif tax_type == "간이과세자 (부가세 약 1.5% 반영)":
        denominator = 1.0 - fee_rate - 0.015 - target_margin
    else:
        denominator = 1.0 - fee_rate - target_margin

    if denominator <= 0.05:
        return 0, 0

    raw_price = base_cost / denominator
    recommended_price = max(0, int(math.ceil(raw_price / 100.0) * 100))

    market_fee = recommended_price * fee_rate
    pre_tax_profit = recommended_price - market_fee - base_cost

    if tax_type == "일반과세자 (부가세 10% 반영)":
        net_profit = pre_tax_profit / 1.1
    elif tax_type == "간이과세자 (부가세 약 1.5% 반영)":
        net_profit = pre_tax_profit - ((recommended_price + extra_shipping_revenue) * 0.015)
    else:
        net_profit = pre_tax_profit

    return recommended_price, int(round(net_profit))


# ==============================================================================
# 3. 사이드바 UI (4대 마켓 수수료·마진·과세·필터 설정)
# ==============================================================================
with st.sidebar:
    st.header("⚙️ 4대 마켓 동시 설정")

    target_margin_pct = st.slider(
        "목표 순마진율 (%)",
        min_value=5,
        max_value=50,
        value=20,
        step=1,
        help="초보 셀러는 최소 15~20% 이상을 권장합니다.",
    )

    tax_type = st.radio(
        "사업자 과세 유형",
        options=[
            "간이과세자 (납부면제/초기)",
            "간이과세자 (부가세 약 1.5% 반영)",
            "일반과세자 (부가세 10% 반영)",
        ],
        index=2,  # 일반과세자 기본값
    )

    shipping_mode = st.radio(
        "배송비 정책",
        options=[
            "유료배송 (고객 별도 부담)",
            "무료배송 (판매가에 택배비 포함)",
        ],
        index=0,
    )
    customer_shipping = 3000
    if shipping_mode == "유료배송 (고객 별도 부담)":
        customer_shipping = st.number_input(
            "고객에게 받을 배송비 (원)",
            min_value=0,
            max_value=10000,
            value=3000,
            step=500,
        )

    st.divider()
    st.subheader("🏪 마켓별 수수료율 (%)")
    fee_smartstore = st.number_input("네이버 스마트스토어 (%)", min_value=0.0, max_value=30.0, value=5.8, step=0.1)
    fee_11st = st.number_input("11번가 (%)", min_value=0.0, max_value=30.0, value=13.0, step=0.1)
    fee_gmarket = st.number_input("G마켓 (%)", min_value=0.0, max_value=30.0, value=13.0, step=0.1)
    fee_auction = st.number_input("옥션 (%)", min_value=0.0, max_value=30.0, value=13.0, step=0.1)

    st.divider()
    st.subheader("🛡️ 리스크 필터 키워드")
    custom_risk_text = st.text_area(
        "제외 키워드 (쉼표로 구분)",
        value=", ".join(DEFAULT_RISK_KEYWORDS),
        height=100,
    )
    custom_kc_text = st.text_area(
        "KC 인증 주의 키워드 (쉼표로 구분)",
        value=", ".join(DEFAULT_KC_KEYWORDS),
        height=80,
    )

risk_kw_list = [k.strip() for k in custom_risk_text.split(",") if k.strip()]
kc_kw_list = [k.strip() for k in custom_kc_text.split(",") if k.strip()]


# ==============================================================================
# 4. 메인 화면 UI
# ==============================================================================
st.title("📦 4대 마켓 동시 도매 소싱·마진 분석기")
st.caption("도매처 원클릭 검색부터 4대 마켓(스마트스토어·11번가·G마켓·옥션) 마진 분석까지 한 번에 처리합니다.")

# --- [신규] 도매처 원클릭 키워드 검색 ---
with st.expander("🔍 도매처 원클릭 키워드 검색 (클릭해서 바로 찾기)", expanded=True):
    search_kw = st.text_input(
        "찾고 싶은 아이템 키워드를 입력하세요",
        placeholder="예: 싱크대 하부장 정리대, 프라이팬 거치대, 냉장고 트레이",
    ).strip()

    if search_kw:
        q_utf8 = quote(search_kw)
        b1, b2, b3, b4, b5, b6 = st.columns(6)
        b1.link_button("도매매 검색", f"https://domemedb.domeggook.com/index/item/supplyList.php?sf=subject&enc=utf8&fromOversea=0&mode=search&sw={q_utf8}&image_file=", use_container_width=True)
        b2.link_button("오너클랜 검색", f"https://www.ownerclan.com/V2/product/search.php?topSearchKeywordInfo=&topSearchKeyword={q_utf8}&topSearchType=all", use_container_width=True)
        b3.link_button("도매토피아 검색", f"https://dometopia.com/goods/search?search_text={q_utf8}", use_container_width=True)
        b4.link_button("스페셜오퍼 검색", f"https://specialoffer.kr/shop/search.php?ss_field=all&ss_tx={q_utf8}", use_container_width=True)
        b5.link_button("온채널 검색", f"https://www.onch3.co.kr/dbcenter_renewal/index.php?channel=product_nm&keyword={q_utf8}", use_container_width=True)
        b6.link_button("네이버 시중가", f"https://search.shopping.naver.com/search/all?query={q_utf8}", use_container_width=True)
    else:
        st.caption("위 입력창에 키워드를 치면 주요 도매사이트와 네이버 쇼핑 검색 버튼이 나타납니다.")

st.divider()

# 데이터 소스 선택 (파일 업로드 vs 샘플 데이터)
col_upload, col_sample = st.columns([3, 1])
with col_upload:
    uploaded_file = st.file_uploader(
        "도매처 상품 파일 업로드 (CSV, XLSX, XLS)",
        type=["csv", "xlsx", "xls"],
    )
with col_sample:
    st.write("")
    st.write("")
    use_sample = st.checkbox("샘플 데이터로 체험하기", value=(uploaded_file is None))

df_raw = None
if uploaded_file is not None and not use_sample:
    try:
        df_raw = load_uploaded_file(uploaded_file)
        st.success(f"파일 로드 완료! 총 {len(df_raw):,}개의 상품이 확인되었습니다.")
    except Exception as err:
        st.error(str(err))
elif use_sample:
    df_raw = create_sample_dataframe()
    st.info("현재 **테스트용 샘플 데이터(8개 상품)**로 실행 중입니다. 상단에서 내 도매처 파일을 올리면 바로 분석됩니다.")

if df_raw is not None and not df_raw.empty:
    st.divider()
    st.subheader("1️⃣ 컬럼 매핑 (도매처 양식 맞추기)")

    columns = list(df_raw.columns)

    def guess_col_index(candidates: list[str], fallback: int = 0) -> int:
        for idx, col in enumerate(columns):
            for cand in candidates:
                if cand in str(col):
                    return idx
        return min(fallback, len(columns) - 1)

    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        col_name = st.selectbox("상품명 열", columns, index=guess_col_index(["상품명", "품명", "제품명", "name"]))
    with m_col2:
        col_supply = st.selectbox("도매 공급가 열", columns, index=guess_col_index(["공급가", "도매가", "단가", "원가", "price"]))
    with m_col3:
        shipping_options = ["(고정 3,000원 적용)"] + columns
        default_ship_idx = 0
        for idx, col in enumerate(columns, start=1):
            if any(k in str(col) for k in ["배송비", "택배비", "운임"]):
                default_ship_idx = idx
                break
        col_shipping = st.selectbox("도매 택배비 열", shipping_options, index=default_ship_idx)
    with m_col4:
        market_price_options = ["(시중가 비교 안 함)"] + columns
        default_mkt_idx = 0
        for idx, col in enumerate(columns, start=1):
            if any(k in str(col) for k in ["시중가", "소비자가", "판매가", "권장가"]):
                default_mkt_idx = idx
                break
        col_market_price = st.selectbox("시중가/시장평균가 열 (선택)", market_price_options, index=default_mkt_idx)

    # 4대 마켓 동시 계산 수행
    df_result = df_raw.copy()
    supply_series = clean_numeric(df_result[col_supply])

    if col_shipping == "(고정 3,000원 적용)":
        shipping_series = pd.Series([3000.0] * len(df_result))
    else:
        shipping_series = clean_numeric(df_result[col_shipping])

    risk_statuses, risk_reasons = [], []
    ss_prices, ss_profits = [], []
    st11_prices, st11_profits = [], []
    gm_prices, gm_profits = [], []
    au_prices, au_profits = [], []

    for idx in range(len(df_result)):
        p_name = df_result.iloc[idx][col_name]
        s_price = float(supply_series.iloc[idx])
        s_ship = float(shipping_series.iloc[idx])

        status, reason = check_product_risk(p_name, risk_kw_list, kc_kw_list)
        risk_statuses.append(status)
        risk_reasons.append(reason)

        # 4개 마켓 각각 계산
        p_ss, pf_ss = calculate_recommended_price(s_price, s_ship, fee_smartstore, target_margin_pct, shipping_mode, customer_shipping, tax_type)
        p_11, pf_11 = calculate_recommended_price(s_price, s_ship, fee_11st, target_margin_pct, shipping_mode, customer_shipping, tax_type)
        p_gm, pf_gm = calculate_recommended_price(s_price, s_ship, fee_gmarket, target_margin_pct, shipping_mode, customer_shipping, tax_type)
        p_au, pf_au = calculate_recommended_price(s_price, s_ship, fee_auction, target_margin_pct, shipping_mode, customer_shipping, tax_type)

        ss_prices.append(p_ss); ss_profits.append(pf_ss)
        st11_prices.append(p_11); st11_profits.append(pf_11)
        gm_prices.append(p_gm); gm_profits.append(pf_gm)
        au_prices.append(p_au); au_profits.append(pf_au)

    df_result["리스크 판정"] = risk_statuses
    df_result["판정 사유"] = risk_reasons
    df_result["스마트스토어 권장가"] = ss_prices
    df_result["스마트스토어 순이익"] = ss_profits
    df_result["11번가 권장가"] = st11_prices
    df_result["11번가 순이익"] = st11_profits
    df_result["G마켓 권장가"] = gm_prices
    df_result["G마켓 순이익"] = gm_profits
    df_result["옥션 권장가"] = au_prices
    df_result["옥션 순이익"] = au_profits

    # 시중가 대비 가격 경쟁력
    if col_market_price != "(시중가 비교 안 함)":
        mkt_series = clean_numeric(df_result[col_market_price])
        comp_list = []
        for p_ss, p_11, mkt_p in zip(ss_prices, st11_prices, mkt_series):
            if mkt_p <= 0:
                comp_list.append("비교불가")
            elif p_11 <= mkt_p:
                comp_list.append("🏆 4대마켓 모두 경쟁력")
            elif p_ss <= mkt_p:
                comp_list.append("✅ 스마트스토어만 경쟁력")
            else:
                comp_list.append("⚠️ 시중가 초과")
        df_result["4대마켓 경쟁력"] = comp_list

    # ==========================================================================
    # 5. 요약 KPI 대시보드
    # ==========================================================================
    st.divider()
    st.subheader(f"2️⃣ 4대 마켓 동시 분석 요약 (목표 마진 {target_margin_pct}%)")

    total_cnt = len(df_result)
    safe_cnt = (df_result["리스크 판정"] == "🟢 안전 통과").sum()
    kc_cnt = (df_result["리스크 판정"] == "🟡 KC 확인 필요").sum()
    risk_cnt = (df_result["리스크 판정"] == "🔴 제외 권장").sum()
    pass_df = df_result[df_result["리스크 판정"] != "🔴 제외 권장"]
    avg_ss_profit = int(pass_df["스마트스토어 순이익"].mean() or 0)
    avg_11_profit = int(pass_df["11번가 순이익"].mean() or 0)

    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("전체 상품 수", f"{total_cnt:,}개")
    kpi2.metric("🟢 안전 통과", f"{safe_cnt:,}개")
    kpi3.metric("🟡 KC 확인 필요", f"{kc_cnt:,}개")
    kpi4.metric("🔴 제외 권장", f"{risk_cnt:,}개")
    kpi5.metric("스토어 평균 순이익", f"{avg_ss_profit:,}원")
    kpi6.metric("오픈마켓 평균 순이익", f"{avg_11_profit:,}원")

    # ==========================================================================
    # 6. 필터링 및 결과 테이블 + CSV 다운로드
    # ==========================================================================
    st.subheader("3️⃣ 소싱 후보 필터링 및 다운로드 (최대 1,000개 표시)")
    f_col1, f_col2, f_col3 = st.columns([2, 2, 2])
    with f_col1:
        selected_statuses = st.multiselect(
            "표시할 리스크 등급",
            options=["🟢 안전 통과", "🟡 KC 확인 필요", "🔴 제외 권장"],
            default=["🟢 안전 통과", "🟡 KC 확인 필요"],
        )
    with f_col2:
        min_supply, max_supply = st.slider(
            "도매 공급가 범위 필터 (원)",
            min_value=0,
            max_value=max(50000, int(supply_series.max())),
            value=(0, max(50000, int(supply_series.max()))),
            step=1000,
        )
    with f_col3:
        file_search_kw = st.text_input(
            "🔎 업로드 파일 내 상품명 검색",
            placeholder="예: 하부장, 선반 (띄어쓰기로 여러 단어)",
        ).strip()

    mask = (
        df_result["리스크 판정"].isin(selected_statuses)
        & (supply_series >= min_supply)
        & (supply_series <= max_supply)
    )

    # 파일 내 키워드 검색 적용 (입력한 모든 단어가 포함된 상품 필터링)
    if file_search_kw:
        for word in file_search_kw.split():
            mask = mask & df_result[col_name].astype(str).str.contains(word, case=False, na=False)

    df_filtered = df_result[mask]

    st.dataframe(df_filtered.head(1000), use_container_width=True, hide_index=True, height=700)

    csv_bytes = df_filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label=f"📥 4대 마켓 분석 결과 {len(df_filtered):,}개 CSV 다운로드 (엑셀 호환)",
        data=csv_bytes,
        file_name="4market_sourcing_result.csv",
        mime="text/csv",
    )