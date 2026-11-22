streamlit
pandas
openpyxl

"""
초보 오픈마켓 셀러를 위한 도매 상품 소싱 & 마진 자동 분석기 (Streamlit)
-------------------------------------------------------------------------
주요 기능:
1. 도매처(도매매, 오너클랜, 온채널 등) CSV/엑셀 파일 업로드 및 한국어 인코딩 자동 감지
2. 테스트용 샘플 데이터 내장 (파일이 없어도 즉시 실행 및 체험 가능)
3. 초보 셀러 회피 품목(유아·의료·지재권 위험) 및 KC 인증 주의 키워드 자동 필터링
4. 5대 오픈마켓(스마트스토어, 11번가, G마켓, 옥션, 카카오 톡딜) 수수료 및 과세유형 반영
5. 목표 마진율 기준 권장 판매가 자동 산출 및 필터링된 소싱 리스트 CSV 다운로드

실행 방법:
    pip install streamlit pandas openpyxl
    streamlit run sourcing_analyzer.py
"""

import io
import math
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. 기본 설정 및 상수 정의
# ==============================================================================
st.set_page_config(
    page_title="도매 소싱 마진·리스크 자동 분석기",
    page_icon="📦",
    layout="wide",
)

# 5대 오픈마켓 평균 수수료율 (카테고리별 상이할 수 있으며, 사이드바에서 수정 가능)
DEFAULT_MARKET_FEES = {
    "네이버 스마트스토어 (쇼핑연동 기준)": 5.8,
    "11번가 (오픈마켓 평균)": 13.0,
    "G마켓 (ESM+ 평균)": 13.0,
    "옥션 (ESM+ 평균)": 13.0,
    "카카오 톡딜 (기본+결제 평균)": 12.0,
}

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
# 2. 유틸리티 함수 (파일 로드, 샘플 생성, 리스크 판별, 마진 계산)
# ==============================================================================
def create_sample_dataframe() -> pd.DataFrame:
    """도매처 CSV가 없는 초보 셀러가 즉시 테스트할 수 있는 샘플 데이터 생성"""
    sample_data = [
        {"도매처": "오너클랜", "상품코드": "W1001", "상품명": "차량용 성에제거 브러시 2in1 고급형", "공급가": 11500, "택배비": 3000, "시중가": 21900},
        {"도매처": "도매매", "상품코드": "D2045", "상품명": "사무실 극세사 무릎담요 대형 1+1 세트", "공급가": 13000, "택배비": 3000, "시중가": 24900},
        {"도매처": "온채널", "상품코드": "O3312", "상품명": "현관문 외풍차단 문풍지 틈새막이 대용량", "공급가": 8500, "택배비": 2500, "시중가": 16500},
        {"도매처": "오너클랜", "상품코드": "W1088", "상품명": "USB 충전식 차량용 열선 방석 시트", "공급가": 16000, "택배비": 3000, "시중가": 29900},
        {"도매처": "도매토피아", "상품코드": "T5521", "상품명": "유아용 실리콘 캐릭터 빨대컵 (디즈니)", "공급가": 7000, "택배비": 3000, "시중가": 15900},
        {"도매처": "도매매", "상품코드": "D2099", "상품명": "겨울 패딩 이불 압축팩 특대형 5매 세트", "공급가": 9900, "택배비": 3000, "시중가": 18900},
        {"도매처": "스페셜오퍼", "상품코드": "S7710", "상품명": "다용도 욕실 미끄럼방지 매트 6P 세트", "공급가": 12000, "택배비": 3000, "시중가": 22500},
        {"도매처": "도매매", "상품코드": "D2110", "상품명": "체지방 감소 다이어트 건강 보조 밴드", "공급가": 14000, "택배비": 3000, "시중가": 28000},
    ]
    return pd.DataFrame(sample_data)


def load_uploaded_file(uploaded_file) -> pd.DataFrame:
    """
    업로드된 CSV 또는 Excel 파일을 안전하게 로드합니다.
    한국 도매사이트 특유의 인코딩(cp949, euc-kr, utf-8-sig) 오류를 자동으로 처리합니다.
    """
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
    """콤마(,)나 '원' 문자가 섞인 가격 컬럼을 안전하게 숫자로 변환"""
    cleaned = (
        series.astype(str)
        .str.replace(r"[^0-9.]", "", regex=True)
        .replace("", "0")
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def check_product_risk(name: str, risk_keywords: list, kc_keywords: list) -> tuple[str, str]:
    """
    상품명에서 금지어(지재권/유아/의료)와 KC 인증 주의 키워드를 탐지합니다.
    반환값: (리스크 등급, 탐지된 사유)
    """
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
) -> tuple[int, int, float]:
    """
    목표 순마진율을 달성하기 위한 권장 판매가(100원 단위 올림)와
    실제 예상 순이익, 실제 마진율을 계산합니다.
    """
    fee_rate = fee_rate_pct / 100.0
    target_margin = target_margin_pct / 100.0

    # 배송비 부담 구조 계산
    # - 무료배송: 도매 택배비 전액을 판매가에 녹임
    # - 유료배송: 고객에게 받은 배송비와 도매 택배비 차액 + 배송비 결제수수료(약 3.63%) 반영
    if shipping_mode == "무료배송 (판매가에 택배비 포함)":
        net_shipping_cost = shipping_cost
        extra_shipping_revenue = 0.0
        shipping_fee = 0.0
    else:
        net_shipping_cost = shipping_cost
        extra_shipping_revenue = customer_shipping
        shipping_fee = customer_shipping * 0.0363

    # 고정 원가 성격의 순비용
    base_cost = supply_price + net_shipping_cost + shipping_fee - extra_shipping_revenue

    # 과세 유형에 따른 분모 계산
    if tax_type == "일반과세자 (부가세 10% 반영)":
        # 순이익 = (판매가 - 판매가*수수료 - base_cost) / 1.1 >= 판매가 * 목표마진율
        denominator = 1.0 - fee_rate - (1.1 * target_margin)
    elif tax_type == "간이과세자 (부가세 약 1.5% 반영)":
        # 순이익 = 판매가 - 판매가*수수료 - 판매가*0.015 - base_cost >= 판매가 * 목표마진율
        denominator = 1.0 - fee_rate - 0.015 - target_margin
    else:
        # 간이과세자 (연매출 4,800만 미만 납부면제 기준)
        denominator = 1.0 - fee_rate - target_margin

    if denominator <= 0.05:
        # 수수료 + 목표 마진이 지나치게 높아 계산 불가한 경우 안전 처리
        return 0, 0, 0.0

    raw_price = base_cost / denominator
    # 100원 단위 올림 처리 (예: 21,340원 -> 21,400원)
    recommended_price = max(0, int(math.ceil(raw_price / 100.0) * 100))

    # 올림된 권장 판매가 기준으로 실제 순이익 및 마진율 역산
    market_fee = recommended_price * fee_rate
    pre_tax_profit = recommended_price - market_fee - base_cost

    if tax_type == "일반과세자 (부가세 10% 반영)":
        net_profit = pre_tax_profit / 1.1
    elif tax_type == "간이과세자 (부가세 약 1.5% 반영)":
        net_profit = pre_tax_profit - ((recommended_price + extra_shipping_revenue) * 0.015)
    else:
        net_profit = pre_tax_profit

    actual_margin_pct = (net_profit / recommended_price * 100.0) if recommended_price > 0 else 0.0
    return recommended_price, int(round(net_profit)), round(actual_margin_pct, 1)


# ==============================================================================
# 3. 사이드바 UI (마켓·마진·과세·필터 설정)
# ==============================================================================
with st.sidebar:
    st.header("⚙️ 소싱 & 마진 설정")

    selected_market = st.selectbox(
        "판매할 마켓 선택",
        options=list(DEFAULT_MARKET_FEES.keys()),
        index=0,
    )
    fee_rate_pct = st.number_input(
        "마켓 수수료율 (%)",
        min_value=0.0,
        max_value=30.0,
        value=float(DEFAULT_MARKET_FEES[selected_market]),
        step=0.1,
        help="카테고리별 실제 수수료에 맞게 조정할 수 있습니다.",
    )

    target_margin_pct = st.slider(
        "목표 순마진율 (%)",
        min_value=5,
        max_value=50,
        value=25,
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
        index=2,  # 일반과세자 기본값 설정
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
st.title("📦 도매 소싱 마진·리스크 자동 분석기")
st.caption("도매처 상품 엑셀/CSV를 올리면 초보 셀러 리스크 필터링과 5대 마켓 권장 판매가를 1초 만에 계산합니다.")

# 데이터 소스 선택 (파일 업로드 vs 샘플 데이터)
col_upload, col_sample = st.columns([3, 1])
with col_upload:
    uploaded_file = st.file_uploader(
        "도매처 상품 파일 업로드 (CSV, XLSX, XLS)",
        type=["csv", "xlsx", "xls"],
        help="도매매, 오너클랜, 온채널 등에서 다운로드한 상품 목록을 올려주세요.",
    )
with col_sample:
    st.write("")  # 여백 맞춤
    st.write("")
    use_sample = st.checkbox("샘플 데이터로 바로 체험하기", value=(uploaded_file is None))

# 데이터프레임 로드
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

# 데이터가 준비된 경우 컬럼 매핑 및 분석 실행
if df_raw is not None and not df_raw.empty:
    st.divider()
    st.subheader("1️⃣ 컬럼 매핑 (도매처 양식 맞추기)")
    st.caption("도매처마다 열 이름이 다르므로, 아래에서 해당하는 열을 선택해 주세요.")

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

    # 분석 계산 수행
    df_result = df_raw.copy()
    supply_series = clean_numeric(df_result[col_supply])

    if col_shipping == "(고정 3,000원 적용)":
        shipping_series = pd.Series([3000.0] * len(df_result))
    else:
        shipping_series = clean_numeric(df_result[col_shipping])

    risk_statuses, risk_reasons = [], []
    rec_prices, net_profits, margin_rates = [], [], []

    for idx in range(len(df_result)):
        p_name = df_result.iloc[idx][col_name]
        s_price = float(supply_series.iloc[idx])
        s_ship = float(shipping_series.iloc[idx])

        status, reason = check_product_risk(p_name, risk_kw_list, kc_kw_list)
        rec_p, profit, m_rate = calculate_recommended_price(
            supply_price=s_price,
            shipping_cost=s_ship,
            fee_rate_pct=fee_rate_pct,
            target_margin_pct=target_margin_pct,
            shipping_mode=shipping_mode,
            customer_shipping=customer_shipping,
            tax_type=tax_type,
        )

        risk_statuses.append(status)
        risk_reasons.append(reason)
        rec_prices.append(rec_p)
        net_profits.append(profit)
        margin_rates.append(m_rate)

    df_result["리스크 판정"] = risk_statuses
    df_result["판정 사유"] = risk_reasons
    df_result["권장 판매가(원)"] = rec_prices
    df_result["예상 순이익(원)"] = net_profits
    df_result["예상 마진율(%)"] = margin_rates

    # 시중가 대비 가격 경쟁력 자동 판정 (시중가 열이 선택된 경우)
    if col_market_price != "(시중가 비교 안 함)":
        mkt_series = clean_numeric(df_result[col_market_price])
        competitiveness = []
        for rec_p, mkt_p in zip(rec_prices, mkt_series):
            if mkt_p <= 0:
                competitiveness.append("비교불가")
            elif rec_p <= mkt_p:
                diff_pct = round((mkt_p - rec_p) / mkt_p * 100, 1)
                competitiveness.append(f"✅ 경쟁력 있음 (-{diff_pct}%)")
            else:
                diff_pct = round((rec_p - mkt_p) / mkt_p * 100, 1)
                competitiveness.append(f"⚠️ 시중가 초과 (+{diff_pct}%)")
        df_result["가격 경쟁력"] = competitiveness

    # ==========================================================================
    # 5. 요약 KPI 대시보드
    # ==========================================================================
    st.divider()
    st.subheader(f"2️⃣ 분석 결과 요약 ({selected_market})")

    total_cnt = len(df_result)
    safe_cnt = (df_result["리스크 판정"] == "🟢 안전 통과").sum()
    kc_cnt = (df_result["리스크 판정"] == "🟡 KC 확인 필요").sum()
    risk_cnt = (df_result["리스크 판정"] == "🔴 제외 권장").sum()
    avg_profit = int(df_result[df_result["리스크 판정"] != "🔴 제외 권장"]["예상 순이익(원)"].mean() or 0)

    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("전체 상품 수", f"{total_cnt:,}개")
    kpi2.metric("🟢 안전 통과", f"{safe_cnt:,}개")
    kpi3.metric("🟡 KC 확인 필요", f"{kc_cnt:,}개")
    kpi4.metric("🔴 제외 권장", f"{risk_cnt:,}개")
    kpi5.metric("통과 상품 평균 순이익", f"{avg_profit:,}원")

    # ==========================================================================
    # 6. 필터링 및 결과 테이블 + CSV 다운로드
    # ==========================================================================
    st.subheader("3️⃣ 소싱 후보 필터링 및 다운로드")
    f_col1, f_col2 = st.columns([2, 2])
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

    mask = (
        df_result["리스크 판정"].isin(selected_statuses)
        & (supply_series >= min_supply)
        & (supply_series <= max_supply)
    )
    df_filtered = df_result[mask]

    # 최대 1,000개까지 표시 및 표 높이 확대
    st.dataframe(df_filtered.head(1000), use_container_width=True, hide_index=True, height=700)

    # 엑셀에서 한글이 깨지지 않도록 utf-8-sig 인코딩으로 CSV 다운로드 제공
    csv_bytes = df_filtered.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label=f"📥 필터링된 {len(df_filtered):,}개 상품 CSV 다운로드 (엑셀 호환)",
        data=csv_bytes,
        file_name="sourcing_filtered_result.csv",
        mime="text/csv",
    )