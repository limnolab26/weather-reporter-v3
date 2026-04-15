# app.py — 기상자료 분석 시스템 v4.0
# 실행: streamlit run app.py

import io
import datetime as dt

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import streamlit as st

from data_processor import WeatherDataProcessor, add_time_columns, NUMERIC_COLUMNS

# 분석 모듈 임포트
import analysis_temp
import analysis_precip
import analysis_solar
import analysis_wind
import analysis_agri
import analysis_climate
from chart_utils import chart_download_btn

# Plotly 전역 글씨 크기 설정 (모든 그래프에 적용)
_base = pio.templates["plotly"]
_base.layout.font = dict(size=14)
_base.layout.title = dict(font=dict(size=16))
pio.templates.default = "plotly"

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 페이지 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.set_page_config(
    page_title="기상자료 분석 시스템 v4.0",
    layout="wide"
)

# 글씨 크기 확대 (시니어 연구자용)
st.markdown("""
<style>
/* ── 탭 메뉴 글씨: 버튼 자체 + 내부 p 태그 모두 지정 ── */
[data-baseweb="tab"],
[data-testid="stTab"],
.stTabs button[role="tab"] {
    font-size: 17px !important;
    font-weight: 600;
    padding: 8px 18px !important;
}
/* 실제 텍스트가 담긴 p 태그를 직접 타겟 */
[data-baseweb="tab"] p,
[data-testid="stTab"] p,
[data-testid="stTab"] [data-testid="stMarkdownContainer"] p {
    font-size: 17px !important;
    font-weight: 600;
    line-height: 1.3;
}
.stTabs [data-baseweb="tab-list"],
div[data-testid="stTabs"] [role="tablist"] {
    gap: 4px;
    flex-wrap: wrap;
}
/* ── 사이드바 텍스트 ── */
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown span {
    font-size: 15px !important;
}
/* ── 데이터프레임 표 글씨 ── */
[data-testid="stDataFrame"] div[role="gridcell"],
[data-testid="stDataFrame"] div[role="columnheader"] {
    font-size: 14px !important;
}
/* ── 일반 본문 글씨 ── */
.stMarkdown p, .stMarkdown li, .stText {
    font-size: 15px !important;
}
</style>
""", unsafe_allow_html=True)

st.title("🌦️ 기상자료 분석 시스템 v4.0")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 세션 상태
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

for key in ["df", "monthly_df", "climate_df"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 공통 상수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ELEMENT_LABELS = {
    "temp_avg": "평균기온 (°C)",
    "temp_max": "최고기온 (°C)",
    "temp_min": "최저기온 (°C)",
    "dew_point": "이슬점온도 (°C)",
    "frost_temp": "초상온도 (°C)",
    "precipitation": "강수량 (mm)",
    "precip_1hr_max": "1시간최다강수 (mm)",
    "humidity": "습도 (%)",
    "humidity_min": "최소습도 (%)",
    "vapor_pressure": "증기압 (hPa)",
    "wind_speed": "평균풍속 (m/s)",
    "wind_max": "최대풍속 (m/s)",
    "wind_gust": "최대순간풍속 (m/s)",
    "wind_run_100m": "풍정합100m (m)",
    "pressure_sea": "해면기압 (hPa)",
    "pressure_local": "현지기압 (hPa)",
    "sunshine": "일조시간 (hr)",
    "daylight_hours": "가조시간 (hr)",
    "solar_rad": "일사량 (MJ/m²)",
    "cloud_cover": "전운량 (1/10)",
    "snowfall": "신적설 (cm)",
    "snow_depth": "적설깊이 (cm)",
    "evaporation_large": "대형증발량 (mm)",
    "evaporation_small": "소형증발량 (mm)",
    "soil_temp_surface": "지면온도 (°C)",
    "soil_temp_5cm": "5cm 지중온도 (°C)",
    "soil_temp_10cm": "10cm 지중온도 (°C)",
    "soil_temp_20cm": "20cm 지중온도 (°C)",
    "soil_temp_30cm": "30cm 지중온도 (°C)",
    "soil_temp_50cm": "0.5m 지중온도 (°C)",
    "soil_temp_100cm": "1.0m 지중온도 (°C)",
    "soil_temp_150cm": "1.5m 지중온도 (°C)",
    "soil_temp_300cm": "3.0m 지중온도 (°C)",
    "soil_temp_500cm": "5.0m 지중온도 (°C)",
    "fog_duration": "안개시간 (hr)",
}

ELEMENT_OPTIONS = {v: k for k, v in ELEMENT_LABELS.items()}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 유틸리티 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def extract_station_name(file) -> str:
    name = file.name.split(".")[0]
    return name.split("_")[0]


@st.cache_data
def load_multiple_files(uploaded_files):
    all_df = []
    for file in uploaded_files:
        station_name = extract_station_name(file)

        if file.name.endswith(".csv"):
            try:
                raw_df = pd.read_csv(file, encoding="cp949")
            except UnicodeDecodeError:
                file.seek(0)
                raw_df = pd.read_csv(file, encoding="utf-8")
        elif file.name.endswith(".xlsx"):
            raw_df = pd.read_excel(file)
        else:
            continue

        processor = WeatherDataProcessor()
        processed_df = processor.process(raw_df)

        if "station_name" not in processed_df.columns:
            processed_df = processed_df.copy()
            processed_df["station_name"] = station_name

        all_df.append(processed_df)

    if not all_df:
        return None

    return pd.concat(all_df, ignore_index=True)


@st.cache_data
def create_monthly_data(df: pd.DataFrame) -> pd.DataFrame:
    all_agg = {
        "temp_avg": "mean", "temp_max": "mean", "temp_min": "mean",
        "dew_point": "mean", "frost_temp": "mean",
        "precipitation": "sum", "precip_1hr_max": "max",
        "humidity": "mean", "humidity_min": "mean", "vapor_pressure": "mean",
        "wind_speed": "mean", "wind_max": "mean", "wind_gust": "mean",
        "wind_run_100m": "sum",
        "pressure_sea": "mean", "pressure_local": "mean",
        "sunshine": "sum", "daylight_hours": "sum", "solar_rad": "sum",
        "cloud_cover": "mean", "cloud_low": "mean",
        "snowfall": "sum", "snow_depth": "max",
        "evaporation_large": "sum", "evaporation_small": "sum",
        "soil_temp_surface": "mean", "soil_temp_5cm": "mean",
        "soil_temp_10cm": "mean", "soil_temp_20cm": "mean",
        "soil_temp_30cm": "mean", "soil_temp_50cm": "mean",
        "soil_temp_100cm": "mean", "soil_temp_150cm": "mean",
        "soil_temp_300cm": "mean", "soil_temp_500cm": "mean",
    }
    agg = {col: func for col, func in all_agg.items() if col in df.columns}

    monthly = (
        df.groupby(["station_name", "year", "month"])
        .agg(agg)
        .reset_index()
    )
    monthly["date"] = pd.to_datetime(
        monthly["year"].astype(str) + "-" + monthly["month"].astype(str) + "-01"
    )
    return monthly


@st.cache_data
def calculate_climate_normals(df: pd.DataFrame, start_year: int, end_year: int) -> pd.DataFrame:
    climate_df = df[(df["year"] >= start_year) & (df["year"] <= end_year)]
    all_agg = {
        "temp_avg": "mean", "temp_max": "mean", "temp_min": "mean",
        "precipitation": "mean", "humidity": "mean", "wind_speed": "mean",
        "wind_max": "mean", "sunshine": "mean", "solar_rad": "mean",
        "snowfall": "mean", "cloud_cover": "mean",
        "pressure_sea": "mean", "pressure_local": "mean",
    }
    agg = {col: func for col, func in all_agg.items() if col in climate_df.columns}
    return (
        climate_df.groupby(["station_name", "month"])
        .agg(agg)
        .reset_index()
    )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 한글 폰트 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def setup_korean_font():
    candidates = [
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "NanumGothic.ttf",
        "malgun.ttf",
    ]
    for path in candidates:
        try:
            fm.fontManager.addfont(path)
            prop = fm.FontProperties(fname=path)
            plt.rcParams["font.family"] = prop.get_name()
            plt.rcParams["axes.unicode_minus"] = False
            return
        except Exception:
            continue
    for font in ["NanumGothic", "Malgun Gothic", "AppleGothic"]:
        if font in [f.name for f in fm.fontManager.ttflist]:
            plt.rcParams["font.family"] = font
            break
    plt.rcParams["axes.unicode_minus"] = False


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Sidebar
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.sidebar.header("📂 데이터 업로드")

uploaded_files = st.sidebar.file_uploader(
    "ASOS CSV 또는 XLSX 업로드",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

st.sidebar.link_button(
    "🌐 기상자료개방포털 ASOS 자료",
    "https://data.kma.go.kr/data/grnd/selectAsosRltmList.do?pgmNo=36"
)

st.sidebar.header("📆 평년기간 설정")
climate_start = st.sidebar.number_input("평년 시작연도", min_value=1950, max_value=2100, value=1991)
climate_end = st.sidebar.number_input("평년 종료연도", min_value=1950, max_value=2100, value=2020)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 데이터 로딩
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if uploaded_files:
    df = load_multiple_files(uploaded_files)
    if df is not None:
        st.session_state.df = df
        st.session_state.monthly_df = create_monthly_data(df)
        st.session_state.climate_df = calculate_climate_normals(
            st.session_state.monthly_df, climate_start, climate_end
        )

# 날짜 필터
if st.session_state.df is not None:
    min_date = st.session_state.df["date"].min()
    max_date = st.session_state.df["date"].max()
    st.sidebar.info(f"📅 자료 기간: {min_date.strftime('%Y.%m.%d')} ~ {max_date.strftime('%Y.%m.%d')}")
    date_range = st.sidebar.date_input(
        "분석 기간 선택",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
    if len(date_range) == 2:
        start_date, end_date = date_range
        mask = (
            (st.session_state.df["date"] >= pd.to_datetime(start_date)) &
            (st.session_state.df["date"] <= pd.to_datetime(end_date))
        )
        filtered_df = st.session_state.df[mask].copy()
    else:
        filtered_df = st.session_state.df.copy()
else:
    filtered_df = None

# 관련 사이트 — 사이드바 제일 하단
st.sidebar.markdown("---")
st.sidebar.markdown("**🔗 관련 사이트**")
st.sidebar.link_button("🌤️ 기상청 날씨누리", "https://www.weather.go.kr/w/index.do")
st.sidebar.link_button("💧 국가가뭄정보포털", "https://www.drought.go.kr/main.do")
st.sidebar.link_button("🌍 기후변화 상황지도", "https://climate.go.kr/atlas/")
st.sidebar.link_button("🌾 농업가뭄관리시스템", "https://adms.ekr.or.kr/page/main.do")
st.sidebar.link_button("🌐 earth 전세계 기상지도", "https://earth.nullschool.net/ko/#current/particulates/surface/level/overlay=pm2.5/orthographic=-238.18,22.58,577")
st.sidebar.link_button("⛰️ 산악기상정보시스템", "https://mtweather.nifos.go.kr/")
st.sidebar.link_button("💨 에어코리아", "https://www.airkorea.or.kr/web/realSearch?pMENU_NO=97")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 탭 생성 (8탭)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

(tab_general, tab_temp, tab_precip, tab_wind,
 tab_solar, tab_agri, tab_climate, tab_download, tab_help) = st.tabs([
    "📊 종합 분석",
    "🌡️ 기온 분석",
    "🌧️ 강수량 분석",
    "💨 바람 분석",
    "☀️ 태양광 분석",
    "🌱 농업기상 분석",
    "📈 기후변화 분석",
    "⬇️ 보고서 다운로드",
    "ℹ️ 사용 방법",
])


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB1 — 종합 분석 (대화형 차트 + 평년 분석)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def prepare_chart_data(df, element, freq):
    if freq == "D":
        grouped = df.groupby(["station_name", "date"])[element].mean().reset_index()
    elif freq == "M":
        grouped = df.groupby(["station_name", "year", "month"])[element].mean().reset_index()
        grouped["date"] = pd.to_datetime(
            grouped["year"].astype(str) + "-" + grouped["month"].astype(str) + "-01"
        )
    else:  # Y
        grouped = df.groupby(["station_name", "year"])[element].mean().reset_index()
        grouped["date"] = pd.to_datetime(grouped["year"].astype(str) + "-01-01")
    return grouped


def create_plotly_chart(df, element, chart_type):
    y_label = ELEMENT_LABELS.get(element, element)
    if chart_type == "선":
        fig = px.line(df, x="date", y=element, color="station_name",
                      markers=True, labels={"date": "날짜", element: y_label})
    else:
        fig = px.bar(df, x="date", y=element, color="station_name",
                     labels={"date": "날짜", element: y_label})
    fig.update_layout(height=420)
    return fig


def create_temp_precip_combo(df):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    for stn in df["station_name"].unique():
        sdf = df[df["station_name"] == stn]
        fig.add_trace(go.Scatter(x=sdf["date"], y=sdf.get("temp_avg"),
                                 name=f"{stn} 평균기온", mode="lines+markers"),
                      secondary_y=False)
        fig.add_trace(go.Bar(x=sdf["date"], y=sdf.get("precipitation"),
                             name=f"{stn} 강수량", opacity=0.5), secondary_y=True)
    fig.update_yaxes(title_text="기온 (°C)", secondary_y=False)
    fig.update_yaxes(title_text="강수량 (mm)", secondary_y=True)
    fig.update_layout(height=420, legend_title="관측소")
    return fig


def create_comparison_chart(df, selected_elements, element_chart_types, freq):
    fig = go.Figure()
    for element in selected_elements:
        if element not in df.columns:
            continue
        el_df = prepare_chart_data(df, element, freq)
        label = ELEMENT_LABELS.get(element, element)
        chart_t = element_chart_types.get(element, "선")
        for stn in df["station_name"].unique():
            sdf = el_df[el_df["station_name"] == stn]
            name = f"{stn} {label}"
            if chart_t == "선":
                fig.add_trace(go.Scatter(x=sdf["date"], y=sdf[element],
                                         name=name, mode="lines+markers"))
            else:
                fig.add_trace(go.Bar(x=sdf["date"], y=sdf[element],
                                     name=name, opacity=0.75))
    fig.update_layout(height=450, legend_title="관측소/요소", barmode="group")
    return fig


def calculate_anomaly(df, climate_df, element):
    merged = pd.merge(df, climate_df, on=["station_name", "month"], suffixes=("", "_normal"))
    merged = merged.copy()
    merged["anomaly"] = merged[element] - merged[f"{element}_normal"]
    return merged


with tab_general:
    if filtered_df is None:
        st.markdown("""
<style>
.hero-wrap {
    width: 100%;
    border-radius: 16px;
    overflow: hidden;
    background: linear-gradient(135deg, #0a2342 0%, #1a4a7a 40%, #2d7dd2 70%, #4ab3f4 100%);
    padding: 60px 48px 50px 48px;
    margin-bottom: 24px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.18);
    position: relative;
}
.hero-title {
    font-size: 2.6rem;
    font-weight: 800;
    color: #ffffff;
    margin: 0 0 12px 0;
    text-shadow: 0 2px 8px rgba(0,0,0,0.3);
    letter-spacing: -0.5px;
}
.hero-sub {
    font-size: 1.15rem;
    color: #cce8ff;
    margin: 0 0 36px 0;
    line-height: 1.7;
}
.hero-icons {
    font-size: 2.2rem;
    letter-spacing: 10px;
    margin-bottom: 32px;
    display: block;
}
.hero-steps {
    display: flex;
    gap: 20px;
    flex-wrap: wrap;
}
.hero-step {
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 12px;
    padding: 16px 22px;
    color: #ffffff;
    font-size: 0.95rem;
    flex: 1;
    min-width: 160px;
    backdrop-filter: blur(4px);
}
.hero-step-num {
    font-size: 1.6rem;
    font-weight: 700;
    color: #7dd3fc;
    display: block;
    margin-bottom: 4px;
}
.hero-deco {
    position: absolute;
    right: 40px;
    top: 30px;
    font-size: 7rem;
    opacity: 0.15;
    line-height: 1;
    pointer-events: none;
    user-select: none;
}
</style>
<div class="hero-wrap">
    <div class="hero-deco">🌦️</div>
    <span class="hero-icons">🌡️ 🌧️ ☀️ 💨 🌱 📈</span>
    <div class="hero-title">기상자료 분석 시스템</div>
    <div class="hero-sub">
        기상청 ASOS 관측자료를 업로드하면 기온·강수·바람·일사·농업기상·기후변화<br>
        등 다양한 분석 결과와 보고서를 자동으로 생성합니다.
    </div>
    <div class="hero-steps">
        <div class="hero-step">
            <span class="hero-step-num">① 업로드</span>
            왼쪽 사이드바에서<br>CSV / XLSX 파일 선택
        </div>
        <div class="hero-step">
            <span class="hero-step-num">② 분석</span>
            상단 탭에서<br>원하는 분석 항목 선택
        </div>
        <div class="hero-step">
            <span class="hero-step-num">③ 다운로드</span>
            Excel·DOCX 보고서<br>또는 차트 XLSX 저장
        </div>
    </div>
</div>
""", unsafe_allow_html=True)
    else:
        gen_t1, gen_t2, gen_t3, gen_t4 = st.tabs(["📊 대화형 차트", "📆 평년 분석", "📅 월평균 분석", "📅 연평균 분석"])

        # ── 대화형 차트 ──
        with gen_t1:
            st.subheader("📊 대화형 차트")
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                stations = sorted(filtered_df["station_name"].unique())
                sel_stations = st.multiselect("관측소 선택", stations, default=stations, key="gen_stations")

            with col2:
                available_elements = {
                    label: col for label, col in ELEMENT_OPTIONS.items()
                    if col in filtered_df.columns
                }
                sel_elem_names = st.multiselect(
                    "기상요소 선택",
                    list(available_elements.keys()),
                    default=[list(available_elements.keys())[0]] if available_elements else [],
                    key="gen_element"
                )
                sel_elements = [available_elements[n] for n in sel_elem_names]

            with col3:
                freq_label = st.selectbox("집계 단위", ["일별", "월별", "연별"], key="gen_freq")
                freq = {"일별": "D", "월별": "M", "연별": "Y"}[freq_label]

            with col4:
                chart_type = st.selectbox(
                    "차트 유형", ["선", "막대", "복합(기온+강수)", "비교(다중요소)"],
                    key="gen_chart_type"
                )

            chart_df = filtered_df[filtered_df["station_name"].isin(sel_stations)]

            if len(chart_df) == 0:
                st.warning("선택된 관측소 데이터가 없습니다.")
            elif not sel_elements and chart_type != "복합(기온+강수)":
                st.info("기상요소를 하나 이상 선택하세요.")
            elif chart_type == "복합(기온+강수)":
                agg_cols = {k: v for k, v in {"temp_avg": "mean", "precipitation": "sum"}.items() if k in chart_df.columns}
                combo_df = chart_df.groupby(["station_name", "year", "month"]).agg(agg_cols).reset_index()
                combo_df["date"] = pd.to_datetime(
                    combo_df["year"].astype(str) + "-" + combo_df["month"].astype(str) + "-01"
                )
                _fig_combo = create_temp_precip_combo(combo_df)
                st.plotly_chart(_fig_combo, use_container_width=True)
                chart_download_btn(_fig_combo, key="app_combo_chart", filename="temp_precip_combo")

            elif chart_type == "비교(다중요소)":
                if sel_elements:
                    with st.expander("요소별 차트 종류 설정", expanded=True):
                        el_types = {}
                        cols = st.columns(min(len(sel_elem_names), 5))
                        for i, name in enumerate(sel_elem_names):
                            with cols[i % 5]:
                                t = st.selectbox(name, ["선", "막대"], key=f"gen_eltype_{name}")
                                el_types[available_elements[name]] = t
                    available_sel = [e for e in sel_elements if e in chart_df.columns]
                    if available_sel:
                        _fig_cmp = create_comparison_chart(chart_df, available_sel, el_types, freq)
                        st.plotly_chart(_fig_cmp, use_container_width=True)
                        chart_download_btn(_fig_cmp, key="app_compare_chart", filename="element_comparison")
            else:
                for element in sel_elements:
                    if element not in chart_df.columns:
                        st.warning(f"'{ELEMENT_LABELS.get(element, element)}' 데이터가 없습니다.")
                        continue
                    el_df = prepare_chart_data(chart_df, element, freq)
                    _fig_el = create_plotly_chart(el_df, element, chart_type)
                    st.plotly_chart(_fig_el, use_container_width=True)
                    chart_download_btn(_fig_el, key=f"app_el_{element}_chart", filename=f"chart_{element}")

        # ── 평년 분석 ──
        with gen_t2:
            st.subheader("📆 평년 분석")
            climate_df = st.session_state.climate_df
            if climate_df is None:
                st.warning("평년 데이터가 없습니다.")
            else:
                available_cl = {
                    label: col for label, col in ELEMENT_OPTIONS.items()
                    if col in climate_df.columns
                }
                cl_elem_name = st.selectbox("기상요소 선택", list(available_cl.keys()), key="cl_element")
                cl_element = available_cl[cl_elem_name]

                # 월별 평년표
                st.markdown("### 📋 월별 평년값")
                pivot = climate_df.pivot_table(values=cl_element, index="month", columns="station_name")
                st.dataframe(pivot.style.format("{:.2f}"), use_container_width=True, height=460)

                csv_cl = pivot.to_csv(encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button("⬇️ 평년 CSV", csv_cl, "climate_normals.csv", "text/csv", key="cl_csv")

                # 평년 그래프
                st.markdown("### 📈 평년 그래프")
                fig_cl = px.line(climate_df, x="month", y=cl_element, color="station_name",
                                 markers=True, labels={"month": "월"})
                fig_cl.update_layout(height=380)
                st.plotly_chart(fig_cl, use_container_width=True)
                chart_download_btn(fig_cl, key="app_climate_normal_chart", filename="climate_normals_chart")

                # 편차
                st.markdown("### 📉 평년 대비 편차")
                if st.session_state.monthly_df is not None and cl_element in st.session_state.monthly_df.columns:
                    anomaly_df = calculate_anomaly(st.session_state.monthly_df, climate_df, cl_element)
                    fig_an = px.line(anomaly_df, x="date", y="anomaly", color="station_name",
                                     labels={"date": "날짜", "anomaly": "편차"})
                    fig_an.update_layout(height=380)
                    st.plotly_chart(fig_an, use_container_width=True)
                    chart_download_btn(fig_an, key="app_anomaly_chart", filename="climate_anomaly")

        # ── 월평균 분석 ──
        with gen_t3:
            st.subheader("📅 월평균 분석")
            available_m = {
                label: col for label, col in ELEMENT_OPTIONS.items()
                if col in filtered_df.columns
            }
            if not available_m:
                st.warning("분석 가능한 기상요소가 없습니다.")
            else:
                m_elem_name = st.selectbox("기상요소 선택", list(available_m.keys()), key="gen_t3_elem")
                m_element = available_m[m_elem_name]

                # 집계 방법 결정
                sum_cols = {"precipitation", "sunshine", "solar_rad", "snowfall",
                            "wind_run_100m", "daylight_hours", "evaporation_large", "evaporation_small"}
                agg_func = "sum" if m_element in sum_cols else "mean"

                group_cols = ["station_name", "year", "month"] if "station_name" in filtered_df.columns else ["year", "month"]
                m_df = (
                    filtered_df.groupby(group_cols)[m_element]
                    .agg(agg_func)
                    .reset_index()
                )

                st.markdown("### 📋 월별 평균표 (연도별)")
                stations_m = sorted(filtered_df["station_name"].unique()) if "station_name" in filtered_df.columns else ["전체"]
                sel_stn_m = st.selectbox("관측소", stations_m, key="gen_t3_stn")
                pivot_m = (
                    m_df[m_df["station_name"] == sel_stn_m]
                    .pivot_table(values=m_element, index="month", columns="year")
                    .round(2)
                )
                pivot_m.index.name = "월"
                pivot_m["평균"] = pivot_m.mean(axis=1).round(2)
                st.dataframe(pivot_m.style.format("{:.2f}"), use_container_width=True, height=460)

                csv_m = pivot_m.to_csv(encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button("⬇️ 월평균 CSV", csv_m, f"monthly_{m_element}.csv", "text/csv", key="gen_t3_csv")

                st.markdown("### 📈 월별 추이")
                fig_m = px.line(
                    m_df[m_df["station_name"] == sel_stn_m] if "station_name" in m_df.columns else m_df,
                    x="month", y=m_element, color="year",
                    labels={"month": "월", m_element: m_elem_name, "year": "연도"},
                    markers=True,
                )
                fig_m.update_layout(height=400, xaxis=dict(tickmode="linear", dtick=1))
                st.plotly_chart(fig_m, use_container_width=True)
                chart_download_btn(fig_m, key="app_monthly_trend_chart", filename="monthly_trend_chart")

        # ── 연평균 분석 ──
        with gen_t4:
            st.subheader("📅 연평균 분석")
            available_y = {
                label: col for label, col in ELEMENT_OPTIONS.items()
                if col in filtered_df.columns
            }
            if not available_y:
                st.warning("분석 가능한 기상요소가 없습니다.")
            else:
                y_elem_names = st.multiselect(
                    "기상요소 선택 (복수 가능)",
                    list(available_y.keys()),
                    default=[list(available_y.keys())[0]] if available_y else [],
                    key="gen_t4_elem"
                )
                if not y_elem_names:
                    st.info("기상요소를 하나 이상 선택하세요.")
                else:
                    sum_cols_y = {"precipitation", "sunshine", "solar_rad", "snowfall",
                                  "wind_run_100m", "daylight_hours", "evaporation_large", "evaporation_small"}

                    records = []
                    for y_elem_name in y_elem_names:
                        y_element = available_y[y_elem_name]
                        agg_f = "sum" if y_element in sum_cols_y else "mean"
                        group_y = ["station_name", "year"] if "station_name" in filtered_df.columns else ["year"]
                        annual_y = (
                            filtered_df.groupby(group_y)[y_element]
                            .agg(agg_f)
                            .reset_index()
                        )
                        annual_y["요소"] = y_elem_name
                        annual_y = annual_y.rename(columns={y_element: "값"})
                        records.append(annual_y)

                    if records:
                        annual_all = pd.concat(records, ignore_index=True)

                        st.markdown("### 📈 연별 추이")
                        if "station_name" in annual_all.columns:
                            fig_y = px.line(
                                annual_all, x="year", y="값",
                                color="station_name", facet_row="요소",
                                labels={"year": "연도", "값": "값"},
                                markers=True,
                            )
                        else:
                            fig_y = px.line(
                                annual_all, x="year", y="값",
                                facet_row="요소",
                                labels={"year": "연도", "값": "값"},
                                markers=True,
                            )
                        fig_y.update_layout(height=max(350, 320 * len(y_elem_names)))
                        fig_y.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
                        st.plotly_chart(fig_y, use_container_width=True)
                        chart_download_btn(fig_y, key="app_annual_trend_chart", filename="annual_trend_chart")

                        st.markdown("### 📋 연별 통계표")
                        for y_elem_name in y_elem_names:
                            y_element = available_y[y_elem_name]
                            agg_f = "sum" if y_element in sum_cols_y else "mean"
                            group_y = ["station_name", "year"] if "station_name" in filtered_df.columns else ["year"]
                            tbl = (
                                filtered_df.groupby(group_y)[y_element]
                                .agg(agg_f)
                                .reset_index()
                                .round(2)
                            )
                            if "station_name" in tbl.columns:
                                tbl = tbl.pivot(index="year", columns="station_name", values=y_element)
                            else:
                                tbl = tbl.set_index("year")
                            tbl.index.name = "연도"
                            tbl["평균"] = tbl.mean(axis=1).round(2)
                            st.markdown(f"**{y_elem_name}**")
                            st.dataframe(tbl.style.format("{:.2f}"), use_container_width=True)

                        csv_y = annual_all.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                        st.download_button("⬇️ 연평균 CSV", csv_y, "annual_stats.csv", "text/csv", key="gen_t4_csv")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB2~6 — 전문 분석 탭 (모듈 호출)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab_temp:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        analysis_temp.render(filtered_df)

with tab_precip:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        analysis_precip.render(filtered_df)

with tab_wind:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        analysis_wind.render(filtered_df)

with tab_solar:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        analysis_solar.render(filtered_df)

with tab_agri:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        analysis_agri.render(filtered_df)

with tab_climate:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        analysis_climate.render(filtered_df)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB8 — 보고서 다운로드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab_download:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        st.subheader("⬇️ 보고서 다운로드")
        st.markdown("선택한 데이터 기반으로 **Excel 또는 DOCX 보고서**를 생성합니다.")

        col1, col2 = st.columns(2)

        with col1:
            if st.button("📊 Excel 보고서 생성", key="excel_btn"):
                try:
                    from excel_generator import ExcelReportGenerator
                    gen = ExcelReportGenerator()

                    progress_bar = st.progress(0, text="Excel 보고서 생성 준비 중...")

                    def _on_progress(value, msg):
                        progress_bar.progress(
                            min(value, 1.0),
                            text=f"📊 {msg}" if msg else "생성 중...",
                        )

                    excel_bytes = gen.generate_excel(
                        df=st.session_state.df,
                        monthly_df=st.session_state.monthly_df,
                        climate_df=st.session_state.climate_df,
                        progress_callback=_on_progress,
                    )
                    progress_bar.empty()
                    st.download_button(
                        "⬇️ Excel 다운로드", excel_bytes,
                        gen.generate_filename(df=st.session_state.df),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="excel_dl"
                    )
                except Exception as e:
                    st.error(f"Excel 생성 오류: {e}")

        with col2:
            st.markdown(
                "**Word 보고서 (DOCX)**  \n"
                "표지·요약표·추이 차트·월별/연별 통계표·자동 해석 포함"
            )
            if st.button("📝 DOCX 보고서 생성", key="docx_btn"):
                try:
                    from docx_generator import DocxReportGenerator
                    docx_gen = DocxReportGenerator()
                    with st.spinner("DOCX 보고서 생성 중..."):
                        docx_bytes = docx_gen.generate_docx(
                            df=st.session_state.df,
                            monthly_df=st.session_state.monthly_df,
                            climate_df=st.session_state.climate_df,
                        )
                    st.download_button(
                        "⬇️ DOCX 다운로드", docx_bytes,
                        docx_gen.generate_filename(df=st.session_state.df),
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="docx_dl"
                    )
                except ImportError:
                    st.error("python-docx 패키지가 필요합니다. requirements.txt를 확인하세요.")
                except Exception as e:
                    st.error(f"DOCX 생성 오류: {e}")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB8 — 사용 방법
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab_help:
    st.subheader("ℹ️ 사용 방법")
    st.markdown("""
## 🌦️ 기상자료 분석 시스템 v4.0

### 1️⃣ 지원 데이터 형식

#### 🔵 ASOS (종관기상관측) — 주 분석 데이터

**기상자료개방포털** (https://data.kma.go.kr) 에서 무료로 다운로드:

1. 로그인 (회원가입 필요)
2. 기후 → 기상관측 → **지상** → **ASOS** 선택
3. 조회 조건 설정:
   - **자료 종류**: 일자료 선택
   - **기간 설정**: 원하는 시작일 ~ 종료일 입력
   - **지점 설정**: 분석 대상 기상관측소 선택 (복수 선택 가능)
   - **항목**: 전체 선택 (62개 컬럼 포함)
4. **조회** 버튼 클릭 후 **CSV 다운로드**
5. 다운로드한 파일을 왼쪽 업로드 창에 올리면 자동 분석

> 파일명 앞부분이 관측소명으로 인식됩니다. 예: `춘천_2000_2023.csv` → 춘천

#### 🟢 AWS (방재기상관측) — 추가 지점 보조 자료

**기상자료개방포털** 에서 같은 방법으로 다운로드 가능:

1. 기후 → 기상관측 → **지상** → **AWS** 선택
2. **자료 종류**: 일자료 선택 후 다운로드
3. ASOS 파일과 **함께** 업로드하면 복수 지점으로 자동 인식

> AWS는 ASOS 대비 측정 항목이 제한됩니다. 아래 지원 현황을 참고하세요.

| 항목 | ASOS | AWS |
|------|:----:|:---:|
| 기온 (평균·최고·최저) | ✅ | ✅ |
| 강수량 | ✅ | ✅ |
| 풍속 (평균·순간최대) | ✅ | ✅ |
| 습도 | ✅ | ❌ |
| 일조·일사량 | ✅ | ❌ |
| 기압 | ✅ | ❌ |
| 지중온도·적설 | ✅ | ❌ |

> 기온·강수·바람은 두 형식 모두 지원합니다. 일조·농업기상 ET₀ 분석은 ASOS 자료 필요.

---

### 2️⃣ 복수 지점 비교 분석

ASOS·AWS 관계없이 **여러 파일을 동시에 업로드**하면 자동으로 복수 관측소로 처리됩니다.

- 각 분석 탭 상단의 **관측소 선택** 필터에서 비교 대상 조절
- 기온·강수 추이 차트에서 관측소별 데이터가 함께 표시됨
- 겹치는 관측 기간이 다를 경우, 해당 기간 내 데이터만 비교됨

---

### 3️⃣ 분석 메뉴 안내

| 탭 | 주요 내용 | 활용 목적 |
|----|-----------|-----------|
| 📊 종합 분석 | 대화형 차트, 평년 분석, 월평균/연평균 | 데이터 전체 조망, 기초 현황 파악 |
| 🌡️ 기온 분석 | 폭염·열대야·한파 일수, 추세, 히트맵 | 기온 극값 통계 및 장기 기온 변화 분석 |
| 🌧️ 강수량 분석 | 강도별 일수, 누적강수, 여름집중도, 강우일수 | 홍수·가뭄 위험도 평가, 수자원 계획 |
| 💨 바람 분석 | 바람장미, Weibull 분포, 풍속등급, 기압 | 풍력발전 입지 검토, 환경영향평가 |
| ☀️ 태양광 분석 | 일사량, 일조시간, 전운량, 일조율 | 태양광 발전량 예측, 농업 광이용 분석 |
| 🌱 농업기상 분석 | 특이일수, ET₀, GDD, 지중온도, PAR | 작물 재배 적지 평가, 관개 계획 수립 |
| 📈 기후변화 분석 | 연평균기온·강수 추이, 이동평균, 전후반기 비교 | 기후변화 영향 평가, 장기 추세 분석 |
| ⬇️ 보고서 | Excel·DOCX 다운로드 (자동 해석 포함) | 보고서 작성, 데이터 공유 |

---

### 4️⃣ ASOS 데이터 구조 (62개 컬럼 자동 인식)

- **기온**: 평균·최고·최저·이슬점·초상온도
- **강수**: 일강수량·1시간최다·10분최다강수
- **바람**: 평균·최대·최대순간풍속, 풍향, 풍정합100m
- **일사·운량**: 일사량·일조·가조시간·전운량·하층운량
- **기압·습도**: 해면·현지기압, 상대습도, 증기압, 최소습도
- **지중온도**: 지면·5cm·10cm·20cm·30cm·0.5m·1m·1.5m·3m·5m
- **증발량**: 대형·소형 증발량, 안개시간

---

### 5️⃣ 권장 사항

- **평년기간**: 1991~2020 (WMO 기준 제30차 평년값)
- **추세 분석**: 10년 이상 데이터 권장
- **비교 분석**: 여러 관측소는 동일 기간 데이터 사용
- **기후변화 분석**: 30년 이상 자료에서 신뢰도 높은 결과 도출
- **AWS 활용**: ASOS 미설치 지역의 보조 자료로 활용하되, 종합적인 분석에는 ASOS 병행 권장

---
""")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Footer
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.markdown("---")
st.caption("기상자료 분석 시스템 v4.0 | ASOS 62컬럼 · AWS 방재기상 지원 | Streamlit 기반 Web App")
