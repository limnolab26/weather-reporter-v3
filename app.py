# app.py — 기상자료 보고서 생성기 Web App v3.0
# 실행: streamlit run app.py

import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime as dt

# Plotly
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Matplotlib
import matplotlib.pyplot as plt
import seaborn as sns

# 기존 모듈
from data_processor import WeatherDataProcessor
from excel_generator import ExcelReportGenerator

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 기본 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.set_page_config(
    page_title="기상자료 분석 시스템 v3.0",
    layout="wide"
)

st.title("🌦️ 기상자료 분석 시스템 v3.0")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 세션 상태 초기화
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if "df" not in st.session_state:
    st.session_state.df = None

if "monthly_df" not in st.session_state:
    st.session_state.monthly_df = None

if "climate_df" not in st.session_state:
    st.session_state.climate_df = None

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 유틸리티 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def extract_station_name(file):
    """
    파일명에서 관측소명 추출
    예: 춘천_2000_2023.csv → 춘천
    """
    name = file.name.split(".")[0]
    station = name.split("_")[0]
    return station


def add_time_columns(df):
    """연도/월/일/계절 컬럼 추가"""

    df["date"] = pd.to_datetime(df["date"])

    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month
    df["day"] = df["date"].dt.day

    def month_to_season(m):
        if m in [3,4,5]:
            return "봄"
        elif m in [6,7,8]:
            return "여름"
        elif m in [9,10,11]:
            return "가을"
        else:
            return "겨울"

    df["season"] = df["month"].apply(month_to_season)

    return df


@st.cache_data
def load_multiple_files(uploaded_files):
    """여러 CSV/XLSX 파일 병합"""

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
            processed_df["station_name"] = station_name

        processed_df = add_time_columns(processed_df)

        all_df.append(processed_df)

    if len(all_df) == 0:
        return None

    merged_df = pd.concat(all_df, ignore_index=True)

    return merged_df


@st.cache_data
def create_monthly_data(df):
    """월자료 생성"""

    all_agg = {
        "temp_avg": "mean",
        "temp_max": "mean",
        "temp_min": "mean",
        "precipitation": "sum",
        "humidity": "mean",
        "wind_speed": "mean",
        "wind_max": "mean",
        "sunshine": "sum",
        "solar_rad": "sum",
        "snowfall": "sum",
    }
    agg = {col: func for col, func in all_agg.items() if col in df.columns}

    monthly = (
        df
        .groupby(["station_name","year","month"])
        .agg(agg)
        .reset_index()
    )

    monthly["date"] = pd.to_datetime(
        monthly["year"].astype(str)
        + "-"
        + monthly["month"].astype(str)
        + "-01"
    )

    return monthly


@st.cache_data
def calculate_climate_normals(df, start_year, end_year):
    """평년 계산"""

    climate_df = df[
        (df["year"] >= start_year) &
        (df["year"] <= end_year)
    ]

    all_agg = {
        "temp_avg": "mean",
        "temp_max": "mean",
        "temp_min": "mean",
        "precipitation": "mean",
        "humidity": "mean",
        "wind_speed": "mean",
        "wind_max": "mean",
        "sunshine": "mean",
        "solar_rad": "mean",
        "snowfall": "mean",
    }
    agg = {col: func for col, func in all_agg.items() if col in climate_df.columns}

    climate = (
        climate_df
        .groupby(["station_name","month"])
        .agg(agg)
        .reset_index()
    )

    return climate


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Sidebar — 파일 업로드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.sidebar.header("📂 데이터 업로드")

uploaded_files = st.sidebar.file_uploader(
    "CSV 또는 XLSX 파일 업로드",
    type=["csv","xlsx"],
    accept_multiple_files=True
)

st.sidebar.link_button(
    "🌐 기상자료개방포털 ASOS 자료",
    "https://data.kma.go.kr/data/grnd/selectAsosRltmList.do?pgmNo=36"
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Sidebar — 평년기간 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.sidebar.header("📆 평년기간 설정")

climate_start = st.sidebar.number_input(
    "평년 시작연도",
    min_value=1950,
    max_value=2100,
    value=1991
)

climate_end = st.sidebar.number_input(
    "평년 종료연도",
    min_value=1950,
    max_value=2100,
    value=2020
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 데이터 로딩
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if uploaded_files:

    df = load_multiple_files(uploaded_files)

    if df is not None:

        st.session_state.df = df

        monthly_df = create_monthly_data(df)
        st.session_state.monthly_df = monthly_df

        climate_df = calculate_climate_normals(
            monthly_df,
            climate_start,
            climate_end
        )

        st.session_state.climate_df = climate_df

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 날짜 필터
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if st.session_state.df is not None:

    min_date = st.session_state.df["date"].min()
    max_date = st.session_state.df["date"].max()

    start_date, end_date = st.sidebar.date_input(
        "📅 분석 기간 선택",
        value=(min_date, max_date)
    )

    mask = (
        (st.session_state.df["date"] >= pd.to_datetime(start_date)) &
        (st.session_state.df["date"] <= pd.to_datetime(end_date))
    )

    filtered_df = st.session_state.df[mask]

else:
    filtered_df = None

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 탭 생성
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 대화형 차트",
    "🐍 Python 차트",
    "📆 평년 분석",
    "⬇️ 보고서 다운로드",
    "ℹ️ 사용 방법"
])
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB1 — 대화형 차트 (Plotly)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def prepare_chart_data(df, element, freq):
    """
    차트용 데이터 집계
    freq:
        D = 일별
        M = 월별
        Y = 연별
    """

    if freq == "D":

        grouped = (
            df
            .groupby(["station_name","date"])[element]
            .mean()
            .reset_index()
        )

    elif freq == "M":

        grouped = (
            df
            .groupby(["station_name","year","month"])[element]
            .mean()
            .reset_index()
        )

        grouped["date"] = pd.to_datetime(
            grouped["year"].astype(str)
            + "-"
            + grouped["month"].astype(str)
            + "-01"
        )

    elif freq == "Y":

        grouped = (
            df
            .groupby(["station_name","year"])[element]
            .mean()
            .reset_index()
        )

        grouped["date"] = pd.to_datetime(
            grouped["year"].astype(str) + "-01-01"
        )

    return grouped


ELEMENT_LABELS = {
    "temp_avg": "평균기온 (°C)",
    "temp_max": "최고기온 (°C)",
    "temp_min": "최저기온 (°C)",
    "precipitation": "강수량 (mm)",
    "humidity": "습도 (%)",
    "wind_speed": "평균풍속 (m/s)",
    "wind_max": "최대풍속 (m/s)",
    "sunshine": "일조시간 (hr)",
    "solar_rad": "일사량 (MJ/m²)",
    "snowfall": "적설 (cm)",
}

ELEMENT_OPTIONS = {
    "평균기온": "temp_avg",
    "최고기온": "temp_max",
    "최저기온": "temp_min",
    "강수량": "precipitation",
    "습도": "humidity",
    "평균풍속": "wind_speed",
    "최대풍속": "wind_max",
    "일조시간": "sunshine",
    "일사량": "solar_rad",
    "적설": "snowfall",
}


def create_plotly_chart(df, element, chart_type):

    y_label = ELEMENT_LABELS.get(element, element)

    if chart_type == "선":

        fig = px.line(
            df,
            x="date",
            y=element,
            color="station_name",
            markers=True,
            labels={"date": "날짜", element: y_label}
        )

    elif chart_type == "막대":

        fig = px.bar(
            df,
            x="date",
            y=element,
            color="station_name",
            labels={"date": "날짜", element: y_label}
        )

    fig.update_layout(height=420)

    return fig


def create_comparison_chart(df, selected_elements, element_chart_types, freq):
    """여러 기상요소를 하나의 차트에 비교"""

    fig = go.Figure()

    stations = df["station_name"].unique()

    for element in selected_elements:

        if element not in df.columns:
            continue

        el_df = prepare_chart_data(df, element, freq)
        label = ELEMENT_LABELS.get(element, element)
        chart_t = element_chart_types.get(element, "선")

        for stn in stations:

            stn_df = el_df[el_df["station_name"] == stn]
            trace_name = f"{stn} {label}"

            if chart_t == "선":
                fig.add_trace(go.Scatter(
                    x=stn_df["date"],
                    y=stn_df[element],
                    name=trace_name,
                    mode="lines+markers"
                ))
            else:
                fig.add_trace(go.Bar(
                    x=stn_df["date"],
                    y=stn_df[element],
                    name=trace_name,
                    opacity=0.75
                ))

    fig.update_layout(
        height=450,
        legend_title="관측소/요소",
        barmode="group"
    )

    return fig


def create_temp_precip_combo(df):

    fig = make_subplots(
        specs=[[{"secondary_y": True}]]
    )

    stations = df["station_name"].unique()

    for stn in stations:

        temp_df = df[
            df["station_name"] == stn
        ]

        # 평균기온 선

        fig.add_trace(
            go.Scatter(
                x=temp_df["date"],
                y=temp_df["temp_avg"],
                name=f"{stn} 평균기온",
                mode="lines+markers"
            ),
            secondary_y=False
        )

        # 강수량 막대

        fig.add_trace(
            go.Bar(
                x=temp_df["date"],
                y=temp_df["precipitation"],
                name=f"{stn} 강수량",
                opacity=0.5
            ),
            secondary_y=True
        )

    fig.update_yaxes(
        title_text="기온 (°C)",
        secondary_y=False
    )

    fig.update_yaxes(
        title_text="강수량 (mm)",
        secondary_y=True
    )

    fig.update_layout(
        legend_title="관측소"
    )

    return fig


with tab1:

    if filtered_df is None:

        st.info("먼저 데이터를 업로드하세요.")

    else:

        st.subheader("📊 대화형 차트")

        col1, col2, col3, col4 = st.columns(4)

        # ━━━━━━━━━━━━━━━━━━━
        # 관측소 선택
        # ━━━━━━━━━━━━━━━━━━━

        with col1:

            stations = sorted(
                filtered_df["station_name"]
                .unique()
            )

            selected_stations = st.multiselect(
                "관측소 선택",
                stations,
                default=stations
            )

        # ━━━━━━━━━━━━━━━━━━━
        # 기상요소 선택
        # ━━━━━━━━━━━━━━━━━━━

        with col2:

            selected_element_names = st.multiselect(
                "기상요소 선택",
                list(ELEMENT_OPTIONS.keys()),
                default=[list(ELEMENT_OPTIONS.keys())[0]],
                key="tab1_element"
            )

            selected_elements = [
                ELEMENT_OPTIONS[n] for n in selected_element_names
            ]

        # ━━━━━━━━━━━━━━━━━━━
        # 집계 단위
        # ━━━━━━━━━━━━━━━━━━━

        with col3:

            freq_options = {
                "일별":"D",
                "월별":"M",
                "연별":"Y"
            }

            freq_label = st.selectbox(
                "집계 단위",
                list(freq_options.keys())
            )

            freq = freq_options[freq_label]

        # ━━━━━━━━━━━━━━━━━━━
        # 차트 유형
        # ━━━━━━━━━━━━━━━━━━━

        with col4:

            chart_type = st.selectbox(
                "차트 유형",
                ["선", "막대", "복합(기온+강수)", "비교(다중요소)"]
            )

        # ━━━━━━━━━━━━━━━━━━━
        # 데이터 필터링
        # ━━━━━━━━━━━━━━━━━━━

        chart_df = filtered_df[
            filtered_df["station_name"]
            .isin(selected_stations)
        ]

        if len(chart_df) == 0:

            st.warning("선택된 관측소 데이터가 없습니다.")

        elif not selected_elements:

            st.info("기상요소를 하나 이상 선택하세요.")

        else:

            if chart_type == "복합(기온+강수)":

                # ━━━━━━━━━━━━━━━━━━━
                # 복합 차트 (기온+강수)
                # ━━━━━━━━━━━━━━━━━━━

                agg_cols = {k: v for k, v in {"temp_avg": "mean", "precipitation": "sum"}.items() if k in chart_df.columns}

                combo_df = (
                    chart_df
                    .groupby(["station_name", "year", "month"])
                    .agg(agg_cols)
                    .reset_index()
                )

                combo_df["date"] = pd.to_datetime(
                    combo_df["year"].astype(str)
                    + "-"
                    + combo_df["month"].astype(str)
                    + "-01"
                )

                fig = create_temp_precip_combo(combo_df)
                fig.update_layout(height=420)

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                    config={"displayModeBar": True}
                )

            elif chart_type == "비교(다중요소)":

                # ━━━━━━━━━━━━━━━━━━━
                # 다중요소 비교 차트
                # ━━━━━━━━━━━━━━━━━━━

                if not selected_elements:
                    st.info("기상요소를 하나 이상 선택하세요.")
                else:
                    with st.expander("요소별 차트 종류 설정", expanded=True):
                        element_chart_types = {}
                        n = len(selected_element_names)
                        type_cols = st.columns(min(n, 5))
                        for i, el_name in enumerate(selected_element_names):
                            with type_cols[i % 5]:
                                el_type = st.selectbox(
                                    el_name,
                                    ["선", "막대"],
                                    key=f"eltype_{el_name}"
                                )
                                element_chart_types[ELEMENT_OPTIONS[el_name]] = el_type

                    available = [e for e in selected_elements if e in chart_df.columns]
                    missing = [e for e in selected_elements if e not in chart_df.columns]
                    for m in missing:
                        st.warning(f"'{ELEMENT_LABELS.get(m, m)}' 데이터가 없습니다.")

                    if available:
                        fig = create_comparison_chart(
                            chart_df, available, element_chart_types, freq
                        )
                        st.plotly_chart(
                            fig,
                            use_container_width=True,
                            config={"displayModeBar": True}
                        )

            else:

                # ━━━━━━━━━━━━━━━━━━━
                # 요소별 개별 차트
                # ━━━━━━━━━━━━━━━━━━━

                for element in selected_elements:

                    if element not in chart_df.columns:
                        st.warning(f"'{ELEMENT_LABELS.get(element, element)}' 데이터가 없습니다.")
                        continue

                    el_df = prepare_chart_data(chart_df, element, freq)

                    fig = create_plotly_chart(el_df, element, chart_type)

                    st.plotly_chart(
                        fig,
                        use_container_width=True,
                        config={"displayModeBar": True}
                    )
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB2 — Python 고품질 차트 (Matplotlib / Seaborn)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

import matplotlib.font_manager as fm


def setup_korean_font():
    """한글 폰트 자동 설정 (파일 경로 우선 → 이름 검색 순)"""

    font_file_candidates = [
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",  # Streamlit Cloud (Linux)
        "NanumGothic.ttf",
        "malgun.ttf",
    ]

    for path in font_file_candidates:
        try:
            fm.fontManager.addfont(path)
            prop = fm.FontProperties(fname=path)
            plt.rcParams["font.family"] = prop.get_name()
            plt.rcParams["axes.unicode_minus"] = False
            return
        except Exception:
            continue

    font_name_candidates = ["NanumGothic", "Malgun Gothic", "AppleGothic"]
    available = [f.name for f in fm.fontManager.ttflist]
    for font in font_name_candidates:
        if font in available:
            plt.rcParams["font.family"] = font
            break

    plt.rcParams["axes.unicode_minus"] = False


def get_chart_bytes(fig):
    """Figure → PNG bytes"""

    buf = io.BytesIO()

    fig.savefig(
        buf,
        format="png",
        dpi=150,
        bbox_inches="tight"
    )

    buf.seek(0)

    return buf.read()


def create_temp_timeseries(df):

    setup_korean_font()

    fig, ax = plt.subplots(
        figsize=(10,4),
        dpi=100
    )

    stations = df["station_name"].unique()

    for stn in stations:

        stn_df = df[
            df["station_name"] == stn
        ]

        ax.plot(
            stn_df["date"],
            stn_df["temp_avg"],
            label=f"{stn} 평균기온"
        )

        ax.fill_between(
            stn_df["date"],
            stn_df["temp_min"],
            stn_df["temp_max"],
            alpha=0.15
        )

    ax.set_title("기온 시계열")

    ax.set_xlabel("날짜")
    ax.set_ylabel("기온 (°C)")

    ax.legend()
    ax.grid(True, alpha=0.3)

    plt.tight_layout()

    return fig


def create_monthly_precip_bar(df):

    setup_korean_font()

    monthly = (
        df
        .groupby(["year","month"])
        ["precipitation"]
        .sum()
        .reset_index()
    )

    fig, ax = plt.subplots(
        figsize=(10,4),
        dpi=100
    )

    ax.bar(
        monthly["month"],
        monthly["precipitation"]
    )

    ax.set_title("월별 강수량")

    ax.set_xlabel("월")
    ax.set_ylabel("강수량 (mm)")

    plt.tight_layout()

    return fig


def create_temp_boxplot(df):

    setup_korean_font()

    fig, ax = plt.subplots(
        figsize=(10,4),
        dpi=100
    )

    sns.boxplot(
        data=df,
        x="month",
        y="temp_avg",
        ax=ax
    )

    ax.set_title("월별 기온 분포 (Box Plot)")

    plt.tight_layout()

    return fig


def create_heatmap(df):

    setup_korean_font()

    pivot = df.pivot_table(
        values="temp_avg",
        index="year",
        columns="month",
        aggfunc="mean"
    )

    fig, ax = plt.subplots(
        figsize=(10,4),
        dpi=100
    )

    sns.heatmap(
        pivot,
        annot=False,
        cmap="RdYlBu_r",
        ax=ax
    )

    ax.set_title(
        "연도 × 월 평균기온 히트맵"
    )

    plt.tight_layout()

    return fig


def create_scatter(df):

    setup_korean_font()

    fig, ax = plt.subplots(
        figsize=(10,4),
        dpi=100
    )

    seasons = df["season"].unique()

    for s in seasons:

        s_df = df[
            df["season"] == s
        ]

        ax.scatter(
            s_df["temp_avg"],
            s_df["precipitation"],
            label=s,
            alpha=0.6
        )

    ax.set_xlabel("기온 (°C)")
    ax.set_ylabel("강수량 (mm)")

    ax.set_title(
        "기온 vs 강수량 산점도"
    )

    ax.legend()

    plt.tight_layout()

    return fig


with tab2:

    if filtered_df is None:

        st.info("먼저 데이터를 업로드하세요.")

    else:

        st.subheader(
            "🐍 Python 고품질 차트"
        )

        chart_type = st.selectbox(
            "차트 종류 선택",
            [
                "기온 시계열",
                "월별 강수량",
                "기온 Box Plot",
                "히트맵",
                "산점도"
            ]
        )

        if st.button("📈 차트 생성"):

            try:

                if chart_type == "기온 시계열":

                    fig = create_temp_timeseries(
                        filtered_df
                    )

                elif chart_type == "월별 강수량":

                    fig = create_monthly_precip_bar(
                        filtered_df
                    )

                elif chart_type == "기온 Box Plot":

                    fig = create_temp_boxplot(
                        filtered_df
                    )

                elif chart_type == "히트맵":

                    fig = create_heatmap(
                        filtered_df
                    )

                elif chart_type == "산점도":

                    fig = create_scatter(
                        filtered_df
                    )

                st.pyplot(fig)

                png_bytes = get_chart_bytes(fig)

                st.download_button(
                    label="🖼️ PNG 다운로드",
                    data=png_bytes,
                    file_name=f"{chart_type}.png",
                    mime="image/png"
                )

            except Exception as e:

                st.error(
                    f"차트 생성 중 오류 발생: {e}"
                )
                # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB3 — 평년 분석
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def calculate_anomaly(df, climate_df, element):
    """편차(Anomaly) 계산"""

    merged = pd.merge(
        df,
        climate_df,
        on=["station_name","month"],
        suffixes=("", "_normal")
    )

    merged["anomaly"] = (
        merged[element] - merged[f"{element}_normal"]
    )

    return merged


def create_climate_table(climate_df, element):
    """월별 평년표 생성"""

    pivot = climate_df.pivot_table(
        values=element,
        index="month",
        columns="station_name"
    )

    return pivot


def create_climate_chart(climate_df, element):

    fig = px.line(
        climate_df,
        x="month",
        y=element,
        color="station_name",
        markers=True,
        labels={
            "month":"월",
            element:"값"
        }
    )

    fig.update_layout(
        title="월별 평년값",
        height=380
    )

    return fig


def create_anomaly_chart(df, element):

    fig = px.line(
        df,
        x="date",
        y="anomaly",
        color="station_name",
        labels={
            "date":"날짜",
            "anomaly":"편차"
        }
    )

    fig.update_layout(
        title="평년 대비 편차",
        height=380
    )

    return fig


with tab3:

    if filtered_df is None:

        st.info("먼저 데이터를 업로드하세요.")

    else:

        st.subheader("📆 평년 분석")

        # ━━━━━━━━━━━━━━━━━━━
        # 요소 선택
        # ━━━━━━━━━━━━━━━━━━━

        element_name = st.selectbox(
            "기상요소 선택",
            list(ELEMENT_OPTIONS.keys()),
            key="tab4_element"
        )

        element = ELEMENT_OPTIONS[element_name]

        # ━━━━━━━━━━━━━━━━━━━
        # 평년 데이터
        # ━━━━━━━━━━━━━━━━━━━

        climate_df = st.session_state.climate_df

        if climate_df is None:

            st.warning("평년 데이터가 없습니다.")
            st.stop()

        # ━━━━━━━━━━━━━━━━━━━
        # 월별 평년표
        # ━━━━━━━━━━━━━━━━━━━

        st.markdown("### 📋 월별 평년값")

        climate_table = create_climate_table(
            climate_df,
            element
        )

        st.dataframe(
            climate_table.style.format("{:.2f}"),
            use_container_width=True,
            height=460
        )

        # 다운로드

        csv_data = climate_table.to_csv(
            encoding="utf-8-sig"
        ).encode("utf-8-sig")

        st.download_button(
            label="⬇️ 평년 CSV 다운로드",
            data=csv_data,
            file_name="climate_normals.csv",
            mime="text/csv"
        )

        # ━━━━━━━━━━━━━━━━━━━
        # 평년 그래프
        # ━━━━━━━━━━━━━━━━━━━

        st.markdown("### 📈 평년 그래프")

        climate_chart = create_climate_chart(
            climate_df,
            element
        )

        st.plotly_chart(
            climate_chart,
            use_container_width=True
        )

        # ━━━━━━━━━━━━━━━━━━━
        # 편차 계산
        # ━━━━━━━━━━━━━━━━━━━

        st.markdown("### 📉 평년 대비 편차")

        anomaly_df = calculate_anomaly(
            st.session_state.monthly_df,
            climate_df,
            element
        )

        anomaly_chart = create_anomaly_chart(
            anomaly_df,
            element
        )

        st.plotly_chart(
            anomaly_chart,
            use_container_width=True
        )
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB4 — 보고서 다운로드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab4:

    if filtered_df is None:

        st.info("먼저 데이터를 업로드하세요.")

    else:

        st.subheader("⬇️ 보고서 다운로드")

        st.markdown(
            """
            선택한 데이터 기반으로  
            **Excel 또는 PDF 보고서**를 생성할 수 있습니다.
            """
        )

        col1, col2 = st.columns(2)

        # ━━━━━━━━━━━━━━━━━━━
        # Excel 다운로드
        # ━━━━━━━━━━━━━━━━━━━

        with col1:

            if st.button("📊 Excel 보고서 생성"):

                try:

                    from excel_generator import ExcelReportGenerator

                    excel_gen = ExcelReportGenerator()

                    summary_df = excel_gen.generate_summary_table(
                        st.session_state.df
                    )

                    excel_bytes = excel_gen.generate_excel(
                        df=st.session_state.df,
                        pivot_df=None,
                        summary_df=summary_df
                    )

                    st.download_button(
                        label="⬇️ Excel 다운로드",
                        data=excel_bytes,
                        file_name=excel_gen.generate_filename(),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:

                    st.error(
                        f"Excel 생성 중 오류 발생: {e}"
                    )

        # ━━━━━━━━━━━━━━━━━━━
        # PDF 다운로드
        # ━━━━━━━━━━━━━━━━━━━

        with col2:

            if st.button("📄 PDF 보고서 생성"):

                try:

                    from pdf_generator import PDFReportGenerator

                    pdf_gen = PDFReportGenerator()

                    pdf_bytes = pdf_gen.generate_pdf(
                        df=st.session_state.df,
                        pivot_df=None,
                        summary_df=None
                    )

                    st.download_button(
                        label="⬇️ PDF 다운로드",
                        data=pdf_bytes,
                        file_name=pdf_gen.generate_filename(),
                        mime="application/pdf"
                    )

                except Exception as e:

                    st.error(
                        f"PDF 생성 중 오류 발생: {e}"
                    )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB5 — 사용 방법
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab5:

    st.subheader("ℹ️ 사용 방법")

    st.markdown(
        """
        ## 🌦️ 기상자료 분석 시스템 사용 안내

        ### 1️⃣ 데이터 업로드

        - 기상청 ASOS 일자료 CSV 또는 XLSX 업로드
        - 여러 관측소 파일 동시 업로드 가능
        - 파일명 예:
            - 춘천_2000_2023.csv
            - 서울_1991_2020.xlsx

        ---

        ### 2️⃣ 기간 설정

        좌측 사이드바에서:

        - 분석 기간 선택
        - 평년기간 설정 (예: 1991–2020)

        ---

        ### 3️⃣ 주요 기능

        #### 📊 대화형 차트

        - 관측소 복수 선택
        - 기상요소 복수 선택 (개별 또는 비교 차트)
        - 일/월/연 집계
        - 선·막대·복합(기온+강수)·비교(다중요소) 차트 지원

        #### 🐍 Python 차트

        보고서용 고품질 그래프:

        - 기온 시계열
        - 월별 강수량
        - Box Plot
        - Heatmap
        - Scatter

        PNG 다운로드 가능

        #### 📆 평년 분석

        핵심 기능:

        - 월별 평년 계산
        - 평년 그래프
        - 평년 대비 편차 분석

        #### ⬇️ 보고서 다운로드

        생성 가능:

        - Excel 보고서
        - PDF 보고서

        ---

        ### 📌 권장 데이터 형식

        필수 컬럼:

        - date
        - temp_avg
        - temp_max
        - temp_min
        - precipitation

        권장 컬럼:

        - humidity
        - wind_speed
        - sunshine

        ---
        """
    )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Footer
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.markdown("---")

st.caption(
    "기상자료 분석 시스템 v3.0 | Streamlit 기반 Web App"
)
