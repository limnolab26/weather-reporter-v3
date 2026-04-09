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
        "sunshine": "mean"
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
        "sunshine": "mean"
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

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 대화형 차트",
    "📋 동적 피벗표",
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


def create_plotly_chart(df, element, chart_type):

    element_labels = {
        "temp_avg":"평균기온 (°C)",
        "temp_max":"최고기온 (°C)",
        "temp_min":"최저기온 (°C)",
        "precipitation":"강수량 (mm)",
        "humidity":"습도 (%)",
        "wind_speed":"풍속 (m/s)",
        "sunshine":"일조시간 (hr)"
    }

    y_label = element_labels.get(element, element)

    if chart_type == "선":

        fig = px.line(
            df,
            x="date",
            y=element,
            color="station_name",
            markers=True,
            labels={
                "date":"날짜",
                element:y_label
            }
        )

    elif chart_type == "막대":

        fig = px.bar(
            df,
            x="date",
            y=element,
            color="station_name",
            labels={
                "date":"날짜",
                element:y_label
            }
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

            element_options = {
                "평균기온":"temp_avg",
                "최고기온":"temp_max",
                "최저기온":"temp_min",
                "강수량":"precipitation",
                "습도":"humidity",
                "풍속":"wind_speed",
                "일조시간":"sunshine"
            }

            selected_element_name = st.selectbox(
                "기상요소 선택",
                list(element_options.keys()),
                key="tab1_element"
            )

            selected_element = element_options[
                selected_element_name
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
                ["선","막대","복합(기온+강수)"]
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

        else:

            # ━━━━━━━━━━━━━━━━━━━
            # 데이터 집계
            # ━━━━━━━━━━━━━━━━━━━

            if chart_type != "복합(기온+강수)":

                chart_df = prepare_chart_data(
                    chart_df,
                    selected_element,
                    freq
                )

            else:

                chart_df = (
                    chart_df
                    .groupby(
                        ["station_name","year","month"]
                    )
                    .agg({
                        "temp_avg":"mean",
                        "precipitation":"sum"
                    })
                    .reset_index()
                )

                chart_df["date"] = pd.to_datetime(
                    chart_df["year"].astype(str)
                    + "-"
                    + chart_df["month"].astype(str)
                    + "-01"
                )

            # ━━━━━━━━━━━━━━━━━━━
            # 차트 생성
            # ━━━━━━━━━━━━━━━━━━━

            if chart_type == "복합(기온+강수)":

                fig = create_temp_precip_combo(
                    chart_df
                )

            else:

                fig = create_plotly_chart(
                    chart_df,
                    selected_element,
                    chart_type
                )

            # ━━━━━━━━━━━━━━━━━━━
            # 차트 표시
            # ━━━━━━━━━━━━━━━━━━━

            st.plotly_chart(
                fig,
                use_container_width=True,
                config={
                    "displayModeBar": True
                }
            )
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB2 — 동적 Pivot Table
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def create_pivot_table(
    df,
    row_field,
    col_field,
    value_field,
    agg_func
):
    """Pivot Table 생성 함수"""

    pivot_df = pd.pivot_table(
        df,
        index=row_field,
        columns=col_field,
        values=value_field,
        aggfunc=agg_func
    )

    return pivot_df


def convert_df_to_csv(df):
    """DataFrame → CSV bytes"""

    return df.to_csv(
        encoding="utf-8-sig"
    ).encode("utf-8-sig")


with tab2:

    if filtered_df is None:

        st.info("먼저 데이터를 업로드하세요.")

    else:

        st.subheader("📋 동적 Pivot Table")

        # ━━━━━━━━━━━━━━━━━━━
        # 필드 선택 UI
        # ━━━━━━━━━━━━━━━━━━━

        col1, col2, col3, col4 = st.columns(4)

        with col1:

            row_options = [
                "station_name",
                "year",
                "month",
                "season"
            ]

            row_field = st.selectbox(
                "행(Row)",
                row_options
            )

        with col2:

            col_options = [
                "station_name",
                "year",
                "month",
                "season"
            ]

            col_field = st.selectbox(
                "열(Column)",
                col_options,
                index=1
            )

        with col3:

            value_options = {
                "평균기온":"temp_avg",
                "최고기온":"temp_max",
                "최저기온":"temp_min",
                "강수량":"precipitation",
                "습도":"humidity",
                "풍속":"wind_speed",
                "일조시간":"sunshine"
            }

            value_name = st.selectbox(
                "값(Value)",
                list(value_options.keys())
            )

            value_field = value_options[
                value_name
            ]

        with col4:

            agg_options = {
                "평균":"mean",
                "합계":"sum",
                "최대":"max",
                "최소":"min",
                "건수":"count"
            }

            agg_name = st.selectbox(
                "집계 함수",
                list(agg_options.keys())
            )

            agg_func = agg_options[
                agg_name
            ]

        # ━━━━━━━━━━━━━━━━━━━
        # Pivot 생성 버튼
        # ━━━━━━━━━━━━━━━━━━━

        if st.button("📊 Pivot 생성"):

            try:

                pivot_df = create_pivot_table(
                    filtered_df,
                    row_field,
                    col_field,
                    value_field,
                    agg_func
                )

                # 결측값 처리
                pivot_df = pivot_df.fillna("—")

                st.success("Pivot Table 생성 완료")

                # ━━━━━━━━━━━━━━━━━━━
                # Pivot 표시
                # ━━━━━━━━━━━━━━━━━━━

                st.dataframe(
                    pivot_df,
                    use_container_width=True
                )

                # ━━━━━━━━━━━━━━━━━━━
                # CSV 다운로드
                # ━━━━━━━━━━━━━━━━━━━

                csv_data = convert_df_to_csv(
                    pivot_df
                )

                st.download_button(
                    label="⬇️ Pivot CSV 다운로드",
                    data=csv_data,
                    file_name="pivot_table.csv",
                    mime="text/csv"
                )

            except Exception as e:

                st.error(
                    f"Pivot 생성 중 오류 발생: {e}"
                )

        # ━━━━━━━━━━━━━━━━━━━
        # 빠른 예시 안내
        # ━━━━━━━━━━━━━━━━━━━

        with st.expander("💡 Pivot 사용 예시"):

            st.markdown(
                """
                **연 × 월 평균기온**

                Row: year  
                Column: month  
                Value: 평균기온  
                Agg: 평균  

                **관측소별 강수량 합계**

                Row: station_name  
                Column: year  
                Value: 강수량  
                Agg: 합계  
                """
            )
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB3 — Python 고품질 차트 (Matplotlib / Seaborn)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

import matplotlib.font_manager as fm


def setup_korean_font():
    """한글 폰트 자동 설정"""

    font_list = [
        "NanumGothic",
        "Malgun Gothic",
        "AppleGothic"
    ]

    available_fonts = [
        f.name for f in fm.fontManager.ttflist
    ]

    for font in font_list:
        if font in available_fonts:
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
        figsize=(12,6),
        dpi=150
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
        figsize=(12,6),
        dpi=150
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
        figsize=(12,6),
        dpi=150
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
        figsize=(12,6),
        dpi=150
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
        figsize=(10,6),
        dpi=150
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


with tab3:

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
# TAB4 — 평년 분석
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
        title="월별 평년값"
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
        title="평년 대비 편차"
    )

    return fig


with tab4:

    if filtered_df is None:

        st.info("먼저 데이터를 업로드하세요.")

    else:

        st.subheader("📆 평년 분석")

        # ━━━━━━━━━━━━━━━━━━━
        # 요소 선택
        # ━━━━━━━━━━━━━━━━━━━

        element_options = {
            "평균기온":"temp_avg",
            "최고기온":"temp_max",
            "최저기온":"temp_min",
            "강수량":"precipitation",
            "습도":"humidity",
            "풍속":"wind_speed",
            "일조시간":"sunshine"
        }

        element_name = st.selectbox(
            "기상요소 선택",
            list(element_options.keys()),
            key="tab4_element"
        )

        element = element_options[element_name]

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
            climate_table,
            use_container_width=True
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
# TAB5 — 보고서 다운로드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab5:

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
# TAB6 — 사용 방법
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab6:

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

        - 관측소 선택
        - 기상요소 선택
        - 일/월/연 집계
        - 복합 그래프 지원

        #### 📋 Pivot Table

        자유로운 통계 생성:

        예:

        - Row: year
        - Column: month
        - Value: precipitation
        - Agg: sum

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
