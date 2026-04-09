# app.py — 기상자료 분석 시스템 v4.0
# 실행: streamlit run app.py

import io
import datetime as dt

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 페이지 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.set_page_config(
    page_title="기상자료 분석 시스템 v4.0",
    layout="wide"
)

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
    date_range = st.sidebar.date_input("📅 분석 기간 선택", value=(min_date, max_date))
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

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 탭 생성 (8탭)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

(tab_general, tab_temp, tab_precip, tab_wind,
 tab_solar, tab_agri, tab_download, tab_help) = st.tabs([
    "📊 종합 분석",
    "🌡️ 기온 분석",
    "🌧️ 강수량 분석",
    "💨 바람 분석",
    "☀️ 태양광 분석",
    "🌱 농업기상 분석",
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
        st.info("먼저 데이터를 업로드하세요.")
    else:
        gen_t1, gen_t2 = st.tabs(["📊 대화형 차트", "📆 평년 분석"])

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
                st.plotly_chart(create_temp_precip_combo(combo_df), use_container_width=True)

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
                        st.plotly_chart(
                            create_comparison_chart(chart_df, available_sel, el_types, freq),
                            use_container_width=True
                        )
            else:
                for element in sel_elements:
                    if element not in chart_df.columns:
                        st.warning(f"'{ELEMENT_LABELS.get(element, element)}' 데이터가 없습니다.")
                        continue
                    el_df = prepare_chart_data(chart_df, element, freq)
                    st.plotly_chart(create_plotly_chart(el_df, element, chart_type), use_container_width=True)

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

                # 편차
                st.markdown("### 📉 평년 대비 편차")
                if st.session_state.monthly_df is not None and cl_element in st.session_state.monthly_df.columns:
                    anomaly_df = calculate_anomaly(st.session_state.monthly_df, climate_df, cl_element)
                    fig_an = px.line(anomaly_df, x="date", y="anomaly", color="station_name",
                                     labels={"date": "날짜", "anomaly": "편차"})
                    fig_an.update_layout(height=380)
                    st.plotly_chart(fig_an, use_container_width=True)

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

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB7 — 보고서 다운로드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab_download:
    if filtered_df is None:
        st.info("먼저 데이터를 업로드하세요.")
    else:
        st.subheader("⬇️ 보고서 다운로드")
        st.markdown("선택한 데이터 기반으로 **Excel 또는 PDF 보고서**를 생성합니다.")

        col1, col2 = st.columns(2)

        with col1:
            if st.button("📊 Excel 보고서 생성", key="excel_btn"):
                try:
                    from excel_generator import ExcelReportGenerator
                    gen = ExcelReportGenerator()
                    excel_bytes = gen.generate_excel(
                        df=st.session_state.df,
                        monthly_df=st.session_state.monthly_df,
                        climate_df=st.session_state.climate_df,
                    )
                    st.download_button(
                        "⬇️ Excel 다운로드", excel_bytes,
                        gen.generate_filename(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="excel_dl"
                    )
                except Exception as e:
                    st.error(f"Excel 생성 오류: {e}")

        with col2:
            if st.button("📄 PDF 보고서 생성", key="pdf_btn"):
                try:
                    from pdf_generator import PDFReportGenerator
                    pdf_gen = PDFReportGenerator()
                    pdf_bytes = pdf_gen.generate_pdf(
                        df=st.session_state.df,
                        pivot_df=None,
                        summary_df=None
                    )
                    st.download_button(
                        "⬇️ PDF 다운로드", pdf_bytes,
                        pdf_gen.generate_filename(), "application/pdf",
                        key="pdf_dl"
                    )
                except Exception as e:
                    st.error(f"PDF 생성 오류: {e}")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB8 — 사용 방법
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

with tab_help:
    st.subheader("ℹ️ 사용 방법")
    st.markdown("""
## 🌦️ 기상자료 분석 시스템 v4.0

### 1️⃣ 데이터 업로드
- 기상청 ASOS 일자료 CSV (cp949 인코딩) 또는 XLSX 업로드
- 여러 관측소 파일 동시 업로드 가능 (자동 병합)
- 파일명 첫 단어를 관측소명으로 인식 (예: `춘천_2000_2023.csv` → 춘천)

### 2️⃣ 분석 메뉴

| 탭 | 주요 내용 |
|----|-----------|
| 📊 종합 분석 | 대화형 차트 (선/막대/복합/비교) + 평년 분석 |
| 🌡️ 기온 분석 | 폭염·열대야·한파·추세·히트맵 |
| 🌧️ 강수량 분석 | 강도별 일수·연강수·연속무강수·히트맵 |
| 💨 바람 분석 | 바람장미·Weibull·풍속등급·기압 |
| ☀️ 태양광 분석 | 일사량·일조·전운량·일조율 |
| 🌱 농업기상 분석 | 특이일수·ET₀·GDD·토양수지·지중온도·PAR |
| ⬇️ 보고서 | Excel·PDF 다운로드 |

### 3️⃣ ASOS 데이터 구조
기상청 ASOS 일자료 62개 컬럼 자동 인식:
- 기온 (평균·최고·최저·이슬점·초상온도)
- 강수 (일강수량·1시간최다·10분최다)
- 바람 (평균·최대·최대순간풍속, 풍향, 풍정합)
- 일사·운량 (일사량·일조·가조시간·전운량)
- 기압·습도 (해면·현지기압, 상대습도, 증기압)
- 지중온도 (지면·5cm·10cm·20cm·30cm·0.5·1·1.5·3·5m)
- 증발량 (대형·소형 증발량)

### 4️⃣ 권장 사항
- 평년기간: 1991~2020 (WMO 기준 제30차 평년값)
- 10년 이상 데이터 권장 (추세 분석)
- 여러 관측소 비교 시 동일 기간 데이터 사용

---
""")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Footer
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

st.markdown("---")
st.caption("기상자료 분석 시스템 v4.0 | ASOS 62컬럼 지원 | Streamlit 기반 Web App")
