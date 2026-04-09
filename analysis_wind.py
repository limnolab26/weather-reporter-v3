# analysis_wind.py — 바람 분석 모듈
# render(df) 함수: 필터링된 일별 DataFrame을 받아 5개 서브탭 렌더링

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

try:
    from scipy import stats as scipy_stats
    _SCIPY_AVAILABLE = True
except ImportError:
    _SCIPY_AVAILABLE = False

CHART_HEIGHT = 400
COMPASS_16 = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
              "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
SEASON_ORDER = ["봄", "여름", "가을", "겨울"]


def _has_col(df: pd.DataFrame, col: str) -> bool:
    return col in df.columns and df[col].notna().any()


def _get_stations(df: pd.DataFrame) -> list:
    if "station_name" not in df.columns:
        return []
    return sorted(df["station_name"].dropna().unique().tolist())


def _filter_stations(df: pd.DataFrame, selected: list) -> pd.DataFrame:
    if not selected or "station_name" not in df.columns:
        return df
    return df[df["station_name"].isin(selected)]


def _deg_to_compass(deg: float) -> str:
    """풍향 각도(0-360)를 16방위 문자열로 변환."""
    idx = int((deg + 11.25) / 22.5) % 16
    return COMPASS_16[idx]


# ──────────────────────────────────────────────────────────────
# Tab 1: 풍속 추이
# ──────────────────────────────────────────────────────────────

def _tab_wind_trend(df: pd.DataFrame) -> None:
    st.subheader("풍속 추이")

    if not _has_col(df, "wind_speed"):
        st.warning("wind_speed 컬럼이 없습니다.")
        return

    if "station_name" not in df.columns:
        st.warning("station_name 컬럼이 없습니다.")
        return

    freq_option = st.radio(
        "집계 단위",
        options=["월별", "연별"],
        horizontal=True,
        key="wind_trend_freq",
    )

    has_max = _has_col(df, "wind_max")

    if freq_option == "월별":
        if "year" not in df.columns or "month" not in df.columns:
            st.warning("year, month 컬럼이 필요합니다.")
            return

        agg_cols = {"wind_speed": ("wind_speed", "mean")}
        if has_max:
            agg_cols["wind_max"] = ("wind_max", "mean")

        grouped = df.groupby(["year", "month", "station_name"], as_index=False).agg(**agg_cols)
        grouped = grouped.assign(
            period=grouped["year"].astype(str) + "-" + grouped["month"].astype(str).str.zfill(2)
        ).sort_values(["station_name", "year", "month"])
        x_col = "period"
        x_label = "연월"
    else:
        if "year" not in df.columns:
            st.warning("year 컬럼이 필요합니다.")
            return

        agg_cols = {"wind_speed": ("wind_speed", "mean")}
        if has_max:
            agg_cols["wind_max"] = ("wind_max", "mean")

        grouped = df.groupby(["year", "station_name"], as_index=False).agg(**agg_cols)
        grouped = grouped.sort_values(["station_name", "year"])
        x_col = "year"
        x_label = "연도"

    fig = go.Figure()
    color_seq = px.colors.qualitative.Plotly
    station_list = grouped["station_name"].unique().tolist()

    for idx, station in enumerate(station_list):
        sdf = grouped[grouped["station_name"] == station]
        color = color_seq[idx % len(color_seq)]

        fig.add_trace(go.Scatter(
            x=sdf[x_col],
            y=sdf["wind_speed"],
            name=f"{station} 평균풍속",
            mode="lines+markers",
            line={"color": color, "width": 2},
            marker={"size": 4},
        ))

        if has_max:
            fig.add_trace(go.Scatter(
                x=sdf[x_col],
                y=sdf["wind_max"],
                name=f"{station} 최대풍속",
                mode="lines+markers",
                line={"color": color, "width": 2, "dash": "dash"},
                marker={"size": 4, "symbol": "diamond"},
            ))

    fig.update_layout(
        title=f"{freq_option} 평균/최대 풍속 추이",
        xaxis={"title": x_label},
        yaxis={"title": "풍속 (m/s)"},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.25},
    )
    st.plotly_chart(fig, use_container_width=True)


# ──────────────────────────────────────────────────────────────
# Tab 2: 바람장미도
# ──────────────────────────────────────────────────────────────

def _tab_wind_rose(df: pd.DataFrame) -> None:
    st.subheader("바람장미도")

    if not _has_col(df, "wind_dir"):
        st.warning("wind_dir 컬럼이 없습니다.")
        return

    stations = _get_stations(df)
    if not stations:
        st.info("station_name 컬럼이 없습니다.")
        return

    col1, col2 = st.columns(2)
    with col1:
        selected_station = st.selectbox(
            "관측소 선택",
            options=stations,
            key="wind_rose_station",
        )
    with col2:
        season_options = ["전체"] + SEASON_ORDER
        selected_season = st.selectbox(
            "계절 필터",
            options=season_options,
            key="wind_rose_season",
        )

    filtered = df[df["station_name"] == selected_station].copy() if "station_name" in df.columns else df.copy()

    if selected_season != "전체" and "season" in filtered.columns:
        filtered = filtered[filtered["season"] == selected_season]

    wind_data = filtered["wind_dir"].dropna()
    if wind_data.empty:
        st.info("선택한 조건에 해당하는 풍향 데이터가 없습니다.")
        return

    compass_counts = wind_data.apply(_deg_to_compass).value_counts()
    total = compass_counts.sum()
    freq_pct = (compass_counts / total * 100).reindex(COMPASS_16, fill_value=0.0)

    fig = go.Figure(go.Barpolar(
        r=freq_pct.values.tolist(),
        theta=COMPASS_16,
        marker_color=freq_pct.values.tolist(),
        marker_colorscale="Blues",
        opacity=0.85,
    ))

    season_label = f" ({selected_season})" if selected_season != "전체" else ""
    fig.update_layout(
        title=f"{selected_station} 바람장미도{season_label}",
        polar={
            "angularaxis": {"direction": "clockwise", "rotation": 90},
            "radialaxis": {"title": "빈도 (%)"},
        },
        height=CHART_HEIGHT,
    )
    st.plotly_chart(fig, use_container_width=True)


# ──────────────────────────────────────────────────────────────
# Tab 3: Weibull 분포 분석
# ──────────────────────────────────────────────────────────────

def _tab_weibull(df: pd.DataFrame) -> None:
    st.subheader("Weibull 분포 분석")

    if not _has_col(df, "wind_speed"):
        st.warning("wind_speed 컬럼이 없습니다.")
        return

    if not _SCIPY_AVAILABLE:
        st.error("scipy가 설치되지 않아 Weibull 분포 분석을 사용할 수 없습니다.")
        return

    stations = _get_stations(df)
    if not stations:
        st.info("station_name 컬럼이 없습니다.")
        return

    selected_station = st.selectbox(
        "관측소 선택",
        options=stations,
        key="wind_weibull_station",
    )

    filtered = df[df["station_name"] == selected_station].copy() if "station_name" in df.columns else df.copy()
    wind_vals = filtered["wind_speed"].dropna().values
    wind_vals = wind_vals[wind_vals > 0]

    if len(wind_vals) < 10:
        st.warning("Weibull 분포 추정에 필요한 데이터가 부족합니다 (최소 10개).")
        return

    # Weibull 피팅
    try:
        shape, loc, scale = scipy_stats.weibull_min.fit(wind_vals, floc=0)
        k = float(shape)   # 형상 모수
        A = float(scale)   # 척도 모수 (scale = A)
    except Exception as exc:
        st.error(f"Weibull 분포 피팅 실패: {exc}")
        return

    # 용량인수 추정 (Vci=3, Vco=25, Vr=12)
    Vci, Vco = 3.0, 25.0
    try:
        cf = float(np.exp(-((Vci / A) ** k)) - np.exp(-((Vco / A) ** k)))
        cf = max(0.0, min(1.0, cf))
    except Exception:
        cf = float("nan")

    mean_ws = float(np.mean(wind_vals))

    # 지표 표시
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("형상 모수 k", f"{k:.3f}")
    m2.metric("척도 모수 A (m/s)", f"{A:.3f}")
    m3.metric("평균 풍속 (m/s)", f"{mean_ws:.2f}")
    m4.metric("이론 용량인수", f"{cf:.3f}" if not np.isnan(cf) else "N/A")

    # 히스토그램 + PDF 곡선
    x_range = np.linspace(0, wind_vals.max() * 1.1, 300)
    pdf_vals = scipy_stats.weibull_min.pdf(x_range, k, loc=0, scale=A)

    bin_count = min(50, max(20, len(wind_vals) // 20))
    counts, bin_edges = np.histogram(wind_vals, bins=bin_count, density=True)
    bin_centers = (bin_edges[:-1] + bin_edges[1:]) / 2

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=bin_centers.tolist(),
        y=counts.tolist(),
        name="풍속 빈도",
        marker_color="rgba(100, 149, 237, 0.6)",
        width=float(bin_edges[1] - bin_edges[0]) * 0.9,
    ))
    fig.add_trace(go.Scatter(
        x=x_range.tolist(),
        y=pdf_vals.tolist(),
        name=f"Weibull PDF (k={k:.2f}, A={A:.2f})",
        mode="lines",
        line={"color": "crimson", "width": 2},
    ))

    fig.update_layout(
        title=f"{selected_station} — Weibull 풍속 분포",
        xaxis={"title": "풍속 (m/s)"},
        yaxis={"title": "확률밀도"},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.2},
    )
    st.plotly_chart(fig, use_container_width=True)


# ──────────────────────────────────────────────────────────────
# Tab 4: 풍속 등급별 발생일수
# ──────────────────────────────────────────────────────────────

def _tab_wind_class(df: pd.DataFrame) -> None:
    st.subheader("풍속 등급별 발생일수")

    if not _has_col(df, "wind_speed"):
        st.warning("wind_speed 컬럼이 없습니다.")
        return

    if "year" not in df.columns:
        st.warning("year 컬럼이 없습니다.")
        return

    def classify_wind(v: float) -> str:
        if v < 1:
            return "Calm (<1)"
        elif v < 5:
            return "Light (1~5)"
        elif v < 10:
            return "Moderate (5~10)"
        elif v < 15:
            return "Strong (10~15)"
        else:
            return "Gale (≥15)"

    CLASS_ORDER = ["Calm (<1)", "Light (1~5)", "Moderate (5~10)", "Strong (10~15)", "Gale (≥15)"]
    CLASS_COLORS = {
        "Calm (<1)": "#AED6F1",
        "Light (1~5)": "#82E0AA",
        "Moderate (5~10)": "#F9E79F",
        "Strong (10~15)": "#F0B27A",
        "Gale (≥15)": "#EC7063",
    }

    wind_df = df[["year", "wind_speed"]].dropna().copy()
    wind_df = wind_df.assign(wind_class=wind_df["wind_speed"].apply(classify_wind))

    class_counts = (
        wind_df.groupby(["year", "wind_class"], as_index=False)
        .size()
        .rename(columns={"size": "days"})
    )

    years_sorted = sorted(class_counts["year"].unique().tolist())

    fig = go.Figure()
    for cls in CLASS_ORDER:
        cdf = class_counts[class_counts["wind_class"] == cls]
        year_day_map = dict(zip(cdf["year"], cdf["days"]))
        y_vals = [year_day_map.get(yr, 0) for yr in years_sorted]

        fig.add_trace(go.Bar(
            x=years_sorted,
            y=y_vals,
            name=cls,
            marker_color=CLASS_COLORS[cls],
        ))

    fig.update_layout(
        barmode="stack",
        title="연도별 풍속 등급 발생일수",
        xaxis={"title": "연도"},
        yaxis={"title": "일수 (일)"},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.2},
    )
    st.plotly_chart(fig, use_container_width=True)


# ──────────────────────────────────────────────────────────────
# Tab 5: 기압 분석
# ──────────────────────────────────────────────────────────────

def _tab_pressure(df: pd.DataFrame) -> None:
    st.subheader("기압 분석")

    has_sea = _has_col(df, "pressure_sea")
    has_local = _has_col(df, "pressure_local")

    if not has_sea and not has_local:
        st.warning("pressure_sea 또는 pressure_local 컬럼이 없습니다.")
        return

    if "year" not in df.columns or "month" not in df.columns or "station_name" not in df.columns:
        st.warning("year, month, station_name 컬럼이 필요합니다.")
        return

    stations = _get_stations(df)
    selected = st.multiselect(
        "관측소 선택",
        options=stations,
        default=stations,
        key="wind_pressure_stations",
    )
    filtered = _filter_stations(df, selected)

    # 월별 평균 기압 시계열
    agg_cols = {}
    if has_sea:
        agg_cols["pressure_sea"] = ("pressure_sea", "mean")
    if has_local:
        agg_cols["pressure_local"] = ("pressure_local", "mean")

    monthly_p = filtered.groupby(["year", "month", "station_name"], as_index=False).agg(**agg_cols)
    monthly_p = monthly_p.assign(
        period=monthly_p["year"].astype(str) + "-" + monthly_p["month"].astype(str).str.zfill(2)
    ).sort_values(["station_name", "year", "month"])

    fig_monthly = go.Figure()
    color_seq = px.colors.qualitative.Plotly
    station_list = monthly_p["station_name"].unique().tolist()

    for idx, station in enumerate(station_list):
        sdf = monthly_p[monthly_p["station_name"] == station]
        color = color_seq[idx % len(color_seq)]

        if has_sea:
            fig_monthly.add_trace(go.Scatter(
                x=sdf["period"].tolist(),
                y=sdf["pressure_sea"].tolist(),
                name=f"{station} 해면기압",
                mode="lines",
                line={"color": color, "width": 1.5},
            ))
        if has_local:
            fig_monthly.add_trace(go.Scatter(
                x=sdf["period"].tolist(),
                y=sdf["pressure_local"].tolist(),
                name=f"{station} 현지기압",
                mode="lines",
                line={"color": color, "width": 1.5, "dash": "dot"},
            ))

    fig_monthly.update_layout(
        title="월별 평균 기압 시계열",
        xaxis={"title": "연월"},
        yaxis={"title": "기압 (hPa)"},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.25},
    )
    st.plotly_chart(fig_monthly, use_container_width=True)

    # 연간 기압 추이 + 회귀선
    pressure_col = "pressure_sea" if has_sea else "pressure_local"
    pressure_label = "해면기압 (hPa)" if has_sea else "현지기압 (hPa)"

    annual_p = filtered.groupby(["year", "station_name"], as_index=False).agg(
        pressure=(pressure_col, "mean")
    )

    fig_annual = go.Figure()

    for idx, station in enumerate(station_list):
        sdf = annual_p[annual_p["station_name"] == station].sort_values("year")
        if sdf.empty:
            continue
        color = color_seq[idx % len(color_seq)]
        x_vals = sdf["year"].values
        y_vals = sdf["pressure"].values

        fig_annual.add_trace(go.Scatter(
            x=x_vals.tolist(),
            y=y_vals.tolist(),
            name=station,
            mode="lines+markers",
            line={"color": color, "width": 2},
            marker={"size": 5},
        ))

        # 회귀선
        if len(x_vals) >= 2:
            coeffs = np.polyfit(x_vals, y_vals, 1)
            reg_y = np.polyval(coeffs, x_vals)
            fig_annual.add_trace(go.Scatter(
                x=x_vals.tolist(),
                y=reg_y.tolist(),
                name=f"{station} 추세",
                mode="lines",
                line={"color": color, "width": 1, "dash": "dash"},
                showlegend=False,
            ))

    fig_annual.update_layout(
        title=f"연간 {pressure_label} 추이 (회귀선 포함)",
        xaxis={"title": "연도"},
        yaxis={"title": pressure_label},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.2},
    )
    st.plotly_chart(fig_annual, use_container_width=True)


# ──────────────────────────────────────────────────────────────
# Public entry point
# ──────────────────────────────────────────────────────────────

def render(df: pd.DataFrame) -> None:
    """바람 분석 탭 렌더링. df는 필터링된 일별 DataFrame."""

    if df is None or df.empty:
        st.warning("데이터가 없습니다.")
        return

    stations = _get_stations(df)
    if stations:
        selected_global = st.multiselect(
            "분석 대상 관측소",
            options=stations,
            default=stations,
            key="wind_stations",
        )
        working_df = _filter_stations(df, selected_global)
    else:
        working_df = df.copy()

    tabs = st.tabs([
        "풍속 추이",
        "바람장미도",
        "Weibull 분포 분석",
        "풍속 등급별 발생일수",
        "기압 분석",
    ])

    with tabs[0]:
        _tab_wind_trend(working_df)

    with tabs[1]:
        _tab_wind_rose(working_df)

    with tabs[2]:
        _tab_weibull(working_df)

    with tabs[3]:
        _tab_wind_class(working_df)

    with tabs[4]:
        _tab_pressure(working_df)
