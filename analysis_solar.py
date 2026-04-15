# analysis_solar.py — 태양광 분석 모듈
# render(df) 함수: 필터링된 일별 DataFrame을 받아 5개 서브탭 렌더링

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from chart_utils import chart_download_btn

SEASON_ORDER = ["봄", "여름", "가을", "겨울"]
CHART_HEIGHT = 400


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


# ──────────────────────────────────────────────────────────────
# Tab 1: 월별 일사량·일조시간
# ──────────────────────────────────────────────────────────────

def _tab_monthly_solar(df: pd.DataFrame) -> None:
    st.subheader("월별 일사량 · 일조시간")

    has_solar = _has_col(df, "solar_rad")
    has_sunshine = _has_col(df, "sunshine")

    if not has_solar and not has_sunshine:
        st.warning("solar_rad 또는 sunshine 컬럼이 없습니다.")
        return

    stations = _get_stations(df)
    if not stations:
        st.info("station_name 컬럼이 없습니다.")
        return

    selected_station = st.selectbox(
        "관측소 선택",
        options=stations,
        key="solar_monthly_station",
    )

    filtered = df[df["station_name"] == selected_station].copy() if "station_name" in df.columns else df.copy()

    if "month" not in filtered.columns:
        st.warning("month 컬럼이 없습니다.")
        return

    agg: dict = {"month": filtered["month"]}
    if has_solar:
        agg["solar_rad"] = filtered["solar_rad"]
    if has_sunshine:
        agg["sunshine"] = filtered["sunshine"]

    monthly = filtered.groupby("month", as_index=False).agg(
        **{k: (k, "mean") for k in (["solar_rad"] if has_solar else []) + (["sunshine"] if has_sunshine else [])}
    )
    monthly = monthly.sort_values("month")

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    if has_solar:
        fig.add_trace(
            go.Bar(
                x=monthly["month"],
                y=monthly["solar_rad"],
                name="평균 일사량 (MJ/m²)",
                marker_color="rgba(255, 165, 0, 0.7)",
            ),
            secondary_y=False,
        )

    if has_sunshine:
        fig.add_trace(
            go.Scatter(
                x=monthly["month"],
                y=monthly["sunshine"],
                name="평균 일조시간 (hr)",
                mode="lines+markers",
                line={"color": "royalblue", "width": 2},
                marker={"size": 6},
            ),
            secondary_y=True,
        )

    fig.update_layout(
        title=f"{selected_station} — 월별 평균 일사량·일조시간",
        xaxis={"title": "월", "tickvals": list(range(1, 13))},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.2},
    )
    fig.update_yaxes(title_text="일사량 (MJ/m²)", secondary_y=False)
    fig.update_yaxes(title_text="일조시간 (hr)", secondary_y=True)

    st.plotly_chart(fig, use_container_width=True)
    chart_download_btn(fig, key="solar_monthly_chart", filename="monthly_solar_sunshine")


# ──────────────────────────────────────────────────────────────
# Tab 2: 연간 합계 추이
# ──────────────────────────────────────────────────────────────

def _tab_annual_trend(df: pd.DataFrame) -> None:
    st.subheader("연간 합계 추이")

    has_solar = _has_col(df, "solar_rad")
    has_sunshine = _has_col(df, "sunshine")

    if not has_solar and not has_sunshine:
        st.warning("solar_rad 또는 sunshine 컬럼이 없습니다.")
        return

    if "year" not in df.columns or "station_name" not in df.columns:
        st.warning("year 또는 station_name 컬럼이 없습니다.")
        return

    stations = _get_stations(df)
    selected = st.multiselect(
        "관측소 선택",
        options=stations,
        default=stations,
        key="solar_annual_stations",
    )
    filtered = _filter_stations(df, selected)

    agg_cols = {}
    if has_solar:
        agg_cols["solar_rad"] = ("solar_rad", "sum")
    if has_sunshine:
        agg_cols["sunshine"] = ("sunshine", "sum")

    annual = filtered.groupby(["year", "station_name"], as_index=False).agg(**agg_cols)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    color_seq = px.colors.qualitative.Plotly
    station_list = annual["station_name"].unique().tolist()

    for idx, station in enumerate(station_list):
        sdf = annual[annual["station_name"] == station].sort_values("year")
        color = color_seq[idx % len(color_seq)]

        if has_solar:
            fig.add_trace(
                go.Scatter(
                    x=sdf["year"],
                    y=sdf["solar_rad"],
                    name=f"{station} 일사량",
                    mode="lines+markers",
                    line={"color": color, "width": 2},
                    marker={"size": 5},
                ),
                secondary_y=False,
            )

        if has_sunshine:
            fig.add_trace(
                go.Scatter(
                    x=sdf["year"],
                    y=sdf["sunshine"],
                    name=f"{station} 일조시간",
                    mode="lines+markers",
                    line={"color": color, "width": 2, "dash": "dash"},
                    marker={"size": 5, "symbol": "diamond"},
                ),
                secondary_y=True,
            )

    fig.update_layout(
        title="연간 합계 추이",
        xaxis={"title": "연도"},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.25},
    )
    fig.update_yaxes(title_text="일사량 합계 (MJ/m²)", secondary_y=False)
    fig.update_yaxes(title_text="일조시간 합계 (hr)", secondary_y=True)

    st.plotly_chart(fig, use_container_width=True)
    chart_download_btn(fig, key="solar_annual_chart", filename="annual_solar_sunshine")


# ──────────────────────────────────────────────────────────────
# Tab 3: 계절별 분석
# ──────────────────────────────────────────────────────────────

def _tab_seasonal(df: pd.DataFrame) -> None:
    st.subheader("계절별 분석")

    if not _has_col(df, "solar_rad"):
        st.warning("solar_rad 컬럼이 없습니다.")
        return

    if "season" not in df.columns or "station_name" not in df.columns:
        st.warning("season 또는 station_name 컬럼이 없습니다.")
        return

    stations = _get_stations(df)
    selected = st.multiselect(
        "관측소 선택",
        options=stations,
        default=stations,
        key="solar_seasonal_stations",
    )
    filtered = _filter_stations(df, selected)

    fig = px.box(
        filtered,
        x="season",
        y="solar_rad",
        color="station_name",
        category_orders={"season": SEASON_ORDER},
        labels={"solar_rad": "일사량 (MJ/m²)", "season": "계절", "station_name": "관측소"},
        title="계절별 일사량 분포",
        height=CHART_HEIGHT,
    )
    fig.update_layout(legend={"orientation": "h", "y": -0.2})
    st.plotly_chart(fig, use_container_width=True)
    chart_download_btn(fig, key="solar_seasonal_chart", filename="seasonal_solar_boxplot")


# ──────────────────────────────────────────────────────────────
# Tab 4: 전운량 분석
# ──────────────────────────────────────────────────────────────

def _tab_cloud(df: pd.DataFrame) -> None:
    st.subheader("전운량 분석")

    if not _has_col(df, "cloud_cover"):
        st.warning("cloud_cover 컬럼이 없습니다.")
        return

    if "year" not in df.columns or "station_name" not in df.columns:
        st.warning("year 또는 station_name 컬럼이 없습니다.")
        return

    stations = _get_stations(df)
    selected = st.multiselect(
        "관측소 선택",
        options=stations,
        default=stations,
        key="solar_cloud_stations",
    )
    filtered = _filter_stations(df, selected)

    # 연간 평균 전운량 추이
    annual_cloud = filtered.groupby(["year", "station_name"], as_index=False).agg(
        cloud_cover=("cloud_cover", "mean")
    )

    fig_line = px.line(
        annual_cloud,
        x="year",
        y="cloud_cover",
        color="station_name",
        labels={"cloud_cover": "평균 전운량 (1/10)", "year": "연도", "station_name": "관측소"},
        title="연간 평균 전운량 추이",
        height=CHART_HEIGHT,
        markers=True,
    )
    fig_line.update_layout(legend={"orientation": "h", "y": -0.2})
    st.plotly_chart(fig_line, use_container_width=True)
    chart_download_btn(fig_line, key="solar_cloud_trend_chart", filename="annual_cloud_cover")

    # 운량 등급 분류 (10분위 기준)
    def classify_cloud(val: float) -> str:
        if val <= 2:
            return "맑음"
        elif val <= 5:
            return "구름조금"
        elif val <= 8:
            return "구름많음"
        else:
            return "흐림"

    cloud_df = filtered[["year", "station_name", "cloud_cover"]].dropna().copy()
    cloud_df = cloud_df.assign(cloud_class=cloud_df["cloud_cover"].apply(classify_cloud))

    class_order = ["맑음", "구름조금", "구름많음", "흐림"]
    color_map = {
        "맑음": "#87CEEB",
        "구름조금": "#FFD700",
        "구름많음": "#A9A9A9",
        "흐림": "#696969",
    }

    cloud_counts = (
        cloud_df.groupby(["year", "cloud_class"], as_index=False)
        .size()
        .rename(columns={"size": "days"})
    )

    fig_bar = go.Figure()
    for cls in class_order:
        cdf = cloud_counts[cloud_counts["cloud_class"] == cls]
        fig_bar.add_trace(
            go.Bar(
                x=cdf["year"],
                y=cdf["days"],
                name=cls,
                marker_color=color_map[cls],
            )
        )

    fig_bar.update_layout(
        barmode="stack",
        title="연도별 운량 등급 발생일수",
        xaxis={"title": "연도"},
        yaxis={"title": "일수 (일)"},
        height=CHART_HEIGHT,
        legend={"orientation": "h", "y": -0.2},
    )
    st.plotly_chart(fig_bar, use_container_width=True)
    chart_download_btn(fig_bar, key="solar_cloud_class_chart", filename="cloud_classification")


# ──────────────────────────────────────────────────────────────
# Tab 5: 일조율 분석
# ──────────────────────────────────────────────────────────────

def _tab_sunshine_ratio(df: pd.DataFrame) -> None:
    st.subheader("일조율 분석")

    if not _has_col(df, "sunshine") or not _has_col(df, "daylight_hours"):
        st.warning("sunshine 또는 daylight_hours 컬럼이 없거나 데이터가 부족합니다.")
        return

    if "month" not in df.columns or "year" not in df.columns or "station_name" not in df.columns:
        st.warning("month, year, station_name 컬럼이 필요합니다.")
        return

    stations = _get_stations(df)
    selected = st.multiselect(
        "관측소 선택",
        options=stations,
        default=stations,
        key="solar_ratio_stations",
    )
    filtered = _filter_stations(df, selected)

    ratio_df = filtered[filtered["daylight_hours"] > 0].copy()
    ratio_df = ratio_df.assign(
        sunshine_ratio=(ratio_df["sunshine"] / ratio_df["daylight_hours"] * 100).clip(0, 100)
    )

    # 월별 평균 일조율
    monthly_ratio = ratio_df.groupby(["month", "station_name"], as_index=False).agg(
        sunshine_ratio=("sunshine_ratio", "mean")
    )

    fig_bar = px.bar(
        monthly_ratio,
        x="month",
        y="sunshine_ratio",
        color="station_name",
        barmode="group",
        labels={"sunshine_ratio": "일조율 (%)", "month": "월", "station_name": "관측소"},
        title="월별 평균 일조율",
        height=CHART_HEIGHT,
    )
    fig_bar.update_layout(
        xaxis={"tickvals": list(range(1, 13))},
        legend={"orientation": "h", "y": -0.2},
    )
    st.plotly_chart(fig_bar, use_container_width=True)
    chart_download_btn(fig_bar, key="solar_monthly_ratio_chart", filename="monthly_sunshine_ratio")

    # 연간 일조율 추이
    annual_ratio = ratio_df.groupby(["year", "station_name"], as_index=False).agg(
        sunshine_ratio=("sunshine_ratio", "mean")
    )

    fig_line = px.line(
        annual_ratio,
        x="year",
        y="sunshine_ratio",
        color="station_name",
        labels={"sunshine_ratio": "연간 평균 일조율 (%)", "year": "연도", "station_name": "관측소"},
        title="연간 평균 일조율 추이",
        height=CHART_HEIGHT,
        markers=True,
    )
    fig_line.update_layout(legend={"orientation": "h", "y": -0.2})
    st.plotly_chart(fig_line, use_container_width=True)
    chart_download_btn(fig_line, key="solar_annual_ratio_chart", filename="annual_sunshine_ratio")


# ──────────────────────────────────────────────────────────────
# Public entry point
# ──────────────────────────────────────────────────────────────

def render(df: pd.DataFrame) -> None:
    """태양광 분석 탭 렌더링. df는 필터링된 일별 DataFrame."""

    if df is None or df.empty:
        st.warning("데이터가 없습니다.")
        return

    stations = _get_stations(df)
    if stations:
        selected_global = st.multiselect(
            "분석 대상 관측소",
            options=stations,
            default=stations,
            key="solar_stations",
        )
        working_df = _filter_stations(df, selected_global)
    else:
        working_df = df.copy()

    tabs = st.tabs([
        "월별 일사량·일조시간",
        "연간 합계 추이",
        "계절별 분석",
        "전운량 분석",
        "일조율 분석",
    ])

    with tabs[0]:
        _tab_monthly_solar(working_df)

    with tabs[1]:
        _tab_annual_trend(working_df)

    with tabs[2]:
        _tab_seasonal(working_df)

    with tabs[3]:
        _tab_cloud(working_df)

    with tabs[4]:
        _tab_sunshine_ratio(working_df)
