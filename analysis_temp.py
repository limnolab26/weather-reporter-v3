# analysis_temp.py — 기온 분석 모듈
# render(df) 함수 하나를 export합니다.

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats
import io


# ── 내부 헬퍼 ──────────────────────────────────────────────────────────────

def _has_col(df: pd.DataFrame, *cols: str) -> bool:
    return all(c in df.columns for c in cols)


def _filter_stations(df: pd.DataFrame, stations: list) -> pd.DataFrame:
    if not stations:
        return df
    return df[df["station_name"].isin(stations)]


def _csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
    return buf.getvalue()


# ── Sub-tab 1: 연도별 통계표 ────────────────────────────────────────────────

def _tab_yearly_stats(df: pd.DataFrame) -> None:
    required = {"station_name", "year"}
    if not required.issubset(df.columns):
        st.warning("station_name 또는 year 컬럼이 없습니다.")
        return

    rows = []
    for (station, year), g in df.groupby(["station_name", "year"]):
        row = {"관측소": station, "연도": year}

        row["평균기온"] = round(g["temp_avg"].mean(), 1) if "temp_avg" in g else np.nan
        row["연최고기온"] = round(g["temp_max"].max(), 1) if "temp_max" in g else np.nan
        row["연최저기온"] = round(g["temp_min"].min(), 1) if "temp_min" in g else np.nan

        if "temp_max" in g.columns:
            row["폭염일수(Tmax≥33°C)"] = int((g["temp_max"] >= 33).sum())
        else:
            row["폭염일수(Tmax≥33°C)"] = np.nan

        if "temp_min" in g.columns:
            row["열대야일수(Tmin≥25°C)"] = int((g["temp_min"] >= 25).sum())
            row["한파일수(Tmin≤-12°C)"] = int((g["temp_min"] <= -12).sum())
            row["서리가능일수(Tmin≤0°C)"] = int((g["temp_min"] <= 0).sum())
        else:
            row["열대야일수(Tmin≥25°C)"] = np.nan
            row["한파일수(Tmin≤-12°C)"] = np.nan
            row["서리가능일수(Tmin≤0°C)"] = np.nan

        rows.append(row)

    if not rows:
        st.warning("표시할 데이터가 없습니다.")
        return

    result_df = pd.DataFrame(rows)
    st.dataframe(result_df, use_container_width=True)
    st.download_button(
        label="CSV 다운로드",
        data=_csv_bytes(result_df),
        file_name="연도별_기온통계.csv",
        mime="text/csv",
        key="dl_yearly_temp_stats",
    )


# ── Sub-tab 2: 극한기온 발생일수 ────────────────────────────────────────────

def _tab_extreme_days(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "station_name"):
        st.warning("year 또는 station_name 컬럼이 없습니다.")
        return

    metric_opts = []
    if "temp_max" in df.columns:
        metric_opts += ["폭염일수(Tmax≥33°C)"]
    if "temp_min" in df.columns:
        metric_opts += ["열대야일수(Tmin≥25°C)", "한파일수(Tmin≤-12°C)", "서리가능일수(Tmin≤0°C)"]

    if not metric_opts:
        st.warning("기온 컬럼(temp_max / temp_min)이 없습니다.")
        return

    selected_metrics = st.multiselect(
        "표시할 지표 선택",
        options=metric_opts,
        default=metric_opts[:2],
        key="extreme_metrics_select",
    )

    stations = df["station_name"].unique().tolist()
    selected_station = st.selectbox(
        "관측소 선택",
        options=stations,
        key="extreme_station_select",
    )

    g = df[df["station_name"] == selected_station]

    rows = []
    for year, yg in g.groupby("year"):
        row = {"year": year}
        if "폭염일수(Tmax≥33°C)" in selected_metrics and "temp_max" in yg.columns:
            row["폭염일수(Tmax≥33°C)"] = int((yg["temp_max"] >= 33).sum())
        if "열대야일수(Tmin≥25°C)" in selected_metrics and "temp_min" in yg.columns:
            row["열대야일수(Tmin≥25°C)"] = int((yg["temp_min"] >= 25).sum())
        if "한파일수(Tmin≤-12°C)" in selected_metrics and "temp_min" in yg.columns:
            row["한파일수(Tmin≤-12°C)"] = int((yg["temp_min"] <= -12).sum())
        if "서리가능일수(Tmin≤0°C)" in selected_metrics and "temp_min" in yg.columns:
            row["서리가능일수(Tmin≤0°C)"] = int((yg["temp_min"] <= 0).sum())
        rows.append(row)

    if not rows or not selected_metrics:
        st.info("선택된 지표 또는 데이터가 없습니다.")
        return

    plot_df = pd.DataFrame(rows)
    value_cols = [c for c in selected_metrics if c in plot_df.columns]
    melted = plot_df.melt(id_vars="year", value_vars=value_cols, var_name="지표", value_name="일수")

    fig = px.bar(
        melted,
        x="year",
        y="일수",
        color="지표",
        barmode="group",
        title=f"{selected_station} — 극한기온 발생일수",
        height=400,
    )
    fig.update_xaxes(title="연도")
    fig.update_yaxes(title="일수 (일)")
    st.plotly_chart(fig, use_container_width=True)


# ── Sub-tab 3: 월별 기온 분포 ────────────────────────────────────────────────

def _tab_monthly_boxplot(df: pd.DataFrame) -> None:
    if not _has_col(df, "month", "station_name"):
        st.warning("month 또는 station_name 컬럼이 없습니다.")
        return

    temp_options = {
        "평균기온 (temp_avg)": "temp_avg",
        "최고기온 (temp_max)": "temp_max",
        "최저기온 (temp_min)": "temp_min",
    }
    available = {k: v for k, v in temp_options.items() if v in df.columns}
    if not available:
        st.warning("기온 컬럼이 없습니다.")
        return

    selected_label = st.selectbox(
        "기온 요소 선택",
        options=list(available.keys()),
        key="monthly_box_temp_select",
    )
    y_col = available[selected_label]

    fig = px.box(
        df,
        x="month",
        y=y_col,
        color="station_name",
        title=f"월별 {selected_label} 분포",
        height=400,
        labels={"month": "월", y_col: selected_label, "station_name": "관측소"},
    )
    fig.update_xaxes(tickvals=list(range(1, 13)), title="월")
    st.plotly_chart(fig, use_container_width=True)


# ── Sub-tab 4: 장기 추세 ─────────────────────────────────────────────────────

def _tab_long_term_trend(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "temp_avg", "station_name"):
        st.warning("year, temp_avg, station_name 컬럼이 필요합니다.")
        return

    annual = (
        df.groupby(["station_name", "year"])["temp_avg"]
        .mean()
        .reset_index()
        .rename(columns={"temp_avg": "연평균기온"})
    )

    fig = go.Figure()
    colors = px.colors.qualitative.Set1

    for idx, station in enumerate(annual["station_name"].unique()):
        s_df = annual[annual["station_name"] == station].dropna(subset=["연평균기온"])
        if len(s_df) < 2:
            continue

        color = colors[idx % len(colors)]
        x_vals = s_df["year"].values
        y_vals = s_df["연평균기온"].values

        fig.add_trace(go.Scatter(
            x=x_vals,
            y=y_vals,
            mode="markers+lines",
            name=station,
            line={"color": color},
            marker={"size": 5},
        ))

        slope, intercept, r_value, _p, _se = stats.linregress(x_vals, y_vals)
        y_trend = intercept + slope * x_vals
        r2 = round(r_value ** 2, 3)
        slope_label = f"{slope:+.3f} °C/yr"

        fig.add_trace(go.Scatter(
            x=x_vals,
            y=y_trend,
            mode="lines",
            name=f"{station} 추세 ({slope_label}, R²={r2})",
            line={"color": color, "dash": "dash", "width": 1.5},
            showlegend=True,
        ))

    fig.update_layout(
        title="연평균기온 장기 추세",
        xaxis_title="연도",
        yaxis_title="연평균기온 (°C)",
        height=400,
        legend={"orientation": "h", "yanchor": "bottom", "y": -0.3},
    )
    st.plotly_chart(fig, use_container_width=True)


# ── Sub-tab 5: 연도×월 히트맵 ──────────────────────────────────────────────

def _tab_year_month_heatmap(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "month", "temp_avg"):
        st.warning("year, month, temp_avg 컬럼이 필요합니다.")
        return

    stations = df["station_name"].unique().tolist() if "station_name" in df.columns else ["전체"]

    selected_station = st.selectbox(
        "관측소 선택 (히트맵)",
        options=stations,
        key="heatmap_station_select",
    )

    if "station_name" in df.columns and selected_station != "전체":
        plot_df = df[df["station_name"] == selected_station]
    else:
        plot_df = df

    pivot = (
        plot_df.groupby(["year", "month"])["temp_avg"]
        .mean()
        .unstack(level="month")
    )

    if pivot.empty:
        st.warning("히트맵을 그릴 데이터가 없습니다.")
        return

    pivot.columns = [f"{m}월" for m in pivot.columns]

    fig = px.imshow(
        pivot,
        color_continuous_scale="RdYlBu_r",
        title=f"{selected_station} — 연도×월 평균기온 히트맵 (°C)",
        labels={"x": "월", "y": "연도", "color": "기온(°C)"},
        aspect="auto",
        height=400,
    )
    fig.update_xaxes(side="bottom")
    st.plotly_chart(fig, use_container_width=True)


# ── 공개 render 함수 ─────────────────────────────────────────────────────────

def render(df: pd.DataFrame) -> None:
    """기온 분석 탭 렌더링. df는 필터링된 일자료 DataFrame."""
    st.subheader("기온 분석")

    if df is None or df.empty:
        st.warning("데이터가 없습니다.")
        return

    # 관측소 필터
    if "station_name" in df.columns:
        all_stations = sorted(df["station_name"].unique().tolist())
        selected_stations = st.multiselect(
            "관측소 선택 (전체 탭 적용)",
            options=all_stations,
            default=all_stations,
            key="temp_station_multiselect",
        )
        filtered = _filter_stations(df, selected_stations)
    else:
        filtered = df

    if filtered.empty:
        st.warning("선택된 관측소에 데이터가 없습니다.")
        return

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "연도별 통계표",
        "극한기온 발생일수",
        "월별 기온 분포",
        "장기 추세",
        "연도×월 히트맵",
    ])

    with tab1:
        _tab_yearly_stats(filtered)

    with tab2:
        _tab_extreme_days(filtered)

    with tab3:
        _tab_monthly_boxplot(filtered)

    with tab4:
        _tab_long_term_trend(filtered)

    with tab5:
        _tab_year_month_heatmap(filtered)
