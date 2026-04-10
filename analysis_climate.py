import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats


def _check_col(df: pd.DataFrame, col: str) -> bool:
    if col not in df.columns:
        st.warning(f"컬럼 '{col}'이 데이터에 없습니다.")
        return False
    return True


def _station_select(df: pd.DataFrame) -> pd.DataFrame:
    if "station_name" not in df.columns:
        return df
    stations = sorted(df["station_name"].dropna().unique().tolist())
    selected = st.multiselect(
        "관측소 선택",
        options=stations,
        default=stations,
        key="climate_station_select",
    )
    if not selected:
        st.warning("관측소를 하나 이상 선택하세요.")
        return df
    return df[df["station_name"].isin(selected)]


def _annual_mean(df: pd.DataFrame, col: str) -> pd.DataFrame:
    group_cols = ["year", "station_name"] if "station_name" in df.columns else ["year"]
    return df.groupby(group_cols)[col].mean().reset_index()


def _annual_sum(df: pd.DataFrame, col: str) -> pd.DataFrame:
    group_cols = ["year", "station_name"] if "station_name" in df.columns else ["year"]
    return df.groupby(group_cols)[col].sum().reset_index()


def _add_moving_averages(series: pd.Series) -> tuple[pd.Series, pd.Series]:
    ma5 = series.rolling(5, min_periods=3).mean()
    ma10 = series.rolling(10, min_periods=5).mean()
    return ma5, ma10


def _regression_trace(x: np.ndarray, y: np.ndarray, name: str, color: str) -> tuple[go.Scatter, float, float]:
    mask = ~np.isnan(y)
    xv, yv = x[mask], y[mask]
    if len(xv) < 3:
        return None, np.nan, np.nan
    slope, intercept, r_value, _, _ = stats.linregress(xv, yv)
    y_pred = slope * x + intercept
    slope_per_decade = slope * 10
    trace = go.Scatter(
        x=x,
        y=y_pred,
        mode="lines",
        name=f"{name} 추세",
        line=dict(color=color, dash="dash", width=1.5),
        showlegend=True,
    )
    return trace, slope_per_decade, r_value ** 2


def _render_temp_trend(df: pd.DataFrame) -> None:
    if not _check_col(df, "temp_avg") or not _check_col(df, "year"):
        return

    annual = _annual_mean(df, "temp_avg")
    stations = annual["station_name"].unique() if "station_name" in annual.columns else ["전체"]
    colors = px.colors.qualitative.Set1

    fig = go.Figure()
    slope_records = []

    for i, station in enumerate(stations):
        color = colors[i % len(colors)]
        if "station_name" in annual.columns:
            sdf = annual[annual["station_name"] == station].sort_values("year")
        else:
            sdf = annual.sort_values("year")

        x = sdf["year"].values
        y = sdf["temp_avg"].values

        fig.add_trace(go.Scatter(
            x=x, y=y, mode="lines+markers",
            name=f"{station} 실측값",
            line=dict(color=color, width=1.5),
            marker=dict(size=5),
        ))

        ma5, ma10 = _add_moving_averages(pd.Series(y))
        fig.add_trace(go.Scatter(
            x=x, y=ma5.values, mode="lines",
            name=f"{station} 5년 이동평균",
            line=dict(color=color, dash="dash", width=1),
        ))
        fig.add_trace(go.Scatter(
            x=x, y=ma10.values, mode="lines",
            name=f"{station} 10년 이동평균",
            line=dict(color=color, dash="dot", width=1),
        ))

        trend_trace, slope_dec, r2 = _regression_trace(x.astype(float), y, station, color)
        if trend_trace is not None:
            fig.add_trace(trend_trace)
            if not np.isnan(slope_dec):
                x_mid = x[len(x) // 2]
                y_mid = (slope_dec / 10) * x_mid + (np.nanmean(y) - (slope_dec / 10) * np.nanmean(x))
                fig.add_annotation(
                    x=x_mid, y=float(y_mid),
                    text=f"{station}: {slope_dec:+.2f}°C/10년, R²={r2:.2f}",
                    showarrow=False, font=dict(size=10, color=color),
                    xanchor="left",
                )
            slope_records.append({"관측소": station, "기온변화(°C/10년)": f"{slope_dec:+.3f}", "R²": f"{r2:.3f}"})

    fig.update_layout(
        title="연평균 기온 추이",
        xaxis_title="연도",
        yaxis_title="기온 (°C)",
        height=420,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig, use_container_width=True)

    if slope_records:
        st.subheader("추세 요약")
        cols = st.columns(len(slope_records))
        for col, rec in zip(cols, slope_records):
            col.metric(
                label=rec["관측소"],
                value=rec["기온변화(°C/10년)"] + " °C/10년",
                delta=f"R²={rec['R²']}",
            )


def _render_precip_trend(df: pd.DataFrame) -> None:
    if not _check_col(df, "precipitation") or not _check_col(df, "year"):
        return

    annual = _annual_sum(df, "precipitation")
    stations = annual["station_name"].unique() if "station_name" in annual.columns else ["전체"]
    colors = px.colors.qualitative.Set2

    fig = go.Figure()
    slope_records = []

    for i, station in enumerate(stations):
        color = colors[i % len(colors)]
        if "station_name" in annual.columns:
            sdf = annual[annual["station_name"] == station].sort_values("year")
        else:
            sdf = annual.sort_values("year")

        x = sdf["year"].values
        y = sdf["precipitation"].values

        fig.add_trace(go.Bar(
            x=x, y=y,
            name=f"{station} 연강수량",
            marker_color=color,
            opacity=0.6,
        ))

        ma5, ma10 = _add_moving_averages(pd.Series(y))
        fig.add_trace(go.Scatter(
            x=x, y=ma5.values, mode="lines",
            name=f"{station} 5년 이동평균",
            line=dict(color=color, dash="dash", width=1.5),
        ))
        fig.add_trace(go.Scatter(
            x=x, y=ma10.values, mode="lines",
            name=f"{station} 10년 이동평균",
            line=dict(color=color, dash="dot", width=1.5),
        ))

        trend_trace, slope_dec, r2 = _regression_trace(x.astype(float), y, station, color)
        if trend_trace is not None:
            fig.add_trace(trend_trace)
            if not np.isnan(slope_dec):
                slope_records.append({"관측소": station, "강수량변화(mm/10년)": f"{slope_dec:+.1f}", "R²": f"{r2:.3f}"})

    fig.update_layout(
        title="연강수량 추이",
        xaxis_title="연도",
        yaxis_title="강수량 (mm)",
        barmode="group",
        height=420,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig, use_container_width=True)

    if slope_records:
        st.subheader("추세 요약")
        cols = st.columns(len(slope_records))
        for col, rec in zip(cols, slope_records):
            col.metric(
                label=rec["관측소"],
                value=rec["강수량변화(mm/10년)"] + " mm/10년",
                delta=f"R²={rec['R²']}",
            )


def _render_period_comparison(df: pd.DataFrame) -> None:
    if not _check_col(df, "year"):
        return

    years = sorted(df["year"].dropna().unique())
    if len(years) < 4:
        st.warning("전후반기 비교를 위해 최소 4개 연도 데이터가 필요합니다.")
        return

    mid_default = int(years[len(years) // 2])
    split_year = st.slider(
        "전후반기 분기점 (해당 연도부터 후반기)",
        min_value=int(years[0]),
        max_value=int(years[-1]),
        value=mid_default,
        key="climate_split_year",
    )

    first_half = df[df["year"] < split_year]
    second_half = df[df["year"] >= split_year]

    if first_half.empty or second_half.empty:
        st.warning("전반기 또는 후반기 데이터가 없습니다. 분기점을 조정하세요.")
        return

    def extreme_days(sub: pd.DataFrame, col: str, op: str, threshold: float) -> float:
        if col not in sub.columns:
            return np.nan
        if op == "ge":
            return sub.groupby("year")[col].apply(lambda s: (s >= threshold).sum()).mean()
        return sub.groupby("year")[col].apply(lambda s: (s <= threshold).sum()).mean()

    def annual_precip_mean(sub: pd.DataFrame) -> float:
        if "precipitation" not in sub.columns:
            return np.nan
        return sub.groupby("year")["precipitation"].sum().mean()

    elements = [
        ("평균기온 (°C)", "temp_avg", "mean"),
        ("최고기온 (°C)", "temp_max", "mean"),
        ("최저기온 (°C)", "temp_min", "mean"),
        ("강수량 (mm)", "precipitation", "mean"),
        ("습도 (%)", "humidity", "mean"),
        ("풍속 (m/s)", "wind_speed", "mean"),
        ("폭염일수 (일/년)", None, "heatwave"),
        ("열대야일수 (일/년)", None, "tropical_night"),
        ("한파일수 (일/년)", None, "cold_wave"),
        ("연강수량 (mm/년)", None, "annual_precip"),
    ]

    rows = []
    for label, col, stat_type in elements:
        if stat_type == "mean":
            if col not in df.columns:
                continue
            v1 = first_half[col].mean()
            v2 = second_half[col].mean()
        elif stat_type == "heatwave":
            v1 = extreme_days(first_half, "temp_max", "ge", 33)
            v2 = extreme_days(second_half, "temp_max", "ge", 33)
        elif stat_type == "tropical_night":
            v1 = extreme_days(first_half, "temp_min", "ge", 25)
            v2 = extreme_days(second_half, "temp_min", "ge", 25)
        elif stat_type == "cold_wave":
            v1 = extreme_days(first_half, "temp_min", "le", -12)
            v2 = extreme_days(second_half, "temp_min", "le", -12)
        elif stat_type == "annual_precip":
            v1 = annual_precip_mean(first_half)
            v2 = annual_precip_mean(second_half)
        else:
            continue

        if np.isnan(v1) or np.isnan(v2):
            continue

        change = v2 - v1
        rate = (change / v1 * 100) if v1 != 0 else np.nan
        rows.append({
            "항목": label,
            f"전반기 평균 (~{split_year-1})": round(v1, 2),
            f"후반기 평균 ({split_year}~)": round(v2, 2),
            "변화량": round(change, 2),
            "변화율 (%)": f"{rate:+.1f}%" if not np.isnan(rate) else "-",
        })

    if not rows:
        st.warning("비교할 데이터가 없습니다.")
        return

    result_df = pd.DataFrame(rows)
    st.dataframe(result_df, use_container_width=True, hide_index=True)


def _render_climate_reference(df: pd.DataFrame) -> None:
    st.info(
        "업로드된 자료 기반 분석 결과입니다. 공식 평년값은 아래 기상청 페이지를 참조하세요."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.link_button(
            "🌐 기상청 기후통계 (지역별 평년값)",
            "https://www.weather.go.kr/w/climate/statistics/region.do",
        )
    with col2:
        st.link_button(
            "🌐 기후변화 상황지도",
            "https://climate.go.kr/atlas/",
        )
    with col3:
        st.link_button(
            "🌐 국가가뭄정보포털",
            "https://www.drought.go.kr/main.do",
        )

    st.subheader("업로드 자료 기반 월별 기후 평균값")

    numeric_cols = [
        c for c in ["temp_avg", "temp_max", "temp_min", "precipitation", "humidity", "wind_speed"]
        if c in df.columns
    ]
    if not numeric_cols or "month" not in df.columns:
        st.warning("월별 기후 평균값을 계산하기 위한 컬럼이 없습니다.")
        return

    group_cols = ["station_name", "month"] if "station_name" in df.columns else ["month"]
    monthly = df.groupby(group_cols)[numeric_cols].mean().round(2).reset_index()

    col_rename = {
        "temp_avg": "평균기온(°C)",
        "temp_max": "최고기온(°C)",
        "temp_min": "최저기온(°C)",
        "precipitation": "강수량(mm)",
        "humidity": "습도(%)",
        "wind_speed": "풍속(m/s)",
        "month": "월",
        "station_name": "관측소",
    }
    monthly = monthly.rename(columns=col_rename)
    st.dataframe(monthly, use_container_width=True, hide_index=True)


def render(df: pd.DataFrame) -> None:
    st.header("기후변화 분석")

    if df is None or df.empty:
        st.warning("표시할 데이터가 없습니다.")
        return

    df = _station_select(df)
    if df.empty:
        return

    tab1, tab2, tab3, tab4 = st.tabs([
        "연평균 기온 추이",
        "연강수량 추이",
        "전후반기 비교",
        "기후통계 참조",
    ])

    with tab1:
        _render_temp_trend(df)

    with tab2:
        _render_precip_trend(df)

    with tab3:
        _render_period_comparison(df)

    with tab4:
        _render_climate_reference(df)
