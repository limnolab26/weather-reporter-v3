import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats

# ── 통계 검정 상수 ────────────────────────────────────────────────
_MK_LARGE_N_THRESHOLD = 100  # Mann-Kendall n>100이면 sampling 경고
_PETTITT_P_SIG = 0.05        # 유의수준
_MK_P_SIG = 0.05


def _mann_kendall(x: np.ndarray) -> tuple:
    """Mann-Kendall 추세 검정 (numpy only, scipy.stats.norm 사용).

    Returns:
        (S, Z_stat, p_value, trend_direction)
    """
    n = len(x)
    s = 0
    for k in range(n - 1):
        for j in range(k + 1, n):
            s += np.sign(x[j] - x[k])
    var_s = n * (n - 1) * (2 * n + 5) / 18
    if s > 0:
        z = (s - 1) / np.sqrt(var_s)
    elif s < 0:
        z = (s + 1) / np.sqrt(var_s)
    else:
        z = 0.0
    p = 2 * (1 - stats.norm.cdf(abs(z)))
    if p < _MK_P_SIG:
        direction = "증가" if s > 0 else "감소"
    else:
        direction = "없음 (유의하지 않음)"
    return float(s), float(z), float(p), direction


def _sens_slope(x: np.ndarray, y: np.ndarray) -> float:
    """Sen's Slope: 모든 쌍의 기울기 중앙값."""
    slopes = []
    n = len(y)
    for i in range(n - 1):
        for j in range(i + 1, n):
            if x[j] != x[i]:
                slopes.append((y[j] - y[i]) / (x[j] - x[i]))
    return float(np.median(slopes)) if slopes else np.nan


def _pettitt_test(x: np.ndarray) -> tuple:
    """Pettitt 변화점 탐지.

    Returns:
        (change_point_index, K_stat, p_value)
    """
    n = len(x)
    K = np.zeros(n)
    for t in range(n):
        s = 0
        for i in range(t + 1):
            for j in range(t + 1, n):
                s += np.sign(x[i] - x[j])
        K[t] = abs(s)
    k_max = np.max(K)
    t_star = int(np.argmax(K))
    p = 2 * np.exp(-6 * k_max ** 2 / (n ** 3 + n ** 2))
    return t_star, float(k_max), float(p)


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

        # 자동 해석 텍스트
        for rec in slope_records:
            slope_val = float(rec["기온변화(°C/10년)"])
            r2_val = float(rec["R²"])
            direction = "상승" if slope_val > 0 else "하강"
            fit_desc = "비교적 강한 선형 관계" if r2_val >= 0.5 else "변동성이 있는 선형 추세"
            st.info(
                f"**{rec['관측소']}**: 분석 기간 중 평균기온은 선형 추세 기준 10년당 "
                f"{slope_val:+.2f}°C {direction}하였으며 (R²={r2_val:.2f}, {fit_desc}), "
                "이는 기후변화에 따른 전반적 온난화 경향과 비교할 수 있습니다."
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

        # 자동 해석 텍스트
        for rec in slope_records:
            slope_val = float(rec["강수량변화(mm/10년)"])
            r2_val = float(rec["R²"])
            direction = "증가" if slope_val > 0 else "감소"
            variability = "변동성이 매우 큰 편" if r2_val < 0.2 else ("변동성이 있는 편" if r2_val < 0.5 else "비교적 안정적인 추세")
            st.info(
                f"**{rec['관측소']}**: 연강수량은 10년당 {slope_val:+.1f}mm {direction} 추세이나, "
                f"R²={r2_val:.2f}로 {variability}입니다."
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
    _csv = result_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", _csv, "climate_detection.csv", "text/csv", key="climate_det_csv")


def _build_mk_results(df: pd.DataFrame, col: str, agg: str) -> pd.DataFrame:
    """관측소별 Mann-Kendall / Sen's Slope / Pettitt 결과 DataFrame 반환."""
    if "station_name" not in df.columns:
        stations = ["전체"]
    else:
        stations = sorted(df["station_name"].dropna().unique().tolist())

    rows = []
    for station in stations:
        if "station_name" in df.columns:
            sdf = df[df["station_name"] == station]
        else:
            sdf = df

        if agg == "mean":
            annual = sdf.groupby("year")[col].mean().reset_index().sort_values("year")
        else:
            annual = sdf.groupby("year")[col].sum().reset_index().sort_values("year")

        x = annual["year"].values.astype(float)
        y = annual[col].values.astype(float)

        if len(y) < 4:
            continue

        s_val, z_val, p_val, direction = _mann_kendall(y)
        sen = _sens_slope(x, y)
        t_star, k_max, p_pettitt = _pettitt_test(y)
        cp_year = int(annual["year"].iloc[t_star]) if t_star < len(annual) else None

        rows.append({
            "관측소": station,
            "추세방향": direction,
            "Mann-Kendall Z": round(z_val, 3),
            "p값": round(p_val, 4),
            "유의성(p<0.05)": "O" if p_val < _MK_P_SIG else "X",
            "Sen's Slope(단위/년)": round(sen, 4) if not np.isnan(sen) else None,
            "Pettitt 변화점(연도)": cp_year,
            "Pettitt p값": round(p_pettitt, 4),
        })
    return pd.DataFrame(rows)


def _mk_interpretation(row: dict, unit: str) -> str:
    """단일 관측소 결과를 자연어로 설명."""
    station = row["관측소"]
    z = row["Mann-Kendall Z"]
    p = row["p값"]
    sen = row["Sen's Slope(단위/년)"]
    cp_year = row["Pettitt 변화점(연도)"]
    p_pettitt = row["Pettitt p값"]

    if row["유의성(p<0.05)"] == "O":
        trend_msg = (
            f"Mann-Kendall 검정 결과 유의한 **{row['추세방향']}** 추세 "
            f"(Z={z:.3f}, p={p:.4f})가 확인되었으며, "
            f"Sen's Slope는 {sen:+.4f}{unit}/년 ({sen * 10:+.3f}{unit}/10년)입니다."
        )
    else:
        trend_msg = (
            f"Mann-Kendall 검정 결과 유의한 추세가 확인되지 않았습니다 "
            f"(Z={z:.3f}, p={p:.4f})."
        )

    if cp_year is not None and p_pettitt < _PETTITT_P_SIG:
        cp_msg = f" Pettitt 변화점 검정 결과 **{cp_year}년**을 기점으로 유의하게 변화하였습니다 (p={p_pettitt:.4f})."
    elif cp_year is not None:
        cp_msg = f" Pettitt 변화점으로 {cp_year}년이 탐지되었으나 통계적으로 유의하지 않습니다 (p={p_pettitt:.4f})."
    else:
        cp_msg = ""

    return f"**{station}** 관측소: {trend_msg}{cp_msg}"


def _mk_chart(df: pd.DataFrame, col: str, agg: str, result_df: pd.DataFrame) -> None:
    """시계열 + Pettitt 변화점 수직선 차트."""
    if "station_name" not in df.columns:
        stations = ["전체"]
    else:
        stations = sorted(df["station_name"].dropna().unique().tolist())

    colors = px.colors.qualitative.Set1
    fig = go.Figure()

    for i, station in enumerate(stations):
        color = colors[i % len(colors)]
        if "station_name" in df.columns:
            sdf = df[df["station_name"] == station]
        else:
            sdf = df

        if agg == "mean":
            annual = sdf.groupby("year")[col].mean().reset_index().sort_values("year")
        else:
            annual = sdf.groupby("year")[col].sum().reset_index().sort_values("year")

        fig.add_trace(go.Scatter(
            x=annual["year"], y=annual[col],
            mode="lines+markers",
            name=station,
            line=dict(color=color, width=1.5),
            marker=dict(size=4),
        ))

        # 변화점 수직선
        match = result_df[result_df["관측소"] == station]
        if not match.empty:
            cp_year = match.iloc[0]["Pettitt 변화점(연도)"]
            p_pettitt = match.iloc[0]["Pettitt p값"]
            if cp_year is not None and not np.isnan(float(p_pettitt)):
                line_color = "red" if float(p_pettitt) < _PETTITT_P_SIG else "orange"
                fig.add_vline(
                    x=cp_year,
                    line_dash="dash",
                    line_color=line_color,
                    annotation_text=f"{station} 변화점({cp_year})",
                    annotation_position="top right",
                )

    yaxis_label = "기온 (°C)" if col == "temp_avg" else "강수량 (mm)"
    fig.update_layout(
        xaxis_title="연도",
        yaxis_title=yaxis_label,
        height=380,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig, use_container_width=True)


def _render_mk_test(df: pd.DataFrame) -> None:
    """추세 통계 검정 탭 렌더링 (Mann-Kendall / Sen's Slope / Pettitt)."""
    st.subheader("추세 통계 검정")

    if not _check_col(df, "year"):
        return

    var_option = st.radio(
        "분석 변수",
        options=["기온", "강수량"],
        horizontal=True,
        key="mk_var_option",
    )

    if var_option == "기온":
        col, agg, unit = "temp_avg", "mean", "°C"
        if not _check_col(df, col):
            return
    else:
        col, agg, unit = "precipitation", "sum", "mm"
        if not _check_col(df, col):
            return

    # n>100 경고
    n_years = df["year"].nunique()
    if n_years > _MK_LARGE_N_THRESHOLD:
        st.warning(
            f"연도 수({n_years}개)가 {_MK_LARGE_N_THRESHOLD}개를 초과합니다. "
            "Mann-Kendall / Pettitt 검정은 O(n²) 연산으로 속도가 느릴 수 있습니다."
        )

    with st.spinner("통계 검정 계산 중..."):
        result_df = _build_mk_results(df, col, agg)

    if result_df.empty:
        st.warning("검정 결과를 계산할 수 없습니다. 데이터를 확인하세요.")
        return

    st.dataframe(result_df, use_container_width=True, hide_index=True)
    _csv = result_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", _csv, "mann_kendall.csv", "text/csv", key="climate_mk_csv")

    # 차트
    _mk_chart(df, col, agg, result_df)

    # 자동 해석
    st.markdown("**자동 해석**")
    for _, row in result_df.iterrows():
        st.info(_mk_interpretation(row.to_dict(), unit))


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
    # 12월까지 스크롤 없이 한 번에 보이도록 충분한 높이 설정
    row_count = len(monthly)
    table_height = min(max(row_count * 38 + 60, 520), 900)
    st.dataframe(monthly, use_container_width=True, hide_index=True, height=table_height)
    _csv = monthly.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", _csv, "climate_monthly.csv", "text/csv", key="climate_monthly_csv")


def render(df: pd.DataFrame) -> None:
    st.header("기후변화 분석")

    if df is None or df.empty:
        st.warning("표시할 데이터가 없습니다.")
        return

    df = _station_select(df)
    if df.empty:
        return

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "연평균 기온 추이",
        "연강수량 추이",
        "전후반기 비교",
        "추세 통계 검정",
        "기후통계 참조",
    ])

    with tab1:
        _render_temp_trend(df)

    with tab2:
        _render_precip_trend(df)

    with tab3:
        _render_period_comparison(df)

    with tab4:
        _render_mk_test(df)

    with tab5:
        _render_climate_reference(df)
