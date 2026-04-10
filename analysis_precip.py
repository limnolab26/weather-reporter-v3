# analysis_precip.py — 강수량 분석 모듈
# render(df) 함수 하나를 export합니다.

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from scipy import stats


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


def _get_stations(df: pd.DataFrame) -> list:
    if "station_name" in df.columns:
        return sorted(df["station_name"].unique().tolist())
    return []


# ── SPI 계산 헬퍼 ───────────────────────────────────────────────────────────

def _calc_spi(monthly_precip: pd.Series, scale: int) -> pd.Series:
    """SPI (Standardized Precipitation Index) 계산.

    monthly_precip: 월별 강수량 Series
    scale: 집계 기간 (1, 3, 6, 12개월)
    Returns: SPI 값 Series (같은 인덱스)
    """
    rolled = monthly_precip.rolling(scale, min_periods=scale).sum()
    spi = pd.Series(np.nan, index=rolled.index)
    valid = rolled.dropna()
    if len(valid) < 10:
        return spi

    non_zero = valid[valid > 0]
    p_zero = (valid == 0).sum() / len(valid)
    if len(non_zero) < 5:
        return spi

    try:
        shape, loc, scale_param = stats.gamma.fit(non_zero, floc=0)
        for idx in valid.index:
            v = valid[idx]
            if v == 0:
                p = p_zero / 2
            else:
                p = p_zero + (1 - p_zero) * stats.gamma.cdf(
                    v, shape, loc=0, scale=scale_param
                )
            p = np.clip(p, 0.0013, 0.9987)
            spi[idx] = stats.norm.ppf(p)
    except Exception:
        pass

    return spi.round(2)


def _monthly_precip_by_station(df: pd.DataFrame) -> dict[str, pd.Series]:
    """관측소별 월별 강수량 Series dict 반환."""
    result = {}
    if not _has_col(df, "year", "month", "precipitation", "station_name"):
        return result
    for stn, sdf in df.groupby("station_name"):
        monthly = (
            sdf.groupby(["year", "month"])["precipitation"]
            .sum()
            .reset_index()
        )
        monthly["date"] = pd.to_datetime(
            monthly["year"].astype(str) + "-" + monthly["month"].astype(str).str.zfill(2)
        )
        monthly = monthly.sort_values("date").set_index("date")["precipitation"]
        result[stn] = monthly
    return result


def _spi_drought_stats(spi_series: pd.Series, stn: str, scale: int) -> dict:
    """SPI ≤ -1.0, ≤ -2.0 월 수와 가뭄 발생률 계산."""
    valid = spi_series.dropna()
    total = len(valid)
    mild = int((valid <= -1.0).sum())
    severe = int((valid <= -2.0).sum())
    rate = round(mild / total * 100, 1) if total > 0 else 0.0
    return {
        "관측소": stn,
        "스케일": f"SPI-{scale}",
        "유효 월수": total,
        "SPI≤-1.0 (월)": mild,
        "SPI≤-2.0 (월)": severe,
        "가뭄 발생률(%)": rate,
    }


def _draw_spi_chart(spi_series: pd.Series, stn: str, scale: int) -> go.Figure:
    """SPI 시계열 Plotly 차트 생성."""
    valid = spi_series.dropna().reset_index()
    valid.columns = ["date", "spi"]

    colors = np.where(valid["spi"] <= -1.0, "#e74c3c",
              np.where(valid["spi"] >= 1.0, "#2980b9", "#95a5a6"))

    fig = go.Figure()

    # 배경 색상 밴드
    band_defs = [
        (2.0, 4.0,   "rgba(41,128,185,0.15)",  "심한 습윤"),
        (1.5, 2.0,   "rgba(93,173,226,0.15)",   "보통 습윤"),
        (-1.5, 1.5,  "rgba(200,200,200,0.08)",  "정상"),
        (-2.0, -1.5, "rgba(230,126,34,0.15)",   "보통 가뭄"),
        (-4.0, -2.0, "rgba(231,76,60,0.15)",    "심한 가뭄"),
    ]
    for y0, y1, color, label in band_defs:
        fig.add_hrect(y0=y0, y1=y1, fillcolor=color, line_width=0, annotation_text=label,
                      annotation_position="right", annotation_font_size=10)

    fig.add_trace(go.Bar(
        x=valid["date"], y=valid["spi"],
        marker_color=colors.tolist(),
        name=f"SPI-{scale}",
    ))
    fig.add_hline(y=0, line_color="black", line_width=1)

    fig.update_layout(
        title=f"{stn} — SPI-{scale} 시계열",
        xaxis_title="연월",
        yaxis_title="SPI 값",
        height=380,
        yaxis=dict(range=[-4, 4]),
    )
    return fig


def _tab_spi(df: pd.DataFrame) -> None:
    """SPI 가뭄지수 분석 서브탭."""
    st.markdown("### 가뭄지수 (SPI) 분석")

    if not _has_col(df, "year", "month", "precipitation", "station_name"):
        st.warning("year, month, precipitation, station_name 컬럼이 필요합니다.")
        return

    stations = _get_stations(df)
    selected = st.multiselect(
        "관측소 선택",
        options=stations,
        default=stations[:1] if stations else [],
        key="spi_stations",
    )
    if not selected:
        st.info("관측소를 1개 이상 선택하세요.")
        return

    scale_options = {
        "SPI-1 (1개월)": 1,
        "SPI-3 (3개월)": 3,
        "SPI-6 (6개월)": 6,
        "SPI-12 (12개월)": 12,
    }
    chosen_labels = st.multiselect(
        "집계 기간 선택",
        options=list(scale_options.keys()),
        default=["SPI-1 (1개월)", "SPI-3 (3개월)"],
        key="spi_scales",
    )
    if not chosen_labels:
        st.info("집계 기간을 1개 이상 선택하세요.")
        return

    monthly_by_stn = _monthly_precip_by_station(df)
    stats_rows = []

    for stn in selected:
        monthly = monthly_by_stn.get(stn)
        if monthly is None or monthly.empty:
            st.warning(f"{stn}: 월별 강수 데이터 없음")
            continue

        # 결측월 비율 경고
        expected_months = (monthly.index[-1].year - monthly.index[0].year) * 12 + \
                          (monthly.index[-1].month - monthly.index[0].month) + 1
        missing_rate = max(0, 1 - len(monthly) / expected_months) if expected_months > 0 else 0
        if missing_rate > 0.2:
            st.warning(f"{stn}: 결측월 비율 {missing_rate:.0%} — SPI 정확도가 낮을 수 있습니다.")

        for label in chosen_labels:
            scale = scale_options[label]
            spi = _calc_spi(monthly, scale)

            if spi.dropna().empty:
                st.warning(f"{stn} SPI-{scale}: 유효 데이터 부족 (최소 10개월 필요)")
                continue

            fig = _draw_spi_chart(spi, stn, scale)
            st.plotly_chart(fig, use_container_width=True)
            stats_rows.append(_spi_drought_stats(spi, stn, scale))

    # 가뭄 통계 요약 테이블
    if stats_rows:
        st.markdown("#### 가뭄 통계 요약")
        stats_df = pd.DataFrame(stats_rows)
        st.dataframe(stats_df, use_container_width=True)

    # SPI 분류 기준 안내
    with st.expander("SPI 분류 기준"):
        st.table(pd.DataFrame({
            "SPI 값": ["≥ 2.0", "1.5 ~ 2.0", "1.0 ~ 1.5", "-1.0 ~ 1.0",
                       "-1.5 ~ -1.0", "-2.0 ~ -1.5", "≤ -2.0"],
            "분류":   ["심한 습윤", "보통 습윤", "약한 습윤", "정상",
                       "약한 가뭄", "보통 가뭄", "심한 가뭄"],
        }))


# ── Sub-tab 1: 연강수량 추이 ────────────────────────────────────────────────

def _tab_annual_precip(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "precipitation", "station_name"):
        st.warning("year, precipitation, station_name 컬럼이 필요합니다.")
        return

    annual = (
        df.groupby(["station_name", "year"])["precipitation"]
        .sum()
        .reset_index()
        .rename(columns={"precipitation": "연강수량(mm)"})
    )

    if annual.empty:
        st.warning("연강수량 데이터가 없습니다.")
        return

    overall_mean = annual["연강수량(mm)"].mean()

    fig = px.bar(
        annual,
        x="year",
        y="연강수량(mm)",
        color="station_name",
        barmode="group",
        title="연강수량 추이",
        height=400,
        labels={"year": "연도", "연강수량(mm)": "연강수량 (mm)", "station_name": "관측소"},
    )
    fig.add_hline(
        y=overall_mean,
        line_dash="dash",
        line_color="red",
        annotation_text=f"전체 평균 {overall_mean:.0f}mm",
        annotation_position="top left",
    )
    st.plotly_chart(fig, use_container_width=True)


# ── Sub-tab 2: 강우강도 분석 ────────────────────────────────────────────────

_INTENSITY_BINS = [
    ("무강우(0mm)", lambda s: s == 0),
    ("<3mm",        lambda s: (s > 0) & (s < 3)),
    ("3~10mm",      lambda s: (s >= 3) & (s < 10)),
    ("10~30mm",     lambda s: (s >= 10) & (s < 30)),
    ("30~80mm",     lambda s: (s >= 30) & (s < 80)),
    ("≥80mm",       lambda s: s >= 80),
]


def _tab_intensity(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "precipitation", "station_name"):
        st.warning("year, precipitation, station_name 컬럼이 필요합니다.")
        return

    stations = _get_stations(df)
    selected_station = st.selectbox(
        "관측소 선택 (강우강도)",
        options=stations,
        key="intensity_station_select",
    )

    s_df = df[df["station_name"] == selected_station].copy()
    # NaN → 0으로 채워 카테고리 분류
    precip = s_df["precipitation"].fillna(0)

    rows = []
    for year, yg in s_df.groupby("year"):
        p = yg["precipitation"].fillna(0)
        row = {"연도": year}
        for label, cond in _INTENSITY_BINS:
            row[label] = int(cond(p).sum())
        rows.append(row)

    if not rows:
        st.warning("데이터가 없습니다.")
        return

    intensity_df = pd.DataFrame(rows)
    category_cols = [label for label, _ in _INTENSITY_BINS]

    fig = go.Figure()
    colors = px.colors.sequential.Blues[1:]  # 밝은 쪽부터
    for i, cat in enumerate(category_cols):
        color = colors[min(i, len(colors) - 1)]
        fig.add_trace(go.Bar(
            name=cat,
            x=intensity_df["연도"],
            y=intensity_df[cat],
            marker_color=color,
        ))

    fig.update_layout(
        barmode="stack",
        title=f"{selected_station} — 강우강도별 발생 일수",
        xaxis_title="연도",
        yaxis_title="일수 (일)",
        height=400,
        legend={"orientation": "h", "yanchor": "bottom", "y": -0.3},
    )
    st.plotly_chart(fig, use_container_width=True)

    # 요약 통계표
    st.caption("강우강도 카테고리별 요약 통계 (단위: 일)")
    summary = intensity_df[category_cols].describe().round(1)
    st.dataframe(summary, use_container_width=True)


# ── Sub-tab 3: 연속 무강수일 ────────────────────────────────────────────────

def _max_consecutive_dry(series: pd.Series) -> int:
    """0 또는 NaN이 연속으로 이어지는 최대 길이를 반환."""
    is_dry = (series.fillna(0) == 0).astype(int)
    max_run = 0
    current = 0
    for val in is_dry:
        if val:
            current += 1
            max_run = max(max_run, current)
        else:
            current = 0
    return max_run


def _dry_spells(series: pd.Series, dates: pd.Series) -> list:
    """연속 무강수 구간 목록 [(시작일, 종료일, 일수), ...] 반환."""
    is_dry = (series.fillna(0) == 0).values
    spells = []
    start_idx = None

    for i, dry in enumerate(is_dry):
        if dry:
            if start_idx is None:
                start_idx = i
        else:
            if start_idx is not None:
                length = i - start_idx
                spells.append({
                    "시작일": dates.iloc[start_idx],
                    "종료일": dates.iloc[i - 1],
                    "연속 무강수일": length,
                })
                start_idx = None

    # 끝까지 이어진 경우
    if start_idx is not None:
        length = len(is_dry) - start_idx
        spells.append({
            "시작일": dates.iloc[start_idx],
            "종료일": dates.iloc[-1],
            "연속 무강수일": length,
        })

    return spells


def _tab_dry_days(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "precipitation", "station_name"):
        st.warning("year, precipitation, station_name 컬럼이 필요합니다.")
        return

    stations = _get_stations(df)
    selected_station = st.selectbox(
        "관측소 선택 (무강수일)",
        options=stations,
        key="dry_station_select",
    )

    s_df = df[df["station_name"] == selected_station].sort_values("date") if "date" in df.columns \
        else df[df["station_name"] == selected_station]

    # 연간 최대 연속 무강수일
    annual_max = (
        s_df.groupby("year")
        .apply(lambda g: _max_consecutive_dry(g["precipitation"]))
        .reset_index()
        .rename(columns={0: "최대연속무강수일"})
    )

    fig = px.bar(
        annual_max,
        x="year",
        y="최대연속무강수일",
        title=f"{selected_station} — 연간 최대 연속 무강수일",
        height=400,
        labels={"year": "연도", "최대연속무강수일": "연속 무강수일 (일)"},
        color_discrete_sequence=["#4A90D9"],
    )
    st.plotly_chart(fig, use_container_width=True)

    # Top-10 긴 건조 구간
    if "date" in s_df.columns:
        all_spells = _dry_spells(s_df["precipitation"], s_df["date"].astype(str))
        if all_spells:
            spells_df = (
                pd.DataFrame(all_spells)
                .sort_values("연속 무강수일", ascending=False)
                .head(10)
                .reset_index(drop=True)
            )
            spells_df.index += 1
            st.caption("역대 Top-10 연속 무강수 구간")
            st.dataframe(spells_df, use_container_width=True)
    else:
        st.caption("date 컬럼이 없어 Top-10 구간 표시를 건너뜁니다.")


# ── Sub-tab 4: 월별 강수 히트맵 ─────────────────────────────────────────────

def _tab_monthly_heatmap(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "month", "precipitation"):
        st.warning("year, month, precipitation 컬럼이 필요합니다.")
        return

    stations = _get_stations(df)
    selected_station = st.selectbox(
        "관측소 선택 (히트맵)",
        options=stations if stations else ["전체"],
        key="precip_heatmap_station_select",
    )

    if stations and selected_station != "전체":
        plot_df = df[df["station_name"] == selected_station]
    else:
        plot_df = df

    pivot = (
        plot_df.groupby(["year", "month"])["precipitation"]
        .sum()
        .unstack(level="month")
    )

    if pivot.empty:
        st.warning("히트맵을 그릴 데이터가 없습니다.")
        return

    pivot.columns = [f"{m}월" for m in pivot.columns]

    fig = px.imshow(
        pivot,
        color_continuous_scale="Blues",
        title=f"{selected_station} — 연도×월 강수량 히트맵 (mm)",
        labels={"x": "월", "y": "연도", "color": "강수량(mm)"},
        aspect="auto",
        height=400,
    )
    fig.update_xaxes(side="bottom")
    st.plotly_chart(fig, use_container_width=True)


# ── Sub-tab 5: 최대일강수 추이 ──────────────────────────────────────────────

def _tab_max_daily_precip(df: pd.DataFrame) -> None:
    if not _has_col(df, "year", "precipitation", "station_name"):
        st.warning("year, precipitation, station_name 컬럼이 필요합니다.")
        return

    # max 계산 시 NaN 제외 (기본 동작)
    annual_max = (
        df.groupby(["station_name", "year"])["precipitation"]
        .max()
        .reset_index()
        .rename(columns={"precipitation": "최대일강수량(mm)"})
    )

    if annual_max.empty:
        st.warning("데이터가 없습니다.")
        return

    fig = px.bar(
        annual_max,
        x="year",
        y="최대일강수량(mm)",
        color="station_name",
        barmode="group",
        title="연간 최대일강수량 추이",
        height=400,
        labels={"year": "연도", "최대일강수량(mm)": "최대일강수량 (mm)", "station_name": "관측소"},
    )
    fig.add_hline(
        y=80,
        line_dash="dash",
        line_color="red",
        annotation_text="호우 기준 80mm",
        annotation_position="top left",
    )
    st.plotly_chart(fig, use_container_width=True)


# ── 공개 render 함수 ─────────────────────────────────────────────────────────

def render(df: pd.DataFrame) -> None:
    """강수량 분석 탭 렌더링. df는 필터링된 일자료 DataFrame."""
    st.subheader("강수량 분석")

    if df is None or df.empty:
        st.warning("데이터가 없습니다.")
        return

    if "precipitation" not in df.columns:
        st.warning("precipitation(일강수량) 컬럼이 없어 강수량 분석을 표시할 수 없습니다.")
        return

    # 관측소 필터
    if "station_name" in df.columns:
        all_stations = sorted(df["station_name"].unique().tolist())
        selected_stations = st.multiselect(
            "관측소 선택 (전체 탭 적용)",
            options=all_stations,
            default=all_stations,
            key="precip_station_multiselect",
        )
        filtered = _filter_stations(df, selected_stations)
    else:
        filtered = df

    if filtered.empty:
        st.warning("선택된 관측소에 데이터가 없습니다.")
        return

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "연강수량 추이",
        "강우강도 분석",
        "연속 무강수일",
        "월별 강수 히트맵",
        "최대일강수 추이",
        "누적강수량 분석",
        "여름 강수 집중도",
        "강우일수 분석",
        "SPI 가뭄지수",
    ])

    with tab1:
        _tab_annual_precip(filtered)

    with tab2:
        _tab_intensity(filtered)

    with tab3:
        _tab_dry_days(filtered)

    with tab4:
        _tab_monthly_heatmap(filtered)

    with tab5:
        _tab_max_daily_precip(filtered)

    with tab6:
        _tab_cumulative(filtered)

    with tab7:
        _tab_summer_concentration(filtered)

    with tab8:
        _tab_rain_days(filtered)

    with tab9:
        _tab_spi(filtered)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 추가 분석 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _calc_mv_ratio(df: pd.DataFrame, stn: str) -> pd.DataFrame:
    """
    MV 커브 계산 — 연도별 월별 누적강수량 비율 (0~1).
    Returns DataFrame: index=year, columns=month(1~12)
    """
    sdf = df[df["station_name"] == stn].copy()
    monthly = sdf.groupby(["year", "month"])["precipitation"].sum().unstack("month")  # year × month
    monthly = monthly.reindex(columns=range(1, 13))  # 1~12월 보장
    annual = monthly.sum(axis=1)
    cumulative = monthly.cumsum(axis=1)
    ratio = cumulative.div(annual, axis=0).round(4)
    return ratio


def _plot_mv_curves(ratio: pd.DataFrame, stn: str) -> go.Figure:
    """
    MV 커브 Plotly 차트.
    ratio: index=year, columns=month(1~12)
    각 연도 = 회색 가는 선, 전체·10년·30년 평균 = 굵은 색 선
    """
    months = list(range(1, 13))
    month_labels = [f"{m}월" for m in months]
    years = sorted(ratio.index.tolist())
    n = len(years)
    recent_10 = [y for y in years if y >= years[-1] - 9]
    recent_30 = [y for y in years if y >= years[-1] - 29]

    fig = go.Figure()

    # 개별 연도 선 (회색, 반투명)
    for yr in years:
        row = ratio.loc[yr]
        vals = [row.get(m, float("nan")) for m in months]
        fig.add_trace(go.Scatter(
            x=month_labels, y=vals,
            mode="lines",
            line=dict(color="lightgray", width=0.8),
            opacity=0.6,
            name=str(yr),
            legendgroup="years",
            showlegend=False,
            hovertemplate=f"{yr}년: %{{y:.3f}}<extra></extra>",
        ))

    # 전체 평균
    mean_all = ratio.mean()
    fig.add_trace(go.Scatter(
        x=month_labels, y=[mean_all.get(m, float("nan")) for m in months],
        mode="lines+markers",
        line=dict(color="#1F4E79", width=2.5),
        marker=dict(size=6),
        name=f"전체 평균 ({years[0]}~{years[-1]})",
        hovertemplate="전체평균: %{y:.3f}<extra></extra>",
    ))

    # 10년 평균
    if len(recent_10) >= 3:
        mean_10 = ratio.loc[recent_10].mean()
        fig.add_trace(go.Scatter(
            x=month_labels, y=[mean_10.get(m, float("nan")) for m in months],
            mode="lines+markers",
            line=dict(color="#E74C3C", width=2.5, dash="dash"),
            marker=dict(size=6),
            name=f"최근 10년 평균 ({recent_10[0]}~{recent_10[-1]})",
            hovertemplate="최근10년: %{y:.3f}<extra></extra>",
        ))

    # 30년 평균 (평년)
    if len(recent_30) >= 10 and len(recent_30) != n:
        mean_30 = ratio.loc[recent_30].mean()
        fig.add_trace(go.Scatter(
            x=month_labels, y=[mean_30.get(m, float("nan")) for m in months],
            mode="lines+markers",
            line=dict(color="#27AE60", width=2.5, dash="dot"),
            marker=dict(size=6),
            name=f"최근 30년 평균 ({recent_30[0]}~{recent_30[-1]})",
            hovertemplate="최근30년: %{y:.3f}<extra></extra>",
        ))

    # 표준편차 밴드 (전체 평균 ±1σ)
    std_all = ratio.std()
    upper = (mean_all + std_all).clip(upper=1.0)
    lower = (mean_all - std_all).clip(lower=0.0)
    u_vals = [upper.get(m, float("nan")) for m in months]
    l_vals = [lower.get(m, float("nan")) for m in months]
    fig.add_trace(go.Scatter(
        x=month_labels + month_labels[::-1],
        y=u_vals + l_vals[::-1],
        fill="toself",
        fillcolor="rgba(31,78,121,0.08)",
        line=dict(color="rgba(0,0,0,0)"),
        name="±1σ 범위",
        hoverinfo="skip",
        showlegend=True,
    ))

    fig.update_layout(
        title=dict(text=f"{stn} MV 커브 (누적강수량 비율)", font=dict(size=16)),
        xaxis=dict(title="월", tickmode="array", tickvals=month_labels, ticktext=month_labels),
        yaxis=dict(title="누적강수량 / 연강수량", range=[-0.02, 1.05],
                   tickformat=".1%"),
        height=460,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
    )
    return fig


def _tab_cumulative(df: pd.DataFrame) -> None:
    """누적강수량 분석 — 월별 × 연도별 매트릭스 + MV 커브"""
    st.markdown("### 📊 누적강수량 분석")

    if not _has_col(df, "precipitation", "year", "month", "station_name"):
        st.warning("강수량 데이터가 없습니다.")
        return

    stations = df["station_name"].unique().tolist()
    stn = st.selectbox("관측소", stations, key="precip_cumul_stn")
    sdf = df[df["station_name"] == stn].copy()

    # ── MV 커브 ─────────────────────────────────────────────
    st.markdown("#### 📈 MV 커브 (누적강수 비율 곡선)")
    st.caption(
        "x축: 월(1~12), y축: 해당 월까지의 누적강수량 ÷ 연강수량. "
        "회색 선 = 개별 연도, 색 선 = 기간별 평균."
    )

    ratio = _calc_mv_ratio(df, stn)
    if not ratio.empty:
        fig_mv = _plot_mv_curves(ratio, stn)
        st.plotly_chart(fig_mv, use_container_width=True)

        # MV 커브 비율 테이블 (expander)
        with st.expander("MV 커브 수치 데이터 (연도별 월별 누적 비율)"):
            display_ratio = ratio.copy()
            display_ratio.index.name = "연도"
            display_ratio.columns = [f"{m}월" for m in display_ratio.columns]
            # 전체 평균·최근10년·최근30년 행 추가
            years = sorted(ratio.index.tolist())
            r10 = [y for y in years if y >= years[-1] - 9]
            r30 = [y for y in years if y >= years[-1] - 29]
            avg_row = ratio.mean().rename("전체평균")
            display_ratio.loc["전체평균"] = avg_row.values
            if len(r10) >= 3:
                display_ratio.loc[f"최근10년평균"] = ratio.loc[r10].mean().values
            if len(r30) >= 10:
                display_ratio.loc[f"최근30년평균"] = ratio.loc[r30].mean().values
            display_ratio.columns = [f"{m}월" for m in range(1, 13)]
            st.dataframe(
                display_ratio.style.format("{:.3f}", na_rep="-"),
                use_container_width=True,
                height=min(420, len(display_ratio) * 35 + 60),
            )
            mv_csv = display_ratio.to_csv(encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "⬇️ MV 커브 CSV", mv_csv, f"{stn}_MV커브.csv", "text/csv",
                key="mv_csv"
            )
    else:
        st.warning("MV 커브 계산에 필요한 데이터가 부족합니다.")

    st.markdown("---")

    # ── 기존: 월별 강수량 매트릭스 & 히트맵 ──────────────────
    monthly = sdf.groupby(["year", "month"])["precipitation"].sum().reset_index()
    pivot = monthly.pivot(index="month", columns="year", values="precipitation").round(1)
    pivot.index = [f"{m}월" for m in pivot.index]
    pivot.index.name = "월"
    pivot.columns = [str(c) for c in pivot.columns]
    pivot.columns.name = "연도"
    pivot.loc["합계"] = pivot.sum()
    pivot["월평균"] = pivot.iloc[:-1].mean(axis=1).round(1)
    pivot.loc["합계", "월평균"] = float("nan")

    st.markdown("#### 연도별 월별 강수량 (mm)")
    st.dataframe(pivot.style.format("{:.1f}", na_rep="-"), use_container_width=True, height=480)

    csv = pivot.to_csv(encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", csv, f"{stn}_누적강수량.csv", "text/csv", key="precip_cumul_csv")

    # 히트맵
    hm_data = monthly.pivot(index="year", columns="month", values="precipitation").fillna(0)
    hm_data.columns = [f"{m}월" for m in hm_data.columns]
    fig = px.imshow(
        hm_data,
        labels=dict(x="월", y="연도", color="강수량(mm)"),
        color_continuous_scale="Blues",
        aspect="auto",
        title=f"{stn} 연도별 월별 강수량 히트맵"
    )
    fig.update_layout(height=max(350, len(hm_data) * 15 + 120))
    st.plotly_chart(fig, use_container_width=True)


def _tab_summer_concentration(df: pd.DataFrame) -> None:
    """여름 강수 집중도 분석 (7+8, 6+7+8, 7+8+9 세 기간)"""
    st.markdown("### ☔ 여름철 강수 집중도 분석")

    if not _has_col(df, "precipitation", "year", "month", "station_name"):
        st.warning("강수량 데이터가 없습니다.")
        return

    annual = df.groupby(["station_name", "year"])["precipitation"].sum().reset_index()
    annual.columns = ["관측소", "연도", "연강수량"]

    # 세 가지 여름 기간 정의
    periods = {
        "7+8월 (성하기)":    [7, 8],
        "6+7+8월 (초여름~성하)": [6, 7, 8],
        "7+8+9월 (한여름~초가을)": [7, 8, 9],
    }

    all_merged = []
    for period_name, months in periods.items():
        sub = (
            df[df["month"].isin(months)]
            .groupby(["station_name", "year"])["precipitation"]
            .sum()
            .reset_index()
        )
        sub.columns = ["관측소", "연도", "기간강수량"]
        merged = pd.merge(annual, sub, on=["관측소", "연도"])
        merged["기간"] = period_name
        merged["집중도(%)"] = (merged["기간강수량"] / merged["연강수량"] * 100).round(1)
        all_merged.append(merged)

    all_df = pd.concat(all_merged, ignore_index=True)

    # 기간 선택
    selected_period = st.selectbox(
        "분석 기간",
        list(periods.keys()),
        key="summer_conc_period"
    )
    show_df = all_df[all_df["기간"] == selected_period]

    # 요약 통계
    st.markdown(f"#### 관측소별 {selected_period} 강수 집중도 요약")
    summary = show_df.groupby("관측소").agg(
        연평균강수량=("연강수량", "mean"),
        기간평균강수량=("기간강수량", "mean"),
        평균집중도=("집중도(%)", "mean"),
        최대집중도=("집중도(%)", "max"),
        최소집중도=("집중도(%)", "min"),
    ).round(1)
    st.dataframe(summary, use_container_width=True)

    # 세 기간 동시 비교표
    st.markdown("#### 기간별 연평균 집중도 비교")
    compare = all_df.groupby(["관측소", "기간"])["집중도(%)"].mean().round(1).unstack("기간")
    st.dataframe(compare, use_container_width=True)

    # 연도별 집중도 추이
    fig = px.line(
        show_df, x="연도", y="집중도(%)",
        color="관측소", markers=True,
        labels={"연도": "연도", "집중도(%)": "집중도 (%)"},
        title=f"연도별 {selected_period} 강수 집중도"
    )
    fig.add_hline(
        y=show_df["집중도(%)"].mean(),
        line_dash="dash", line_color="gray",
        annotation_text=f"전체평균 {show_df['집중도(%)'].mean():.1f}%"
    )
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

    # 연강수량 vs 기간강수량 산점도 (numpy 추세선, statsmodels 불필요)
    fig2 = px.scatter(
        show_df, x="연강수량", y="기간강수량",
        color="관측소",
        labels={"연강수량": "연강수량 (mm)", "기간강수량": f"{selected_period} 강수량 (mm)"},
        title=f"연강수량 vs {selected_period} 강수량"
    )
    # 관측소별 numpy polyfit 추세선 추가
    colors_list = px.colors.qualitative.Set2
    for i, (stn_name, grp) in enumerate(show_df.groupby("관측소")):
        x_arr = grp["연강수량"].values.astype(float)
        y_arr = grp["기간강수량"].values.astype(float)
        mask = ~(np.isnan(x_arr) | np.isnan(y_arr))
        if mask.sum() >= 2:
            coeffs = np.polyfit(x_arr[mask], y_arr[mask], 1)
            x_line = np.linspace(x_arr[mask].min(), x_arr[mask].max(), 50)
            y_line = np.polyval(coeffs, x_line)
            fig2.add_trace(go.Scatter(
                x=x_line, y=y_line, mode="lines",
                name=f"{stn_name} 추세",
                line=dict(color=colors_list[i % len(colors_list)], dash="dash", width=1.5),
                showlegend=False,
            ))
    fig2.update_layout(height=380)
    st.plotly_chart(fig2, use_container_width=True)

    csv = show_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", csv, "여름강수집중도.csv", "text/csv", key="summer_conc_csv")


def _tab_rain_days(df: pd.DataFrame) -> None:
    """연속강우·강우일수·무강우일수 분석"""
    st.markdown("### 🌂 강우일수 및 연속강우 분석")

    if not _has_col(df, "precipitation", "year", "station_name"):
        st.warning("강수량 데이터가 없습니다.")
        return

    records = []
    for stn, sdf in df.groupby("station_name"):
        for year, ydf in sdf.groupby("year"):
            if "date" in ydf.columns:
                ydf = ydf.sort_values("date")
            p = ydf["precipitation"].fillna(0)

            rain_days = int((p > 0).sum())
            dry_days = int((p == 0).sum())
            total_days = len(p)

            # 강우 강도별
            d_lt3   = int(((p > 0) & (p < 3)).sum())
            d_3_10  = int(((p >= 3)  & (p < 10)).sum())
            d_10_30 = int(((p >= 10) & (p < 30)).sum())
            d_30_80 = int(((p >= 30) & (p < 80)).sum())
            d_ge80  = int((p >= 80).sum())

            # 최장 연속 강우일
            is_rain = (p > 0).astype(int).reset_index(drop=True)
            groups = (is_rain != is_rain.shift()).cumsum()
            consec_rain = is_rain.groupby(groups).sum()
            max_consec_rain = int(consec_rain[consec_rain > 0].max()) if rain_days > 0 else 0

            # 최장 연속 무강우일
            is_dry = (p == 0).astype(int).reset_index(drop=True)
            groups2 = (is_dry != is_dry.shift()).cumsum()
            consec_dry = is_dry.groupby(groups2).sum()
            max_consec_dry = int(consec_dry[consec_dry > 0].max()) if dry_days > 0 else 0

            records.append({
                "관측소": stn, "연도": int(year),
                "강우일수": rain_days, "무강우일수": dry_days, "총일수": total_days,
                "<3mm": d_lt3, "3~10mm": d_3_10, "10~30mm": d_10_30,
                "30~80mm": d_30_80, "≥80mm": d_ge80,
                "최장연속강우(일)": max_consec_rain,
                "최장연속무강우(일)": max_consec_dry,
            })

    tbl = pd.DataFrame(records)

    st.markdown("#### 연도별 강우일수 통계표")
    st.dataframe(tbl, use_container_width=True, height=440)

    csv = tbl.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", csv, "강우일수분석.csv", "text/csv", key="raindays_csv")

    # ── 다중 관측소 추이 (라인차트) ──
    st.markdown("#### 연도별 추이 (다중 관측소 비교)")
    trend_metric = st.selectbox(
        "추이 항목 선택",
        ["강우일수", "무강우일수", "최장연속강우(일)", "최장연속무강우(일)"],
        key="raindays_trend_metric"
    )
    fig_trend = px.line(
        tbl, x="연도", y=trend_metric,
        color="관측소", markers=True,
        labels={"연도": "연도", trend_metric: f"{trend_metric}"},
        title=f"연도별 {trend_metric} 추이"
    )
    # 관측소별 평균 기준선
    for stn_name, sdf2 in tbl.groupby("관측소"):
        avg = sdf2[trend_metric].mean()
        fig_trend.add_hline(
            y=avg, line_dash="dot", opacity=0.4,
            annotation_text=f"{stn_name} 평균 {avg:.1f}일",
            annotation_position="top right",
        )
    fig_trend.update_layout(height=420)
    st.plotly_chart(fig_trend, use_container_width=True)

    # ── 단일 관측소 상세 막대 ──
    st.markdown("#### 관측소별 상세 분석")
    col1, col2 = st.columns(2)
    with col1:
        stations = tbl["관측소"].unique().tolist()
        stn2 = st.selectbox("관측소", stations, key="raindays_stn")
    with col2:
        metric = st.selectbox(
            "표시 항목",
            ["강우일수", "무강우일수", "최장연속강우(일)", "최장연속무강우(일)"],
            key="raindays_metric"
        )

    plot_df = tbl[tbl["관측소"] == stn2]
    fig = px.bar(
        plot_df, x="연도", y=metric,
        title=f"{stn2} 연도별 {metric}",
        color_discrete_sequence=["#3498db"]
    )
    fig.update_layout(height=380)
    st.plotly_chart(fig, use_container_width=True)

    # 강우강도 스택 막대
    st.markdown("#### 강우강도 등급별 연도별 일수")
    intensity_cols = ["<3mm", "3~10mm", "10~30mm", "30~80mm", "≥80mm"]
    fig2 = go.Figure()
    colors = ["#aed6f1", "#3498db", "#1a5276", "#f39c12", "#e74c3c"]
    for col, color in zip(intensity_cols, colors):
        fig2.add_trace(go.Bar(
            x=plot_df["연도"], y=plot_df[col],
            name=col, marker_color=color
        ))
    fig2.update_layout(barmode="stack", height=380, xaxis_title="연도", yaxis_title="일수")
    st.plotly_chart(fig2, use_container_width=True)
