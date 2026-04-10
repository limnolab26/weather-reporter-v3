# analysis_precip.py — 강수량 분석 모듈
# render(df) 함수 하나를 export합니다.

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
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


def _get_stations(df: pd.DataFrame) -> list:
    if "station_name" in df.columns:
        return sorted(df["station_name"].unique().tolist())
    return []


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

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "연강수량 추이",
        "강우강도 분석",
        "연속 무강수일",
        "월별 강수 히트맵",
        "최대일강수 추이",
        "누적강수량 분석",
        "여름 강수 집중도",
        "강우일수 분석",
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


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 추가 분석 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _tab_cumulative(df: pd.DataFrame) -> None:
    """누적강수량 분석 — 월별 × 연도별 매트릭스"""
    st.markdown("### 📊 누적강수량 분석 (연도별 월별 강수량)")

    if not _has_col(df, "precipitation", "year", "month", "station_name"):
        st.warning("강수량 데이터가 없습니다.")
        return

    stations = df["station_name"].unique().tolist()
    stn = st.selectbox("관측소", stations, key="precip_cumul_stn")
    sdf = df[df["station_name"] == stn].copy()

    monthly = sdf.groupby(["year", "month"])["precipitation"].sum().reset_index()
    pivot = monthly.pivot(index="month", columns="year", values="precipitation").round(1)
    pivot.index.name = "월"

    # 행합계(연합계), 열합계(월합계) 추가
    pivot.loc["합계"] = pivot.sum()
    pivot["월평균"] = pivot.iloc[:-1].mean(axis=1).round(1)
    pivot.loc["합계", "월평균"] = float("nan")

    st.markdown("#### 연도별 월별 강수량 (mm)")
    st.dataframe(pivot.style.format("{:.1f}", na_rep="-"), use_container_width=True, height=480)

    csv = pivot.to_csv(encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", csv, f"{stn}_누적강수량.csv", "text/csv", key="precip_cumul_csv")

    # 히트맵 (합계행 제외)
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
    """여름 강수 집중도 분석 (6~8월 / 연강수량)"""
    st.markdown("### ☔ 여름철 강수 집중도 분석")

    if not _has_col(df, "precipitation", "year", "month", "station_name"):
        st.warning("강수량 데이터가 없습니다.")
        return

    annual = df.groupby(["station_name", "year"])["precipitation"].sum().reset_index()
    annual.columns = ["관측소", "연도", "연강수량"]

    summer = (
        df[df["month"].isin([6, 7, 8])]
        .groupby(["station_name", "year"])["precipitation"]
        .sum()
        .reset_index()
    )
    summer.columns = ["관측소", "연도", "여름강수량"]

    merged = pd.merge(annual, summer, on=["관측소", "연도"])
    merged["여름집중도(%)"] = (merged["여름강수량"] / merged["연강수량"] * 100).round(1)

    # 관측소별 평균
    summary = merged.groupby("관측소").agg(
        연평균강수량=("연강수량", "mean"),
        여름평균강수량=("여름강수량", "mean"),
        평균집중도=("여름집중도(%)", "mean"),
        최대집중도=("여름집중도(%)", "max"),
        최소집중도=("여름집중도(%)", "min"),
    ).round(1)
    st.markdown("#### 관측소별 여름(6~8월) 강수 집중도 요약")
    st.dataframe(summary, use_container_width=True)

    # 연도별 집중도 추이
    fig = px.line(
        merged, x="연도", y="여름집중도(%)",
        color="관측소", markers=True,
        labels={"연도": "연도", "여름집중도(%)": "집중도 (%)"},
        title="연도별 여름(6~8월) 강수 집중도"
    )
    fig.add_hline(
        y=merged["여름집중도(%)"].mean(),
        line_dash="dash", line_color="gray",
        annotation_text=f"전체평균 {merged['여름집중도(%)'].mean():.1f}%"
    )
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

    # 연강수량 vs 여름강수량 산점도
    fig2 = px.scatter(
        merged, x="연강수량", y="여름강수량",
        color="관측소", trendline="ols",
        labels={"연강수량": "연강수량 (mm)", "여름강수량": "여름강수량 (mm)"},
        title="연강수량 vs 여름강수량"
    )
    fig2.update_layout(height=380)
    st.plotly_chart(fig2, use_container_width=True)


def _tab_rain_days(df: pd.DataFrame) -> None:
    """연속강우·강우일수·무강우일수 분석"""
    st.markdown("### 🌂 강우일수 및 연속강우 분석")

    if not _has_col(df, "precipitation", "year", "station_name"):
        st.warning("강수량 데이터가 없습니다.")
        return

    records = []
    for stn, sdf in df.groupby("station_name"):
        for year, ydf in sdf.groupby("year"):
            ydf = ydf.sort_values("date")
            p = ydf["precipitation"].fillna(0)

            rain_days = int((p > 0).sum())
            dry_days = int((p == 0).sum())
            total_days = len(p)

            # 강우 강도별
            d_lt3 = int(((p > 0) & (p < 3)).sum())
            d_3_10 = int(((p >= 3) & (p < 10)).sum())
            d_10_30 = int(((p >= 10) & (p < 30)).sum())
            d_30_80 = int(((p >= 30) & (p < 80)).sum())
            d_ge80 = int((p >= 80).sum())

            # 최장 연속 강우일
            is_rain = (p > 0).astype(int)
            groups = (is_rain != is_rain.shift()).cumsum()
            consec_rain = is_rain.groupby(groups).sum()
            max_consec_rain = int(consec_rain[consec_rain > 0].max()) if rain_days > 0 else 0

            # 최장 연속 무강우일
            is_dry = (p == 0).astype(int)
            groups2 = (is_dry != is_dry.shift()).cumsum()
            consec_dry = is_dry.groupby(groups2).sum()
            max_consec_dry = int(consec_dry[consec_dry > 0].max()) if dry_days > 0 else 0

            records.append({
                "관측소": stn, "연도": int(year),
                "강우일수": rain_days, "무강우일": dry_days, "총일수": total_days,
                "<3mm": d_lt3, "3~10mm": d_3_10, "10~30mm": d_10_30,
                "30~80mm": d_30_80, "≥80mm": d_ge80,
                "최장연속강우(일)": max_consec_rain,
                "최장연속무강우(일)": max_consec_dry,
            })

    tbl = pd.DataFrame(records)
    st.dataframe(tbl, use_container_width=True, height=420)

    csv = tbl.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", csv, "강우일수분석.csv", "text/csv", key="raindays_csv")

    col1, col2 = st.columns(2)
    with col1:
        stations = tbl["관측소"].unique().tolist()
        stn2 = st.selectbox("관측소", stations, key="raindays_stn")
    with col2:
        metric = st.selectbox(
            "표시 항목",
            ["강우일수", "무강우일", "최장연속강우(일)", "최장연속무강우(일)"],
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
