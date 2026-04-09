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

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "연강수량 추이",
        "강우강도 분석",
        "연속 무강수일",
        "월별 강수 히트맵",
        "최대일강수 추이",
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
