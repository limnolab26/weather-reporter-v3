# analysis_soiltemp.py — 지중온도 분석 모듈
# render(df) 함수 하나를 export합니다.

import io

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy import stats

from chart_utils import chart_download_btn

SOIL_DEPTH_COLS = [
    {"col": "soil_temp_surface", "label": "지면(0cm)",  "depth_m": 0.0},
    {"col": "soil_temp_5cm",     "label": "5cm",        "depth_m": 0.05},
    {"col": "soil_temp_10cm",    "label": "10cm",       "depth_m": 0.10},
    {"col": "soil_temp_20cm",    "label": "20cm",       "depth_m": 0.20},
    {"col": "soil_temp_30cm",    "label": "30cm",       "depth_m": 0.30},
    {"col": "soil_temp_50cm",    "label": "50cm",       "depth_m": 0.50},
    {"col": "soil_temp_100cm",   "label": "100cm",      "depth_m": 1.00},
    {"col": "soil_temp_150cm",   "label": "150cm",      "depth_m": 1.50},
    {"col": "soil_temp_300cm",   "label": "300cm",      "depth_m": 3.00},
    {"col": "soil_temp_500cm",   "label": "500cm",      "depth_m": 5.00},
]


# ──────────────────────────────────────────────────────────────
# 헬퍼
# ──────────────────────────────────────────────────────────────

def _has_col(df, *cols):
    return all(c in df.columns for c in cols)


def _filter_stations(df, stations):
    if not stations:
        return df
    return df[df["station_name"].isin(stations)]


def _csv_bytes(df):
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
    return buf.getvalue()


def _available_depths(df):
    """df에 실제 존재하는 지중온도 컬럼만 반환"""
    return [d for d in SOIL_DEPTH_COLS if d["col"] in df.columns]


# ──────────────────────────────────────────────────────────────
# 분야별 활용 안내
# ──────────────────────────────────────────────────────────────

def _section_guide():
    with st.expander("📖 이 데이터로 무엇을 알 수 있나요? — 분야별 활용 안내", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
#### 🌤️ 기상 전문가
대기와 지표면 사이의 **열 교환 메커니즘**을 이해하는 핵심 자료입니다.

- **위상지연(Phase Lag)**: 지표 온도 변화가 각 깊이에 도달하는 지연 시간을
  분석해 토양의 **열확산계수**를 추정합니다
- **열 침투 깊이**: 계절별로 온도 변화가 어느 깊이까지 전달되는지 추적합니다
- **기후변화 신호**: 지중온도의 장기 트렌드는 대기온도보다 노이즈가 적어
  기후변화 탐지에 유리합니다
- **서리·역전층 예측**: 야간 지면온도 급락은 복사냉각·안개의 전조 신호입니다
""")
            st.markdown("""
#### ⚡ 에너지 전문가
**지열 히트펌프 시스템** 설계 및 효율 평가의 핵심 입력값입니다.

- **항온층 분석**: 연변동폭이 2°C 이하로 안정되는 깊이가 히트펌프 최적 열원입니다
- **지열 포텐셜 등급**: A~E 등급으로 지열 시스템 경제성을 빠르게 판단합니다
- **냉난방 전환 적기**: 지중온도가 외기온보다 유리한 시기를 자동 계산합니다
- **COP 예측**: 지중온도와 난방 목표온도 차이로 계절별 효율을 추정합니다
""")
        with col2:
            st.markdown("""
#### 🌾 농업 전문가
**파종 시기 결정, 생육 예측, 동해 방지**에 직접 활용합니다.

- **파종 적기 캘린더**: 작물별 발아 기준온도 도달일을 연도별로 추적합니다
- **토양 적산온도(sGDD)**: 기온 GDD보다 실제 뿌리 생육에 더 정확한 예측 도구입니다
- **동해·냉해 진단**: 지중온도 0°C 이하 날짜와 추정 동결심도를 분석합니다

| 작물 | 발아 기준온도 | 분석 깊이 |
|------|------------|---------|
| 벼   | 13°C       | 10cm    |
| 옥수수 | 10°C     | 5cm     |
| 감자 | 7°C        | 10cm    |
| 콩   | 12°C       | 5cm     |
| 고추 | 15°C       | 5cm     |
""")
            st.markdown("""
#### 🌿 환경 전문가
**토양 생태계, 탄소 순환, 기후변화 모니터링**에 활용합니다.

- **토양 호흡 지수**: Q10 모델로 온도 기반 미생물 활동량을 간접 추정합니다
  *(온도 10°C 상승 시 미생물 활동 약 2배 증가)*
- **기후변화 침투 깊이**: 기온 상승 신호가 지하 몇 미터까지 도달했는지 추적합니다
- **생태활동 온도 구간**: 10~30°C 유지 일수로 토양 생물 서식 환경을 평가합니다
""")


# ──────────────────────────────────────────────────────────────
# 서브탭 1: 깊이-시간 히트맵
# ──────────────────────────────────────────────────────────────

def _tab_depth_heatmap(df):
    st.markdown("#### 🗺️ 깊이-시간 온도 히트맵")
    st.caption(
        "열이 지표에서 지하로 전달되는 패턴을 직관적으로 보여줍니다. "
        "붉은 색이 아래로 퍼져가는 시기가 여름 열파의 지중 침투 과정입니다."
    )

    depths = _available_depths(df)
    if not depths:
        st.info("지중온도 컬럼이 없습니다.")
        return

    agg_unit = st.radio("집계 단위", ["일별", "월별"], horizontal=True, key="soil_heatmap_agg")

    if agg_unit == "월별":
        plot_df = df.groupby(["year", "month"])[[d["col"] for d in depths]].mean().reset_index()
        x_col = "month_label"
        plot_df[x_col] = (
            plot_df["year"].astype(str) + "-" + plot_df["month"].astype(str).str.zfill(2)
        )
    else:
        plot_df = df.copy()
        x_col = "date"

    z_data = []
    y_labels = []
    for d in depths:
        if d["col"] in plot_df.columns:
            z_data.append(plot_df[d["col"]].values)
            y_labels.append(d["label"])

    if not z_data:
        st.warning("표시할 데이터가 없습니다.")
        return

    x_vals = plot_df[x_col].astype(str).values

    fig = go.Figure(data=go.Heatmap(
        z=z_data,
        x=x_vals,
        y=y_labels,
        colorscale="RdBu_r",
        colorbar={"title": "온도(°C)"},
        hovertemplate="날짜: %{x}<br>깊이: %{y}<br>온도: %{z:.1f}°C<extra></extra>",
    ))
    fig.update_layout(
        title="깊이-시간 온도 히트맵",
        xaxis_title="날짜",
        yaxis_title="깊이",
        height=420,
        yaxis={"autorange": "reversed"},
    )
    st.plotly_chart(fig, width='stretch')
    chart_download_btn(fig, key="soil_heatmap_chart", filename="soil_depth_time_heatmap")


# ──────────────────────────────────────────────────────────────
# 서브탭 2: 수직 온도 프로파일
# ──────────────────────────────────────────────────────────────

def _tab_vertical_profile(df):
    st.markdown("#### 📏 깊이별 수직 온도 프로파일")
    st.caption(
        "특정 날짜 또는 월의 지하 온도 분포입니다. "
        "겨울→봄→여름으로 이동할 때 온도 곡선이 오른쪽으로 이동하며 깊이에 따라 지연됩니다."
    )

    depths = _available_depths(df)
    if not depths:
        st.info("지중온도 컬럼이 없습니다.")
        return

    view_mode = st.radio("표시 기준", ["월 평균", "특정 날짜"], horizontal=True, key="soil_profile_mode")
    fig = go.Figure()

    if view_mode == "월 평균":
        months = sorted(df["month"].unique())
        selected_months = st.multiselect(
            "월 선택 (최대 6개)", months, default=months[:4], key="soil_profile_months",
        )
        month_labels = {
            1:"1월", 2:"2월", 3:"3월", 4:"4월", 5:"5월", 6:"6월",
            7:"7월", 8:"8월", 9:"9월", 10:"10월", 11:"11월", 12:"12월",
        }
        colors = px.colors.qualitative.Set2
        for i, m in enumerate(selected_months):
            mdf = df[df["month"] == m]
            temps   = [mdf[d["col"]].mean()  for d in depths if d["col"] in mdf.columns]
            d_labs  = [d["depth_m"]           for d in depths if d["col"] in mdf.columns]
            fig.add_trace(go.Scatter(
                x=temps, y=d_labs, mode="markers+lines",
                name=month_labels.get(m, f"{m}월"),
                line={"color": colors[i % len(colors)]},
                marker={"size": 7},
            ))
    else:
        all_dates = df["date"].dt.date.unique()
        selected_date = st.selectbox("날짜 선택", sorted(all_dates), key="soil_profile_date")
        ddf = df[df["date"].dt.date == selected_date]
        temps  = [ddf[d["col"]].mean() for d in depths if d["col"] in ddf.columns]
        d_labs = [d["depth_m"]          for d in depths if d["col"] in ddf.columns]
        fig.add_trace(go.Scatter(
            x=temps, y=d_labs, mode="markers+lines",
            name=str(selected_date),
            line={"color": "#E74C3C"}, marker={"size": 8},
        ))

    fig.update_layout(
        title="깊이별 수직 온도 프로파일",
        xaxis_title="온도 (°C)", yaxis_title="깊이 (m)",
        yaxis={"autorange": "reversed"},
        height=450,
        legend={"orientation": "h", "yanchor": "bottom", "y": -0.3},
    )
    st.plotly_chart(fig, width='stretch')
    chart_download_btn(fig, key="soil_profile_chart", filename="soil_vertical_profile")


# ──────────────────────────────────────────────────────────────
# 서브탭 3: 시계열 비교
# ──────────────────────────────────────────────────────────────

def _tab_timeseries(df):
    st.markdown("#### 📈 기온 vs 지중온도 시계열 비교")
    st.caption("기온과 여러 깊이의 지중온도를 겹쳐 표시합니다. 깊이가 깊을수록 온도 변화가 완만해지고 시간이 지연됩니다.")

    depths = _available_depths(df)
    depth_options = {d["label"]: d["col"] for d in depths}
    selected_labels = st.multiselect(
        "표시할 깊이 선택",
        options=list(depth_options.keys()),
        default=list(depth_options.keys())[:4],
        key="soil_ts_depth_select",
    )

    cols_to_plot = []
    rename_map = {}
    if "temp_avg" in df.columns:
        cols_to_plot.append("temp_avg")
        rename_map["temp_avg"] = "기온(평균)"
    for label in selected_labels:
        col = depth_options[label]
        if col in df.columns:
            cols_to_plot.append(col)
            rename_map[col] = label

    if not cols_to_plot:
        st.warning("표시할 컬럼이 없습니다.")
        return

    plot_df = df[["date"] + cols_to_plot].rename(columns=rename_map)
    melted = plot_df.melt(id_vars="date", var_name="항목", value_name="온도(°C)")

    fig = px.line(
        melted, x="date", y="온도(°C)", color="항목",
        title="기온 vs 지중온도 시계열",
        height=400,
    )
    fig.update_traces(line={"width": 1.2})
    st.plotly_chart(fig, width='stretch')
    chart_download_btn(fig, key="soil_ts_chart", filename="soil_temperature_timeseries")


# ──────────────────────────────────────────────────────────────
# 서브탭 4: 위상지연 분석
# ──────────────────────────────────────────────────────────────

def _tab_phase_lag(df):
    st.markdown("#### ⏱️ 위상지연 분석")
    st.caption(
        "지면온도의 연간 최고값 발생일 대비 각 깊이의 최고온도 도달 지연 일수입니다. "
        "깊을수록 열이 늦게 도달하며, 이 지연 시간으로 토양의 열전도 특성을 추정합니다."
    )

    if not _has_col(df, "year", "soil_temp_surface"):
        st.info("지면온도(soil_temp_surface) 또는 year 컬럼이 없어 위상지연 분석을 수행할 수 없습니다.")
        return

    years = df["year"].unique()
    if len(years) < 2:
        st.warning("⚠️ 데이터 기간이 2년 미만입니다. 위상지연 분석 신뢰도가 낮습니다.")

    depths = _available_depths(df)
    lag_rows = []
    for d in depths:
        lags = []
        for year, ydf in df.groupby("year"):
            if "soil_temp_surface" not in ydf.columns or d["col"] not in ydf.columns:
                continue
            surf_series = ydf["soil_temp_surface"].dropna()
            tgt_series  = ydf[d["col"]].dropna()
            if surf_series.empty or tgt_series.empty:
                continue
            ref_idx = surf_series.idxmax()
            tgt_idx = tgt_series.idxmax()
            if pd.notna(ref_idx) and pd.notna(tgt_idx) and "date" in ydf.columns:
                lag = (ydf.loc[tgt_idx, "date"] - ydf.loc[ref_idx, "date"]).days
                lags.append(lag)
        if lags:
            lag_rows.append({
                "깊이": d["label"],
                "depth_m": d["depth_m"],
                "평균 지연 일수": round(np.mean(lags), 1),
            })

    if not lag_rows:
        st.warning("위상지연 계산 결과가 없습니다.")
        return

    lag_df = pd.DataFrame(lag_rows)
    fig = px.bar(
        lag_df, x="평균 지연 일수", y="깊이",
        orientation="h",
        title="깊이별 평균 위상지연 일수 (지면온도 연간 최고 기준)",
        height=380,
        text="평균 지연 일수",
    )
    fig.update_traces(texttemplate="%{text}일", textposition="outside")
    fig.update_layout(yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig, width='stretch')
    chart_download_btn(fig, key="soil_phaselag_chart", filename="soil_phase_lag")

    st.dataframe(lag_df[["깊이", "평균 지연 일수"]], width='stretch')
    st.download_button(
        label="CSV 다운로드",
        data=_csv_bytes(lag_df[["깊이", "평균 지연 일수"]]),
        file_name="위상지연_분석.csv", mime="text/csv",
        key="dl_soil_phaselag",
    )


# ──────────────────────────────────────────────────────────────
# 서브탭 5: 깊이별 통계표
# ──────────────────────────────────────────────────────────────

def _tab_stats_table(df):
    st.markdown("#### 📋 깊이별 통계 요약표")

    period_unit = st.selectbox(
        "집계 단위", ["전체 기간", "연별", "월별", "계절별"],
        key="soil_stats_period",
    )

    depths = _available_depths(df)
    if not depths:
        st.info("지중온도 컬럼이 없습니다.")
        return

    def make_stats(sub_df):
        rows = {}
        for d in depths:
            if d["col"] not in sub_df.columns:
                continue
            s = sub_df[d["col"]].dropna()
            rows[d["label"]] = {
                "평균": round(s.mean(), 1),
                "최고": round(s.max(), 1),
                "최저": round(s.min(), 1),
                "표준편차": round(s.std(), 2),
                "변동폭": round(s.max() - s.min(), 1),
                "유효관측수": len(s),
            }
        return pd.DataFrame(rows).T

    if period_unit == "전체 기간":
        result = make_stats(df)
        st.dataframe(result, width='stretch')
        st.download_button(
            label="CSV 다운로드",
            data=_csv_bytes(result.reset_index().rename(columns={"index": "깊이"})),
            file_name="지중온도_통계.csv", mime="text/csv",
            key="dl_soil_stats_total",
        )
    elif period_unit == "연별":
        for year, ydf in df.groupby("year"):
            st.markdown(f"**{year}년**")
            st.dataframe(make_stats(ydf), width='stretch')
    elif period_unit == "월별":
        for month, mdf in df.groupby("month"):
            st.markdown(f"**{month}월**")
            st.dataframe(make_stats(mdf), width='stretch')
    else:  # 계절별
        for season, sdf in df.groupby("season"):
            st.markdown(f"**{season}**")
            st.dataframe(make_stats(sdf), width='stretch')


# ──────────────────────────────────────────────────────────────
# 서브탭 6: 기후 분석 (장기 트렌드 + 진폭 감쇠)
# ──────────────────────────────────────────────────────────────

def _tab_climate(df):
    st.markdown("#### 🌡️ 장기 추세 및 진폭 감쇠 분석")

    depths = _available_depths(df)
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**깊이별 연간 온도 변동폭 (진폭 감쇠)**")
        st.caption(
            "깊이가 깊어질수록 연간 온도 변동폭이 감소합니다. "
            "변동폭 2°C 이하가 되는 깊이가 지열 히트펌프 최적 설치 깊이입니다."
        )
        amp_rows = []
        for d in depths:
            if d["col"] not in df.columns:
                continue
            amp = df[d["col"]].max() - df[d["col"]].min()
            amp_rows.append({"깊이": d["label"], "depth_m": d["depth_m"], "변동폭(°C)": round(amp, 1)})
        if amp_rows:
            amp_df = pd.DataFrame(amp_rows)
            fig_amp = px.bar(
                amp_df, x="깊이", y="변동폭(°C)",
                title="깊이별 온도 변동폭",
                height=350,
                color="변동폭(°C)", color_continuous_scale="Reds",
            )
            fig_amp.add_hline(y=2, line_dash="dash", line_color="blue",
                              annotation_text="항온층 기준(2°C)")
            st.plotly_chart(fig_amp, width='stretch')
            chart_download_btn(fig_amp, key="soil_amplitude_chart", filename="soil_amplitude")

    with col2:
        st.markdown("**깊이별 장기 트렌드 (°C/10년)**")
        st.caption(
            "지표에서 깊어질수록 트렌드 기울기가 감소하면 "
            "기후변화 신호가 아직 깊은 지중까지 침투하지 않은 것입니다."
        )
        if not _has_col(df, "year"):
            st.info("year 컬럼이 없습니다.")
            return

        years_count = df["year"].nunique()
        if years_count < 5:
            st.warning(f"⚠️ 데이터 연수가 {years_count}년으로 트렌드 신뢰도가 낮습니다. (권장: 5년 이상)")

        trend_rows = []
        for d in depths:
            if d["col"] not in df.columns:
                continue
            annual = df.groupby("year")[d["col"]].mean().dropna()
            if len(annual) < 3:
                continue
            slope, _, r_value, p_value, _ = stats.linregress(annual.index, annual.values)
            trend_rows.append({
                "깊이": d["label"],
                "트렌드(°C/10년)": round(slope * 10, 3),
                "R²": round(r_value ** 2, 3),
                "p-value": round(p_value, 4),
            })

        if trend_rows:
            trend_df = pd.DataFrame(trend_rows)
            fig_trend = px.bar(
                trend_df, x="트렌드(°C/10년)", y="깊이",
                orientation="h",
                title="깊이별 장기 온도 트렌드",
                height=350,
                color="트렌드(°C/10년)",
                color_continuous_scale="RdBu_r",
                color_continuous_midpoint=0,
            )
            st.plotly_chart(fig_trend, width='stretch')
            chart_download_btn(fig_trend, key="soil_trend_chart", filename="soil_trend")
            st.dataframe(trend_df, width='stretch')
            st.download_button(
                label="CSV 다운로드",
                data=_csv_bytes(trend_df),
                file_name="지중온도_트렌드.csv", mime="text/csv",
                key="dl_soil_trend",
            )


# ──────────────────────────────────────────────────────────────
# 서브탭 7: 에너지 분석 (지열 포텐셜)
# ──────────────────────────────────────────────────────────────

def _tab_energy(df):
    st.markdown("#### ⚡ 지열 포텐셜 분석")

    depths = _available_depths(df)

    # 지열 포텐셜 등급 산출
    basis_col = None
    basis_label = None
    for candidate in ["soil_temp_300cm", "soil_temp_500cm", "soil_temp_150cm"]:
        if candidate in df.columns:
            basis_col = candidate
            basis_label = next(d["label"] for d in SOIL_DEPTH_COLS if d["col"] == candidate)
            break

    if basis_col:
        annual_mean = df[basis_col].mean()
        annual_amp  = df[basis_col].max() - df[basis_col].min()

        if annual_mean >= 14 and annual_amp <= 3:
            grade, grade_desc, grade_color = "A", "지열 활용 최우수", "🟢"
        elif annual_mean >= 12 and annual_amp <= 5:
            grade, grade_desc, grade_color = "B", "우수", "🔵"
        elif annual_mean >= 10 and annual_amp <= 7:
            grade, grade_desc, grade_color = "C", "양호", "🟡"
        elif annual_mean >= 8:
            grade, grade_desc, grade_color = "D", "보통", "🟠"
        else:
            grade, grade_desc, grade_color = "E", "불량", "🔴"

        st.markdown(f"### {grade_color} **{grade} 등급** — {grade_desc}")
        st.caption(
            f"분석 기준: {basis_label} | 연평균 {annual_mean:.1f}°C | 연변동폭 {annual_amp:.1f}°C\n\n"
            "A등급: 연평균 ≥14°C이면서 연변동폭 ≤3°C / B: ≥12°C, ≤5°C / "
            "C: ≥10°C, ≤7°C / D: ≥8°C / E: 기타"
        )
        st.divider()

    # 기온 vs 지중온도 비교
    if "temp_avg" in df.columns and depths:
        st.markdown("**기온 vs 지중온도 비교 — 냉난방 전환 적기**")
        st.caption(
            "지중온도 > 기온인 기간(동절기)은 히트펌프 난방에 유리하고, "
            "지중온도 < 기온인 기간(하절기)은 냉방에 유리합니다."
        )
        cmp_col_label = st.selectbox(
            "비교할 깊이 선택",
            options=[d["label"] for d in depths],
            index=min(6, len(depths) - 1),
            key="soil_energy_cmp_depth",
        )
        cmp_dcol = next(d["col"] for d in depths if d["label"] == cmp_col_label)
        cmp_df = df[["date", "temp_avg", cmp_dcol]].rename(
            columns={"temp_avg": "기온(평균)", cmp_dcol: f"지중온도({cmp_col_label})"}
        )
        melted_cmp = cmp_df.melt(id_vars="date", var_name="항목", value_name="온도(°C)")
        fig_cmp = px.line(
            melted_cmp, x="date", y="온도(°C)", color="항목",
            title=f"기온 vs 지중온도 {cmp_col_label} 비교",
            height=380,
        )
        st.plotly_chart(fig_cmp, width='stretch')
        chart_download_btn(fig_cmp, key="soil_energy_cmp_chart", filename="soil_energy_comparison")


# ──────────────────────────────────────────────────────────────
# 서브탭 8: 농업 분석
# ──────────────────────────────────────────────────────────────

def _tab_agri(df):
    st.markdown("#### 🌱 농업 활용 — 파종 적기 및 토양 적산온도")

    depths = _available_depths(df)
    if not depths:
        st.info("지중온도 컬럼이 없습니다.")
        return

    CROP_PRESETS = {
        "벼":       {"t_base": 13.0, "depth_label": "10cm"},
        "옥수수":   {"t_base": 10.0, "depth_label": "5cm"},
        "감자":     {"t_base":  7.0, "depth_label": "10cm"},
        "콩":       {"t_base": 12.0, "depth_label": "5cm"},
        "고추":     {"t_base": 15.0, "depth_label": "5cm"},
        "보리·밀":  {"t_base":  3.0, "depth_label": "10cm"},
        "직접 입력":{"t_base": 10.0, "depth_label": "10cm"},
    }

    col1, col2, col3 = st.columns(3)
    with col1:
        crop = st.selectbox("작물", list(CROP_PRESETS.keys()), key="soil_agri_crop")
    with col2:
        t_base = st.number_input(
            "기준온도(°C)", value=float(CROP_PRESETS[crop]["t_base"]),
            min_value=-5.0, max_value=30.0, step=0.5, key="soil_agri_tbase",
        )
    with col3:
        default_depth = CROP_PRESETS[crop]["depth_label"]
        depth_labels  = [d["label"] for d in depths]
        default_idx   = next((i for i, l in enumerate(depth_labels) if default_depth in l), 0)
        sel_depth_label = st.selectbox(
            "분석 깊이", depth_labels, index=default_idx, key="soil_agri_depth",
        )

    sel_col = next(d["col"] for d in depths if d["label"] == sel_depth_label)

    st.caption(
        f"'{crop}' 선택 — 기준온도 {t_base}°C, 분석 깊이 {sel_depth_label} 기준으로 "
        "파종 적기 도달일과 토양 적산온도(sGDD)를 계산합니다."
    )

    if sel_col not in df.columns:
        st.warning(f"{sel_depth_label} 컬럼이 없습니다.")
        return

    # 파종 적기 달성일
    st.markdown("**파종 적기 달성일 (연도별)**")
    sowing_rows = []
    for year, ydf in df.groupby("year"):
        ydf = ydf.sort_values("date")
        above    = (ydf[sel_col] >= t_base).astype(int)
        rolling5 = above.rolling(5).sum()
        idx5     = rolling5[rolling5 >= 5].index
        sowing_date = ydf.loc[idx5[0], "date"] if len(idx5) > 0 else None
        first_idx   = ydf[ydf[sel_col] >= t_base].index
        first_date  = ydf.loc[first_idx[0], "date"] if len(first_idx) > 0 else None
        sowing_rows.append({
            "연도": year,
            "최초 도달일": first_date.strftime("%m-%d") if first_date else "-",
            "5일 연속 유지 시작일": sowing_date.strftime("%m-%d") if sowing_date else "-",
        })
    sowing_df = pd.DataFrame(sowing_rows)
    st.dataframe(sowing_df, width='stretch')
    st.download_button(
        label="CSV 다운로드",
        data=_csv_bytes(sowing_df),
        file_name=f"파종적기_{crop}.csv", mime="text/csv",
        key="dl_soil_sowing",
    )

    # sGDD 누적 곡선
    st.markdown("**토양 적산온도(sGDD) 누적 곡선 — 연도별 비교**")
    st.caption(
        f"공식: sGDD_일 = max(T_지중 − {t_base}°C, 0)를 연도 초부터 누적한 값입니다. "
        "연도가 위로 올라갈수록 해당 연도의 적산온도 누적이 빠른 것입니다."
    )
    fig_gdd = go.Figure()
    colors = px.colors.qualitative.Set1
    for i, (year, ydf) in enumerate(df.groupby("year")):
        ydf = ydf.sort_values("date")
        daily_gdd  = np.maximum(ydf[sel_col].values - t_base, 0)
        cumgdd     = np.nancumsum(daily_gdd)
        day_of_year = ydf["date"].dt.dayofyear.values
        fig_gdd.add_trace(go.Scatter(
            x=day_of_year, y=cumgdd, mode="lines", name=str(year),
            line={"color": colors[i % len(colors)], "width": 1.5},
        ))
    fig_gdd.update_layout(
        title=f"토양 적산온도(sGDD) 누적 곡선 — {crop} (기준온도 {t_base}°C, {sel_depth_label})",
        xaxis_title="연중 일수 (일)", yaxis_title="누적 sGDD (°C·일)",
        height=400,
        legend={"orientation": "h", "yanchor": "bottom", "y": -0.35},
    )
    st.plotly_chart(fig_gdd, width='stretch')
    chart_download_btn(fig_gdd, key="soil_sgdd_chart", filename="soil_sgdd_cumulative")

    # 동결 분석
    st.divider()
    st.markdown("**🧊 동결 관련 지표 (지면온도 기준)**")
    st.caption(
        "지면온도(soil_temp_surface) 0°C 이하인 날을 분석합니다. "
        "월동 작물의 동해 피해 가능 시기 파악에 활용합니다."
    )
    if "soil_temp_surface" in df.columns:
        frost_rows = []
        for year, ydf in df.groupby("year"):
            frost_days = int((ydf["soil_temp_surface"] < 0).sum())
            min_temp   = ydf["soil_temp_surface"].min()
            frost_rows.append({
                "연도": year,
                "지면온도 0°C 미만 일수": frost_days,
                "지면온도 최저(°C)": round(min_temp, 1) if pd.notna(min_temp) else "-",
            })
        frost_df = pd.DataFrame(frost_rows)
        st.dataframe(frost_df, width='stretch')
        st.download_button(
            label="CSV 다운로드",
            data=_csv_bytes(frost_df),
            file_name="동결일수_분석.csv", mime="text/csv",
            key="dl_soil_frost",
        )
    else:
        st.info("지면온도(soil_temp_surface) 컬럼이 없습니다.")


# ──────────────────────────────────────────────────────────────
# 공개 render() 진입점
# ──────────────────────────────────────────────────────────────

def render(df: pd.DataFrame) -> None:
    """지중온도 분석 탭 렌더링. df는 필터링된 일자료 DataFrame."""

    st.subheader("🌡️ 지중온도 분석")

    if df is None or df.empty:
        st.warning("데이터가 없습니다.")
        return

    available = _available_depths(df)
    if not available:
        st.info(
            "현재 데이터에 지중온도 컬럼이 없습니다. "
            "기상청 ASOS 일자료에 지중온도 항목(지면온도, 5cm~5.0m 지중온도)이 "
            "포함된 CSV를 업로드하면 이 분석이 활성화됩니다."
        )
        return

    st.caption(f"✅ 지중온도 컬럼 {len(available)}개 감지: {', '.join(d['label'] for d in available)}")

    _section_guide()
    st.divider()

    # 관측소 필터
    if "station_name" in df.columns:
        all_stations = sorted(df["station_name"].unique().tolist())
        selected_stations = st.multiselect(
            "관측소 선택", options=all_stations, default=all_stations,
            key="soil_station_multiselect",
        )
        filtered = _filter_stations(df, selected_stations)
    else:
        filtered = df

    if filtered.empty:
        st.warning("선택된 관측소에 데이터가 없습니다.")
        return

    # 서브탭 구성
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "🗺️ 깊이-시간 히트맵",
        "📏 수직 프로파일",
        "📈 시계열 비교",
        "⏱️ 위상지연",
        "📋 통계표",
        "🌍 기후 분석",
        "⚡ 에너지",
        "🌾 농업",
    ])

    with tab1: _tab_depth_heatmap(filtered)
    with tab2: _tab_vertical_profile(filtered)
    with tab3: _tab_timeseries(filtered)
    with tab4: _tab_phase_lag(filtered)
    with tab5: _tab_stats_table(filtered)
    with tab6: _tab_climate(filtered)
    with tab7: _tab_energy(filtered)
    with tab8: _tab_agri(filtered)
