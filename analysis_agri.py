# analysis_agri.py — 농업기상 분석 탭
# 서리·폭염·집중홍수·일사부족 / ET₀ / GDD / 토양수분수지 / 지중온도 / PAR

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ET₀ 계산 (FAO-56 Penman-Monteith)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def calc_et0(df: pd.DataFrame) -> pd.Series:
    """
    FAO-56 Penman-Monteith ET₀ (mm/day)
    필요: temp_avg, temp_max, temp_min, humidity, solar_rad, wind_speed
    """
    required = ["temp_avg"]
    if not all(c in df.columns for c in required):
        return pd.Series(np.nan, index=df.index)

    T = df["temp_avg"].copy()
    Tmax = df["temp_max"].copy() if "temp_max" in df.columns else T + 3
    Tmin = df["temp_min"].copy() if "temp_min" in df.columns else T - 3
    RH = df["humidity"].copy().fillna(70) if "humidity" in df.columns else pd.Series(70, index=df.index)
    Rs = df["solar_rad"].copy().fillna(0) if "solar_rad" in df.columns else pd.Series(0.0, index=df.index)
    u10 = df["wind_speed"].copy().fillna(2.0) if "wind_speed" in df.columns else pd.Series(2.0, index=df.index)

    Tmax = Tmax.fillna(T + 3)
    Tmin = Tmin.fillna(T - 3)

    # 2m 높이 풍속 환산
    u2 = u10 * (4.87 / np.log(67.8 * 10 - 5.42))

    # 포화/실제 증기압 (kPa)
    es_max = 0.6108 * np.exp(17.27 * Tmax / (Tmax + 237.3))
    es_min = 0.6108 * np.exp(17.27 * Tmin / (Tmin + 237.3))
    es = (es_max + es_min) / 2
    ea = (es * RH / 100).clip(lower=0.001)

    # 포화증기압 기울기 Δ
    delta = 4098 * 0.6108 * np.exp(17.27 * T / (T + 237.3)) / (T + 237.3) ** 2

    # 건습계 상수
    gamma = 0.0665  # kPa/°C (고도 ~0m 기준)

    # 순복사량
    Rns = 0.77 * Rs  # 단파 순복사 (알베도 0.23)

    # 일조율 추정
    if "sunshine" in df.columns and "daylight_hours" in df.columns:
        n_N = (df["sunshine"] / df["daylight_hours"].replace(0, np.nan)).fillna(0.5).clip(0, 1)
    elif "sunshine" in df.columns:
        n_N = (df["sunshine"] / 12.0).clip(0, 1)
    else:
        n_N = pd.Series(0.5, index=df.index)

    # 장파 순복사 (FAO-56 eq. 39)
    sigma = 4.903e-9  # MJ/m²/day/K⁴
    T_K4 = ((Tmax + 273.16) ** 4 + (Tmin + 273.16) ** 4) / 2
    Rnl = sigma * T_K4 * (0.34 - 0.14 * np.sqrt(ea)) * (1.35 * n_N - 0.35)

    Rn = (Rns - Rnl).clip(lower=-5)

    # PM 방정식
    num = 0.408 * delta * Rn + gamma * (900 / (T + 273)) * u2 * (es - ea)
    den = delta + gamma * (1 + 0.34 * u2)
    et0 = (num / den).clip(lower=0)

    return et0.round(2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 적산온도 (Growing Degree Days)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def calc_gdd(df: pd.DataFrame, t_base: float = 10.0) -> pd.Series:
    """일별 적산온도 기여값 = max(T_avg - T_base, 0)"""
    if "temp_avg" not in df.columns:
        return pd.Series(np.nan, index=df.index)
    return (df["temp_avg"] - t_base).clip(lower=0).round(2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 토양수분수지 (단순 P - ET₀ 누적)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def calc_soil_water_balance(df: pd.DataFrame) -> pd.DataFrame:
    """
    단순 토양수분수지
    SWB = 누적(P - ET₀), 상한=100mm(포장용수량), 하한=0
    """
    result = df.copy()
    result["et0"] = calc_et0(df)
    precip = df["precipitation"].fillna(0) if "precipitation" in df.columns else pd.Series(0.0, index=df.index)
    result["wb_daily"] = (precip - result["et0"]).round(2)
    result["wb_cumul"] = result.groupby("station_name")["wb_daily"].cumsum().clip(-300, 300).round(2)
    return result


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PAR 추정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def calc_par(df: pd.DataFrame) -> pd.Series:
    """
    PAR(mol/m²/day) ≈ 일사량(MJ/m²/day) × 2.02
    (가시광선 비율 0.48 × MJ→J×10⁶ / (E_photon≈218000J/mol))
    """
    if "solar_rad" not in df.columns:
        return pd.Series(np.nan, index=df.index)
    return (df["solar_rad"] * 2.02).round(2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 메인 렌더 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def render(df: pd.DataFrame) -> None:
    """농업기상 분석 탭 렌더링"""

    st.subheader("🌱 농업기상 분석")

    if df is None or len(df) == 0:
        st.info("먼저 데이터를 업로드하세요.")
        return

    stations = sorted(df["station_name"].unique())
    selected = st.multiselect("관측소 선택", stations, default=stations, key="agri_stations")
    if not selected:
        st.warning("관측소를 선택하세요.")
        return

    df = df[df["station_name"].isin(selected)].copy()

    t1, t2, t3, t4, t5, t6 = st.tabs([
        "특이일수 통계",
        "증발산량 ET₀",
        "적산온도 GDD",
        "토양수분수지",
        "지중온도",
        "PAR · 일차생산량",
    ])

    with t1:
        _tab_special_days(df)

    with t2:
        _tab_et0(df)

    with t3:
        _tab_gdd(df)

    with t4:
        _tab_soil_water(df)

    with t5:
        _tab_soil_temp(df)

    with t6:
        _tab_par(df)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 서브탭 렌더 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _tab_special_days(df: pd.DataFrame) -> None:
    st.markdown("### 📋 연도별 농업기상 특이일수")

    records = []
    for stn, sdf in df.groupby("station_name"):
        for year, ydf in sdf.groupby("year"):
            row = {"관측소": stn, "연도": int(year)}

            if "temp_max" in ydf.columns:
                row["폭염일수(≥33°C)"] = int((ydf["temp_max"] >= 33).sum())
            if "temp_min" in ydf.columns:
                row["서리위험일수(≤0°C)"] = int((ydf["temp_min"] <= 0).sum())
                row["열대야일수(≥25°C)"] = int((ydf["temp_min"] >= 25).sum())
            if "frost_temp" in ydf.columns:
                row["초상온도서리(≤0°C)"] = int((ydf["frost_temp"] <= 0).sum())
            if "precipitation" in ydf.columns:
                row["집중홍수위험일(≥80mm)"] = int((ydf["precipitation"] >= 80).sum())
            if "sunshine" in ydf.columns:
                row["일사부족일(≤3hr)"] = int((ydf["sunshine"] <= 3).sum())
            if "wind_max" in ydf.columns:
                row["강풍일(≥14m/s)"] = int((ydf["wind_max"] >= 14).sum())

            records.append(row)

    if not records:
        st.warning("계산에 필요한 데이터가 없습니다.")
        return

    tbl = pd.DataFrame(records)
    st.dataframe(tbl, use_container_width=True, height=420)

    csv = tbl.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("⬇️ CSV 다운로드", csv, "농업기상_특이일수.csv", "text/csv", key="agri_special_csv")

    # 항목별 막대 차트
    numeric_cols = [c for c in tbl.columns if c not in ["관측소", "연도"]]
    if numeric_cols:
        st.markdown("#### 항목별 연도별 추이")
        sel_col = st.selectbox("항목", numeric_cols, key="agri_special_col")
        sel_stn = st.selectbox("관측소", tbl["관측소"].unique().tolist(), key="agri_special_stn")

        plot_df = tbl[tbl["관측소"] == sel_stn]
        fig = px.bar(
            plot_df, x="연도", y=sel_col,
            labels={"연도": "연도", sel_col: "일수"},
            color_discrete_sequence=["#27ae60"]
        )
        fig.update_layout(height=380)
        st.plotly_chart(fig, use_container_width=True)


def _tab_et0(df: pd.DataFrame) -> None:
    st.markdown("### 💧 기준증발산량 ET₀ (FAO Penman-Monteith)")

    required_cols = ["temp_avg"]
    if not any(c in df.columns for c in required_cols):
        st.warning("기온 데이터가 필요합니다.")
        return

    df = df.copy()
    df["et0"] = calc_et0(df)

    col1, col2 = st.columns(2)
    with col1:
        freq = st.radio("집계 단위", ["월별", "연별"], horizontal=True, key="agri_et0_freq")
    with col2:
        stn = st.selectbox("관측소", df["station_name"].unique().tolist(), key="agri_et0_stn")

    sdf = df[df["station_name"] == stn].copy()

    if freq == "월별":
        grouped = sdf.groupby(["year", "month"])["et0"].sum().reset_index()
        grouped["date"] = pd.to_datetime(
            grouped["year"].astype(str) + "-" + grouped["month"].astype(str) + "-01"
        )
        fig = px.line(
            grouped, x="date", y="et0",
            labels={"date": "날짜", "et0": "ET₀ (mm/월)"},
            title=f"{stn} 월별 ET₀"
        )
    else:
        grouped = sdf.groupby("year")["et0"].sum().reset_index()
        fig = px.bar(
            grouped, x="year", y="et0",
            labels={"year": "연도", "et0": "ET₀ (mm/년)"},
            title=f"{stn} 연별 ET₀",
            color_discrete_sequence=["#2980b9"]
        )

    fig.update_layout(height=380)
    st.plotly_chart(fig, use_container_width=True)

    # 실측 증발량과 비교
    if "evaporation_large" in df.columns or "evaporation_small" in df.columns:
        st.markdown("#### ET₀ vs 실측 증발량 비교 (월합계)")
        evap_col = "evaporation_large" if "evaporation_large" in df.columns else "evaporation_small"
        monthly = sdf.groupby(["year", "month"]).agg(
            et0=("et0", "sum"),
            evap=(evap_col, "sum")
        ).reset_index()
        monthly["date"] = pd.to_datetime(
            monthly["year"].astype(str) + "-" + monthly["month"].astype(str) + "-01"
        )
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=monthly["date"], y=monthly["et0"].round(1),
                                  name="ET₀ (PM)", mode="lines"))
        fig2.add_trace(go.Scatter(x=monthly["date"], y=monthly["evap"].round(1),
                                  name=f"실측 ({evap_col})", mode="lines",
                                  line=dict(dash="dash")))
        fig2.update_layout(height=350, yaxis_title="증발산량 (mm/월)")
        st.plotly_chart(fig2, use_container_width=True)


def _tab_gdd(df: pd.DataFrame) -> None:
    st.markdown("### 🌿 적산온도 (Growing Degree Days)")

    if "temp_avg" not in df.columns:
        st.warning("평균기온 데이터가 필요합니다.")
        return

    col1, col2 = st.columns(2)
    with col1:
        t_base = st.selectbox("기준온도 (°C)", [0, 5, 10, 15], index=2, key="agri_gdd_tbase")
    with col2:
        stn = st.selectbox("관측소", df["station_name"].unique().tolist(), key="agri_gdd_stn")

    sdf = df[df["station_name"] == stn].copy()
    sdf["gdd"] = calc_gdd(sdf, float(t_base))

    # 연도별 누적 적산온도 (1월 1일부터 누적)
    annual_gdd = sdf.groupby("year")["gdd"].sum().reset_index()
    annual_gdd.columns = ["연도", f"연간 적산온도(T_base={t_base}°C)"]

    st.dataframe(annual_gdd, use_container_width=True, height=300)

    fig = px.bar(
        annual_gdd, x="연도", y=f"연간 적산온도(T_base={t_base}°C)",
        labels={"연도": "연도"},
        title=f"{stn} 연간 적산온도 (기준온도 {t_base}°C)",
        color_discrete_sequence=["#e67e22"]
    )
    fig.update_layout(height=380)
    st.plotly_chart(fig, use_container_width=True)

    # 특정 연도 일별 누적 곡선
    st.markdown("#### 연도별 일별 누적 적산온도")
    years = sorted(sdf["year"].unique())
    sel_years = st.multiselect(
        "비교 연도 선택", years,
        default=years[-3:] if len(years) >= 3 else years,
        key="agri_gdd_years"
    )

    if sel_years:
        fig2 = go.Figure()
        for yr in sel_years:
            yr_df = sdf[sdf["year"] == yr].copy()
            yr_df = yr_df.sort_values("date")
            yr_df["cumul_gdd"] = yr_df["gdd"].cumsum()
            yr_df["doy"] = yr_df["date"].dt.dayofyear
            fig2.add_trace(go.Scatter(
                x=yr_df["doy"],
                y=yr_df["cumul_gdd"].round(1),
                name=str(yr),
                mode="lines"
            ))
        fig2.update_layout(
            height=380,
            xaxis_title="일수 (Day of Year)",
            yaxis_title=f"누적 적산온도 (°C·day)",
        )
        st.plotly_chart(fig2, use_container_width=True)


def _tab_soil_water(df: pd.DataFrame) -> None:
    st.markdown("### 💧 토양수분수지 (단순 P - ET₀ 모델)")

    if "temp_avg" not in df.columns:
        st.warning("기온 데이터가 필요합니다.")
        return

    stn = st.selectbox("관측소", df["station_name"].unique().tolist(), key="agri_swb_stn")
    sdf = df[df["station_name"] == stn].copy().sort_values("date")

    swb = calc_soil_water_balance(sdf)

    # 월별 P, ET₀, P-ET₀ 비교
    monthly = swb.groupby(["year", "month"]).agg(
        precip=("precipitation", "sum") if "precipitation" in swb.columns else ("et0", lambda x: np.nan),
        et0=("et0", "sum"),
        wb=("wb_daily", "sum")
    ).reset_index()

    if "precipitation" not in swb.columns:
        monthly["precip"] = np.nan

    monthly["date"] = pd.to_datetime(
        monthly["year"].astype(str) + "-" + monthly["month"].astype(str) + "-01"
    )

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=monthly["date"], y=monthly["precip"].round(1),
        name="강수량 (mm)", opacity=0.6
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=monthly["date"], y=monthly["et0"].round(1),
        name="ET₀ (mm)", mode="lines", line=dict(color="red")
    ), secondary_y=False)
    fig.add_trace(go.Bar(
        x=monthly["date"], y=monthly["wb"].round(1),
        name="P-ET₀ (mm)", opacity=0.5,
        marker_color=monthly["wb"].apply(lambda v: "#3498db" if v >= 0 else "#e74c3c")
    ), secondary_y=True)

    fig.update_yaxes(title_text="강수·ET₀ (mm)", secondary_y=False)
    fig.update_yaxes(title_text="물수지 P-ET₀ (mm)", secondary_y=True)
    fig.update_layout(height=400, barmode="overlay")
    st.plotly_chart(fig, use_container_width=True)

    st.caption("※ 단순 수지 모델 (토양 저장 미고려). 음수=토양수분 부족, 양수=잉여.")


def _tab_soil_temp(df: pd.DataFrame) -> None:
    st.markdown("### 🌍 지중온도 깊이별 분석")

    soil_cols = {
        "지면온도": "soil_temp_surface",
        "5cm": "soil_temp_5cm",
        "10cm": "soil_temp_10cm",
        "20cm": "soil_temp_20cm",
        "30cm": "soil_temp_30cm",
        "0.5m": "soil_temp_50cm",
        "1.0m": "soil_temp_100cm",
        "1.5m": "soil_temp_150cm",
        "3.0m": "soil_temp_300cm",
        "5.0m": "soil_temp_500cm",
    }
    available = {k: v for k, v in soil_cols.items() if v in df.columns}

    if not available:
        st.warning("지중온도 데이터가 없습니다. (ASOS 파일에 지중온도 항목 포함 필요)")
        return

    stn = st.selectbox("관측소", df["station_name"].unique().tolist(), key="agri_soil_stn")
    sdf = df[df["station_name"] == stn].copy()

    # 월별 평균 깊이별 온도
    st.markdown("#### 월별 평균 지중온도 (깊이별 비교)")

    sel_depths = st.multiselect(
        "깊이 선택", list(available.keys()),
        default=list(available.keys())[:5],
        key="agri_soil_depths"
    )

    if not sel_depths:
        return

    monthly = sdf.groupby(["year", "month"])[
        [available[d] for d in sel_depths]
    ].mean().reset_index()

    monthly["date"] = pd.to_datetime(
        monthly["year"].astype(str) + "-" + monthly["month"].astype(str) + "-01"
    )

    fig = go.Figure()
    colors = px.colors.qualitative.Plotly
    for i, depth in enumerate(sel_depths):
        col = available[depth]
        if col in monthly.columns:
            fig.add_trace(go.Scatter(
                x=monthly["date"],
                y=monthly[col].round(2),
                name=depth,
                mode="lines",
                line=dict(color=colors[i % len(colors)])
            ))

    fig.update_layout(
        height=400,
        xaxis_title="날짜",
        yaxis_title="온도 (°C)",
        legend_title="깊이"
    )
    st.plotly_chart(fig, use_container_width=True)

    # 연중 깊이별 분포 (특정 연도)
    st.markdown("#### 연중 지중온도 깊이별 프로파일 (월평균)")
    years = sorted(sdf["year"].dropna().unique().tolist())
    sel_year = st.selectbox("연도", years, index=len(years) - 1, key="agri_soil_year")

    yr_df = sdf[sdf["year"] == sel_year]
    monthly_profile = yr_df.groupby("month")[
        [available[d] for d in available]
    ].mean().round(2)
    monthly_profile.index.name = "월"
    monthly_profile.columns = list(available.keys())[:len(monthly_profile.columns)]

    st.dataframe(monthly_profile, use_container_width=True)


def _tab_par(df: pd.DataFrame) -> None:
    st.markdown("### ☀️ 개략 PAR 및 일차생산량 추정")

    if "solar_rad" not in df.columns:
        st.warning("일사량 데이터가 필요합니다.")
        return

    stn = st.selectbox("관측소", df["station_name"].unique().tolist(), key="agri_par_stn")
    sdf = df[df["station_name"] == stn].copy()
    sdf["par"] = calc_par(sdf)

    # 월별 PAR
    monthly = sdf.groupby(["year", "month"]).agg(
        par_sum=("par", "sum"),
        solar_sum=("solar_rad", "sum")
    ).reset_index()
    monthly["date"] = pd.to_datetime(
        monthly["year"].astype(str) + "-" + monthly["month"].astype(str) + "-01"
    )

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=monthly["date"], y=monthly["par_sum"].round(1),
        name="월합계 PAR (mol/m²)",
        opacity=0.75, marker_color="#f39c12"
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=monthly["date"], y=monthly["solar_sum"].round(1),
        name="월합계 일사량 (MJ/m²)",
        mode="lines", line=dict(color="#8e44ad")
    ), secondary_y=True)

    fig.update_yaxes(title_text="PAR (mol/m²/월)", secondary_y=False)
    fig.update_yaxes(title_text="일사량 (MJ/m²/월)", secondary_y=True)
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)

    # 일차생산량 개략 추정
    st.markdown("#### 개략 일차생산량 (GPP) 추정")

    col1, col2 = st.columns(2)
    with col1:
        fapar = st.slider("fAPAR (식생 흡수 PAR 비율)", 0.1, 1.0, 0.6, 0.05,
                          key="agri_par_fapar",
                          help="작물·삼림별 차이. 작물 ~0.5~0.7, 삼림 ~0.8")
    with col2:
        lue = st.slider("LUE (광이용효율, gC/mol PAR)", 0.5, 5.0, 2.5, 0.1,
                        key="agri_par_lue",
                        help="작물 ~2~4 gC/mol, 삼림 ~1~2 gC/mol")

    annual_par = sdf.groupby("year")["par"].sum().reset_index()
    annual_par["gpp_gC_m2_yr"] = (annual_par["par"] * fapar * lue).round(1)
    annual_par["gpp_tC_ha_yr"] = (annual_par["gpp_gC_m2_yr"] / 100).round(2)
    annual_par.columns = ["연도", "연간PAR(mol/m²)", "GPP(gC/m²/yr)", "GPP(tC/ha/yr)"]

    st.dataframe(annual_par, use_container_width=True)

    fig2 = px.bar(
        annual_par, x="연도", y="GPP(gC/m²/yr)",
        title="개략 연간 일차생산량 추정",
        color_discrete_sequence=["#27ae60"]
    )
    fig2.update_layout(height=350)
    st.plotly_chart(fig2, use_container_width=True)

    st.caption(
        "※ GPP = PAR × fAPAR × LUE 단순 LUE 모델 추정값. "
        "fAPAR·LUE는 식생 유형에 따라 조정하세요. "
        "정밀 분석에는 위성 원격탐사 데이터 병행 권장."
    )
