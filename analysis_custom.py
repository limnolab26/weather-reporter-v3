# analysis_custom.py — 맞춤 분석 모듈
# render(df) 함수 하나를 export합니다.
# 사용자가 컬럼·집계·차트를 직접 조립하고 결과를 다운로드합니다.

import io
import json

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

from chart_utils import chart_download_btn

# ── 컬럼 한국어 레이블 (data_processor.py COLUMN_MAPPING 역매핑) ──────────
# 사용자에게 보여주는 이름 → 내부 컬럼명
COL_LABELS: dict[str, str] = {
    "평균기온 (°C)":        "temp_avg",
    "최고기온 (°C)":        "temp_max",
    "최저기온 (°C)":        "temp_min",
    "강수량 (mm)":          "precipitation",
    "평균풍속 (m/s)":       "wind_speed",
    "최대풍속 (m/s)":       "wind_max",
    "최대순간풍속 (m/s)":   "wind_gust",
    "평균습도 (%)":         "humidity",
    "최소습도 (%)":         "humidity_min",
    "일조시간 (hr)":        "sunshine",
    "일사량 (MJ/m²)":      "solar_rad",
    "적설 (cm)":            "snowfall",
    "적설심 (cm)":          "snow_depth",
    "현지기압 (hPa)":      "pressure_local",
    "해면기압 (hPa)":      "pressure_sea",
    "이슬점온도 (°C)":     "dew_point",
    "초상온도 (°C)":       "frost_temp",
    "증기압 (hPa)":        "vapor_pressure",
    "대형증발량 (mm)":     "evaporation_large",
    "소형증발량 (mm)":     "evaporation_small",
    "전운량 (1/10)":       "cloud_cover",
    "안개지속시간 (hr)":   "fog_duration",
    # 지중온도
    "지면온도 (°C)":       "soil_temp_surface",
    "지중 5cm (°C)":       "soil_temp_5cm",
    "지중 10cm (°C)":      "soil_temp_10cm",
    "지중 20cm (°C)":      "soil_temp_20cm",
    "지중 30cm (°C)":      "soil_temp_30cm",
    "지중 50cm (°C)":      "soil_temp_50cm",
    "지중 100cm (°C)":     "soil_temp_100cm",
    "지중 150cm (°C)":     "soil_temp_150cm",
    "지중 300cm (°C)":     "soil_temp_300cm",
    "지중 500cm (°C)":     "soil_temp_500cm",
}

# X축 옵션
X_AXIS_OPTIONS: dict[str, str] = {
    "날짜 (일별)":   "date",
    "월":            "month",
    "연도":          "year",
    "계절":          "season",
    "연도-월":       "year_month",
}

# 집계 함수
AGG_FUNCTIONS: dict[str, str] = {
    "평균":   "mean",
    "합계":   "sum",
    "최대":   "max",
    "최소":   "min",
    "개수":   "count",
    "표준편차": "std",
}

# 차트 유형
CHART_TYPES = ["선 그래프", "막대 그래프", "산점도", "박스플롯", "히트맵(연×월)"]


# ── 헬퍼 ──────────────────────────────────────────────────────────────────

def _available_cols(df: pd.DataFrame) -> dict[str, str]:
    """df에 실제 존재하는 컬럼만 COL_LABELS에서 필터링하여 반환"""
    return {label: col for label, col in COL_LABELS.items() if col in df.columns}


def _csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def _json_bytes(obj: dict) -> bytes:
    return json.dumps(obj, ensure_ascii=False, indent=2).encode("utf-8")


def _aggregate(df: pd.DataFrame, x_col: str, y_cols: list[str], agg_fn: str) -> pd.DataFrame:
    """
    x_col 기준으로 y_cols를 agg_fn으로 집계한다.
    날짜(date) 그룹핑 시에는 그대로 반환, 나머지는 groupby 사용.
    """
    if x_col == "date":
        return df[["date", "station_name"] + y_cols].copy()

    group_cols = ["station_name", x_col]
    existing = [c for c in group_cols if c in df.columns]
    result = df.groupby(existing)[y_cols].agg(agg_fn).reset_index()
    return result


def _load_config(uploaded_json) -> dict:
    """사용자가 업로드한 JSON 설정 파일 파싱"""
    try:
        return json.loads(uploaded_json.read().decode("utf-8"))
    except Exception:
        return {}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 서브탭 1: 맞춤 차트 빌더
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _tab_chart_builder(df: pd.DataFrame) -> None:
    """
    사용자가 X축·Y축·집계·차트유형을 고르면
    Plotly 차트가 실시간으로 업데이트된다.
    """
    st.markdown("#### 📊 맞춤 차트 빌더")
    st.caption(
        "원하는 기상요소와 집계 방식을 선택하면 차트가 즉시 생성됩니다. "
        "전문 분야에 따라 지중온도·기온·강수 등 어떤 조합이든 가능합니다."
    )

    avail = _available_cols(df)
    if not avail:
        st.info("분석 가능한 수치 컬럼이 없습니다.")
        return

    # ── 설정 패널 ──
    with st.container(border=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            x_label = st.selectbox(
                "X축", list(X_AXIS_OPTIONS.keys()),
                key="cst_chart_x",
            )
            x_col = X_AXIS_OPTIONS[x_label]

        with col2:
            agg_label = st.selectbox(
                "집계 방식", list(AGG_FUNCTIONS.keys()),
                key="cst_chart_agg",
            )
            agg_fn = AGG_FUNCTIONS[agg_label]

        with col3:
            chart_type = st.selectbox(
                "차트 유형", CHART_TYPES,
                key="cst_chart_type",
            )

        y_labels = st.multiselect(
            "Y축 항목 선택 (복수 가능)",
            options=list(avail.keys()),
            default=list(avail.keys())[:2],
            key="cst_chart_y",
        )

        # 관측소 (단일 선택: 비교는 색상으로)
        stations = sorted(df["station_name"].unique()) if "station_name" in df.columns else []
        if stations:
            sel_stn = st.selectbox("관측소", ["전체 (비교)"] + stations, key="cst_chart_stn")
        else:
            sel_stn = "전체 (비교)"

    if not y_labels:
        st.info("Y축 항목을 하나 이상 선택하세요.")
        return

    y_cols = [avail[l] for l in y_labels]

    # 데이터 필터링
    plot_df = df.copy()
    if sel_stn != "전체 (비교)" and "station_name" in df.columns:
        plot_df = df[df["station_name"] == sel_stn].copy()

    # x_col이 df에 없으면 안내
    if x_col != "date" and x_col not in plot_df.columns:
        st.warning(f"'{x_label}' 컬럼이 없습니다. 날짜 열에서 자동 생성되어야 합니다.")
        return

    # year_month는 문자열로 변환
    if x_col == "year_month" and "year_month" in plot_df.columns:
        plot_df["year_month"] = plot_df["year_month"].astype(str)

    # 집계
    if x_col != "date":
        group_cols = (["station_name", x_col] if "station_name" in plot_df.columns
                      else [x_col])
        agg_df = plot_df.groupby(group_cols)[y_cols].agg(agg_fn).reset_index()
    else:
        agg_df = plot_df[
            (["date", "station_name"] if "station_name" in plot_df.columns else ["date"])
            + y_cols
        ].copy()

    # 차트 생성
    color_col = "station_name" if sel_stn == "전체 (비교)" and "station_name" in agg_df.columns else None
    y_rename = {col: label for label, col in zip(y_labels, y_cols)}

    fig = None

    if chart_type == "선 그래프":
        melted = agg_df.melt(
            id_vars=[x_col] + (["station_name"] if color_col else []),
            value_vars=y_cols, var_name="_항목", value_name="_값"
        )
        melted["_항목"] = melted["_항목"].map(y_rename)
        color_key = "_항목" if len(y_cols) > 1 else color_col
        fig = px.line(
            melted, x=x_col, y="_값",
            color=color_key,
            labels={x_col: x_label, "_값": f"{agg_label} 값"},
            title=f"{' · '.join(y_labels)} — {x_label} × {agg_label}",
            height=420,
        )

    elif chart_type == "막대 그래프":
        melted = agg_df.melt(
            id_vars=[x_col] + (["station_name"] if color_col else []),
            value_vars=y_cols, var_name="_항목", value_name="_값"
        )
        melted["_항목"] = melted["_항목"].map(y_rename)
        color_key = "_항목" if len(y_cols) > 1 else color_col
        fig = px.bar(
            melted, x=x_col, y="_값",
            color=color_key,
            barmode="group",
            labels={x_col: x_label, "_값": f"{agg_label} 값"},
            title=f"{' · '.join(y_labels)} — {x_label} × {agg_label}",
            height=420,
        )

    elif chart_type == "산점도":
        if len(y_cols) < 2:
            st.info("산점도는 Y축 항목을 2개 이상 선택해야 합니다. 첫 번째 vs 두 번째 항목으로 표시됩니다.")
        x_data_col = y_cols[0]
        y_data_col = y_cols[1] if len(y_cols) > 1 else y_cols[0]
        fig = px.scatter(
            agg_df,
            x=x_data_col, y=y_data_col,
            color=color_col,
            labels={x_data_col: y_labels[0], y_data_col: y_labels[1] if len(y_labels) > 1 else y_labels[0]},
            title=f"{y_labels[0]} vs {y_labels[1] if len(y_labels) > 1 else y_labels[0]}",
            trendline="ols",
            height=420,
        )

    elif chart_type == "박스플롯":
        if x_col == "date":
            st.info("박스플롯은 X축을 '월', '계절', '연도' 중 하나로 선택하세요.")
            return
        melted = plot_df.melt(
            id_vars=[x_col] + (["station_name"] if color_col else []),
            value_vars=y_cols, var_name="_항목", value_name="_값"
        )
        melted["_항목"] = melted["_항목"].map(y_rename)
        fig = px.box(
            melted, x=x_col, y="_값",
            color="_항목",
            labels={x_col: x_label, "_값": "값"},
            title=f"{' · '.join(y_labels)} 분포 — {x_label}별",
            height=420,
        )

    elif chart_type == "히트맵(연×월)":
        if len(y_cols) > 1:
            st.info("히트맵은 Y축 항목을 1개만 선택하세요. 첫 번째 항목으로 표시됩니다.")
        heat_col = y_cols[0]
        heat_label = y_labels[0]
        if not all(c in plot_df.columns for c in ["year", "month"]):
            st.warning("year, month 컬럼이 필요합니다.")
            return
        pivot = (
            plot_df.groupby(["year", "month"])[heat_col]
            .agg(agg_fn)
            .unstack(level="month")
        )
        pivot.columns = [f"{m}월" for m in pivot.columns]
        fig = px.imshow(
            pivot,
            color_continuous_scale="RdBu_r",
            labels={"x": "월", "y": "연도", "color": heat_label},
            title=f"{heat_label} 연×월 히트맵 ({agg_label})",
            aspect="auto",
            height=max(300, len(pivot) * 30 + 100),
        )

    if fig is not None:
        st.plotly_chart(fig, use_container_width=True)
        chart_download_btn(fig, key="cst_chart_dl", filename="custom_chart")

        # 집계 결과 데이터 표시 및 CSV 다운로드
        with st.expander("📋 집계 데이터 보기 / 다운로드", expanded=False):
            display_df = agg_df.rename(columns=y_rename)
            st.dataframe(display_df, use_container_width=True)
            st.download_button(
                label="CSV 다운로드",
                data=_csv_bytes(display_df),
                file_name="custom_chart_data.csv",
                mime="text/csv",
                key="cst_chart_csv",
            )

        # 설정 저장 기능
        with st.expander("💾 이 설정 저장하기 (나중에 다시 불러올 수 있습니다)", expanded=False):
            config = {
                "type": "chart",
                "x_label": x_label,
                "agg_label": agg_label,
                "chart_type": chart_type,
                "y_labels": y_labels,
                "station": sel_stn,
            }
            config_name = st.text_input(
                "설정 이름", value="내 맞춤 차트",
                key="cst_chart_config_name",
            )
            config["name"] = config_name
            st.download_button(
                label="⬇️ 설정 JSON 저장",
                data=_json_bytes(config),
                file_name=f"{config_name}.json",
                mime="application/json",
                key="cst_chart_config_dl",
            )
            st.caption("저장한 JSON 파일은 '설정 불러오기' 탭에서 재사용할 수 있습니다.")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 서브탭 2: 맞춤 피벗 표
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _tab_pivot_builder(df: pd.DataFrame) -> None:
    """
    사용자가 행·열·값·집계를 지정하여
    엑셀 피벗 테이블 방식의 요약 표를 생성한다.
    """
    st.markdown("#### 📋 맞춤 피벗 표")
    st.caption(
        "행·열·값을 직접 지정하면 엑셀 피벗처럼 요약 표가 만들어집니다. "
        "CSV로 내보내서 바로 보고서에 활용하세요."
    )

    avail = _available_cols(df)

    # 차원 컬럼 (그룹핑에 쓸 수 있는 것들)
    DIM_OPTIONS: dict[str, str] = {
        "연도":    "year",
        "월":      "month",
        "계절":    "season",
        "관측소":  "station_name",
        "연도-월": "year_month",
    }
    avail_dims = {k: v for k, v in DIM_OPTIONS.items() if v in df.columns}

    with st.container(border=True):
        col1, col2 = st.columns(2)
        with col1:
            row_label = st.selectbox(
                "행(Row)", list(avail_dims.keys()),
                index=0, key="cst_pivot_row",
            )
        with col2:
            col_label = st.selectbox(
                "열(Column)", list(avail_dims.keys()),
                index=1 if len(avail_dims) > 1 else 0,
                key="cst_pivot_col",
            )

        val_label = st.selectbox(
            "값(Value)", list(avail.keys()),
            key="cst_pivot_val",
        )
        agg_label = st.selectbox(
            "집계 방식", list(AGG_FUNCTIONS.keys()),
            key="cst_pivot_agg",
        )

    row_col = avail_dims[row_label]
    col_col = avail_dims[col_label]
    val_col = avail[val_label]
    agg_fn  = AGG_FUNCTIONS[agg_label]

    if row_col == col_col:
        st.warning("행과 열에 같은 항목을 선택할 수 없습니다.")
        return

    # year_month 문자열 변환
    work_df = df.copy()
    if "year_month" in work_df.columns:
        work_df["year_month"] = work_df["year_month"].astype(str)

    # 피벗 생성
    try:
        pivot = work_df.pivot_table(
            index=row_col,
            columns=col_col,
            values=val_col,
            aggfunc=agg_fn,
        ).round(2)
    except Exception as e:
        st.error(f"피벗 테이블 생성 오류: {e}")
        return

    # 열 이름 정리
    if col_col == "month":
        pivot.columns = [f"{c}월" for c in pivot.columns]
    pivot.index.name = row_label
    pivot.columns.name = col_label

    # 행/열 합계 옵션
    show_total = st.checkbox("행·열 합계 추가", value=False, key="cst_pivot_total")
    if show_total:
        pivot.loc["합계"] = pivot.sum()
        pivot["합계"] = pivot.sum(axis=1)

    st.dataframe(pivot, use_container_width=True)

    # CSV 다운로드
    st.download_button(
        label="CSV 다운로드",
        data=_csv_bytes(pivot.reset_index()),
        file_name=f"pivot_{val_label}_{agg_label}.csv",
        mime="text/csv",
        key="cst_pivot_csv",
    )

    # 히트맵으로 시각화
    if st.checkbox("피벗 결과 히트맵으로 보기", value=True, key="cst_pivot_heatmap"):
        display_pivot = pivot.drop(index="합계", errors="ignore")
        display_pivot = display_pivot.drop(columns="합계", errors="ignore")
        fig = px.imshow(
            display_pivot,
            color_continuous_scale="RdBu_r",
            labels={"color": f"{val_label} ({agg_label})"},
            title=f"{val_label} 피벗 히트맵 ({agg_label})",
            aspect="auto",
            height=max(300, len(display_pivot) * 28 + 120),
        )
        st.plotly_chart(fig, use_container_width=True)
        chart_download_btn(fig, key="cst_pivot_heatmap_dl", filename="pivot_heatmap")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 서브탭 3: 데이터 패키지 내보내기
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _tab_export_package(df: pd.DataFrame) -> None:
    """
    여러 분석 결과를 Excel 다중 시트 또는 ZIP으로 묶어 내보낸다.
    사용자가 원하는 항목·기간·집계를 조합해 자신만의 데이터 패키지를 만든다.
    """
    st.markdown("#### 📦 데이터 패키지 내보내기")
    st.caption(
        "원하는 항목과 집계 방식을 선택하면 Excel 파일(다중 시트)로 묶어 드립니다. "
        "기상 전문가라면 기온·지중온도만, 농업 전문가라면 지중온도·강수량만 골라 가세요."
    )

    avail = _available_cols(df)
    if not avail:
        st.info("내보낼 데이터 컬럼이 없습니다.")
        return

    with st.container(border=True):
        st.markdown("**① 포함할 항목 선택**")
        selected_labels = st.multiselect(
            "내보낼 기상요소",
            options=list(avail.keys()),
            default=list(avail.keys())[:5],
            key="cst_export_cols",
        )

        st.markdown("**② 집계 단위 선택 (복수 가능 → 시트별 분리)**")
        agg_units = st.multiselect(
            "집계 단위",
            options=["원본 일자료", "월별 통계", "연별 통계", "계절별 통계"],
            default=["월별 통계", "연별 통계"],
            key="cst_export_units",
        )

        st.markdown("**③ 집계 함수 선택**")
        col1, col2 = st.columns(2)
        with col1:
            export_agg_labels = st.multiselect(
                "포함할 집계값",
                options=["평균", "최대", "최소", "합계", "표준편차"],
                default=["평균", "최대", "최소"],
                key="cst_export_agg",
            )
        with col2:
            include_all_stations = st.checkbox(
                "관측소별로 시트 분리",
                value=False,
                key="cst_export_by_station",
            )

    if not selected_labels or not agg_units:
        st.info("항목과 집계 단위를 하나 이상 선택하세요.")
        return

    selected_cols = [avail[l] for l in selected_labels]
    export_agg_fns = {l: AGG_FUNCTIONS[l] for l in export_agg_labels if l in AGG_FUNCTIONS}

    # Excel 생성 버튼
    if st.button("📦 Excel 패키지 생성", type="primary", key="cst_export_btn"):
        buf = io.BytesIO()

        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            stations = df["station_name"].unique() if "station_name" in df.columns else ["전체"]

            station_groups = (
                [(stn, df[df["station_name"] == stn]) for stn in stations]
                if include_all_stations
                else [("전체", df)]
            )

            for stn_name, stn_df in station_groups:
                prefix = f"{stn_name}_" if include_all_stations else ""

                for unit in agg_units:
                    sheet_name = f"{prefix}{unit}"[:31]  # Excel 시트명 31자 제한

                    if unit == "원본 일자료":
                        out = stn_df[["date"] + selected_cols].copy()
                        if "station_name" in stn_df.columns and not include_all_stations:
                            out.insert(1, "station_name", stn_df["station_name"])
                        out.to_excel(writer, sheet_name=sheet_name, index=False)

                    else:
                        # 집계 단위 → group_col 매핑
                        group_map = {
                            "월별 통계":  ["year", "month"],
                            "연별 통계":  ["year"],
                            "계절별 통계": ["year", "season"],
                        }
                        group_cols_base = group_map.get(unit, ["year"])
                        group_cols = [c for c in group_cols_base if c in stn_df.columns]

                        if not include_all_stations and "station_name" in stn_df.columns:
                            group_cols = ["station_name"] + group_cols

                        # 복수 집계함수 → 멀티인덱스 컬럼 생성
                        agg_dict = {col: list(export_agg_fns.values()) for col in selected_cols if col in stn_df.columns}
                        if not agg_dict:
                            continue

                        try:
                            agg_result = stn_df.groupby(group_cols).agg(agg_dict).round(2)
                            # 멀티인덱스 컬럼 → "컬럼_집계" 형태로 평탄화
                            fn_kr = {v: k for k, v in AGG_FUNCTIONS.items()}
                            agg_result.columns = [
                                f"{col}_{fn_kr.get(fn, fn)}"
                                for col, fn in agg_result.columns
                            ]
                            agg_result.reset_index().to_excel(writer, sheet_name=sheet_name, index=False)
                        except Exception:
                            # 단일 집계함수로 폴백
                            first_fn = list(export_agg_fns.values())[0] if export_agg_fns else "mean"
                            simple = stn_df.groupby(group_cols)[selected_cols].agg(first_fn).round(2)
                            simple.reset_index().to_excel(writer, sheet_name=sheet_name, index=False)

        buf.seek(0)
        st.success("✅ Excel 패키지 생성 완료!")
        st.download_button(
            label="⬇️ Excel 패키지 다운로드",
            data=buf.getvalue(),
            file_name="기상분석_맞춤패키지.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="cst_export_dl",
        )
        st.caption(
            f"포함된 시트: {len(agg_units) * len(station_groups)}개 | "
            f"항목: {', '.join(selected_labels[:5])}{'...' if len(selected_labels) > 5 else ''}"
        )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 서브탭 4: 설정 불러오기
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _tab_load_config(df: pd.DataFrame) -> None:
    """
    이전에 저장한 분석 설정 JSON을 불러와서
    맞춤 차트를 바로 재현한다.
    """
    st.markdown("#### 🔄 저장된 설정 불러오기")
    st.caption(
        "이전에 '맞춤 차트 빌더'에서 저장한 JSON 파일을 업로드하면 "
        "같은 분석 설정이 자동으로 적용됩니다."
    )

    uploaded = st.file_uploader(
        "설정 JSON 파일 업로드",
        type=["json"],
        key="cst_config_upload",
    )

    if uploaded is None:
        st.info("JSON 파일을 업로드하면 설정 내용이 아래에 표시됩니다.")
        return

    config = _load_config(uploaded)
    if not config:
        st.error("JSON 파일을 읽을 수 없습니다.")
        return

    st.success(f"✅ 설정 '{config.get('name', '이름 없음')}' 불러오기 완료")

    with st.expander("설정 내용 확인", expanded=True):
        st.json(config)

    avail = _available_cols(df)
    config_type = config.get("type", "chart")

    if config_type == "chart":
        # 설정값으로 바로 차트 재현
        x_label   = config.get("x_label", "날짜 (일별)")
        agg_label = config.get("agg_label", "평균")
        chart_type = config.get("chart_type", "선 그래프")
        y_labels  = [l for l in config.get("y_labels", []) if l in avail]
        sel_stn   = config.get("station", "전체 (비교)")

        if not y_labels:
            st.warning("저장된 설정의 Y축 항목이 현재 데이터에 없습니다.")
            return

        st.markdown(f"**재현 설정**: X={x_label} / Y={', '.join(y_labels)} / {agg_label} / {chart_type}")
        st.info("동일한 설정으로 차트를 재현하려면 '맞춤 차트 빌더' 탭에서 위 값을 선택하세요.")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 공개 render 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def render(df: pd.DataFrame) -> None:
    """맞춤 분석 탭 렌더링. df는 app.py에서 전달된 필터링된 일자료 DataFrame."""

    st.subheader("🔧 맞춤 분석")

    if df is None or len(df) == 0:
        st.info("먼저 데이터를 업로드하세요.")
        return

    st.caption(
        "전문 분야에 맞게 기상요소를 직접 선택하고, "
        "차트·피벗 표·Excel 패키지로 만들어 가져가세요. "
        "어떤 컬럼 조합이든 가능합니다."
    )

    # 관측소 필터 (analysis_agri.py 패턴 동일)
    if "station_name" in df.columns:
        stations = sorted(df["station_name"].unique())
        selected = st.multiselect(
            "관측소 선택", stations, default=stations,
            key="cst_stations",
        )
        if not selected:
            st.warning("관측소를 선택하세요.")
            return
        filtered = df[df["station_name"].isin(selected)].copy()
    else:
        filtered = df.copy()

    if filtered.empty:
        st.warning("선택된 관측소에 데이터가 없습니다.")
        return

    # 현재 데이터에 있는 컬럼 안내
    avail = _available_cols(filtered)
    st.info(
        f"📊 현재 데이터에서 분석 가능한 항목: **{len(avail)}개** — "
        f"{', '.join(list(avail.keys())[:8])}{'...' if len(avail) > 8 else ''}"
    )

    # 서브탭
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 맞춤 차트 빌더",
        "📋 맞춤 피벗 표",
        "📦 데이터 패키지 내보내기",
        "🔄 설정 불러오기",
    ])

    with tab1:
        _tab_chart_builder(filtered)
    with tab2:
        _tab_pivot_builder(filtered)
    with tab3:
        _tab_export_package(filtered)
    with tab4:
        _tab_load_config(filtered)
