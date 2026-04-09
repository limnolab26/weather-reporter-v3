# excel_generator.py — 엑셀 보고서 생성 모듈

import io
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 컬럼명 한글 매핑
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COLUMN_KR = {
    "date": "날짜",
    "station_name": "관측소명",
    "station_code": "지점번호",
    "temp_avg": "평균기온(°C)",
    "temp_max": "최고기온(°C)",
    "temp_min": "최저기온(°C)",
    "precipitation": "강수량(mm)",
    "humidity": "습도(%)",
    "wind_speed": "평균풍속(m/s)",
    "wind_max": "최대풍속(m/s)",
    "sunshine": "일조시간(hr)",
    "solar_rad": "일사량(MJ/m²)",
    "snowfall": "적설(cm)",
    "year": "연도",
    "month": "월",
    "day": "일",
    "season": "계절",
}

# 원본데이터 시트에 출력할 컬럼 순서 (있는 것만 선택됨)
RAW_COL_ORDER = [
    "date", "station_name", "year", "month", "day", "season",
    "temp_avg", "temp_max", "temp_min",
    "precipitation", "humidity", "wind_speed", "wind_max",
    "sunshine", "solar_rad", "snowfall",
]

# 월별통계 집계 방법
MONTHLY_AGG = {
    "temp_avg": "mean",
    "temp_max": "mean",
    "temp_min": "mean",
    "precipitation": "sum",
    "humidity": "mean",
    "wind_speed": "mean",
    "wind_max": "mean",
    "sunshine": "sum",
    "solar_rad": "sum",
    "snowfall": "sum",
}

# 요약통계 대상 컬럼
SUMMARY_COLS = [
    "temp_avg", "temp_max", "temp_min",
    "precipitation", "humidity", "wind_speed", "wind_max",
    "sunshine", "solar_rad", "snowfall",
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 스타일 헬퍼
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

HEADER_FILL = PatternFill(
    start_color="2F5496",
    end_color="2F5496",
    fill_type="solid"
)

SUBHEADER_FILL = PatternFill(
    start_color="D6E0F5",
    end_color="D6E0F5",
    fill_type="solid"
)

HEADER_FONT = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
SUBHEADER_FONT = Font(bold=True, name="맑은 고딕", size=10)
BODY_FONT = Font(name="맑은 고딕", size=9)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT = Alignment(horizontal="left", vertical="center")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _style_header_row(ws, row_idx: int, n_cols: int):
    """헤더 행에 진한 파란색 배경 + 흰색 볼드체 적용"""
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def _style_body_rows(ws, start_row: int, end_row: int, n_cols: int):
    """데이터 행 기본 스타일"""
    for row_idx in range(start_row, end_row + 1):
        for col_idx in range(1, n_cols + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = CENTER


def _auto_col_width(ws, min_width: int = 8, max_width: int = 25):
    """열 너비 자동 조정"""
    for column_cells in ws.columns:
        length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in column_cells
        )
        col_letter = get_column_letter(column_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, min_width), max_width)


def _freeze_header(ws):
    ws.freeze_panes = ws["A2"]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 엑셀 생성 클래스
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class ExcelReportGenerator:
    """기상자료 Excel 보고서 생성기"""

    def __init__(self):
        pass

    # ─────────────────────────────
    # 메인 Excel 생성 함수
    # ─────────────────────────────

    def generate_excel(
        self,
        df: pd.DataFrame,
        pivot_df: pd.DataFrame = None,
        summary_df: pd.DataFrame = None
    ) -> bytes:

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            # 1️⃣ 원본 데이터 (한글 컬럼명)
            raw_df = self._prepare_raw_sheet(df)
            raw_df.to_excel(writer, sheet_name="원본데이터", index=False)

            # 2️⃣ 월별 통계 (관측소 × 연도 × 월)
            monthly_df = self._build_monthly_stats(df)
            if monthly_df is not None:
                monthly_df.to_excel(writer, sheet_name="월별통계", index=False)

            # 3️⃣ 연별 통계 (관측소 × 연도)
            annual_df = self._build_annual_stats(df)
            if annual_df is not None:
                annual_df.to_excel(writer, sheet_name="연별통계", index=False)

            # 4️⃣ 요약통계
            summ_df = summary_df if summary_df is not None else self.generate_summary_table(df)
            summ_df.to_excel(writer, sheet_name="요약통계")

        # 스타일 적용
        output.seek(0)
        styled = self._apply_styles(output)
        return styled

    # ─────────────────────────────
    # 시트 데이터 준비
    # ─────────────────────────────

    def _prepare_raw_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """원본 데이터 시트용: 컬럼 순서 정리 + 한글 헤더"""
        cols = [c for c in RAW_COL_ORDER if c in df.columns]
        out = df[cols].copy()
        out = out.rename(columns=COLUMN_KR)
        # 숫자 컬럼 소수점 2자리 반올림
        for col in out.columns:
            if out[col].dtype in [np.float64, np.float32]:
                out[col] = out[col].round(2)
        return out

    def _build_monthly_stats(self, df: pd.DataFrame) -> pd.DataFrame | None:
        """관측소별 월별 통계 시트"""
        agg = {col: func for col, func in MONTHLY_AGG.items() if col in df.columns}
        if not agg:
            return None

        group_cols = ["station_name", "year", "month"]
        if not all(c in df.columns for c in group_cols):
            return None

        monthly = (
            df
            .groupby(group_cols)
            .agg(agg)
            .reset_index()
            .sort_values(["station_name", "year", "month"])
        )
        monthly = monthly.rename(columns={**{"station_name": "관측소명", "year": "연도", "month": "월"}, **COLUMN_KR})
        for col in monthly.select_dtypes(include=[np.float64, np.float32]).columns:
            monthly[col] = monthly[col].round(2)
        return monthly

    def _build_annual_stats(self, df: pd.DataFrame) -> pd.DataFrame | None:
        """관측소별 연별 통계 시트"""
        annual_agg = {col: func for col, func in MONTHLY_AGG.items() if col in df.columns}
        if not annual_agg:
            return None

        group_cols = ["station_name", "year"]
        if not all(c in df.columns for c in group_cols):
            return None

        annual = (
            df
            .groupby(group_cols)
            .agg(annual_agg)
            .reset_index()
            .sort_values(["station_name", "year"])
        )
        annual = annual.rename(columns={**{"station_name": "관측소명", "year": "연도"}, **COLUMN_KR})
        for col in annual.select_dtypes(include=[np.float64, np.float32]).columns:
            annual[col] = annual[col].round(2)
        return annual

    # ─────────────────────────────
    # 요약 통계
    # ─────────────────────────────

    def generate_summary_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """관측소별 주요 기상요소 요약"""
        cols = [c for c in SUMMARY_COLS if c in df.columns]
        records = []

        stations = df["station_name"].unique() if "station_name" in df.columns else ["전체"]

        for stn in stations:
            stn_df = df[df["station_name"] == stn] if "station_name" in df.columns else df
            for col in cols:
                kr_name = COLUMN_KR.get(col, col)
                records.append({
                    "관측소": stn,
                    "기상요소": kr_name,
                    "평균": round(stn_df[col].mean(), 2),
                    "최대": round(stn_df[col].max(), 2),
                    "최소": round(stn_df[col].min(), 2),
                    "합계": round(stn_df[col].sum(), 2),
                    "표준편차": round(stn_df[col].std(), 2),
                    "결측수": int(stn_df[col].isna().sum()),
                })

        return pd.DataFrame(records)

    # ─────────────────────────────
    # 스타일 적용
    # ─────────────────────────────

    def _apply_styles(self, excel_bytes: io.BytesIO) -> bytes:

        wb = load_workbook(excel_bytes)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            max_col = ws.max_column
            max_row = ws.max_row

            _style_header_row(ws, 1, max_col)
            _style_body_rows(ws, 2, max_row, max_col)
            _auto_col_width(ws)
            _freeze_header(ws)

            # 숫자 열 오른쪽 정렬
            for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal="right", vertical="center")

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    # ─────────────────────────────
    # 파일명 생성
    # ─────────────────────────────

    def generate_filename(self, prefix: str = "기상보고서") -> str:
        today = datetime.today().strftime("%Y%m%d")
        return f"{prefix}_{today}.xlsx"
