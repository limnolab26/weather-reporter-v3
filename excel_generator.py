# excel_generator.py — 엑셀 보고서 생성 모듈
# 역할:
# 1. 기상 데이터 Excel 보고서 생성
# 2. 요약통계 시트 생성
# 3. 피벗테이블 시트 생성
# 4. 자동 서식 적용

import pandas as pd
import numpy as np
import io
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 엑셀 생성 클래스
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class ExcelReportGenerator:
    """
    기상자료 Excel 보고서 생성기
    """

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
        """
        Excel 파일 생성

        Parameters
        ----------
        df : 원본 데이터
        pivot_df : 피벗테이블 (선택)
        summary_df : 통계요약 (선택)

        Returns
        -------
        bytes
            Excel binary
        """

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 1️⃣ 원본 데이터 시트
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

            df.to_excel(
                writer,
                sheet_name="원본데이터",
                index=False
            )

            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 2️⃣ 피벗 시트 (선택)
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

            if pivot_df is not None:

                pivot_df.to_excel(
                    writer,
                    sheet_name="피벗테이블"
                )

            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 3️⃣ 통계요약 시트 (선택)
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

            if summary_df is not None:

                summary_df.to_excel(
                    writer,
                    sheet_name="통계요약"
                )

        # 스타일 적용
        output.seek(0)

        styled_output = self.apply_excel_style(output)

        return styled_output

    # ─────────────────────────────
    # Excel 스타일 적용
    # ─────────────────────────────

    def apply_excel_style(
        self,
        excel_bytes: io.BytesIO
    ) -> bytes:
        """
        Excel 스타일 자동 적용
        """

        wb = load_workbook(excel_bytes)

        for sheet_name in wb.sheetnames:

            ws = wb[sheet_name]

            # 헤더 스타일
            header_font = Font(
                bold=True
            )

            center_align = Alignment(
                horizontal="center",
                vertical="center"
            )

            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 헤더 서식
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

            for cell in ws[1]:

                cell.font = header_font
                cell.alignment = center_align

            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 열 너비 자동 조정
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

            for column_cells in ws.columns:

                length = max(
                    len(str(cell.value)) if cell.value else 0
                    for cell in column_cells
                )

                adjusted_width = min(
                    length + 2,
                    40
                )

                col_letter = get_column_letter(
                    column_cells[0].column
                )

                ws.column_dimensions[
                    col_letter
                ].width = adjusted_width

        # 저장
        output = io.BytesIO()

        wb.save(output)

        output.seek(0)

        return output.getvalue()

    # ─────────────────────────────
    # 요약 보고서 생성 (자동 계산 포함)
    # ─────────────────────────────

    def generate_summary_table(
        self,
        df: pd.DataFrame
    ) -> pd.DataFrame:
        """
        주요 기상요소 요약 생성
        """

        elements = [
            "temp_avg",
            "temp_max",
            "temp_min",
            "precipitation",
            "humidity",
            "wind_speed"
        ]

        summary_data = {}

        for col in elements:

            if col in df.columns:

                summary_data[col] = {

                    "평균": df[col].mean(),
                    "최대": df[col].max(),
                    "최소": df[col].min(),
                    "합계": df[col].sum(),
                    "표준편차": df[col].std()

                }

        summary_df = pd.DataFrame(
            summary_data
        )

        return summary_df

    # ─────────────────────────────
    # 파일명 생성
    # ─────────────────────────────

    def generate_filename(
        self,
        prefix: str = "기상보고서"
    ) -> str:
        """
        다운로드 파일명 생성
        """

        today = datetime.today().strftime(
            "%Y%m%d"
        )

        filename = f"{prefix}_{today}.xlsx"

        return filename