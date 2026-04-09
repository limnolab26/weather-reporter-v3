# pdf_generator.py — PDF 보고서 생성 모듈
# 역할:
# 1. 기상 데이터 PDF 보고서 생성
# 2. 표 기반 데이터 출력
# 3. 요약 통계 출력
# 4. 다중 페이지 자동 처리
# 5. 한글 폰트 지원

import pandas as pd
import numpy as np
import io
from datetime import datetime

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak
)

from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 한글 폰트 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def register_korean_font():
    """
    한글 폰트 자동 등록

    우선순위:
    1. NanumGothic.ttf
    2. Malgun Gothic (Windows)
    3. 기본 폰트
    """

    font_paths = [
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",  # Streamlit Cloud (Linux)
        "NanumGothic.ttf",
        "malgun.ttf",
    ]

    for font_path in font_paths:

        try:

            pdfmetrics.registerFont(
                TTFont("KoreanFont", font_path)
            )

            return "KoreanFont"

        except:

            continue

    return "Helvetica"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PDF 생성 클래스
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class PDFReportGenerator:
    """
    기상자료 PDF 보고서 생성기
    """

    def __init__(self):

        self.font_name = register_korean_font()

        self.styles = getSampleStyleSheet()

    # ─────────────────────────────
    # 메인 PDF 생성 함수
    # ─────────────────────────────

    def generate_pdf(
        self,
        df: pd.DataFrame,
        pivot_df: pd.DataFrame = None,
        summary_df: pd.DataFrame = None
    ) -> bytes:
        """
        PDF 생성

        Returns
        -------
        bytes
            PDF binary
        """

        buffer = io.BytesIO()

        doc = SimpleDocTemplate(
            buffer,
            pagesize=(595, 842),  # A4
            leftMargin=30,
            rightMargin=30,
            topMargin=30,
            bottomMargin=30
        )

        elements = []

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 제목
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

        title_text = "기상자료 보고서"

        elements.append(
            Paragraph(
                title_text,
                self.styles["Title"]
            )
        )

        elements.append(Spacer(1, 12))

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 생성 날짜
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

        today = datetime.today().strftime(
            "%Y-%m-%d"
        )

        elements.append(
            Paragraph(
                f"생성일: {today}",
                self.styles["Normal"]
            )
        )

        elements.append(Spacer(1, 12))

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 요약 통계
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

        if summary_df is not None:

            elements.append(
                Paragraph(
                    "요약 통계",
                    self.styles["Heading2"]
                )
            )

            elements.append(Spacer(1, 6))

            summary_table = self.create_table_from_dataframe(
                summary_df
            )

            elements.append(summary_table)

            elements.append(PageBreak())

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 피벗 테이블
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

        if pivot_df is not None:

            elements.append(
                Paragraph(
                    "피벗 테이블",
                    self.styles["Heading2"]
                )
            )

            elements.append(Spacer(1, 6))

            pivot_table = self.create_table_from_dataframe(
                pivot_df
            )

            elements.append(pivot_table)

            elements.append(PageBreak())

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 원본 데이터
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━

        elements.append(
            Paragraph(
                "원본 데이터",
                self.styles["Heading2"]
            )
        )

        elements.append(Spacer(1, 6))

        data_table = self.create_table_from_dataframe(
            df
        )

        elements.append(data_table)

        # PDF 빌드
        doc.build(elements)

        pdf_bytes = buffer.getvalue()

        buffer.close()

        return pdf_bytes

    # ─────────────────────────────
    # DataFrame → Table 변환
    # ─────────────────────────────

    def create_table_from_dataframe(
        self,
        df: pd.DataFrame,
        max_rows: int = 40
    ):
        """
        DataFrame을 PDF Table로 변환
        """

        df = df.copy()

        # 너무 긴 경우 분할
        df = df.head(max_rows)

        data = []

        # 헤더 추가
        data.append(
            [str(col) for col in df.columns]
        )

        # 데이터 추가
        for _, row in df.iterrows():

            data.append(
                [
                    self.format_cell(value)
                    for value in row
                ]
            )

        table = Table(
            data,
            repeatRows=1
        )

        table.setStyle(
            TableStyle([

                ("BACKGROUND",
                 (0, 0),
                 (-1, 0),
                 colors.lightgrey),

                ("GRID",
                 (0, 0),
                 (-1, -1),
                 0.5,
                 colors.grey),

                ("FONTNAME",
                 (0, 0),
                 (-1, -1),
                 self.font_name),

                ("FONTSIZE",
                 (0, 0),
                 (-1, -1),
                 8),

                ("ALIGN",
                 (0, 0),
                 (-1, -1),
                 "CENTER")

            ])
        )

        return table

    # ─────────────────────────────
    # 셀 포맷
    # ─────────────────────────────

    def format_cell(
        self,
        value
    ):
        """
        값 포맷 정리
        """

        if pd.isna(value):

            return "-"

        if isinstance(value, float):

            return f"{value:.2f}"

        return str(value)

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

        filename = f"{prefix}_{today}.pdf"

        return filename