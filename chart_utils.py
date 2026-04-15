"""
chart_utils.py — Plotly 차트 데이터 XLSX 추출 유틸리티

모든 분석 모듈에서 공용으로 사용합니다.
st.plotly_chart() 직후에 chart_download_btn()을 호출하면
차트 데이터와 차트 이미지가 포함된 XLSX 파일을 다운로드할 수 있는 버튼이 표시됩니다.
"""
from __future__ import annotations
from typing import Optional, Tuple
import io
import pandas as pd


def fig_to_df(fig) -> Optional[pd.DataFrame]:
    """Plotly figure의 모든 trace 데이터를 DataFrame으로 변환합니다.

    지원 차트 타입:
    - 선/막대/산점도: x, y 컬럼 + 계열명
    - 히트맵: Y축, X축, 값 컬럼
    - 극좌표(바람장미): 방향, 빈도/값 컬럼
    - 박스플롯: 값 컬럼 + 계열명
    - 히스토그램: 값 컬럼 + 계열명
    """
    try:
        rows = []
        for trace in fig.data:
            name = getattr(trace, 'name', '') or ''
            trace_type = getattr(trace, 'type', '') or ''

            # ── 히트맵 ───────────────────────────────────────
            if trace_type == 'heatmap' and getattr(trace, 'z', None) is not None:
                z = trace.z
                x_lbls = list(trace.x) if trace.x is not None else list(range(len(z[0]) if z else 0))
                y_lbls = list(trace.y) if trace.y is not None else list(range(len(z)))
                for yi, row_vals in enumerate(z):
                    for xi, val in enumerate(row_vals):
                        rows.append({'Y축': y_lbls[yi], 'X축': x_lbls[xi], '값': val, '계열': name})

            # ── 극좌표 (바람장미) ────────────────────────────
            elif getattr(trace, 'r', None) is not None and getattr(trace, 'theta', None) is not None:
                for r_val, t_val in zip(trace.r, trace.theta):
                    rows.append({'방향': t_val, '빈도/값': r_val, '계열': name})

            # ── 박스플롯 ─────────────────────────────────────
            elif trace_type == 'box':
                y_vals = getattr(trace, 'y', None)
                x_vals = getattr(trace, 'x', None)
                if y_vals is not None:
                    for i, yv in enumerate(y_vals):
                        row = {'값': yv, '계열': name}
                        if x_vals is not None:
                            row['X'] = x_vals[i] if i < len(x_vals) else ''
                        rows.append(row)

            # ── 히스토그램 ───────────────────────────────────
            elif trace_type == 'histogram':
                x_vals = getattr(trace, 'x', None)
                if x_vals is not None:
                    for xv in x_vals:
                        rows.append({'값': xv, '계열': name})

            # ── 선/막대/산점도 (x·y 공통) ────────────────────
            elif getattr(trace, 'x', None) is not None and getattr(trace, 'y', None) is not None:
                for xv, yv in zip(trace.x, trace.y):
                    rows.append({'X': xv, 'Y': yv, '계열': name})

        return pd.DataFrame(rows) if rows else None
    except Exception:
        return None


def _build_xlsx(fig, filename: str) -> Tuple[Optional[bytes], Optional[str]]:
    """Plotly figure를 데이터 시트 + 차트 이미지 시트가 포함된 XLSX bytes로 변환합니다.

    Returns:
        (xlsx_bytes, chart_error_msg)
        - xlsx_bytes: XLSX 파일 bytes (항상 생성, 실패 시 None)
        - chart_error_msg: 차트 이미지 생성 실패 메시지 (성공 시 None)
    """
    chart_error: Optional[str] = None
    try:
        import openpyxl
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()

        # ── 데이터 시트 ──────────────────────────────────────
        ws_data = wb.active
        ws_data.title = "데이터"

        df = fig_to_df(fig)
        if df is not None and not df.empty:
            # 헤더
            for ci, col in enumerate(df.columns, 1):
                cell = ws_data.cell(row=1, column=ci, value=col)
                cell.font = openpyxl.styles.Font(bold=True)
            # 데이터 행
            for ri, row_tuple in enumerate(df.itertuples(index=False), 2):
                for ci, val in enumerate(row_tuple, 1):
                    ws_data.cell(row=ri, column=ci, value=val)
            # 열 너비 자동 조정
            for ci, col in enumerate(df.columns, 1):
                max_len = max(len(str(col)), df[col].astype(str).str.len().max())
                ws_data.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 40)

        # ── 차트 이미지 시트 ─────────────────────────────────
        try:
            import plotly.io as pio
            img_bytes = pio.to_image(fig, format='png', width=1200, height=600, scale=1.5, engine='kaleido')
            ws_chart = wb.create_sheet("차트")
            img = XLImage(io.BytesIO(img_bytes))
            img.anchor = 'A1'
            ws_chart.add_image(img)
        except Exception as e:
            chart_error = f"{type(e).__name__}: {e}"

        out = io.BytesIO()
        wb.save(out)
        return out.getvalue(), chart_error

    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


def chart_download_btn(fig, key: str, filename: str = "chart_data") -> None:
    """Plotly figure 데이터를 차트 이미지가 포함된 XLSX로 다운로드하는 버튼을 표시합니다.

    - 데이터 시트: trace 데이터 (모든 차트 타입 지원)
    - 차트 시트: kaleido로 렌더링한 PNG 이미지 (kaleido 미설치 시 생략)
    - 완전 실패 시 CSV로 폴백

    Args:
        fig:      st.plotly_chart()에 넘긴 것과 동일한 Plotly figure 객체
        key:      Streamlit 위젯 고유 키 (앱 전체에서 중복 불가)
        filename: 다운로드 파일명 (.xlsx 확장자 자동 추가)
    """
    import streamlit as st

    xlsx_bytes, chart_error = _build_xlsx(fig, filename)

    if xlsx_bytes is not None:
        st.download_button(
            "⬇️ 차트 XLSX 다운로드",
            xlsx_bytes,
            f"{filename}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=key,
        )
        if chart_error:
            st.caption(f"⚠️ 차트 이미지 생성 실패 (데이터 시트만 포함): {chart_error}")
    else:
        # 폴백: CSV
        df = fig_to_df(fig)
        if df is not None and not df.empty:
            csv = df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                "⬇️ 차트 데이터 CSV",
                csv,
                f"{filename}.csv",
                "text/csv",
                key=key,
            )
        if chart_error:
            st.caption(f"⚠️ XLSX 생성 실패: {chart_error}")
