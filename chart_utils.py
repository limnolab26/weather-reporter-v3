"""
chart_utils.py — Plotly 차트 데이터 CSV 추출 유틸리티

모든 분석 모듈에서 공용으로 사용합니다.
st.plotly_chart() 직후에 chart_download_btn()을 호출하면
차트 데이터를 CSV로 다운로드할 수 있는 버튼이 표시됩니다.
"""
from __future__ import annotations
from typing import Optional
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


def chart_download_btn(fig, key: str, filename: str = "chart_data") -> None:
    """Plotly figure 데이터를 CSV로 다운로드하는 버튼을 표시합니다.

    Args:
        fig:      st.plotly_chart()에 넘긴 것과 동일한 Plotly figure 객체
        key:      Streamlit 위젯 고유 키 (앱 전체에서 중복 불가)
        filename: 다운로드 파일명 (.csv 확장자 자동 추가)
    """
    import streamlit as st
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
