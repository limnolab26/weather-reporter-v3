# excel_generator.py — 엑셀 보고서 생성 모듈 (v5.0)

import io
from datetime import datetime

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 색상 팔레트 상수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

C = {
    'dark_blue':    '1F4E79',
    'mid_blue':     '2E75B6',
    'light_blue':   'DEEAF1',
    'white':        'FFFFFF',
    'orange':       'C55A11',
    'light_orange': 'FCE4D6',
    'green':        '375623',
    'light_green':  'E2EFDA',
    'mid_green':    '70AD47',
    'gray':         '595959',
    'light_gray':   'F2F2F2',
    'border':       'BFBFBF',
    'yellow':       'FFE699',
    'light_yellow': 'FFF2CC',
}

FONT = '맑은 고딕'
RAW_SHEET = "'📊 원본 데이터'"

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SUM/AVG 구분
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SUM_ELEMENTS = {
    'precipitation', 'sunshine', 'solar_rad', 'snowfall',
    'wind_run_100m', 'daylight_hours', 'evaporation_large', 'evaporation_small',
}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 요소 레이블 매핑
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ELEMENT_LABELS = {
    'temp_avg':        '평균기온(℃)',
    'temp_max':        '최고기온(℃)',
    'temp_min':        '최저기온(℃)',
    'dew_point':       '이슬점온도(℃)',
    'frost_temp':      '초상온도(℃)',
    'precipitation':   '강수량(mm)',
    'precip_1hr_max':  '1시간최다강수(mm)',
    'humidity':        '습도(%)',
    'humidity_min':    '최소습도(%)',
    'vapor_pressure':  '증기압(hPa)',
    'wind_speed':      '평균풍속(m/s)',
    'wind_max':        '최대풍속(m/s)',
    'wind_gust':       '최대순간풍속(m/s)',
    'wind_run_100m':   '풍정합100m(m)',
    'pressure_sea':    '해면기압(hPa)',
    'pressure_local':  '현지기압(hPa)',
    'sunshine':        '일조시간(hr)',
    'daylight_hours':  '가조시간(hr)',
    'solar_rad':       '일사량(MJ/m²)',
    'cloud_cover':     '전운량(1/10)',
    'snowfall':        '신적설(cm)',
    'snow_depth':      '적설깊이(cm)',
    'evaporation_large':  '대형증발량(mm)',
    'evaporation_small':  '소형증발량(mm)',
    'soil_temp_surface':  '지면온도(℃)',
    'soil_temp_5cm':   '5cm지중온도(℃)',
    'soil_temp_10cm':  '10cm지중온도(℃)',
    'soil_temp_20cm':  '20cm지중온도(℃)',
    'soil_temp_30cm':  '30cm지중온도(℃)',
    'soil_temp_50cm':  '0.5m지중온도(℃)',
    'soil_temp_100cm': '1.0m지중온도(℃)',
    'soil_temp_150cm': '1.5m지중온도(℃)',
    'soil_temp_300cm': '3.0m지중온도(℃)',
    'soil_temp_500cm': '5.0m지중온도(℃)',
    'fog_duration':    '안개시간(hr)',
}

# 기상 특성 검토 대상 요소
CHAR_ELEMS = [
    ('temp_avg',      '평균기온(℃)',     'AVERAGEIFS', '0.0'),
    ('temp_max',      '최고기온(℃)',     'AVERAGEIFS', '0.0'),
    ('temp_min',      '최저기온(℃)',     'AVERAGEIFS', '0.0'),
    ('precipitation', '강수량(mm)',      'SUMIFS_AVG', '#,##0.0'),
    ('humidity',      '습도(%)',         'AVERAGEIFS', '0.0'),
    ('wind_speed',    '풍속(m/s)',       'AVERAGEIFS', '0.0'),
    ('sunshine',      '일조시간(hr)',    'SUMIFS_AVG', '#,##0.0'),
    ('solar_rad',     '일사량(MJ/m²)',   'SUMIFS_AVG', '#,##0.0'),
]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 헬퍼 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _thin(color=None):
    clr = color if color else C['border']
    s = Side(style='thin', color=clr)
    return Border(left=s, right=s, top=s, bottom=s)


def _hc(ws, row, col, val, *, bg=None, fg=None,
        bold=True, sz=10, align='center', wrap=False):
    bg = bg if bg else C['dark_blue']
    fg = fg if fg else C['white']
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name=FONT, bold=bold, color=fg, size=sz)
    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    c.border = _thin()
    return c


def _dc(ws, row, col, val, *, bg=None, bold=False, sz=9,
        align='center', nf=None, wrap=False):
    bg = bg if bg else C['white']
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name=FONT, bold=bold, size=sz, color='000000')
    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    c.border = _thin()
    if nf:
        c.number_format = nf
    return c


def _title(ws, row, cs, ce, text, h=24, bg=None, sz=12):
    bg = bg if bg else C['dark_blue']
    ws.merge_cells(start_row=row, start_column=cs, end_row=row, end_column=ce)
    c = ws.cell(row=row, column=cs, value=text)
    c.font = Font(name=FONT, bold=True, size=sz, color=C['white'])
    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[row].height = h
    return c


def _note(ws, row, cs, ce, text):
    ws.merge_cells(start_row=row, start_column=cs, end_row=row, end_column=ce)
    c = ws.cell(row=row, column=cs, value=text)
    c.font = Font(name=FONT, size=8, color=C['gray'])
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[row].height = 13
    return c


def _safe_val(v):
    """NaN/inf 를 None 으로 변환"""
    if v is None:
        return None
    try:
        if np.isnan(v) or np.isinf(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ExcelReportGenerator
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class ExcelReportGenerator:
    """기상자료 Excel 보고서 생성기"""

    def __init__(self):
        self._col = {}      # 요소키 → Excel 컬럼 레터 (원본데이터 기준)
        self._years = []
        self._n_rows = 0
        self._has_stn = False

    # ──────────────────────────────────
    # 공개 인터페이스
    # ──────────────────────────────────

    def generate_excel(self, df, monthly_df=None, climate_df=None) -> bytes:
        if df is None or df.empty:
            raise ValueError("데이터가 없습니다.")

        df2 = df.copy()
        if 'date' in df2.columns:
            df2['date'] = pd.to_datetime(df2['date'])
        if 'year' not in df2.columns and 'date' in df2.columns:
            df2['year'] = df2['date'].dt.year
        if 'month' not in df2.columns and 'date' in df2.columns:
            df2['month'] = df2['date'].dt.month

        self._years = sorted(df2['year'].dropna().unique().astype(int).tolist())
        self._n_rows = len(df2)
        self._has_stn = 'station_name' in df2.columns

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        self._sheet_summary(wb, df2)
        self._sheet_raw(wb, df2)           # 반드시 두 번째 (_col 세팅)
        self._sheet_weather_chars(wb, df2)
        self._sheet_climate_change(wb, df2)
        self._sheet_precipitation(wb, df2)
        self._sheet_cumulative_precip(wb, df2)
        self._sheet_rainfall_days(wb, df2)
        self._sheet_pivot(wb, df2)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    def generate_filename(self, prefix="기상보고서") -> str:
        today = datetime.today().strftime("%Y%m%d")
        return f"{prefix}_{today}.xlsx"

    # ──────────────────────────────────
    # 1. 보고서 요약 시트
    # ──────────────────────────────────

    def _sheet_summary(self, wb, df):
        ws = wb.create_sheet("📋 보고서 요약")
        ws.sheet_view.showGridLines = False

        # 컬럼 폭
        ws.column_dimensions['A'].width = 2
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 16
        ws.column_dimensions['E'].width = 16
        ws.column_dimensions['F'].width = 16

        # 메인 타이틀
        _title(ws, 1, 2, 6, '기상자료 현황 보고서', h=36, sz=16)

        # 분석 정보
        yr_min = min(self._years) if self._years else '-'
        yr_max = max(self._years) if self._years else '-'
        today_str = datetime.today().strftime('%Y년 %m월 %d일')

        r = 3
        _hc(ws, r, 2, '분석 기간', bg=C['mid_blue'], sz=9)
        _dc(ws, r, 3, f'{yr_min}년 ~ {yr_max}년', bg=C['light_blue'], align='center')
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        _hc(ws, r, 5, '작성일', bg=C['mid_blue'], sz=9)
        _dc(ws, r, 6, today_str, bg=C['light_blue'], align='center')

        r = 5
        _title(ws, r, 2, 6, '▶ 관측소별 주요 기상요소 요약', h=20, bg=C['mid_blue'], sz=10)

        # 헤더
        r = 6
        headers = ['관측소명', '평균기온(℃)', '연강수량(mm)', '평균풍속(m/s)', '일조시간(hr)']
        for ci, h in enumerate(headers, 2):
            _hc(ws, r, ci, h, sz=9)

        # 관측소별 데이터
        stations = df['station_name'].unique() if self._has_stn else ['전체']
        r = 7
        for stn in stations:
            sdf = df[df['station_name'] == stn] if self._has_stn else df
            _dc(ws, r, 2, stn if self._has_stn else '전체', bg=C['light_blue'], bold=True)

            def _avg(key):
                if key in sdf.columns:
                    v = sdf[key].mean()
                    return round(float(v), 1) if pd.notna(v) else None
                return None

            def _sum_yr(key):
                if key in sdf.columns and 'year' in sdf.columns:
                    yr_count = sdf['year'].nunique()
                    if yr_count > 0:
                        v = sdf[key].sum() / yr_count
                        return round(float(v), 1) if pd.notna(v) else None
                return None

            _dc(ws, r, 3, _avg('temp_avg'), nf='0.0')
            _dc(ws, r, 4, _sum_yr('precipitation'), nf='#,##0.0')
            _dc(ws, r, 5, _avg('wind_speed'), nf='0.0')
            _dc(ws, r, 6, _sum_yr('sunshine'), nf='#,##0.0')
            r += 1

        # 주석
        _note(ws, r + 1, 2, 6, '※ 연강수량·일조시간은 연도 수로 나눈 연평균값입니다.')

    # ──────────────────────────────────
    # 2. 원본 데이터 시트 (_col 매핑 세팅)
    # ──────────────────────────────────

    def _sheet_raw(self, wb, df):
        ws = wb.create_sheet("📊 원본 데이터")
        ws.freeze_panes = 'A2'

        # 컬럼 순서 결정
        elem_keys = [k for k in ELEMENT_LABELS if k in df.columns]

        # 헤더 행 작성
        col = 1
        _hc(ws, 1, col, '날짜');      ws.column_dimensions[get_column_letter(col)].width = 12
        col += 1
        _hc(ws, 1, col, '연도');      ws.column_dimensions[get_column_letter(col)].width = 7
        self._col['year'] = get_column_letter(col)
        col += 1
        _hc(ws, 1, col, '월');         ws.column_dimensions[get_column_letter(col)].width = 5
        self._col['month'] = get_column_letter(col)
        col += 1

        if self._has_stn:
            _hc(ws, 1, col, '관측소명'); ws.column_dimensions[get_column_letter(col)].width = 12
            self._col['station'] = get_column_letter(col)
            col += 1

        for key in elem_keys:
            label = ELEMENT_LABELS[key]
            _hc(ws, 1, col, label, wrap=True)
            ws.column_dimensions[get_column_letter(col)].width = max(len(label) * 1.1, 10)
            self._col[key] = get_column_letter(col)
            col += 1

        # 데이터 행 작성
        for ri, (_, row) in enumerate(df.iterrows(), 2):
            c = 1
            # 날짜
            date_val = row.get('date')
            if pd.notna(date_val):
                date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
            else:
                date_str = ''
            _dc(ws, ri, c, date_str, nf='@'); c += 1

            # 연도 (수식)
            if date_str:
                ws.cell(row=ri, column=c, value=f'=YEAR(A{ri})')
            else:
                ws.cell(row=ri, column=c, value=None)
            ws.cell(row=ri, column=c).font = Font(name=FONT, size=9)
            ws.cell(row=ri, column=c).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=ri, column=c).border = _thin()
            ws.cell(row=ri, column=c).number_format = '0'
            c += 1

            # 월 (수식)
            if date_str:
                ws.cell(row=ri, column=c, value=f'=MONTH(A{ri})')
            else:
                ws.cell(row=ri, column=c, value=None)
            ws.cell(row=ri, column=c).font = Font(name=FONT, size=9)
            ws.cell(row=ri, column=c).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=ri, column=c).border = _thin()
            ws.cell(row=ri, column=c).number_format = '0'
            c += 1

            if self._has_stn:
                _dc(ws, ri, c, row.get('station_name', '')); c += 1

            for key in elem_keys:
                v = _safe_val(row.get(key))
                _dc(ws, ri, c, v, nf='#,##0.0')
                c += 1

        # Excel Table 생성
        if self._n_rows > 0:
            total_cols = col - 1
            ref = f"A1:{get_column_letter(total_cols)}{self._n_rows + 1}"
            try:
                tbl = Table(displayName="기상데이터", ref=ref)
                style = TableStyleInfo(
                    name="TableStyleMedium2",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False,
                )
                tbl.tableStyleInfo = style
                ws.add_table(tbl)
            except Exception:
                pass

        ws.row_dimensions[1].height = 30

    # ──────────────────────────────────
    # 3. 기상 특성 검토 시트
    # ──────────────────────────────────

    def _sheet_weather_chars(self, wb, df):
        ws = wb.create_sheet("🌡️ 기상 특성 검토")
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'B5'

        avail = [(k, lbl, method, nf) for k, lbl, method, nf in CHAR_ELEMS if k in self._col]
        if not avail:
            _title(ws, 1, 1, 4, '기상 특성 검토 — 데이터 없음')
            return

        n_yr = len(self._years) if self._years else 1
        mc = self._col.get('month', 'C')

        # 컬럼 폭
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 10

        r = 1
        _title(ws, r, 1, len(avail) + 2, '기상 특성 검토')

        # ── 표 1: 월별 기후 특성 ──────────────────
        r = 2
        _title(ws, r, 1, len(avail) + 2, '표 1. 월별 기후 특성 (월평균/월합산)', h=20, bg=C['mid_blue'], sz=10)

        r = 3
        _hc(ws, r, 1, '구분')
        _hc(ws, r, 2, '월')
        for ci, (k, lbl, method, nf) in enumerate(avail, 3):
            _hc(ws, r, ci, lbl, wrap=True)
            ws.column_dimensions[get_column_letter(ci)].width = 13

        month_rows = {}  # month → row number (표1 기준)
        for m in range(1, 13):
            r += 1
            month_rows[m] = r
            _dc(ws, r, 1, f'{m}월', bg=C['light_blue'], bold=True)
            _dc(ws, r, 2, m, bg=C['light_blue'])
            for ci, (k, lbl, method, nf) in enumerate(avail, 3):
                ec = self._col[k]
                if method == 'AVERAGEIFS':
                    formula = f'=ROUND(AVERAGEIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{mc}:{mc},{m}),1)'
                else:
                    formula = f'=ROUND(SUMIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{mc}:{mc},{m})/{n_yr},1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = nf

        # 연간 합계/평균 행
        r += 1
        annual_row = r
        _dc(ws, r, 1, '연간', bg=C['yellow'], bold=True)
        _dc(ws, r, 2, '-', bg=C['yellow'])
        for ci, (k, lbl, method, nf) in enumerate(avail, 3):
            start_r = month_rows[1]
            end_r = month_rows[12]
            cl = get_column_letter(ci)
            if method == 'AVERAGEIFS':
                formula = f'=ROUND(AVERAGE({cl}{start_r}:{cl}{end_r}),1)'
            else:
                formula = f'=ROUND(SUM({cl}{start_r}:{cl}{end_r}),1)'
            c = ws.cell(row=r, column=ci, value=formula)
            c.font = Font(name=FONT, size=9, bold=True)
            c.fill = PatternFill(start_color=C['yellow'], end_color=C['yellow'], fill_type='solid')
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = _thin()
            c.number_format = nf

        r += 2

        # ── 표 2: 계절별 기상 특성 ──────────────────
        _title(ws, r, 1, len(avail) + 2, '표 2. 계절별 기상 특성', h=20, bg=C['mid_blue'], sz=10)
        r += 1
        _hc(ws, r, 1, '계절')
        _hc(ws, r, 2, '월')
        for ci, (k, lbl, method, nf) in enumerate(avail, 3):
            _hc(ws, r, ci, lbl, wrap=True)

        seasons = [
            ('봄', [3, 4, 5]),
            ('여름', [6, 7, 8]),
            ('가을', [9, 10, 11]),
            ('겨울', [12, 1, 2]),
        ]
        season_bg = {
            '봄': C['light_green'], '여름': C['light_orange'],
            '가을': C['light_yellow'], '겨울': C['light_blue'],
        }

        for sname, months in seasons:
            r += 1
            bg = season_bg[sname]
            _dc(ws, r, 1, sname, bg=bg, bold=True)
            _dc(ws, r, 2, '+'.join(str(m) + '월' for m in months), bg=bg)
            for ci, (k, lbl, method, nf) in enumerate(avail, 3):
                cl = get_column_letter(ci)
                refs = ','.join(f'{cl}{month_rows[m]}' for m in months if m in month_rows)
                if method == 'AVERAGEIFS':
                    formula = f'=ROUND(AVERAGE({refs}),1)'
                else:
                    formula = f'=ROUND(SUM({refs}),1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = nf

        r += 2

        # ── 표 3: 극값 분석 ──────────────────
        _title(ws, r, 1, len(avail) + 2, '표 3. 극값 분석 (전체 기간)', h=20, bg=C['mid_blue'], sz=10)
        r += 1
        _hc(ws, r, 1, '구분')
        _hc(ws, r, 2, '항목')
        for ci, (k, lbl, method, nf) in enumerate(avail, 3):
            _hc(ws, r, ci, lbl, wrap=True)

        for stat_name, func in [('최대값', 'MAX'), ('최솟값', 'MIN')]:
            r += 1
            _dc(ws, r, 1, stat_name, bg=C['light_gray'], bold=True)
            _dc(ws, r, 2, func, bg=C['light_gray'])
            for ci, (k, lbl, method, nf) in enumerate(avail, 3):
                ec = self._col[k]
                formula = f'={func}({RAW_SHEET}!{ec}:{ec})'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=C['light_gray'], end_color=C['light_gray'], fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = nf

        # ── 차트: 월별 기온 + 강수량 ──────────────────
        r += 2
        self._add_weather_chart(ws, avail, month_rows, r)

    def _add_weather_chart(self, ws, avail, month_rows, chart_row):
        """월별 기온 꺾은선 + 강수량 막대 차트"""
        try:
            temp_cols = [ci + 3 for ci, (k, _, _, _) in enumerate(avail) if k in ('temp_avg', 'temp_max', 'temp_min')]
            precip_cols = [ci + 3 for ci, (k, _, _, _) in enumerate(avail) if k == 'precipitation']

            if not temp_cols and not precip_cols:
                return

            start_r = month_rows[1]
            end_r = month_rows[12]

            # 기온 꺾은선
            if temp_cols:
                lc = LineChart()
                lc.title = '월별 기온'
                lc.style = 10
                lc.y_axis.title = '기온(℃)'
                lc.x_axis.title = '월'
                lc.width = 15
                lc.height = 10
                for tc in temp_cols:
                    data = Reference(ws, min_col=tc, min_row=start_r - 1, max_row=end_r)
                    lc.add_data(data, titles_from_data=True)
                cats = Reference(ws, min_col=2, min_row=start_r, max_row=end_r)
                lc.set_categories(cats)
                ws.add_chart(lc, f'A{chart_row}')

            # 강수량 막대
            if precip_cols:
                bc = BarChart()
                bc.title = '월별 강수량'
                bc.style = 10
                bc.y_axis.title = '강수량(mm)'
                bc.x_axis.title = '월'
                bc.width = 15
                bc.height = 10
                for pc in precip_cols:
                    data = Reference(ws, min_col=pc, min_row=start_r - 1, max_row=end_r)
                    bc.add_data(data, titles_from_data=True)
                cats = Reference(ws, min_col=2, min_row=start_r, max_row=end_r)
                bc.set_categories(cats)
                offset_col = 10
                ws.add_chart(bc, f'{get_column_letter(offset_col)}{chart_row}')
        except Exception:
            pass

    # ──────────────────────────────────
    # 4. 기후변화 검토 시트
    # ──────────────────────────────────

    def _sheet_climate_change(self, wb, df):
        ws = wb.create_sheet("🌍 기후변화 검토")
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = 'B5'

        ws.column_dimensions['A'].width = 3
        ws.column_dimensions['B'].width = 8

        r = 1
        _title(ws, r, 1, 9, '기후변화 검토')

        yc = self._col.get('year', 'B')
        mc = self._col.get('month', 'C')

        stat_cols = [
            ('temp_avg',      '연평균기온(℃)',   'AVERAGEIFS', '0.0'),
            ('temp_max',      '최고기온(℃)',     'AVERAGEIFS', '0.0'),
            ('temp_min',      '최저기온(℃)',     'AVERAGEIFS', '0.0'),
            ('precipitation', '총강수량(mm)',    'SUMIFS',     '#,##0.0'),
            ('precipitation', '월평균강수량(mm)', 'SUMIFS_12',  '#,##0.0'),
            ('humidity',      '평균습도(%)',     'AVERAGEIFS', '0.0'),
        ]
        avail_sc = [(k, lbl, meth, nf) for k, lbl, meth, nf in stat_cols if k in self._col]

        # ── 표 1: 연도별 기상 통계 ──────────────────
        r = 2
        _title(ws, r, 1, len(avail_sc) + 2, '표 1. 연도별 기상 통계', h=20, bg=C['mid_blue'], sz=10)
        r = 3
        _hc(ws, r, 1, 'No.')
        _hc(ws, r, 2, '연도')
        for ci, (k, lbl, meth, nf) in enumerate(avail_sc, 3):
            _hc(ws, r, ci, lbl, wrap=True)
            ws.column_dimensions[get_column_letter(ci)].width = 14

        year_rows = {}
        for yi, yr in enumerate(self._years):
            r += 1
            year_rows[yr] = r
            bg = C['light_blue'] if yi % 2 == 0 else C['white']
            _dc(ws, r, 1, yi + 1, bg=bg)
            _dc(ws, r, 2, yr, bg=bg, bold=True)
            for ci, (k, lbl, meth, nf) in enumerate(avail_sc, 3):
                ec = self._col[k]
                if meth == 'AVERAGEIFS':
                    formula = f'=ROUND(AVERAGEIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{yc}:{yc},{yr}),1)'
                elif meth == 'SUMIFS':
                    formula = f'=ROUND(SUMIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{yc}:{yc},{yr}),1)'
                else:  # SUMIFS_12
                    formula = f'=ROUND(SUMIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{yc}:{yc},{yr})/12,1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = nf

        # 집계 행
        if year_rows:
            first_yr_r = year_rows[self._years[0]]
            last_yr_r = year_rows[self._years[-1]]
            for stat_name, bg, func in [
                ('전체평균', C['light_yellow'], 'AVERAGE'),
                ('최대', C['light_orange'], 'MAX'),
                ('최소', C['light_blue'], 'MIN'),
            ]:
                r += 1
                _dc(ws, r, 1, stat_name, bg=bg, bold=True)
                _dc(ws, r, 2, '-', bg=bg)
                for ci, (k, lbl, meth, nf) in enumerate(avail_sc, 3):
                    cl = get_column_letter(ci)
                    formula = f'=ROUND({func}({cl}{first_yr_r}:{cl}{last_yr_r}),1)'
                    c = ws.cell(row=r, column=ci, value=formula)
                    c.font = Font(name=FONT, size=9, bold=True)
                    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                    c.alignment = Alignment(horizontal='center', vertical='center')
                    c.border = _thin()
                    c.number_format = nf

        r += 2

        # ── 표 2: 전반기·후반기 비교 ──────────────────
        if len(self._years) >= 2:
            mid = len(self._years) // 2
            first_half = self._years[:mid]
            second_half = self._years[mid:]
            _title(ws, r, 1, len(avail_sc) + 2,
                   f'표 2. 전반기({first_half[0]}~{first_half[-1]}) · 후반기({second_half[0]}~{second_half[-1]}) 비교',
                   h=20, bg=C['mid_blue'], sz=10)
            r += 1
            _hc(ws, r, 1, '구분')
            _hc(ws, r, 2, '기간')
            for ci, (k, lbl, meth, nf) in enumerate(avail_sc, 3):
                _hc(ws, r, ci, lbl, wrap=True)

            for period_name, years_list in [
                (f'전반기\n({first_half[0]}~{first_half[-1]})', first_half),
                (f'후반기\n({second_half[0]}~{second_half[-1]})', second_half),
            ]:
                r += 1
                is_first = '전반기' in period_name
                bg = C['light_blue'] if is_first else C['light_orange']
                _dc(ws, r, 1, period_name, bg=bg, bold=True, wrap=True)
                _dc(ws, r, 2, f'{years_list[0]}~{years_list[-1]}', bg=bg)
                ws.row_dimensions[r].height = 28
                for ci, (k, lbl, meth, nf) in enumerate(avail_sc, 3):
                    yr_refs = [year_rows[yr] for yr in years_list if yr in year_rows]
                    if yr_refs:
                        cl = get_column_letter(ci)
                        refs_str = ','.join(f'{cl}{rr}' for rr in yr_refs)
                        if meth == 'AVERAGEIFS':
                            formula = f'=ROUND(AVERAGE({refs_str}),1)'
                        else:
                            formula = f'=ROUND(SUM({refs_str})/COUNT({refs_str}),1)'
                        c = ws.cell(row=r, column=ci, value=formula)
                        c.font = Font(name=FONT, size=9)
                        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                        c.alignment = Alignment(horizontal='center', vertical='center')
                        c.border = _thin()
                        c.number_format = nf

        r += 2

        # ── 표 3: 이동평균 분석 ──────────────────
        if 'temp_avg' in self._col:
            tc = self._col['temp_avg']
            _title(ws, r, 1, 6, '표 3. 이동평균 분석 (연평균기온 기준)', h=20, bg=C['mid_blue'], sz=10)
            r += 1
            _hc(ws, r, 1, 'No.')
            _hc(ws, r, 2, '연도')
            _hc(ws, r, 3, '연평균기온(℃)')
            _hc(ws, r, 4, '5년 이동평균')
            _hc(ws, r, 5, '10년 이동평균')
            ws.column_dimensions['C'].width = 14
            ws.column_dimensions['D'].width = 14
            ws.column_dimensions['E'].width = 14

            ma_rows = {}
            for i, yr in enumerate(self._years):
                r += 1
                ma_rows[yr] = r
                bg = C['light_blue'] if i % 2 == 0 else C['white']
                _dc(ws, r, 1, i + 1, bg=bg)
                _dc(ws, r, 2, yr, bg=bg, bold=True)
                formula_yr = f'=ROUND(AVERAGEIFS({RAW_SHEET}!{tc}:{tc},{RAW_SHEET}!{yc}:{yc},{yr}),1)'
                c = ws.cell(row=r, column=3, value=formula_yr)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '0.0'

                # 5년 이동평균
                if i >= 4:
                    prev5 = [ma_rows[self._years[j]] for j in range(i - 4, i + 1)]
                    refs5 = ','.join(f'C{rr}' for rr in prev5)
                    formula_5 = f'=ROUND(AVERAGE({refs5}),1)'
                else:
                    formula_5 = ''
                c5 = ws.cell(row=r, column=4, value=formula_5 if formula_5 else None)
                c5.font = Font(name=FONT, size=9)
                c5.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c5.alignment = Alignment(horizontal='center', vertical='center')
                c5.border = _thin()
                c5.number_format = '0.0'

                # 10년 이동평균
                if i >= 9:
                    prev10 = [ma_rows[self._years[j]] for j in range(i - 9, i + 1)]
                    refs10 = ','.join(f'C{rr}' for rr in prev10)
                    formula_10 = f'=ROUND(AVERAGE({refs10}),1)'
                else:
                    formula_10 = ''
                c10 = ws.cell(row=r, column=5, value=formula_10 if formula_10 else None)
                c10.font = Font(name=FONT, size=9)
                c10.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c10.alignment = Alignment(horizontal='center', vertical='center')
                c10.border = _thin()
                c10.number_format = '0.0'

            # 차트: 연평균기온 추이
            r += 2
            try:
                first_r = ma_rows[self._years[0]]
                last_r = ma_rows[self._years[-1]]
                lc = LineChart()
                lc.title = '연평균기온 추이 및 이동평균'
                lc.style = 10
                lc.y_axis.title = '기온(℃)'
                lc.x_axis.title = '연도'
                lc.width = 20
                lc.height = 12
                for col_n, col_title in [(3, '연평균기온'), (4, '5년이동평균'), (5, '10년이동평균')]:
                    data = Reference(ws, min_col=col_n, min_row=first_r - 1, max_row=last_r)
                    series = lc.series[len(lc.series)] if False else None
                    lc.add_data(data, titles_from_data=False)
                    lc.series[-1].title.value = col_title
                cats = Reference(ws, min_col=2, min_row=first_r, max_row=last_r)
                lc.set_categories(cats)
                ws.add_chart(lc, f'A{r}')
            except Exception:
                pass

    # ──────────────────────────────────
    # 5. 강수량 분석 시트
    # ──────────────────────────────────

    def _sheet_precipitation(self, wb, df):
        ws = wb.create_sheet("🌧️ 강수량 분석")
        ws.sheet_view.showGridLines = False

        if 'precipitation' not in self._col:
            _title(ws, 1, 1, 4, '강수량 분석 — 강수량 데이터 없음')
            return

        pc = self._col['precipitation']
        yc = self._col.get('year', 'B')
        mc = self._col.get('month', 'C')

        ws.column_dimensions['A'].width = 14

        r = 1
        _title(ws, r, 1, len(self._years) + 3, '강수량 분석')

        # ── 표 1: 연도별 기상개황 ──────────────────
        r = 2
        n_yr_cols = len(self._years)
        total_width = n_yr_cols + 3
        _title(ws, r, 1, total_width, '표 1. 연도별 기상개황', h=20, bg=C['mid_blue'], sz=10)
        r = 3
        _hc(ws, r, 1, '항목')
        for ci, yr in enumerate(self._years, 2):
            _hc(ws, r, ci, str(yr))
            ws.column_dimensions[get_column_letter(ci)].width = 10
        avg_col = n_yr_cols + 2
        _hc(ws, r, avg_col, '평균', bg=C['orange'], fg=C['white'])
        ws.column_dimensions[get_column_letter(avg_col)].width = 10

        temp_col = self._col.get('temp_avg')
        overview = [
            ('연평균기온(℃)', temp_col, 'AVERAGEIFS', '0.0'),
            ('총강수량(mm)',  pc, 'SUMIFS', '#,##0.0'),
            ('월평균강수량(mm)', pc, 'SUMIFS_12', '#,##0.0'),
        ]
        overview_rows = {}
        for item_name, ec, meth, nf in overview:
            r += 1
            _dc(ws, r, 1, item_name, bold=True, align='left')
            yr_cell_cols = []
            for ci, yr in enumerate(self._years, 2):
                if ec is None:
                    _dc(ws, r, ci, None)
                    continue
                if meth == 'AVERAGEIFS':
                    formula = f'=ROUND(AVERAGEIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{yc}:{yc},{yr}),1)'
                elif meth == 'SUMIFS':
                    formula = f'=ROUND(SUMIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{yc}:{yc},{yr}),1)'
                else:
                    formula = f'=ROUND(SUMIFS({RAW_SHEET}!{ec}:{ec},{RAW_SHEET}!{yc}:{yc},{yr})/12,1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = nf
                yr_cell_cols.append(get_column_letter(ci))
            # 평균 열
            if yr_cell_cols:
                refs = ','.join(f'{cl}{r}' for cl in yr_cell_cols)
                avg_f = f'=ROUND(AVERAGE({refs}),1)'
                c = ws.cell(row=r, column=avg_col, value=avg_f)
                c.font = Font(name=FONT, size=9, bold=True)
                c.fill = PatternFill(start_color=C['light_orange'], end_color=C['light_orange'], fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = nf
            overview_rows[item_name] = r

        r += 2

        # ── 표 2: 연도별 월별 강수량 ──────────────────
        _title(ws, r, 1, total_width, '표 2. 연도별 월별 강수량 (mm)', h=20, bg=C['mid_blue'], sz=10)
        r += 1
        _hc(ws, r, 1, '월')
        for ci, yr in enumerate(self._years, 2):
            _hc(ws, r, ci, str(yr))
        _hc(ws, r, avg_col, '평균', bg=C['orange'], fg=C['white'])

        month_precip_rows = {}
        for m in range(1, 13):
            r += 1
            month_precip_rows[m] = r
            bg = C['light_blue'] if m % 2 == 0 else C['white']
            _dc(ws, r, 1, f'{m}월', bg=bg, bold=True)
            yr_cells = []
            for ci, yr in enumerate(self._years, 2):
                formula = f'=ROUND(SUMIFS({RAW_SHEET}!{pc}:{pc},{RAW_SHEET}!{yc}:{yc},{yr},{RAW_SHEET}!{mc}:{mc},{m}),1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'
                yr_cells.append(get_column_letter(ci))
            if yr_cells:
                refs = ','.join(f'{cl}{r}' for cl in yr_cells)
                c = ws.cell(row=r, column=avg_col, value=f'=ROUND(AVERAGE({refs}),1)')
                c.font = Font(name=FONT, size=9, bold=True)
                c.fill = PatternFill(start_color=C['light_orange'], end_color=C['light_orange'], fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'

        # 하단 집계
        first_m_r = month_precip_rows[1]
        last_m_r = month_precip_rows[12]
        for stat_name, bg, func in [
            ('평균', C['light_yellow'], 'AVERAGE'),
            ('최대', C['light_orange'], 'MAX'),
            ('최소', C['light_blue'], 'MIN'),
        ]:
            r += 1
            _dc(ws, r, 1, stat_name, bg=bg, bold=True)
            for ci in range(2, avg_col + 1):
                cl = get_column_letter(ci)
                formula = f'=ROUND({func}({cl}{first_m_r}:{cl}{last_m_r}),1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9, bold=True)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'

        r += 2

        # ── 차트 ──────────────────
        try:
            # 연도별 연강수량 막대
            bc = BarChart()
            bc.title = '연도별 총 강수량'
            bc.style = 10
            bc.y_axis.title = '강수량(mm)'
            bc.x_axis.title = '연도'
            bc.width = 18
            bc.height = 10
            total_row = overview_rows.get('총강수량(mm)')
            if total_row:
                data = Reference(ws, min_col=2, max_col=n_yr_cols + 1, min_row=total_row)
                bc.add_data(data)
                cats = Reference(ws, min_col=2, max_col=n_yr_cols + 1, min_row=3)
                bc.set_categories(cats)
                ws.add_chart(bc, f'A{r}')

            # 월별 평균 강수량 막대
            bc2 = BarChart()
            bc2.title = '월별 평균 강수량'
            bc2.style = 10
            bc2.y_axis.title = '강수량(mm)'
            bc2.x_axis.title = '월'
            bc2.width = 18
            bc2.height = 10
            data2 = Reference(ws, min_col=avg_col, min_row=first_m_r - 1, max_row=last_m_r)
            bc2.add_data(data2, titles_from_data=True)
            cats2 = Reference(ws, min_col=1, min_row=first_m_r, max_row=last_m_r)
            bc2.set_categories(cats2)
            ws.add_chart(bc2, f'{get_column_letter(12)}{r}')
        except Exception:
            pass

    # ──────────────────────────────────
    # 6. 누적강수량 분석 시트
    # ──────────────────────────────────

    def _sheet_cumulative_precip(self, wb, df):
        ws = wb.create_sheet("누적강수량 분석")
        ws.freeze_panes = 'B2'

        if 'precipitation' not in self._col:
            _title(ws, 1, 1, 4, '누적강수량 분석 — 강수량 데이터 없음')
            return

        pc = self._col['precipitation']
        yc = self._col.get('year', 'B')
        mc = self._col.get('month', 'C')

        n_yr = len(self._years)
        # 이동평균 기준 연도 목록
        avg10_yrs = self._years[-10:] if len(self._years) >= 10 else self._years
        avg20_yrs = self._years[-20:] if len(self._years) >= 20 else self._years
        avg30_yrs = self._years[-30:] if len(self._years) >= 30 else self._years

        # 컬럼 설정: 항목, 연도들, 10년평균, 20년평균, 30년평균
        ws.column_dimensions['A'].width = 14
        for ci in range(2, n_yr + 5):
            ws.column_dimensions[get_column_letter(ci)].width = 9

        yr_start_col = 2
        avg10_col = n_yr + 2
        avg20_col = n_yr + 3
        avg30_col = n_yr + 4

        def write_header_row(r, section_title):
            _hc(ws, r, 1, section_title, bg=C['mid_blue'])
            for ci, yr in enumerate(self._years, yr_start_col):
                _hc(ws, r, ci, str(yr), sz=8)
            _hc(ws, r, avg10_col, f'10년평균\n({avg10_yrs[0]}~)', sz=8, wrap=True, bg=C['orange'])
            _hc(ws, r, avg20_col, f'20년평균\n({avg20_yrs[0]}~)', sz=8, wrap=True, bg=C['orange'])
            _hc(ws, r, avg30_col, f'30년평균\n({avg30_yrs[0]}~)', sz=8, wrap=True, bg=C['orange'])
            ws.row_dimensions[r].height = 28

        # 여름 조합 정의
        summer_combos = [
            ('6+7월',   [6, 7]),
            ('6+7+8월', [6, 7, 8]),
            ('7+8월',   [7, 8]),
            ('7+8+9월', [7, 8, 9]),
        ]

        # ── 구간①: 연도별 월별 강수량 (절대값) ──────────────────
        r = 1
        write_header_row(r, '월별 강수량(mm)')

        monthly_cell = {}  # (yr_idx, month) → cell address

        for m in range(1, 13):
            r += 1
            bg = C['light_blue'] if m % 2 == 0 else C['white']
            _dc(ws, r, 1, f'{m}월', bg=bg, bold=True)
            yr_cells = []
            for ci, yr in enumerate(self._years, yr_start_col):
                formula = f'=ROUND(SUMIFS({RAW_SHEET}!{pc}:{pc},{RAW_SHEET}!{yc}:{yc},{yr},{RAW_SHEET}!{mc}:{mc},{m}),1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'
                monthly_cell[(ci - yr_start_col, m)] = (get_column_letter(ci), r)
                yr_cells.append((get_column_letter(ci), r))

            # 이동평균 열
            def _avg_formula(yrs_list):
                idxs = [i for i, yr in enumerate(self._years) if yr in yrs_list]
                if not idxs:
                    return ''
                refs = ','.join(f'{get_column_letter(idx + yr_start_col)}{r}' for idx in idxs)
                return f'=ROUND(AVERAGE({refs}),1)'

            for avg_col, avg_yrs in [(avg10_col, avg10_yrs), (avg20_col, avg20_yrs), (avg30_col, avg30_yrs)]:
                formula = _avg_formula(avg_yrs)
                c = ws.cell(row=r, column=avg_col, value=formula if formula else None)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=C['light_orange'], end_color=C['light_orange'], fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'

        # 합계 행
        month_rows_sec1 = {m: r - (12 - m) for m in range(1, 13)}
        # 재계산: 각 월 row 기억
        month_rows_abs = {}
        row_ptr = 2
        for m in range(1, 13):
            month_rows_abs[m] = row_ptr
            row_ptr += 1

        r += 1
        _dc(ws, r, 1, '연합계', bg=C['yellow'], bold=True)
        for ci in range(yr_start_col, avg30_col + 1):
            cl = get_column_letter(ci)
            refs = ','.join(f'{cl}{month_rows_abs[m]}' for m in range(1, 13))
            formula = f'=ROUND(SUM({refs}),1)'
            c = ws.cell(row=r, column=ci, value=formula)
            c.font = Font(name=FONT, size=9, bold=True)
            c.fill = PatternFill(start_color=C['yellow'], end_color=C['yellow'], fill_type='solid')
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = _thin()
            c.number_format = '#,##0.0'
        annual_abs_row = r

        # 여름 조합 행
        for combo_name, months in summer_combos:
            r += 1
            _dc(ws, r, 1, combo_name, bg=C['light_green'], bold=True)
            for ci in range(yr_start_col, avg30_col + 1):
                cl = get_column_letter(ci)
                refs = ','.join(f'{cl}{month_rows_abs[m]}' for m in months)
                formula = f'=ROUND(SUM({refs}),1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=C['light_green'], end_color=C['light_green'], fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'

        r += 2

        # ── 구간②: 월별 누적강수량 ──────────────────
        cum_start_row = r
        write_header_row(r, '누적강수량(mm)')

        cum_rows = {}
        for m in range(1, 13):
            r += 1
            cum_rows[m] = r
            bg = C['light_blue'] if m % 2 == 0 else C['white']
            _dc(ws, r, 1, f'1~{m}월', bg=bg, bold=True)
            for ci in range(yr_start_col, avg30_col + 1):
                cl = get_column_letter(ci)
                refs = ','.join(f'{cl}{month_rows_abs[mm]}' for mm in range(1, m + 1))
                formula = f'=ROUND(SUM({refs}),1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'

        # 여름 조합 누적
        for combo_name, months in summer_combos:
            r += 1
            _dc(ws, r, 1, f'누적 {combo_name}', bg=C['light_green'], bold=True)
            for ci in range(yr_start_col, avg30_col + 1):
                cl = get_column_letter(ci)
                refs = ','.join(f'{cl}{month_rows_abs[m]}' for m in months)
                formula = f'=ROUND(SUM({refs}),1)'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=C['light_green'], end_color=C['light_green'], fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'

        r += 2

        # ── 구간③: 연간 대비 누적강수량 비율 ──────────────────
        write_header_row(r, '누적비율(%)')
        ratio_start = r

        for m in range(1, 13):
            r += 1
            bg = C['light_blue'] if m % 2 == 0 else C['white']
            _dc(ws, r, 1, f'1~{m}월 비율', bg=bg, bold=True)
            for ci in range(yr_start_col, avg30_col + 1):
                cl = get_column_letter(ci)
                cum_r = cum_rows[m]
                ann_r = annual_abs_row
                formula = f'=IFERROR(ROUND({cl}{cum_r}/{cl}{ann_r}*100,1),"")'
                c = ws.cell(row=r, column=ci, value=formula)
                c.font = Font(name=FONT, size=9)
                c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '0.0'

    # ──────────────────────────────────
    # 7. 강우일수 분석 시트
    # ──────────────────────────────────

    def _sheet_rainfall_days(self, wb, df):
        ws = wb.create_sheet("강우일수 분석")
        ws.freeze_panes = 'A2'

        if 'precipitation' not in df.columns:
            _title(ws, 1, 1, 4, '강우일수 분석 — 강수량 데이터 없음')
            return

        col_headers = [
            '연도', '강우일수\n(>0mm)', '무강우일수\n(=0mm)', '총일수',
            '< 3mm', '3~10mm', '10~30mm', '30~80mm', '≥ 80mm',
            '최장연속\n강우(일)', '최장연속\n무강우(일)',
        ]
        n_cols = len(col_headers)

        stations = df['station_name'].unique() if self._has_stn else ['전체']
        r = 1

        for stn in stations:
            sdf = df[df['station_name'] == stn].copy() if self._has_stn else df.copy()

            if 'year' not in sdf.columns:
                continue

            # 섹션 타이틀
            title_text = f'강우일수 분석 — {stn}' if self._has_stn else '강우일수 분석'
            _title(ws, r, 1, n_cols, title_text, h=22, bg=C['dark_blue'], sz=11)
            r += 1

            # 헤더
            for ci, h in enumerate(col_headers, 1):
                _hc(ws, r, ci, h, sz=9, wrap=True)
                ws.column_dimensions[get_column_letter(ci)].width = 10
            ws.row_dimensions[r].height = 28
            r += 1

            records = []
            for yr in sorted(sdf['year'].dropna().unique().astype(int)):
                ydf = sdf[sdf['year'] == yr].copy()
                prec = ydf['precipitation'].fillna(0)

                rain_days = int((prec > 0).sum())
                no_rain_days = int((prec == 0).sum())
                total_days = len(prec)

                lt3 = int(((prec > 0) & (prec < 3)).sum())
                r3_10 = int(((prec >= 3) & (prec < 10)).sum())
                r10_30 = int(((prec >= 10) & (prec < 30)).sum())
                r30_80 = int(((prec >= 30) & (prec < 80)).sum())
                ge80 = int((prec >= 80).sum())

                # 최장 연속 강우
                max_consec_rain = 0
                cur = 0
                for v in prec:
                    if v > 0:
                        cur += 1
                        max_consec_rain = max(max_consec_rain, cur)
                    else:
                        cur = 0

                # 최장 연속 무강우
                max_consec_dry = 0
                cur = 0
                for v in prec:
                    if v == 0:
                        cur += 1
                        max_consec_dry = max(max_consec_dry, cur)
                    else:
                        cur = 0

                records.append([
                    yr, rain_days, no_rain_days, total_days,
                    lt3, r3_10, r10_30, r30_80, ge80,
                    max_consec_rain, max_consec_dry,
                ])

            data_start_row = r
            for i, rec in enumerate(records):
                bg = C['light_blue'] if i % 2 == 0 else C['white']
                for ci, val in enumerate(rec, 1):
                    _dc(ws, r, ci, val, bg=bg)
                r += 1
            data_end_row = r - 1

            # 집계 행
            if records:
                for stat_name, bg, func in [
                    ('평균', C['light_yellow'], 'AVERAGE'),
                    ('최대', C['light_orange'], 'MAX'),
                    ('최소', C['light_blue'], 'MIN'),
                ]:
                    _dc(ws, r, 1, stat_name, bg=bg, bold=True)
                    for ci in range(2, n_cols + 1):
                        cl = get_column_letter(ci)
                        formula = f'=ROUND({func}({cl}{data_start_row}:{cl}{data_end_row}),1)'
                        c = ws.cell(row=r, column=ci, value=formula)
                        c.font = Font(name=FONT, size=9, bold=True)
                        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                        c.alignment = Alignment(horizontal='center', vertical='center')
                        c.border = _thin()
                        c.number_format = '0.0'
                    r += 1

            r += 2

    # ──────────────────────────────────
    # 8. 피벗 분석 시트
    # ──────────────────────────────────

    def _sheet_pivot(self, wb, df):
        ws = wb.create_sheet("🔄 피벗 분석")
        ws.sheet_view.showGridLines = True
        ws.freeze_panes = 'C3'

        yc = self._col.get('year', 'B')
        mc = self._col.get('month', 'C')

        pivot_elems = [(k, lbl) for k, lbl in ELEMENT_LABELS.items() if k in self._col]
        if not pivot_elems:
            _title(ws, 1, 1, 4, '피벗 분석 — 데이터 없음')
            return

        # 컬럼 폭
        ws.column_dimensions['A'].width = 18
        ws.column_dimensions['B'].width = 8
        for ci in range(3, 16):
            ws.column_dimensions[get_column_letter(ci)].width = 9

        r = 1
        for key, elem_lbl in pivot_elems:
            ec = self._col[key]
            is_sum = key in SUM_ELEMENTS

            # 요소 타이틀
            _title(ws, r, 1, 15, f'▶ {elem_lbl}  ({"월합산" if is_sum else "월평균"})', h=20, bg=C['dark_blue'], sz=10)
            r += 1

            # 헤더: 연도, 구분, 1~12월, 연간
            _hc(ws, r, 1, '연도', sz=9)
            _hc(ws, r, 2, '구분', sz=9)
            for m in range(1, 13):
                _hc(ws, r, m + 2, f'{m}월', sz=9)
            _hc(ws, r, 15, '연간', bg=C['orange'], fg=C['white'], sz=9)
            r += 1

            year_block_rows = {}
            for yr in self._years:
                year_block_rows[yr] = r
                bg = C['light_blue']
                _dc(ws, r, 1, yr, bg=bg, bold=True)
                _dc(ws, r, 2, '값', bg=bg)
                month_cells = []
                for m in range(1, 13):
                    if is_sum:
                        formula = (f'=IFERROR(ROUND(SUMIFS({RAW_SHEET}!{ec}:{ec},'
                                   f'{RAW_SHEET}!{yc}:{yc},{yr},'
                                   f'{RAW_SHEET}!{mc}:{mc},{m}),1),"")')
                    else:
                        formula = (f'=IFERROR(ROUND(AVERAGEIFS({RAW_SHEET}!{ec}:{ec},'
                                   f'{RAW_SHEET}!{yc}:{yc},{yr},'
                                   f'{RAW_SHEET}!{mc}:{mc},{m}),1),"")')
                    c = ws.cell(row=r, column=m + 2, value=formula)
                    c.font = Font(name=FONT, size=9)
                    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                    c.alignment = Alignment(horizontal='center', vertical='center')
                    c.border = _thin()
                    c.number_format = '#,##0.0'
                    month_cells.append(f'{get_column_letter(m + 2)}{r}')

                # 연간 열
                refs = ','.join(month_cells)
                if is_sum:
                    annual_f = f'=IFERROR(ROUND(SUM({refs}),1),"")'
                else:
                    annual_f = f'=IFERROR(ROUND(AVERAGE({refs}),1),"")'
                c = ws.cell(row=r, column=15, value=annual_f)
                c.font = Font(name=FONT, size=9, bold=True)
                c.fill = PatternFill(start_color=C['light_orange'], end_color=C['light_orange'], fill_type='solid')
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = _thin()
                c.number_format = '#,##0.0'
                r += 1

            # 집계 행 (전체평균, 최대, 최소)
            if year_block_rows:
                first_yr_r = year_block_rows[self._years[0]]
                last_yr_r = year_block_rows[self._years[-1]]
                for stat_name, bg, func in [
                    ('전체평균', C['light_yellow'], 'AVERAGE'),
                    ('최대', C['light_orange'], 'MAX'),
                    ('최소', C['light_blue'], 'MIN'),
                ]:
                    _dc(ws, r, 1, stat_name, bg=bg, bold=True)
                    _dc(ws, r, 2, func, bg=bg)
                    for ci in range(3, 16):
                        cl = get_column_letter(ci)
                        formula = f'=IFERROR(ROUND({func}({cl}{first_yr_r}:{cl}{last_yr_r}),1),"")'
                        c = ws.cell(row=r, column=ci, value=formula)
                        c.font = Font(name=FONT, size=9, bold=True)
                        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                        c.alignment = Alignment(horizontal='center', vertical='center')
                        c.border = _thin()
                        c.number_format = '#,##0.0'
                    r += 1

            r += 2
