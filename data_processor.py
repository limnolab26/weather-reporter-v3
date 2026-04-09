# data_processor.py — 기상자료 처리 모듈
# 역할:
# 1. 기상청 ASOS CSV 읽기
# 2. 컬럼 표준화
# 3. 날짜/연도/월/계절 컬럼 생성
# 4. 통계 및 집계 지원

import pandas as pd
import numpy as np
from typing import List, Dict

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 표준 컬럼 매핑 (기상청 ASOS → 내부 표준)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COLUMN_MAPPING = {

    # 날짜
    "일시": "date",
    "날짜": "date",

    # 관측소 (지점=관측소번호, 지점명=관측소명)
    "지점명": "station_name",
    "지점": "station_code",

    # 기온
    "평균기온(°C)": "temp_avg",
    "최고기온(°C)": "temp_max",
    "최저기온(°C)": "temp_min",

    # 강수
    "일강수량(mm)": "precipitation",

    # 습도 (구버전: 평균습도, 신버전: 평균 상대습도)
    "평균습도(%)": "humidity",
    "평균 상대습도(%)": "humidity",

    # 풍속 (구버전: 공백 없음, 신버전: 공백 있음)
    "평균풍속(m/s)": "wind_speed",
    "평균 풍속(m/s)": "wind_speed",
    "최대풍속(m/s)": "wind_max",
    "최대 풍속(m/s)": "wind_max",

    # 일조 (구버전: 일조시간, 신버전: 합계 일조시간)
    "일조시간(hr)": "sunshine",
    "합계 일조시간(hr)": "sunshine",

    # 일사 (구버전: 일사량, 신버전: 합계 일사량)
    "일사량(MJ/m2)": "solar_rad",
    "합계 일사량(MJ/m2)": "solar_rad",

    # 적설 (구버전: 적설, 신버전: 일 최심신적설)
    "적설(cm)": "snowfall",
    "일 최심신적설(cm)": "snowfall",

}


STANDARD_COLUMNS = [
    "date",
    "station_name",
    "temp_avg",
    "temp_max",
    "temp_min",
    "precipitation",
    "humidity",
    "wind_speed",
    "wind_max",
    "sunshine",
    "solar_rad",
    "snowfall"
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 계절 생성 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def month_to_season(month: int) -> str:
    """월 → 계절 변환"""

    if month in [3, 4, 5]:
        return "봄"
    elif month in [6, 7, 8]:
        return "여름"
    elif month in [9, 10, 11]:
        return "가을"
    else:
        return "겨울"


def add_time_columns(df: pd.DataFrame) -> pd.DataFrame:
    """연·월·일·계절 컬럼 생성"""

    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month
    df["day"] = df["date"].dt.day

    df["season"] = df["month"].apply(month_to_season)

    df["year_month"] = df["date"].dt.to_period("M")

    return df


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CSV 로딩 및 표준화
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class WeatherDataProcessor:
    """
    기상자료 처리 클래스
    """

    def __init__(self):
        self.dataframes: List[pd.DataFrame] = []

    def process(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        외부 DataFrame 입력 처리 (app.py용)
        """

        df = self.standardize_columns(df)

        df = self.clean_data(df)

        df = add_time_columns(df)

        return df

    # ─────────────────────────────
    # CSV 로딩
    # ─────────────────────────────

    def load_csv(self, file) -> pd.DataFrame:
        """
        CSV 읽기 및 표준화
        """

        try:
            df = pd.read_csv(
                file,
                encoding="cp949"
            )

        except UnicodeDecodeError:

            df = pd.read_csv(
                file,
                encoding="utf-8"
            )

        except:
            df = pd.read_csv(file)
        
        df = self.standardize_columns(df)

        df = self.clean_data(df)

        df = add_time_columns(df)

        self.dataframes.append(df)

        return df

    # ─────────────────────────────
    # 컬럼 표준화
    # ─────────────────────────────

    def standardize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        컬럼명을 내부 표준으로 변환
        """

        rename_dict = {}

        for col in df.columns:

            if col in COLUMN_MAPPING:
                rename_dict[col] = COLUMN_MAPPING[col]

        df = df.rename(columns=rename_dict)

        return df

    # ─────────────────────────────
    # 데이터 정리
    # ─────────────────────────────

    def clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        날짜 변환 및 결측 처리
        """

        # 날짜 변환
        df["date"] = pd.to_datetime(
            df["date"],
            errors="coerce"
        )

        # 숫자 컬럼 변환
        numeric_cols = [
            col for col in STANDARD_COLUMNS
            if col not in ["date", "station_name"]
            and col in df.columns
        ]

        for col in numeric_cols:

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            )

        # 관측소명 없으면 기본값
        if "station_name" not in df.columns:
            df["station_name"] = "관측소"

        # 날짜 없는 행 제거
        df = df.dropna(subset=["date"])

        df = df.sort_values("date")

        return df

    # ─────────────────────────────
    # 전체 데이터 결합
    # ─────────────────────────────

    def get_combined_data(self) -> pd.DataFrame:
        """
        여러 CSV 결합
        """

        if not self.dataframes:
            return pd.DataFrame()

        df = pd.concat(
            self.dataframes,
            ignore_index=True
        )

        return df

    # ─────────────────────────────
    # 기간 필터
    # ─────────────────────────────

    def filter_by_date(
        self,
        df: pd.DataFrame,
        start_date,
        end_date
    ) -> pd.DataFrame:
        """
        날짜 범위 필터
        """

        mask = (
            (df["date"] >= start_date)
            &
            (df["date"] <= end_date)
        )

        return df.loc[mask]

    # ─────────────────────────────
    # 집계 함수
    # ─────────────────────────────

    def aggregate_data(
        self,
        df: pd.DataFrame,
        freq: str = "D",
        element: str = "temp_avg"
    ) -> pd.DataFrame:
        """
        시계열 집계

        freq:
        D  = 일별
        ME = 월별
        YE = 연별
        """

        if df.empty:
            return df

        grouped = (
            df
            .set_index("date")
            .groupby("station_name")[element]
            .resample(freq)
            .mean()
            .reset_index()
        )

        return grouped

    # ─────────────────────────────
    # 피벗용 데이터 생성
    # ─────────────────────────────

    def create_pivot_table(
        self,
        df: pd.DataFrame,
        rows: str,
        cols: str,
        values: str,
        aggfunc: str = "mean"
    ) -> pd.DataFrame:
        """
        동적 피벗테이블 생성
        """

        pivot = pd.pivot_table(
            df,
            index=rows,
            columns=cols,
            values=values,
            aggfunc=aggfunc
        )

        return pivot

    # ─────────────────────────────
    # 기본 통계 요약
    # ─────────────────────────────

    def summary_statistics(
        self,
        df: pd.DataFrame
    ) -> pd.DataFrame:
        """
        주요 기상요소 통계
        """

        numeric_cols = [
            col for col in STANDARD_COLUMNS
            if col in df.columns
            and col not in ["date"]
        ]

        summary = df[numeric_cols].describe()

        return summary


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 외부 유틸 함수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def prepare_chart_data(
    df: pd.DataFrame,
    element: str,
    freq: str
) -> pd.DataFrame:
    """
    Plotly 차트용 데이터 준비

    반환:
    index=날짜
    columns=관측소
    values=집계값
    """

    # freq 안정화
    freq_map = {
        "D": "D",
        "ME": "M",
        "YE": "Y"
        }

    freq = freq_map.get(freq, "D")

    pivot = (
        df
        .set_index("date")
        .groupby("station_name")[element]
        .resample(freq)
        .mean()
        .unstack(0)
    )

    return pivot
