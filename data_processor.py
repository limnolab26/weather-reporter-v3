# data_processor.py — 기상자료 처리 모듈
# 기상청 ASOS 일자료 62개 컬럼 전체 지원

import pandas as pd
import numpy as np
from typing import List

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 기상청 ASOS → 내부 표준 컬럼 매핑
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COLUMN_MAPPING = {

    # 날짜
    "일시": "date",
    "날짜": "date",

    # 관측소
    "지점명": "station_name",
    "지점": "station_code",

    # ── 기온 ─────────────────────────────
    "평균기온(°C)": "temp_avg",
    "최저기온(°C)": "temp_min",
    "최고기온(°C)": "temp_max",
    "평균 이슬점온도(°C)": "dew_point",
    "최저 초상온도(°C)": "frost_temp",

    # ── 강수 ─────────────────────────────
    "일강수량(mm)": "precipitation",
    "강수 계속시간(hr)": "precip_duration",
    "10분 최다 강수량(mm)": "precip_10min_max",
    "1시간 최다강수량(mm)": "precip_1hr_max",
    "9-9강수(mm)": "precip_9to9",

    # ── 바람 ─────────────────────────────
    "평균 풍속(m/s)": "wind_speed",
    "평균풍속(m/s)": "wind_speed",          # 구버전
    "최대 풍속(m/s)": "wind_max",
    "최대풍속(m/s)": "wind_max",            # 구버전
    "최대 풍속 풍향(16방위)": "wind_max_dir",
    "최대 순간 풍속(m/s)": "wind_gust",
    "최대 순간 풍속 풍향(16방위)": "wind_gust_dir",
    "풍정합(100m)": "wind_run_100m",
    "최다풍향(16방위)": "wind_dir",

    # ── 습도 ─────────────────────────────
    "평균 상대습도(%)": "humidity",
    "평균습도(%)": "humidity",              # 구버전
    "최소 상대습도(%)": "humidity_min",
    "평균 증기압(hPa)": "vapor_pressure",

    # ── 기압 ─────────────────────────────
    "평균 현지기압(hPa)": "pressure_local",
    "평균 해면기압(hPa)": "pressure_sea",
    "최고 해면기압(hPa)": "pressure_sea_max",
    "최저 해면기압(hPa)": "pressure_sea_min",

    # ── 일조·일사 ─────────────────────────
    "합계 일조시간(hr)": "sunshine",
    "일조시간(hr)": "sunshine",             # 구버전
    "가조시간(hr)": "daylight_hours",
    "합계 일사량(MJ/m2)": "solar_rad",
    "일사량(MJ/m2)": "solar_rad",          # 구버전
    "1시간 최다일사량(MJ/m2)": "solar_1hr_max",

    # ── 운량 ─────────────────────────────
    "평균 전운량(1/10)": "cloud_cover",
    "평균 중하층운량(1/10)": "cloud_low",

    # ── 적설 ─────────────────────────────
    "일 최심신적설(cm)": "snowfall",
    "적설(cm)": "snowfall",                # 구버전
    "일 최심적설(cm)": "snow_depth",
    "합계 3시간 신적설(cm)": "snow_3hr",

    # ── 증발 ─────────────────────────────
    "합계 대형증발량(mm)": "evaporation_large",
    "합계 소형증발량(mm)": "evaporation_small",

    # ── 지면·지중온도 ─────────────────────
    "평균 지면온도(°C)": "soil_temp_surface",
    "평균 5cm 지중온도(°C)": "soil_temp_5cm",
    "평균 10cm 지중온도(°C)": "soil_temp_10cm",
    "평균 20cm 지중온도(°C)": "soil_temp_20cm",
    "평균 30cm 지중온도(°C)": "soil_temp_30cm",
    "0.5m 지중온도(°C)": "soil_temp_50cm",
    "1.0m 지중온도(°C)": "soil_temp_100cm",
    "1.5m 지중온도(°C)": "soil_temp_150cm",
    "3.0m 지중온도(°C)": "soil_temp_300cm",
    "5.0m 지중온도(°C)": "soil_temp_500cm",

    # ── 기타 ─────────────────────────────
    "안개 계속시간(hr)": "fog_duration",
}

# 수치 변환 대상 컬럼 (내부 표준명)
NUMERIC_COLUMNS = [
    # 기온
    "temp_avg", "temp_min", "temp_max", "dew_point", "frost_temp",
    # 강수
    "precipitation", "precip_duration", "precip_10min_max",
    "precip_1hr_max", "precip_9to9",
    # 바람
    "wind_speed", "wind_max", "wind_gust", "wind_run_100m",
    "wind_dir", "wind_max_dir", "wind_gust_dir",
    # 습도·기압
    "humidity", "humidity_min", "vapor_pressure",
    "pressure_local", "pressure_sea", "pressure_sea_max", "pressure_sea_min",
    # 일조·일사
    "sunshine", "daylight_hours", "solar_rad", "solar_1hr_max",
    # 운량
    "cloud_cover", "cloud_low",
    # 적설
    "snowfall", "snow_depth", "snow_3hr",
    # 증발
    "evaporation_large", "evaporation_small",
    # 지중온도
    "soil_temp_surface", "soil_temp_5cm", "soil_temp_10cm",
    "soil_temp_20cm", "soil_temp_30cm", "soil_temp_50cm",
    "soil_temp_100cm", "soil_temp_150cm", "soil_temp_300cm", "soil_temp_500cm",
    # 기타
    "fog_duration",
]

# 하위 호환용 (app.py 일부에서 참조)
STANDARD_COLUMNS = [
    "date", "station_name",
    "temp_avg", "temp_max", "temp_min",
    "precipitation", "humidity", "wind_speed", "wind_max",
    "sunshine", "solar_rad", "snowfall",
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 계절 헬퍼
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def month_to_season(month: int) -> str:
    if month in [3, 4, 5]:
        return "봄"
    elif month in [6, 7, 8]:
        return "여름"
    elif month in [9, 10, 11]:
        return "가을"
    else:
        return "겨울"


def add_time_columns(df: pd.DataFrame) -> pd.DataFrame:
    """연·월·일·계절·연월 컬럼 생성"""
    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month
    df["day"] = df["date"].dt.day
    df["season"] = df["month"].apply(month_to_season)
    df["year_month"] = df["date"].dt.to_period("M")
    return df


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 데이터 처리 클래스
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class WeatherDataProcessor:
    """기상자료 처리 클래스 (ASOS 62컬럼 지원)"""

    def __init__(self):
        self.dataframes: List[pd.DataFrame] = []

    def process(self, df: pd.DataFrame) -> pd.DataFrame:
        """외부 DataFrame 처리 (app.py 호출용)"""
        df = self.standardize_columns(df)
        df = self.clean_data(df)
        df = add_time_columns(df)
        return df

    def load_csv(self, file) -> pd.DataFrame:
        """CSV 읽기 + 표준화"""
        try:
            df = pd.read_csv(file, encoding="cp949")
        except UnicodeDecodeError:
            df = pd.read_csv(file, encoding="utf-8")
        except Exception:
            df = pd.read_csv(file)

        df = self.standardize_columns(df)
        df = self.clean_data(df)
        df = add_time_columns(df)
        self.dataframes.append(df)
        return df

    def standardize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """기상청 컬럼명 → 내부 표준명 변환"""
        rename_dict = {
            col: COLUMN_MAPPING[col]
            for col in df.columns
            if col in COLUMN_MAPPING
        }
        df = df.rename(columns=rename_dict)

        # 중복 컬럼 제거 (구버전·신버전 동시 존재 시)
        df = df.loc[:, ~df.columns.duplicated(keep="first")]
        return df

    def clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """날짜 파싱 + 수치 변환 + 기본 정리"""

        # 날짜 변환
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

        # 수치 컬럼 변환 (존재하는 것만)
        for col in NUMERIC_COLUMNS:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        # 관측소명 기본값
        if "station_name" not in df.columns:
            df["station_name"] = "관측소"

        # 날짜 없는 행 제거 + 정렬
        df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)

        return df

    def get_combined_data(self) -> pd.DataFrame:
        if not self.dataframes:
            return pd.DataFrame()
        return pd.concat(self.dataframes, ignore_index=True)

    def filter_by_date(self, df: pd.DataFrame, start_date, end_date) -> pd.DataFrame:
        mask = (df["date"] >= start_date) & (df["date"] <= end_date)
        return df.loc[mask]

    def summary_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        numeric_cols = [
            c for c in NUMERIC_COLUMNS
            if c in df.columns
        ]
        return df[numeric_cols].describe()
