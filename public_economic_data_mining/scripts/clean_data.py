#!/usr/bin/env python3
"""Normalize raw public power datasets into tidy CSV files."""

from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUT_DIR = ROOT / "data_clean"
NA_VALUES = ["-", "--", "—", "x", "X", "…", ".."]
ROC_OFFSET = 1911
ROC_MONTH_RE = re.compile(r"(?P<year>\d+)\s*年\s*(?P<month>\d+)\s*月")
OWNER_MAP = {
    "全國": "nation_total",
    "台電": "taipower",
    "民營電廠": "ipp",
    "發電業": "generation_companies",
    "自用發電設備": "captive_generators",
}


def ensure_out_dir() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)


def parse_roc_year(value: object) -> int | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("年", "")
    match = re.search(r"(\d+)", text)
    if not match:
        return None
    year = int(match.group(1))
    if 1911 <= year <= 3000:
        return year
    return year + ROC_OFFSET


def parse_roc_month(value: object) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT
    text = str(value).strip()
    if not text:
        return pd.NaT
    match = ROC_MONTH_RE.search(text)
    if not match:
        return pd.NaT
    year = int(match.group("year"))
    month = int(match.group("month"))
    if year < 1911:
        year += ROC_OFFSET
    return pd.Timestamp(dt.date(year, month, 1))


def parse_compact_date(value: object) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT
    text = str(value).strip()
    if not text:
        return pd.NaT
    return pd.to_datetime(text, format="%Y%m%d", errors="coerce")


def parse_month_number(value: object) -> float | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def convert_to_numeric(
    series: pd.Series, *, multiplier: float = 1.0, allow_int: bool = False
) -> pd.Series:
    values = pd.to_numeric(series, errors="coerce")
    if multiplier != 1.0:
        values = values * multiplier
    if allow_int:
        return values.round().astype("Int64")
    return values


def write_csv(df: pd.DataFrame, filename: str) -> None:
    target = OUT_DIR / filename
    df.to_csv(target, index=False)


def clean_sales_by_usage() -> None:
    df = pd.read_csv(
        DATA_DIR / "售電量(用途別).csv", encoding="utf-8-sig", na_values=NA_VALUES
    )
    df = df.rename(
        columns={
            "年度": "year",
            "用途別": "usage_category",
            "售電量": "sales_kwh",
        }
    )
    df["year"] = convert_to_numeric(df["year"], allow_int=True)
    df["usage_category"] = df["usage_category"].str.strip()
    df["sales_kwh"] = convert_to_numeric(df["sales_kwh"], allow_int=True)
    df = df.dropna(subset=["year", "usage_category", "sales_kwh"])
    write_csv(df, "electricity_sales_by_usage.csv")


def clean_generation_by_energy() -> None:
    df = pd.read_csv(
        DATA_DIR / "發購電量(能源別).csv", encoding="utf-8-sig", na_values=NA_VALUES
    )
    df = df.rename(
        columns={
            "年度": "year",
            "發購電": "producer_type",
            "能源別": "energy_type",
            "發購電量": "generation_kwh",
        }
    )
    df["year"] = convert_to_numeric(df["year"], allow_int=True)
    producer_map = {"台電": "taipower", "購電": "purchased"}
    df["producer_type"] = df["producer_type"].map(producer_map).fillna(
        df["producer_type"]
    )
    df["energy_type"] = df["energy_type"].str.strip()
    df["generation_kwh"] = convert_to_numeric(df["generation_kwh"], allow_int=True)
    df = df.dropna(subset=["year", "producer_type", "energy_type", "generation_kwh"])
    write_csv(df, "generation_by_energy_type.csv")


def clean_consumption_by_sector() -> None:
    df = pd.read_excel(DATA_DIR / "電力消費.xlsx", header=4)
    df = df.dropna(subset=["日期"])
    df["period"] = df["日期"].apply(parse_roc_month)
    df = df.dropna(subset=["period"])
    df["year"] = df["period"].dt.year
    df["month"] = df["period"].dt.month
    column_map = {
        "日期": "raw_period",
        "全國用電量": "total_mwh",
        "工業用電": "industrial_mwh",
        "紡織成衣及服飾業": "textile_apparel_mwh",
        "紙漿、紙及紙製品業": "pulp_paper_mwh",
        "化學材料製造業": "chemical_materials_mwh",
        "塑膠製品製造業": "plastic_products_mwh",
        "非金屬礦物製品製造業": "nonmetal_minerals_mwh",
        "金屬基本工業": "basic_metals_mwh",
        "金屬製品製造業": "fabricated_metals_mwh",
        "電腦通信及視聽電子產品製造業": "electronics_mwh",
        "其他": "industrial_other_mwh",
        "住商用電（含住宅及服務業）": "residential_commercial_mwh",
        "住宅用電": "residential_mwh",
        "服務業用電": "service_sector_mwh",
        "批發及零售業": "wholesale_retail_mwh",
        "住宿及餐飲業": "accommodation_food_mwh",
        "運輸服務業": "transport_services_mwh",
        "倉儲業": "warehousing_mwh",
        "通信業": "telecom_mwh",
        "金融保險及不動產業": "finance_real_estate_mwh",
        "工商服務業": "business_services_mwh",
        "社會服務及個人服務業": "social_personal_services_mwh",
        "公共行政業": "public_admin_mwh",
        "其他.1": "service_other_mwh",
        "運輸用電": "transport_mwh",
        "公路": "road_mwh",
        "鐵路": "rail_mwh",
        "管線運輸": "pipeline_mwh",
        "其他.2": "transport_other_mwh",
        "能源部門自用": "energy_sector_mwh",
        "農業部門": "agriculture_mwh",
    }
    df = df.rename(columns=column_map)
    energy_columns = [col for col in column_map.values() if col.endswith("_mwh")]
    for col in energy_columns:
        df[col] = convert_to_numeric(df[col])
    ordered_cols = [
        "period",
        "year",
        "month",
        "raw_period",
    ] + [col for col in column_map.values() if col not in {"raw_period"}]
    df = df[ordered_cols]
    write_csv(df, "electricity_consumption_by_sector.csv")


def clean_power_supply_and_units() -> None:
    df = pd.read_csv(
        DATA_DIR / "電力供需資訊.csv", encoding="utf-8-sig", na_values=NA_VALUES
    )
    df["date"] = df["日期"].apply(parse_compact_date)
    summary_mapping = {
        "淨尖峰供電能力(萬瓩)": ("net_peak_capacity_mw", 10.0),
        "尖峰負載(萬瓩)": ("peak_load_mw", 10.0),
        "備轉容量(萬瓩)": ("reserve_capacity_mw", 10.0),
        "備轉容量率(%)": ("reserve_margin_pct", 1.0),
        "工業用電(百萬度)": ("industrial_consumption_gwh", 1.0),
        "民生用電(百萬度)": ("residential_consumption_gwh", 1.0),
    }
    summary = df[["date"]].copy()
    for src, (dest, factor) in summary_mapping.items():
        summary[dest] = convert_to_numeric(df[src], multiplier=factor)
    summary = summary.dropna(subset=["date"])
    write_csv(summary, "power_supply_demand_summary.csv")
    plant_columns = [
        col for col in df.columns if col not in set(summary_mapping) | {"日期", "date"}
    ]
    plant_df = df[["date"] + plant_columns].copy()
    plant_df = plant_df.dropna(subset=["date"])
    melted = plant_df.melt(id_vars="date", var_name="unit_raw", value_name="value")
    melted["unit_name"] = (
        melted["unit_raw"].str.replace(r"\(萬瓩\)", "", regex=True).str.strip()
    )
    melted["available_capacity_mw"] = convert_to_numeric(melted["value"], multiplier=10)
    melted = melted.drop(columns=["unit_raw", "value"])
    melted = melted.dropna(subset=["available_capacity_mw"])
    write_csv(melted, "power_unit_available_capacity.csv")


def clean_solar_feed_in() -> None:
    df = pd.read_csv(
        DATA_DIR / "太陽光電購電實績.csv", encoding="utf-8-sig", na_values=NA_VALUES
    )
    df = df.rename(columns={"年度": "roc_year", "月份": "month", "度數(千度)": "energy_mwh"})
    df["year"] = df["roc_year"].apply(parse_roc_year)
    df["month"] = df["month"].apply(parse_month_number)
    df["is_annual_total"] = df["month"].isna()
    df["month"] = df["month"].fillna(1).astype(int)
    df["period"] = pd.to_datetime(
        dict(year=df["year"], month=df["month"], day=1), errors="coerce"
    )
    df["energy_mwh"] = convert_to_numeric(df["energy_mwh"])
    df = df.dropna(subset=["period", "energy_mwh"])
    write_csv(df, "solar_feed_in_records.csv")


def clean_retail_statistics() -> None:
    df = pd.read_csv(
        DATA_DIR / "公用售電業售電統計資料.csv",
        encoding="utf-8-sig",
        na_values=NA_VALUES,
    )
    rename_map = {
        "年度": "year",
        "電燈售電量(度)": "lighting_sales_kwh",
        "電力售電量(度)": "power_sales_kwh",
        "售電量合計(度)": "total_sales_kwh",
        "電燈用戶數(戶)": "lighting_customers",
        "電力用戶數(戶)": "power_customers",
        "用戶數合計(戶)": "total_customers",
        "電燈(非營業用)售電量(度)": "lighting_nonbusiness_kwh",
        "電燈(營業用)售電量(度)": "lighting_business_kwh",
        "電燈(非營業用)用戶數(戶)": "lighting_nonbusiness_customers",
        "電燈(營業用)用戶數(戶)": "lighting_business_customers",
        "電燈平均單價(元)": "lighting_avg_price_ntd",
        "電力平均單價(元)": "power_avg_price_ntd",
        "平均單價合計(元)": "overall_avg_price_ntd",
    }
    df = df.rename(columns=rename_map)
    df["year"] = convert_to_numeric(df["year"], allow_int=True)
    kwh_columns = [col for col in rename_map.values() if col.endswith("_kwh")]
    for col in kwh_columns:
        df[col] = convert_to_numeric(df[col], allow_int=True)
    df["total_customers"] = convert_to_numeric(df["total_customers"], allow_int=True)
    df["lighting_customers"] = convert_to_numeric(df["lighting_customers"], allow_int=True)
    df["power_customers"] = convert_to_numeric(df["power_customers"], allow_int=True)
    df["lighting_nonbusiness_customers"] = convert_to_numeric(
        df["lighting_nonbusiness_customers"], allow_int=True
    )
    df["lighting_business_customers"] = convert_to_numeric(
        df["lighting_business_customers"], allow_int=True
    )
    price_cols = [
        "lighting_avg_price_ntd",
        "power_avg_price_ntd",
        "overall_avg_price_ntd",
    ]
    for col in price_cols:
        df[col] = convert_to_numeric(df[col])
    write_csv(df, "utility_retail_statistics.csv")


def clean_generation_costs() -> None:
    df = pd.read_csv(
        DATA_DIR / "各種發電方式之發電成本.csv", encoding="utf-8-sig", na_values=NA_VALUES
    )
    year_columns = {
        "111年\n審定決算(元/度)": 2022,
        "112年 審定決算(元/度)": 2023,
        "113年 審定決算(元/度)": 2024,
    }
    tidy_rows = []
    level_zero = None
    level_one = None
    for _, row in df.iterrows():
        raw_label = str(row["各種發電方式之發電成本"])
        indent = len(raw_label) - len(raw_label.lstrip())
        label = raw_label.strip()
        if not label:
            continue
        if indent == 0:
            level_zero = label
            level_one = None
        elif indent <= 4:
            level_one = label
        for col, year in year_columns.items():
            value = row[col]
            if pd.isna(value):
                continue
            tidy_rows.append(
                {
                    "category": level_zero,
                    "subcategory": level_one,
                    "item": label,
                    "calendar_year": year,
                    "cost_ntd_per_kwh": float(value),
                }
            )
    tidy_df = pd.DataFrame(tidy_rows)
    write_csv(tidy_df, "generation_costs_by_technology.csv")


def clean_nuclear_performance() -> None:
    df = pd.read_csv(
        DATA_DIR / "核能發電績效及減碳效益.csv", encoding="utf-8-sig", na_values=NA_VALUES
    )
    df = df.rename(
        columns={
            "年度": "roc_year",
            "核能發電量(億度)": "generation_100m_kwh",
            "容量因數平均值(%)": "capacity_factor_pct",
            "減碳量(萬噸)": "co2_reduction_10k_tonnes",
        }
    )
    df["year"] = df["roc_year"].apply(parse_roc_year)
    df["generation_gwh"] = convert_to_numeric(df["generation_100m_kwh"]) * 100
    df["co2_reduction_tonnes"] = convert_to_numeric(
        df["co2_reduction_10k_tonnes"]
    ) * 10000
    df["capacity_factor_pct"] = convert_to_numeric(df["capacity_factor_pct"])
    df = df.drop(columns=["generation_100m_kwh", "co2_reduction_10k_tonnes"])
    write_csv(df, "nuclear_generation_performance.csv")


def clean_radiation_monitoring() -> None:
    df = pd.read_csv(
        DATA_DIR / "核設施輻射監測即時資料.csv", encoding="utf-8-sig", na_values=NA_VALUES
    )
    df = df.rename(
        columns={
            "站名": "station_name",
            "站號": "station_id",
            "劑量率(微西弗/小時)": "dose_rate_uSv_per_hr",
            "日期時間": "timestamp_raw",
            "經度": "longitude",
            "緯度": "latitude",
        }
    )
    df["timestamp"] = pd.to_datetime(
        df["timestamp_raw"], format="%Y%m%dT%H%M%S", errors="coerce"
    )
    df["dose_rate_uSv_per_hr"] = convert_to_numeric(df["dose_rate_uSv_per_hr"])
    write_csv(df, "nuclear_facility_radiation.csv")


def clean_daily_reserve(file_name: str, output_name: str) -> None:
    df = pd.read_csv(DATA_DIR / file_name, encoding="utf-8-sig", na_values=NA_VALUES)
    df["date"] = df["日期"].apply(parse_compact_date)
    df["reserve_capacity_mw"] = convert_to_numeric(
        df["備轉容量(萬瓩)"], multiplier=10
    )
    df["reserve_margin_pct"] = convert_to_numeric(df["備轉容量率(%)"])
    df = df[["date", "reserve_capacity_mw", "reserve_margin_pct"]].dropna(
        subset=["date"]
    )
    write_csv(df, output_name)


def clean_peak_history() -> None:
    df = pd.read_csv(
        DATA_DIR / "歷年尖峰負載及備用容量率.csv",
        encoding="utf-8-sig",
        na_values=NA_VALUES,
    )
    df = df.rename(columns={"年度": "year"})
    df["year"] = df["year"].apply(parse_roc_year)
    df["尖峰負載(MW)"] = convert_to_numeric(df["尖峰負載(MW)"])
    df["備用容量率(％)"] = convert_to_numeric(df["備用容量率(％)"])
    df = df.rename(
        columns={
            "尖峰負載(MW)": "peak_load_mw",
            "備用容量率(％)": "reserve_margin_pct",
        }
    )
    write_csv(df, "annual_peak_load_and_reserve.csv")


def _melt_wide_monthly(
    path: Path, value_label: str, item_label: str
) -> pd.DataFrame:
    df = pd.read_excel(path, header=[4, 5])
    date_col = [col for col in df.columns if col[0] == "日期"][0]
    long = (
        df.set_index(date_col)
        .stack([0, 1], future_stack=True)
        .reset_index()
    )
    long = long.rename(
        columns={
            date_col: "raw_period",
            "level_1": "owner",
            "level_2": item_label,
            0: value_label,
        }
    )
    long["period"] = long["raw_period"].apply(parse_roc_month)
    long["owner"] = long["owner"].str.strip()
    long[item_label] = long[item_label].str.strip()
    long["owner_group"] = long["owner"].map(OWNER_MAP).fillna(long["owner"])
    return long


def clean_generation_fuel() -> None:
    long = _melt_wide_monthly(
        DATA_DIR / "發電燃料耗用.xlsx", value_label="consumption_units", item_label="fuel"
    )
    long["consumption_units"] = convert_to_numeric(long["consumption_units"])
    long = long.dropna(subset=["period", "fuel", "consumption_units"])
    long = long[
        ["period", "raw_period", "owner_group", "owner", "fuel", "consumption_units"]
    ]
    write_csv(long, "generation_fuel_consumption_long.csv")


def clean_generation_capacity() -> None:
    long = _melt_wide_monthly(
        DATA_DIR / "發電裝置容量.xlsx", value_label="capacity_mw", item_label="technology"
    )
    long["capacity_mw"] = convert_to_numeric(long["capacity_mw"])
    long = long.dropna(subset=["period", "technology", "capacity_mw"])
    long = long[
        ["period", "raw_period", "owner_group", "owner", "technology", "capacity_mw"]
    ]
    write_csv(long, "generation_capacity_long.csv")


def main() -> None:
    ensure_out_dir()
    clean_sales_by_usage()
    clean_generation_by_energy()
    clean_consumption_by_sector()
    clean_power_supply_and_units()
    clean_solar_feed_in()
    clean_retail_statistics()
    clean_generation_costs()
    clean_nuclear_performance()
    clean_radiation_monitoring()
    clean_daily_reserve("本年度每日尖峰備轉容量率.csv", "current_year_daily_reserve_margin.csv")
    clean_daily_reserve("近三年每日尖峰備轉容量率.csv", "three_year_daily_reserve_margin.csv")
    clean_peak_history()
    clean_generation_fuel()
    clean_generation_capacity()


if __name__ == "__main__":
    main()
