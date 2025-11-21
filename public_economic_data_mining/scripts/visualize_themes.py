#!/usr/bin/env python3
"""Generate visualizations for the ten analysis themes."""

from __future__ import annotations

import pandas as pd
from pathlib import Path
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

DATA_DIR = Path("data_clean")
OUTPUT_DIR = Path("figures")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

plt.rcParams["font.family"] = ["Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False
plt.style.use("seaborn-v0_8-darkgrid")


def save_current_fig(name: str) -> None:
    path = OUTPUT_DIR / f"{name}.png"
    plt.tight_layout()
    plt.savefig(path, dpi=150)
    plt.close()


def theme1_energy_structure():
    df = pd.read_csv(DATA_DIR / "generation_by_energy_type.csv")
    agg = (
        df.groupby(["year", "energy_type"])["generation_kwh"]
        .sum()
        .reset_index()
    )
    pivot = agg.pivot(index="year", columns="energy_type", values="generation_kwh").fillna(0)
    pivot = pivot / 1e9  # GWh
    pivot.sort_index(inplace=True)
    pivot.plot.area(figsize=(10, 5))
    plt.ylabel("發購電量 (GWh)")
    plt.xlabel("年度")
    plt.title("主題1：能源結構長期趨勢")
    save_current_fig("theme1_energy_structure")


def theme2_demand_structure():
    df = pd.read_csv(DATA_DIR / "electricity_sales_by_usage.csv")
    pivot = df.pivot(index="year", columns="usage_category", values="sales_kwh") / 1e9
    pivot.plot(kind="area", figsize=(10, 5))
    plt.ylabel("售電量 (GWh)")
    plt.title("主題2：用途別售電量佔比")
    save_current_fig("theme2_demand_structure")


def theme3_supply_demand_balance():
    df = pd.read_csv(DATA_DIR / "power_supply_demand_summary.csv", parse_dates=["date"])
    daily = df.set_index("date").resample("W").mean()
    plt.figure(figsize=(10, 5))
    plt.plot(daily.index, daily["net_peak_capacity_mw"], label="淨尖峰供電能力")
    plt.plot(daily.index, daily["peak_load_mw"], label="尖峰負載")
    plt.fill_between(
        daily.index,
        daily["peak_load_mw"],
        daily["net_peak_capacity_mw"],
        where=daily["net_peak_capacity_mw"] >= daily["peak_load_mw"],
        color="green",
        alpha=0.2,
        label="備轉空間",
    )
    plt.ylabel("MW")
    plt.title("主題3：尖峰供需平衡")
    plt.legend()
    save_current_fig("theme3_supply_demand")


def theme4_renewable_growth():
    df = pd.read_csv(DATA_DIR / "solar_feed_in_records.csv")
    df = df[~df["is_annual_total"]].copy()
    df["period"] = pd.to_datetime(df["period"])
    plt.figure(figsize=(10, 4))
    plt.plot(df["period"], df["energy_mwh"] / 1000, label="太陽光電購電 (千MWh)")
    plt.title("主題4：太陽光電購電成長")
    plt.ylabel("千MWh")
    save_current_fig("theme4_renewables")


def theme5_generation_costs():
    df = pd.read_csv(DATA_DIR / "generation_costs_by_technology.csv")
    latest_year = df["calendar_year"].max()
    subset = df[df["calendar_year"] == latest_year]
    major = subset[subset["subcategory"].notna()].copy()
    major = major.groupby("subcategory")["cost_ntd_per_kwh"].mean().sort_values()
    plt.figure(figsize=(8, 5))
    major.plot(kind="barh", color="tab:purple")
    plt.xlabel("元/度")
    plt.title(f"主題5：各種發電方式成本 (年度 {latest_year})")
    save_current_fig("theme5_costs")


def theme6_demand_vs_macro():
    df = pd.read_csv(DATA_DIR / "electricity_consumption_by_sector.csv", parse_dates=["period"])
    plt.figure(figsize=(6, 6))
    plt.scatter(df["industrial_mwh"] / 1000, df["total_mwh"] / 1000, s=10, alpha=0.5)
    plt.xlabel("工業用電 (千MWh)")
    plt.ylabel("全國用電 (千MWh)")
    plt.title("主題6：工業用電 vs 全國用電")
    save_current_fig("theme6_demand_macro")


def theme7_seasonality():
    df = pd.read_csv(DATA_DIR / "three_year_daily_reserve_margin.csv", parse_dates=["date"])
    df["month"] = df["date"].dt.month
    monthly = df.groupby("month")["reserve_margin_pct"].mean()
    plt.figure(figsize=(8, 4))
    plt.plot(monthly.index, monthly.values, marker="o")
    plt.title("主題7：備轉容量率季節性")
    plt.xlabel("月份")
    plt.ylabel("備轉容量率 (%)")
    save_current_fig("theme7_seasonality")


def theme8_energy_security():
    df = pd.read_csv(DATA_DIR / "generation_fuel_consumption_long.csv", parse_dates=["period"])
    latest = df[df["period"] >= df["period"].max() - pd.DateOffset(years=3)]
    pivot = (
        latest.groupby(["period", "fuel"])["consumption_units"].sum().unstack().fillna(0)
    )
    shares = pivot.divide(pivot.sum(axis=1), axis=0) * 100
    shares.rolling(3).mean().plot(figsize=(10, 5))
    plt.ylabel("燃料占比 (%)")
    plt.title("主題8：近年燃料依賴度")
    save_current_fig("theme8_security")


def theme9_emissions():
    df = pd.read_csv(DATA_DIR / "nuclear_generation_performance.csv")
    df = df.sort_values("year")
    plt.figure(figsize=(8, 4))
    plt.plot(df["year"], df["co2_reduction_tonnes"] / 1e6, marker="o")
    plt.ylabel("減碳量 (百萬噸)")
    plt.title("主題9：核能減碳效益")
    save_current_fig("theme9_emissions")


def theme10_dashboard():
    df = pd.read_csv(DATA_DIR / "power_supply_demand_summary.csv", parse_dates=["date"])
    monthly = df.set_index("date").resample("M").mean()
    fig, ax1 = plt.subplots(figsize=(10, 5))
    ax1.plot(monthly.index, monthly["peak_load_mw"], color="tab:red", label="尖峰負載")
    ax1.set_ylabel("尖峰負載 (MW)", color="tab:red")
    ax2 = ax1.twinx()
    ax2.bar(monthly.index, monthly["reserve_margin_pct"], alpha=0.3, color="tab:blue", label="備轉率")
    ax2.set_ylabel("備轉率 (%)", color="tab:blue")
    fig.suptitle("主題10：綜合儀表 (尖峰負載 vs 備轉率)")
    fig.legend(loc="upper right")
    save_current_fig("theme10_dashboard")


def main():
    theme1_energy_structure()
    theme2_demand_structure()
    theme3_supply_demand_balance()
    theme4_renewable_growth()
    theme5_generation_costs()
    theme6_demand_vs_macro()
    theme7_seasonality()
    theme8_energy_security()
    theme9_emissions()
    theme10_dashboard()
    print(f"圖表已輸出至 {OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
