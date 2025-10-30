#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DU post-load processing speed analysis.

Reads data from data/du_tasks_times.csv (UTF-8; semicolon-separated) with columns:
  - date (YYYY-MM-DD)
  - scenario (Типовой | Без дублей обменов | Без дублей обменов + Параллельные портфели)
  - minutes (float)

Produces charts:
  - Daily bars per scenario
  - 7-day rolling averages
  - Acceleration vs baseline Типовой (baseline_minutes / scenario_minutes)
  - Weekly average acceleration
"""

import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_PATH = PROJECT_ROOT / "data" / "du_tasks_times.csv"
OUT_DIR = PROJECT_ROOT / "Документация" / "Reports" / "du_speed_analysis" / "figures"


def ensure_paths() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)


def read_data(csv_path: Path) -> pd.DataFrame:
    if not csv_path.exists():
        raise FileNotFoundError(f"Input CSV not found: {csv_path}")
    df = pd.read_csv(csv_path, sep=";", encoding="utf-8")
    # Normalize
    df["date"] = pd.to_datetime(df["date"], format="%Y-%m-%d", errors="coerce")
    df = df.dropna(subset=["date", "scenario", "minutes"]).copy()
    df["minutes"] = pd.to_numeric(df["minutes"], errors="coerce")
    df = df.dropna(subset=["minutes"]).copy()
    return df


def plot_daily_bars(df: pd.DataFrame) -> None:
    plt.figure(figsize=(16, 6))
    scenarios = df["scenario"].unique()
    for scenario in scenarios:
        sub = df[df["scenario"] == scenario].sort_values("date")
        plt.bar(sub["date"], sub["minutes"], label=scenario, alpha=0.6)
    plt.ylabel("Minutes")
    plt.xlabel("Date")
    plt.title("Daily processing time by scenario")
    plt.legend()
    plt.tight_layout()
    plt.savefig(OUT_DIR / "01_daily_bars.png", dpi=150)
    plt.close()


def plot_rolling_avg(df: pd.DataFrame, window: int = 7) -> None:
    plt.figure(figsize=(16, 6))
    for scenario in df["scenario"].unique():
        sub = df[df["scenario"] == scenario].sort_values("date")
        roll = sub.set_index("date")["minutes"].rolling(window, min_periods=max(1, window // 2)).mean()
        plt.plot(roll.index, roll.values, label=f"{scenario} (MA{window})")
    plt.ylabel("Minutes (rolling avg)")
    plt.xlabel("Date")
    plt.title(f"{window}-day rolling average by scenario")
    plt.legend()
    plt.tight_layout()
    plt.savefig(OUT_DIR / "02_rolling_avg.png", dpi=150)
    plt.close()


def compute_acceleration(df: pd.DataFrame) -> pd.DataFrame:
    # Pivot to align scenarios by date
    pivot = df.pivot_table(index="date", columns="scenario", values="minutes", aggfunc="mean")
    base_col = "Типовой"
    if base_col not in pivot.columns:
        raise ValueError("Baseline 'Типовой' is required in the dataset to compute acceleration.")
    accel_frames = []
    for col in pivot.columns:
        if col == base_col:
            continue
        acc = pivot[base_col] / pivot[col]
        accel_frames.append(acc.rename(col))
    if not accel_frames:
        raise ValueError("No alternative scenarios to compare against baseline.")
    accel = pd.concat(accel_frames, axis=1)
    accel = accel.replace([pd.NA, pd.NaT, float("inf"), -float("inf")], pd.NA).dropna(how="all")
    return accel


def plot_acceleration(accel: pd.DataFrame) -> None:
    plt.figure(figsize=(16, 6))
    for col in accel.columns:
        plt.plot(accel.index, accel[col], label=col)
    plt.axhline(1.0, color="gray", linestyle="--", linewidth=1)
    plt.ylabel("Acceleration vs baseline (×)")
    plt.xlabel("Date")
    plt.title("Acceleration relative to 'Типовой'")
    plt.legend()
    plt.tight_layout()
    plt.savefig(OUT_DIR / "03_acceleration.png", dpi=150)
    plt.close()


def plot_weekly_acceleration(accel: pd.DataFrame) -> None:
    weekly = accel.resample("W").mean()
    plt.figure(figsize=(14, 6))
    for col in weekly.columns:
        plt.plot(weekly.index, weekly[col], marker="o", label=col)
    plt.axhline(1.0, color="gray", linestyle="--", linewidth=1)
    plt.ylabel("Avg weekly acceleration (×)")
    plt.xlabel("Week")
    plt.title("Weekly average acceleration vs baseline")
    plt.legend()
    plt.tight_layout()
    plt.savefig(OUT_DIR / "04_weekly_acceleration.png", dpi=150)
    plt.close()


def main() -> int:
    try:
        ensure_paths()
        df = read_data(DATA_PATH)
        df = df.sort_values("date")
        plot_daily_bars(df)
        plot_rolling_avg(df, window=7)
        accel = compute_acceleration(df)
        # Ensure datetime index for resampling
        accel.index = pd.to_datetime(accel.index)
        plot_acceleration(accel)
        plot_weekly_acceleration(accel)
        print(f"Done. Figures saved to: {OUT_DIR}")
        return 0
    except Exception as exc:
        print(f"Error: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())



