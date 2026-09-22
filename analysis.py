import argparse
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns


def load_data(file_path: Path) -> pd.DataFrame:
    if file_path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(file_path)
    if file_path.suffix.lower() == ".csv":
        return pd.read_csv(file_path)
    raise ValueError("File must be CSV or Excel")


def analyze(file_path: Path) -> dict[str, Any]:
    df = load_data(file_path)
    df["transaction_date"] = pd.to_datetime(df["transaction_date"], format="%m/%d/%y")
    df["transaction_time"] = pd.to_datetime(df["transaction_time"], format="%H:%M:%S")
    df["revenue"] = df["transaction_qty"] * df["unit_price"]
    df["hour"] = df["transaction_time"].dt.hour

    return {
        "df": df,
        "total_revenue": df["revenue"].sum(),
        "transactions": len(df),
        "items_sold": int(df["transaction_qty"].sum()),
        "avg_value": float(np.mean(df["revenue"])),
        "median_value": float(np.median(df["revenue"])),
        "date_range": (df["transaction_date"].min(), df["transaction_date"].max()),
        "by_category": df.groupby("product_category")["revenue"].sum().sort_values(ascending=False),
        "by_store": df.groupby("store_location")["revenue"].sum().sort_values(ascending=False),
        "by_hour": df.groupby("hour")["revenue"].sum().sort_index(),
        "top_products": (
            df.groupby("product_detail")["revenue"].sum()
            .sort_values(ascending=False)
            .head(10)
        ),
    }


def print_report(r: dict[str, Any]) -> None:
    start, end = r["date_range"]
    print(f"Date range     : {start:%Y-%m-%d} to {end:%Y-%m-%d}")
    print(f"Transactions   : {r['transactions']:,}")
    print(f"Items sold     : {r['items_sold']:,}")
    print(f"Total revenue  : ${r['total_revenue']:,.2f}")
    print(f"Avg value      : ${r['avg_value']:,.2f}")
    print(f"Median value   : ${r['median_value']:,.2f}")
    print("\nRevenue by category:\n", r["by_category"].to_string())
    print("\nRevenue by store:\n", r["by_store"].to_string())
    print("\nTop 10 products:\n", r["top_products"].to_string())


def save_charts(r: dict[str, Any], output_dir: Path, show: bool = True) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    sns.set_theme(style="whitegrid")

    charts = [
        ("by_category", "product_category", "Revenue by Product Category", "revenue_by_category.png"),
        ("by_store", "store_location", "Revenue by Store", "revenue_by_store.png"),
    ]

    for key, ylabel, title, fname in charts:
        data = r[key].reset_index()
        data.columns = pd.Index([ylabel, "revenue"])
        fig, ax = plt.subplots(figsize=(10, 6))
        sns.barplot(data=data, x="revenue", y=ylabel, hue=ylabel, legend=False, palette="viridis", ax=ax)
        ax.set_title(title)
        ax.set_xlabel("Revenue ($)")
        ax.set_ylabel("")
        fig.tight_layout()
        fig.savefig(output_dir / fname, dpi=150)
        if not show:
            plt.close(fig)

    # Hourly line chart
    hourly = r["by_hour"].reset_index()
    hourly.columns = pd.Index(["hour", "revenue"])
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.lineplot(data=hourly, x="hour", y="revenue", marker="o", ax=ax)
    ax.set_title("Revenue by Transaction Hour")
    ax.set_xlabel("Hour of Day")
    ax.set_ylabel("Revenue ($)")
    ax.set_xticks(hourly["hour"])
    fig.tight_layout()
    fig.savefig(output_dir / "revenue_by_hour.png", dpi=150)
    if not show:
        plt.close(fig)

    if show:
        plt.show(block=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Coffee shop sales analysis.")
    parser.add_argument("file", nargs="?", type=Path, default=Path("Coffee Shop Sales.csv"))
    parser.add_argument("--output-dir", type=Path, default=Path("analysis_outputs"))
    parser.add_argument("--no-show", action="store_true", help="Save charts without opening windows.")
    args = parser.parse_args()

    results = analyze(args.file)
    print_report(results)
    save_charts(results, args.output_dir, show=not args.no_show)
    print(f"\nCharts saved to: {args.output_dir}")


if __name__ == "__main__":
    main()
