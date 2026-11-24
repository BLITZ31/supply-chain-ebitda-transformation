import os
import sys
import json
import sqlite3
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import seaborn as sns

warnings.filterwarnings("ignore")


# Configuration & Corporate Consulting Palette

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUMMARY_DIR = os.path.join(BASE_DIR, "outputs", "summary_tables")
CHARTS_DIR = os.path.join(BASE_DIR, "outputs", "charts")
os.makedirs(SUMMARY_DIR, exist_ok=True)
os.makedirs(CHARTS_DIR, exist_ok=True)

NAVY = "#00338D"
DARK_SLATE = "#1B365D"
BLUE = "#005EB8"
TEAL = "#0091DA"
LIGHT_BLUE = "#7FA8D9"
GREEN = "#1E7E34"
AMBER = "#D97706"
RED = "#C0392B"
GREY = "#6B7280"
LIGHT_GREY = "#E5E7EB"
BG_COLOR = "#FFFFFF"

plt.rcParams.update({
    "figure.facecolor": BG_COLOR,
    "axes.facecolor": BG_COLOR,
    "axes.edgecolor": "#D1D5DB",
    "axes.linewidth": 0.8,
    "font.family": "DejaVu Sans",
    "axes.titleweight": "bold",
    "axes.titlesize": 13,
    "axes.labelsize": 10,
    "xtick.labelsize": 9.5,
    "ytick.labelsize": 9.5,
})


def locate_dataset() -> str:
    """Locate DataCoSupplyChainDataset.csv in data/ or current directory."""
    candidates = [
        os.path.join(BASE_DIR, "data", "DataCoSupplyChainDataset.csv"),
        os.path.join(BASE_DIR, "DataCoSupplyChainDataset.csv"),
        os.path.join(os.getcwd(), "DataCoSupplyChainDataset.csv"),
    ]
    if len(sys.argv) > 1:
        candidates.insert(0, sys.argv[1])
    for path in candidates:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(
        "DataCoSupplyChainDataset.csv not found. Place it in data/ or pass its path as an argument."
    )


def load_and_clean(csv_path: str) -> pd.DataFrame:
    """Load raw CSV with latin1 encoding and engineer clean analytical fields."""
    df = pd.read_csv(csv_path, encoding="latin1")
    df["order_dt"] = pd.to_datetime(df["order date (DateOrders)"])
    df["ship_dt"] = pd.to_datetime(df["shipping date (DateOrders)"])
    df["order_year"] = df["order_dt"].dt.year
    df["order_ym"] = df["order_dt"].dt.to_period("M").astype(str)
    df["shipping_delay_days"] = (
        df["Days for shipping (real)"] - df["Days for shipment (scheduled)"]
    )
    df["is_loss_making"] = (df["Benefit per order"] < 0).astype(int)
    df["is_canceled_or_fraud"] = (df["Delivery Status"] == "Shipping canceled").astype(int)
    df["is_suspected_fraud"] = (df["Order Status"] == "SUSPECTED_FRAUD").astype(int)
    df["is_canceled"] = (df["Order Status"] == "CANCELED").astype(int)
    df["excess_disc_over_15"] = np.where(
        df["Order Item Discount Rate"] > 0.15,
        (df["Order Item Discount Rate"] - 0.15) * df["Sales"],
        0.0,
    )
    return df


def run_sql_diagnostics(df: pd.DataFrame) -> dict:
    """Load cleaned dataset into SQLite and execute structured SQL aggregations."""
    conn = sqlite3.connect(":memory:")
    sql_cols = [
        "Order Id", "Order Item Id", "order_year", "order_ym", "Type",
        "Shipping Mode", "Delivery Status", "Order Status", "Market",
        "Order Region", "Customer Segment", "Category Name", "Product Name",
        "Days for shipping (real)", "Days for shipment (scheduled)",
        "shipping_delay_days", "Late_delivery_risk", "Order Item Quantity",
        "Sales", "Order Item Discount", "Order Item Discount Rate",
        "Order Item Total", "Benefit per order", "is_loss_making",
        "is_canceled_or_fraud", "is_suspected_fraud", "is_canceled",
        "excess_disc_over_15"
    ]
    df[sql_cols].to_sql("orders", conn, index=False, if_exists="replace")

    # 1. Shipping Mode SLA Diagnostics
    q_shipping = """
    SELECT
        "Shipping Mode" AS shipping_mode,
        COUNT(*) AS order_items,
        SUM("Order Item Total") AS net_sales,
        SUM("Benefit per order") AS operating_profit,
        ROUND(SUM("Benefit per order") * 100.0 / SUM("Order Item Total"), 2) AS margin_pct,
        ROUND(AVG("Days for shipment (scheduled)"), 2) AS avg_scheduled_days,
        ROUND(AVG("Days for shipping (real)"), 2) AS avg_actual_days,
        ROUND(AVG(shipping_delay_days), 2) AS avg_delay_days,
        ROUND(AVG("Late_delivery_risk") * 100.0, 2) AS late_delivery_pct
    FROM orders
    GROUP BY "Shipping Mode"
    ORDER BY late_delivery_pct DESC;
    """
    shipping_df = pd.read_sql_query(q_shipping, conn)

    # 2. Payment Channel Fraud & Cancellation Risk Audit
    q_payment = """
    SELECT
        Type AS payment_type,
        COUNT(*) AS order_items,
        SUM("Order Item Total") AS net_sales,
        SUM("Benefit per order") AS operating_profit,
        SUM(is_suspected_fraud) AS suspected_fraud_items,
        SUM(is_canceled) AS canceled_items,
        ROUND(AVG(is_suspected_fraud) * 100.0, 2) AS fraud_rate_pct,
        ROUND(AVG(is_canceled) * 100.0, 2) AS cancel_rate_pct,
        ROUND(AVG(is_canceled_or_fraud) * 100.0, 2) AS total_failure_rate_pct,
        SUM(CASE WHEN is_canceled_or_fraud = 1 THEN "Order Item Total" ELSE 0 END) AS failed_net_sales
    FROM orders
    GROUP BY Type
    ORDER BY total_failure_rate_pct DESC;
    """
    payment_df = pd.read_sql_query(q_payment, conn)

    # 3. Regional Risk & Profitability Summary
    q_region = """
    SELECT
        Market AS market,
        "Order Region" AS order_region,
        COUNT(*) AS order_items,
        SUM("Order Item Total") AS net_sales,
        SUM("Benefit per order") AS operating_profit,
        ROUND(SUM("Benefit per order") * 100.0 / SUM("Order Item Total"), 2) AS margin_pct,
        ROUND(AVG("Late_delivery_risk") * 100.0, 2) AS late_delivery_pct,
        ROUND(AVG(is_canceled_or_fraud) * 100.0, 2) AS cancel_fraud_pct
    FROM orders
    GROUP BY Market, "Order Region"
    ORDER BY net_sales DESC;
    """
    region_df = pd.read_sql_query(q_region, conn)

    # 4. Market Summary
    q_market = """
    SELECT
        Market AS market,
        COUNT(*) AS order_items,
        SUM("Order Item Total") AS net_sales,
        SUM("Benefit per order") AS operating_profit,
        ROUND(SUM("Benefit per order") * 100.0 / SUM("Order Item Total"), 2) AS margin_pct,
        ROUND(AVG("Late_delivery_risk") * 100.0, 2) AS late_delivery_pct,
        ROUND(AVG(is_loss_making) * 100.0, 2) AS loss_order_pct
    FROM orders
    GROUP BY Market
    ORDER BY net_sales DESC;
    """
    market_df = pd.read_sql_query(q_market, conn)

    # 5. Customer Segment Summary
    q_segment = """
    SELECT
        "Customer Segment" AS customer_segment,
        COUNT(*) AS order_items,
        SUM("Order Item Total") AS net_sales,
        SUM("Benefit per order") AS operating_profit,
        ROUND(SUM("Benefit per order") * 100.0 / SUM("Order Item Total"), 2) AS margin_pct,
        ROUND(AVG("Order Item Discount Rate") * 100.0, 2) AS avg_discount_pct,
        ROUND(AVG("Late_delivery_risk") * 100.0, 2) AS late_delivery_pct
    FROM orders
    GROUP BY "Customer Segment"
    ORDER BY net_sales DESC;
    """
    segment_df = pd.read_sql_query(q_segment, conn)

    # 6. Yearly Financial Performance
    q_yearly = """
    SELECT
        order_year,
        COUNT(DISTINCT "Order Id") AS unique_orders,
        COUNT(*) AS order_items,
        SUM(Sales) AS gross_sales,
        SUM("Order Item Discount") AS total_discounts,
        SUM("Order Item Total") AS net_sales,
        SUM("Benefit per order") AS operating_profit,
        ROUND(SUM("Benefit per order") * 100.0 / SUM("Order Item Total"), 2) AS margin_pct,
        ROUND(AVG("Late_delivery_risk") * 100.0, 2) AS late_delivery_pct
    FROM orders
    GROUP BY order_year
    ORDER BY order_year;
    """
    yearly_df = pd.read_sql_query(q_yearly, conn)

    conn.close()
    return {
        "shipping_sla": shipping_df,
        "payment_risk": payment_df,
        "regional_risk": region_df,
        "market_summary": market_df,
        "segment_summary": segment_df,
        "yearly_summary": yearly_df,
    }


def compute_abc_xyz_and_discounts(df: pd.DataFrame) -> tuple:
    """Compute ABC-XYZ inventory segmentation and discount tier margin leakage."""
    df_3yr = df[df["order_year"].isin([2015, 2016, 2017])].copy()

    cat = (
        df.groupby("Category Name")
        .agg(
            order_items=("Order Item Id", "count"),
            units_sold=("Order Item Quantity", "sum"),
            gross_sales=("Sales", "sum"),
            net_sales=("Order Item Total", "sum"),
            discounts=("Order Item Discount", "sum"),
            operating_profit=("Benefit per order", "sum"),
            late_delivery_pct=("Late_delivery_risk", lambda x: x.mean() * 100.0),
            loss_order_pct=("is_loss_making", lambda x: x.mean() * 100.0),
        )
        .sort_values("net_sales", ascending=False)
        .reset_index()
    )
    cat["rev_share_pct"] = cat["net_sales"] / cat["net_sales"].sum() * 100.0
    cat["cum_rev_share_pct"] = cat["rev_share_pct"].cumsum()
    cat["margin_pct"] = cat["operating_profit"] / cat["net_sales"] * 100.0

    # ABC classification (80 / 15 / 5 revenue contribution)
    cat["abc_class"] = np.where(
        cat["cum_rev_share_pct"] <= 80.0,
        "A",
        np.where(cat["cum_rev_share_pct"] <= 95.0, "B", "C"),
    )

    # Demand volatility (Coefficient of Variation across 36 full months 2015-2017)
    monthly_3yr = (
        df_3yr.groupby(["Category Name", "order_ym"])["Order Item Quantity"]
        .sum()
        .unstack(fill_value=0)
    )
    cv_3yr = (monthly_3yr.std(axis=1) / monthly_3yr.mean(axis=1)).rename("demand_cv")

    monthly_all = (
        df.groupby(["Category Name", "order_ym"])["Order Item Quantity"]
        .sum()
        .unstack(fill_value=0)
    )
    cv_all = (monthly_all.std(axis=1) / monthly_all.mean(axis=1)).rename("demand_cv_all")

    cat = cat.merge(cv_3yr, on="Category Name", how="left").merge(
        cv_all, on="Category Name", how="left"
    )
    cat["demand_cv"] = cat["demand_cv"].fillna(cat["demand_cv_all"]).round(3)
    cat.drop(columns=["demand_cv_all"], inplace=True)

    cat["xyz_class"] = np.where(
        cat["demand_cv"] <= 0.5, "X", np.where(cat["demand_cv"] <= 1.0, "Y", "Z")
    )
    cat["abc_xyz"] = cat["abc_class"] + "-" + cat["xyz_class"]

    abc_xyz_summary = (
        cat.groupby("abc_xyz")
        .agg(
            categories=("Category Name", "count"),
            order_items=("order_items", "sum"),
            net_sales=("net_sales", "sum"),
            operating_profit=("operating_profit", "sum"),
            avg_cv=("demand_cv", "mean"),
            avg_late_pct=("late_delivery_pct", "mean"),
        )
        .reset_index()
    )
    abc_xyz_summary["rev_share_pct"] = (
        abc_xyz_summary["net_sales"] / cat["net_sales"].sum() * 100.0
    )
    abc_xyz_summary["margin_pct"] = (
        abc_xyz_summary["operating_profit"] / abc_xyz_summary["net_sales"] * 100.0
    )

    # Discount Tier Margin Leakage
    bins = [-0.01, 0.0, 0.05, 0.10, 0.15, 0.20, 0.35]
    labels = ["0%", "1-5%", "6-10%", "11-15%", "16-20%", ">20%"]
    df["discount_tier"] = pd.cut(df["Order Item Discount Rate"], bins=bins, labels=labels)
    disc_df = (
        df.groupby("discount_tier", observed=False)
        .agg(
            order_items=("Order Item Id", "count"),
            gross_sales=("Sales", "sum"),
            total_discounts=("Order Item Discount", "sum"),
            net_sales=("Order Item Total", "sum"),
            operating_profit=("Benefit per order", "sum"),
            loss_order_pct=("is_loss_making", lambda x: x.mean() * 100.0),
        )
        .reset_index()
    )
    disc_df["margin_pct"] = disc_df["operating_profit"] / disc_df["net_sales"] * 100.0
    return cat, abc_xyz_summary, disc_df


def compute_baseline_metrics(df: pd.DataFrame) -> dict:
    """Compute exact dataset-wide & 3-year annualized baseline KPIs for Excel model."""
    df_3yr = df[df["order_year"].isin([2015, 2016, 2017])].copy()
    n_years = 3.0

    ann_gross = float(df_3yr["Sales"].sum() / n_years)
    ann_disc = float(df_3yr["Order Item Discount"].sum() / n_years)
    ann_net = float(df_3yr["Order Item Total"].sum() / n_years)
    ann_profit = float(df_3yr["Benefit per order"].sum() / n_years)
    ann_cogs_opex = ann_net - ann_profit

    ann_excess_disc_15 = float(df_3yr["excess_disc_over_15"].sum() / n_years)
    ann_neg_profit_loss = float(-df_3yr.loc[df_3yr["Benefit per order"] < 0, "Benefit per order"].sum() / n_years)
    ann_neg_items = float((df_3yr["Benefit per order"] < 0).sum() / n_years)
    ann_late_items = float((df_3yr["Late_delivery_risk"] == 1).sum() / n_years)

    tf_fail = df_3yr[(df_3yr["Type"] == "TRANSFER") & (df_3yr["is_canceled_or_fraud"] == 1)]
    ann_tf_fail_items = float(len(tf_fail) / n_years)
    ann_tf_fail_sales = float(tf_fail["Order Item Total"].sum() / n_years)
    ann_tf_fail_profit = float(tf_fail["Benefit per order"].sum() / n_years)

    metrics = {
        "total_rows": int(len(df)),
        "unique_orders": int(df["Order Id"].nunique()),
        "unique_categories": int(df["Category Name"].nunique()),
        "unique_products": int(df["Product Card Id"].nunique()),
        "unique_markets": int(df["Market"].nunique()),
        "unique_regions": int(df["Order Region"].nunique()),
        "cum_gross_sales": float(df["Sales"].sum()),
        "cum_discounts": float(df["Order Item Discount"].sum()),
        "cum_net_sales": float(df["Order Item Total"].sum()),
        "cum_operating_profit": float(df["Benefit per order"].sum()),
        "cum_margin_pct": float(df["Benefit per order"].sum() / df["Order Item Total"].sum() * 100.0),
        "overall_late_pct": float(df["Late_delivery_risk"].mean() * 100.0),
        "overall_loss_pct": float(df["is_loss_making"].mean() * 100.0),
        "cum_loss_pool": float(-df.loc[df["Benefit per order"] < 0, "Benefit per order"].sum()),
        "cum_transfer_fail_sales": float(df.loc[(df["Type"] == "TRANSFER") & (df["is_canceled_or_fraud"] == 1), "Order Item Total"].sum()),
        "cum_transfer_fail_items": int(((df["Type"] == "TRANSFER") & (df["is_canceled_or_fraud"] == 1)).sum()),
        # 3-Year Annualized Baseline (FY2015-FY2017 Average)
        "ann_gross_sales": ann_gross,
        "ann_discounts": ann_disc,
        "ann_net_sales": ann_net,
        "ann_operating_profit": ann_profit,
        "ann_cogs_opex": ann_cogs_opex,
        "ann_margin_pct": float(ann_profit / ann_net * 100.0),
        "ann_excess_disc_over_15": ann_excess_disc_15,
        "ann_neg_profit_loss": ann_neg_profit_loss,
        "ann_neg_items": ann_neg_items,
        "ann_late_items": ann_late_items,
        "ann_tf_fail_items": ann_tf_fail_items,
        "ann_tf_fail_sales": ann_tf_fail_sales,
        "ann_tf_fail_profit": ann_tf_fail_profit,
    }
    return metrics


def generate_charts(metrics: dict, shipping_df: pd.DataFrame, abc_xyz_df: pd.DataFrame, disc_df: pd.DataFrame, region_df: pd.DataFrame):
    """Generate 4 executive consulting-grade charts for README.md."""


    # CHART 1: Base-Case Annualized EBITDA Waterfall Bridge
    
    base_ebitda = metrics["ann_operating_profit"] / 1e3  # in $K
    l1 = metrics["ann_excess_disc_over_15"] * 0.55 * 0.85 / 1e3
    l2 = metrics["ann_neg_profit_loss"] * 0.18 / 1e3
    l3 = metrics["ann_late_items"] * 4.00 / 1e3
    l4 = (metrics["ann_tf_fail_profit"] + metrics["ann_tf_fail_items"] * 12.0) * 0.55 / 1e3
    wc_freed = (metrics["ann_cogs_opex"] * 0.75 / 365.0) * 10.0
    l5 = wc_freed * 0.18 / 1e3
    target_ebitda = base_ebitda + l1 + l2 + l3 + l4 + l5

    labels = [
        "Baseline\nEBITDA\n(12.0%)",
        "Pricing &\nDiscount Cap\n(>15%)",
        "Loss-Making\nOrder Floor\nRemediation",
        "Expedited\nSLA & 3PL\nRe-Routing",
        "TRANSFER\nFraud & Cancel\nControls",
        "ABC-XYZ\nInventory\nCarrying Lift",
        "Target\nEBITDA\n(16.6%)",
    ]
    values = [base_ebitda, l1, l2, l3, l4, l5, target_ebitda]
    bottoms = [0, base_ebitda, base_ebitda + l1, base_ebitda + l1 + l2,
               base_ebitda + l1 + l2 + l3, base_ebitda + l1 + l2 + l3 + l4, 0]
    colors = [DARK_SLATE, TEAL, BLUE, TEAL, BLUE, TEAL, GREEN]

    fig, ax = plt.subplots(figsize=(11, 5.6), dpi=200)
    bars = ax.bar(labels, values, bottom=bottoms, color=colors, width=0.58, edgecolor="white", linewidth=1.2)

    # Connector lines
    cum = base_ebitda
    for i in range(len(values) - 1):
        if i > 0:
            cum += values[i]
        ax.plot([i + 0.29, i + 0.71], [cum, cum], color="#9CA3AF", linestyle="--", linewidth=0.9)

    for idx, (bar, val, bot) in enumerate(zip(bars, values, bottoms)):
        y_pos = bot + val + 28
        txt = f"${val:,.0f}K" if idx in (0, len(values) - 1) else f"+${val:,.0f}K"
        ax.text(bar.get_x() + bar.get_width() / 2, y_pos, txt,
                ha="center", va="bottom", fontsize=9.5, fontweight="bold",
                color=DARK_SLATE if idx == 0 else (GREEN if idx == len(values) - 1 else BLUE))

    ax.set_ylim(0, target_ebitda * 1.18)
    ax.yaxis.set_major_formatter(mtick.StrMethodFormatter("${x:,.0f}K"))
    ax.set_ylabel("Annualized EBITDA ($ Thousands, FY15-17 Run-Rate)", fontweight="bold", color=DARK_SLATE)
    ax.set_title("Base-Case EBITDA Bridge: +$509K (+456 bps) Annualized Uplift Across 5 Transformation Levers",
                 loc="left", pad=16, color=DARK_SLATE, fontsize=12.5)
    ax.text(0, 1.01, "Source: DataCo Global Supply Chain Dataset (180,519 order items) | Excludes $197K one-time Working Capital cash release",
            transform=ax.transAxes, fontsize=8.8, color=GREY)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, "01_ebitda_waterfall_bridge.png"), dpi=200)
    plt.close()

    
    # CHART 2: ABC-XYZ Inventory & Revenue Matrix
    
    abc_order = ["A", "B", "C"]
    xyz_order = ["X", "Y", "Z"]
    mat_rev = pd.DataFrame(0.0, index=abc_order, columns=xyz_order)
    mat_cnt = pd.DataFrame(0, index=abc_order, columns=xyz_order)
    mat_mgn = pd.DataFrame(0.0, index=abc_order, columns=xyz_order)

    for _, row in abc_xyz_df.iterrows():
        a, x = row["abc_xyz"].split("-")
        mat_rev.loc[a, x] = row["rev_share_pct"]
        mat_cnt.loc[a, x] = int(row["categories"])
        mat_mgn.loc[a, x] = row["margin_pct"]

    annot = pd.DataFrame("", index=abc_order, columns=xyz_order)
    for a in abc_order:
        for x in xyz_order:
            if mat_cnt.loc[a, x] > 0:
                annot.loc[a, x] = f"{mat_rev.loc[a, x]:.1f}% Rev\n({mat_cnt.loc[a, x]} Cats)\n{mat_mgn.loc[a, x]:.1f}% Mgn"
            else:
                annot.loc[a, x] = "0.0% Rev\n(0 Cats)"

    fig, ax = plt.subplots(figsize=(10.8, 5.8), dpi=200)
    sns.heatmap(
        mat_rev, annot=annot, fmt="", cmap="Blues", linewidths=2, linecolor="white",
        cbar_kws={"label": "Share of Net Sales (%)"}, ax=ax,
        annot_kws={"size": 10.5, "weight": "bold"}, vmin=0, vmax=80
    )
    ax.set_xticklabels(["X: Stable Demand\n(CV <= 0.5)", "Y: Moderate Volatility\n(0.5 < CV <= 1.0)", "Z: Erratic Demand\n(CV > 1.0)"], fontsize=10, fontweight="bold")
    ax.set_yticklabels(["A: High Value\n(Top 80% Rev)", "B: Mid Value\n(Next 15% Rev)", "C: Long Tail\n(Bottom 5% Rev)"], rotation=0, fontsize=10, fontweight="bold")
    ax.set_title("ABC-XYZ Matrix: 7 A-X Categories Drive 77% of Revenue; 26 Erratic Z-Class Categories Bloat Stock",
                 loc="left", pad=22, color=DARK_SLATE, fontsize=11.8)
    ax.text(0, 1.015, "Action: Automate A-X core replenishment; shift 26 erratic B-Z & C-Z categories (7.6% rev, 52% of categories) to Make-to-Order",
            transform=ax.transAxes, fontsize=8.6, color=GREY)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, "02_abc_xyz_inventory_matrix.png"), dpi=200)
    plt.close()

    
    # CHART 3: Dual Panel — Shipping Mode SLA Failure & Discount Margin Leakage
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.2, 5.2), dpi=200)

    # Panel A: Late Delivery Risk by Shipping Mode
    sm_plot = shipping_df.sort_values("late_delivery_pct", ascending=True)
    bar_colors = [BLUE if v < 50 else (AMBER if v < 80 else RED) for v in sm_plot["late_delivery_pct"]]
    bars1 = ax1.barh(sm_plot["shipping_mode"], sm_plot["late_delivery_pct"], color=bar_colors, height=0.55)
    ax1.axvline(metrics["overall_late_pct"], color=DARK_SLATE, linestyle="--", linewidth=1.2, label=f"Overall Avg ({metrics['overall_late_pct']:.1f}%)")
    for bar, (_, row) in zip(bars1, sm_plot.iterrows()):
        ax1.text(bar.get_width() + 1.5, bar.get_y() + bar.get_height() / 2,
                 f"{row['late_delivery_pct']:.1f}%\n({row['avg_actual_days']:.1f}d vs {row['avg_scheduled_days']:.0f}d SLA)",
                 va="center", fontsize=8.6, fontweight="bold", color=DARK_SLATE)
    ax1.set_xlim(0, 122)
    ax1.xaxis.set_major_formatter(mtick.PercentFormatter())
    ax1.set_title("A. Expedited Shipping SLA Crisis", loc="left", color=DARK_SLATE, fontsize=11.5)
    ax1.set_xlabel("Late Delivery Risk (%)", fontweight="bold", color=DARK_SLATE)
    ax1.legend(loc="lower right", frameon=True, fontsize=8.5)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.grid(axis="x", linestyle=":", alpha=0.5)

    # Panel B: Discount Tier vs Margin % & Discount Spend
    x = np.arange(len(disc_df))
    b2 = ax2.bar(x, disc_df["total_discounts"] / 1e6, color=LIGHT_BLUE, width=0.55, label="Total Discounts ($M)")
    ax2.set_xticks(x)
    ax2.set_xticklabels(disc_df["discount_tier"], fontweight="bold")
    ax2.set_ylabel("Cumulative Discounts Granted ($M)", color=BLUE, fontweight="bold")
    ax2.set_ylim(0, 1.75)
    ax2.yaxis.set_major_formatter(mtick.StrMethodFormatter("${x:.2f}M"))

    ax2_twin = ax2.twinx()
    ax2_twin.plot(x, disc_df["margin_pct"], color=RED, marker="o", linewidth=2.4, markersize=7, label="Operating Margin (%)")
    for xi, mgn in zip(x, disc_df["margin_pct"]):
        ax2_twin.text(
            xi, mgn + 0.16, f"{mgn:.2f}%", ha="center", va="bottom",
            fontsize=8.5, fontweight="bold", color=RED,
            bbox=dict(facecolor="white", alpha=0.82, edgecolor="none", pad=1.2)
        )
    ax2_twin.set_ylabel("Operating Profit Margin (%)", color=RED, fontweight="bold")
    ax2_twin.set_ylim(10.8, 14.0)
    ax2_twin.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=1))
    ax2.set_title("B. Discount Leakage: >15% Tiers Drain $2.27M", loc="left", color=DARK_SLATE, fontsize=11.5)
    ax2.set_xlabel("Order Discount Rate Tier", fontweight="bold", color=DARK_SLATE)
    ax2.spines["top"].set_visible(False)
    ax2_twin.spines["top"].set_visible(False)
    ax2.grid(axis="y", linestyle=":", alpha=0.45)

    # Combine legends for panel B
    lines1, labels1 = ax2.get_legend_handles_labels()
    lines2, labels2 = ax2_twin.get_legend_handles_labels()
    ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=8.3, frameon=True)

    fig.suptitle("Operational & Commercial Margin Diagnostics (180,519 Line Items)", fontsize=13.5, fontweight="bold", color=NAVY, y=0.99)
    plt.tight_layout(rect=[0, 0, 1, 0.94])
    plt.savefig(os.path.join(CHARTS_DIR, "03_delivery_delay_vs_margin.png"), dpi=200)
    plt.close()

   
    # CHART 4: Regional Risk & Exposure Matrix (Top 12 Regions by Revenue)
  
    top_reg = region_df.head(12).copy().sort_values("late_delivery_pct", ascending=True)
    fig, ax = plt.subplots(figsize=(10.5, 5.5), dpi=200)
    bubble_sizes = (top_reg["net_sales"] / top_reg["net_sales"].max()) * 900 + 120
    scatter = ax.scatter(
        top_reg["late_delivery_pct"],
        top_reg["margin_pct"],
        s=bubble_sizes,
        c=top_reg["cancel_fraud_pct"],
        cmap="RdYlBu_r",
        alpha=0.85,
        edgecolors=DARK_SLATE,
        linewidth=1.1,
    )
    cbar = plt.colorbar(scatter, ax=ax)
    cbar.set_label("Cancellation & Suspected Fraud Rate (%)", fontweight="bold", fontsize=9)

    ax.axvline(metrics["overall_late_pct"], color=GREY, linestyle="--", linewidth=1.0)
    ax.axhline(metrics["cum_margin_pct"], color=GREY, linestyle="--", linewidth=1.0)

    custom_offsets = {
        "South America": (-12, 16),
        "Southern Europe": (26, -18),
        "Northern Europe": (-18, 12),
        "Central America": (0, -24),
        "Western Europe": (0, -24),
        "South Asia": (4, 14),
        "Southeast Asia": (0, 14),
        "East of USA": (0, 13),
        "West of USA ": (-8, 13),
        "Caribbean": (4, 13),
        "Oceania": (-10, 13),
        "Eastern Asia": (18, -16),
    }
    for _, row in top_reg.iterrows():
        reg_name = row["order_region"]
        ox, oy = custom_offsets.get(reg_name, (0, 12))
        ax.annotate(
            f"{reg_name.strip()} (USD {row['net_sales']/1e6:.1f}M)",
            (row["late_delivery_pct"], row["margin_pct"]),
            xytext=(ox, oy),
            textcoords="offset points",
            ha="center",
            fontsize=8.1,
            fontweight="bold",
            color=DARK_SLATE,
            bbox=dict(facecolor="white", alpha=0.72, edgecolor="none", pad=0.8),
        )

    ax.set_xlabel("Late Delivery Risk (%)", fontweight="bold", color=DARK_SLATE)
    ax.set_ylabel("Operating Margin (%)", fontweight="bold", color=DARK_SLATE)
    ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=1))
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=1))
    ax.set_xlim(top_reg["late_delivery_pct"].min() - 0.65, top_reg["late_delivery_pct"].max() + 0.75)
    ax.set_ylim(top_reg["margin_pct"].min() - 0.38, top_reg["margin_pct"].max() + 0.42)
    ax.set_title("Regional Exposure Matrix: Western Europe (USD 5.3M) & South Asia (USD 1.4M) Show High Delay & Weak Margin",
                 loc="left", pad=18, color=DARK_SLATE, fontsize=10.8)
    ax.text(0, 1.015, "Top 12 Order Regions by Net Sales | Bubble Size = Net Sales (USD M) | Color = Canceled/Fraud Rate (%)",
            transform=ax.transAxes, fontsize=8.5, color=GREY)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, linestyle=":", alpha=0.45)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, "04_regional_risk_heatmap.png"), dpi=200)
    plt.close()


def main():
    csv_path = locate_dataset()
    print(f"[1/4] Loading & cleaning dataset from: {csv_path}")
    df = load_and_clean(csv_path)

    print("[2/4] Running SQLite operational & risk diagnostics...")
    sql_tables = run_sql_diagnostics(df)
    for name, table in sql_tables.items():
        table.to_csv(os.path.join(SUMMARY_DIR, f"{name}.csv"), index=False)

    print("[3/4] Running ABC-XYZ inventory segmentation & discount leakage audit...")
    cat_df, abc_xyz_df, disc_df = compute_abc_xyz_and_discounts(df)
    cat_df.to_csv(os.path.join(SUMMARY_DIR, "category_abc_xyz.csv"), index=False)
    abc_xyz_df.to_csv(os.path.join(SUMMARY_DIR, "abc_xyz_summary.csv"), index=False)
    disc_df.to_csv(os.path.join(SUMMARY_DIR, "discount_tier_leakage.csv"), index=False)

    metrics = compute_baseline_metrics(df)
    with open(os.path.join(SUMMARY_DIR, "baseline_metrics.json"), "w") as f:
        json.dump(metrics, f, indent=2)

    print("[4/4] Generating 4 consulting-grade executive charts...")
    generate_charts(
        metrics,
        sql_tables["shipping_sla"],
        abc_xyz_df,
        disc_df,
        sql_tables["regional_risk"],
    )
    print(f"Done! Summary tables saved to {SUMMARY_DIR} and charts saved to {CHARTS_DIR}.")


if __name__ == "__main__":
    main()
