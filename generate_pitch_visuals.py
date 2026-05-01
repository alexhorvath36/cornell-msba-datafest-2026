from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from textwrap import fill


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "dataset" / "dataset"
OUT = ROOT / "pitch_deck_visuals"
OUT.mkdir(exist_ok=True)


PRIMARY = "#0B1F33"
ACCENT = "#D77A61"
ACCENT_2 = "#5B8E7D"
ACCENT_3 = "#E3B23C"
MUTED = "#7A8A99"
LIGHT = "#E9EEF2"


def configure_style() -> None:
    sns.set_theme(style="whitegrid")
    plt.rcParams.update(
        {
            "figure.facecolor": "white",
            "axes.facecolor": "white",
            "axes.edgecolor": LIGHT,
            "axes.labelcolor": PRIMARY,
            "axes.titlecolor": PRIMARY,
            "xtick.color": PRIMARY,
            "ytick.color": PRIMARY,
            "grid.color": LIGHT,
            "grid.linestyle": "-",
            "grid.linewidth": 0.8,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "font.size": 11,
            "axes.titlesize": 18,
            "axes.labelsize": 11,
            "figure.titlesize": 22,
        }
    )


def pick_mode(series: pd.Series):
    mode = series.mode(dropna=True)
    if len(mode):
        return mode.iloc[0]
    series = series.dropna()
    return series.iloc[0] if len(series) else np.nan


def add_header(fig: plt.Figure, title: str, subtitle: str) -> None:
    fig.text(
        0.06,
        0.965,
        fill(title, width=78),
        ha="left",
        va="top",
        fontsize=20,
        fontweight="bold",
        color=PRIMARY,
    )
    fig.text(
        0.06,
        0.895,
        fill(subtitle, width=120),
        ha="left",
        va="top",
        fontsize=10.5,
        color=MUTED,
    )


def add_panel_title(ax: plt.Axes, title: str) -> None:
    ax.set_title(title, loc="left", pad=10, fontsize=12, fontweight="bold")


def clean_axes(ax: plt.Axes) -> None:
    ax.spines["left"].set_color(LIGHT)
    ax.spines["bottom"].set_color(LIGHT)
    ax.tick_params(colors=PRIMARY)


def save_fig(fig: plt.Figure, name: str) -> None:
    fig.tight_layout(rect=(0, 0, 1, 0.84))
    fig.savefig(OUT / name, dpi=220, bbox_inches="tight")
    plt.close(fig)


def build_order_table() -> pd.DataFrame:
    orders = pd.read_csv(
        DATA / "marketplace_transactions.csv",
        parse_dates=[
            "purchase_timestamp",
            "approved_at",
            "shipped_date",
            "delivered_date",
            "estimated_delivery_date",
        ],
    )
    feedback = (
        pd.read_csv(DATA / "feedback.csv", parse_dates=["feedback_date"])
        .sort_values(["transaction_id", "feedback_date"])
        .drop_duplicates("transaction_id", keep="last")
    )
    details = pd.read_csv(DATA / "transaction_details.csv")
    items = pd.read_csv(DATA / "items.csv")
    translations = pd.read_csv(DATA / "item_category_translation.csv")
    buyers = pd.read_csv(DATA / "buyers.csv")
    artisans = pd.read_csv(DATA / "artisans.csv")

    line = details.merge(
        items[["item_id", "item_category_name", "item_weight_g"]],
        on="item_id",
        how="left",
    )
    line = line.merge(translations, on="item_category_name", how="left")
    line["category"] = line["item_category_name_english"].fillna(line["item_category_name"])
    line = line.merge(
        artisans[["artisan_id", "artisan_state", "artisan_city"]],
        on="artisan_id",
        how="left",
    )

    order_lines = (
        line.groupby("transaction_id")
        .agg(
            items=("transaction_item_id", "size"),
            artisans=("artisan_id", "nunique"),
            order_price=("price", "sum"),
            order_ship=("shipping_cost", "sum"),
            total_weight=("item_weight_g", "sum"),
            category=("category", pick_mode),
            artisan_id=("artisan_id", pick_mode),
            artisan_state=("artisan_state", pick_mode),
            artisan_city=("artisan_city", pick_mode),
        )
        .reset_index()
    )

    order = (
        orders.merge(order_lines, on="transaction_id", how="left")
        .merge(
            buyers[["buyer_id", "buyer_unique_id", "buyer_state", "buyer_city"]],
            on="buyer_id",
            how="left",
        )
        .merge(feedback[["transaction_id", "rating_score"]], on="transaction_id", how="left")
    )

    order["gross_value"] = order["order_price"] + order["order_ship"]
    order["delivery_days"] = (
        order["delivered_date"] - order["purchase_timestamp"]
    ).dt.total_seconds() / 86400
    order["late"] = order["delivered_date"] > order["estimated_delivery_date"]
    order["cross_state"] = order["buyer_state"] != order["artisan_state"]
    order["ship_share"] = order["order_ship"] / order["gross_value"]
    order["purchase_month"] = order["purchase_timestamp"].dt.to_period("M").astype(str)
    return order


def chart_growth(order: pd.DataFrame) -> None:
    monthly = (
        order.groupby("purchase_month")
        .agg(orders=("transaction_id", "size"), gmv=("gross_value", "sum"))
        .reset_index()
    )
    monthly = monthly[(monthly["purchase_month"] >= "2017-01") & (monthly["purchase_month"] <= "2018-08")]

    fig, ax = plt.subplots(figsize=(12.5, 5.5))
    ax.plot(monthly["purchase_month"], monthly["orders"], color=PRIMARY, lw=2.5, marker="o", ms=4)
    ax.fill_between(monthly["purchase_month"], monthly["orders"], color=PRIMARY, alpha=0.08)
    add_header(
        fig,
        "Marketplace demand scaled sharply through 2018",
        "Monthly order volume rose from 800 in Jan 2017 to 6,512 in Aug 2018, increasing operational pressure on fulfillment.",
    )
    ax.set_ylabel("Orders")
    ax.set_xlabel("")
    ax.tick_params(axis="x", rotation=70)
    ax.annotate(
        "Holiday spike",
        xy=("2017-11", 7544),
        xytext=("2017-12", 7900),
        arrowprops={"arrowstyle": "->", "color": ACCENT},
        color=ACCENT,
        fontsize=10,
    )
    ax.set_ylim(0, 8500)
    clean_axes(ax)
    save_fig(fig, "01_growth_in_orders.png")


def chart_cross_state(order: pd.DataFrame) -> None:
    perf = (
        order.dropna(subset=["cross_state"])
        .groupby("cross_state")
        .agg(
            avg_delivery=("delivery_days", "mean"),
            late_rate=("late", "mean"),
            avg_rating=("rating_score", "mean"),
        )
        .reset_index()
    )
    perf["segment"] = perf["cross_state"].map({False: "Same-state", True: "Cross-state"})

    fig, axes = plt.subplots(1, 3, figsize=(14, 4.8))
    metrics = [
        ("avg_delivery", "Delivery Days", PRIMARY, "{:.1f}"),
        ("late_rate", "Late Rate", ACCENT, "{:.1%}"),
        ("avg_rating", "Average Rating", ACCENT_2, "{:.2f}"),
    ]
    for ax, (col, title, color, fmt) in zip(axes, metrics):
        sns.barplot(
            data=perf,
            x="segment",
            y=col,
            hue="segment",
            palette=[ACCENT_2, color],
            legend=False,
            ax=ax,
        )
        add_panel_title(ax, title)
        ax.set_xlabel("")
        ax.set_ylabel("")
        for container in ax.containers:
            labels = [fmt.format(v) for v in container.datavalues]
            ax.bar_label(container, labels=labels, padding=4, color=PRIMARY, fontsize=10)
        clean_axes(ax)
    add_header(
        fig,
        "Cross-state lanes are slower, later, and rated worse",
        "64.1% of orders cross state lines. Those lanes average 15.1 delivery days versus 7.9 for same-state orders.",
    )
    save_fig(fig, "02_cross_state_penalty.png")


def chart_late_rating(order: pd.DataFrame) -> None:
    rated = order.dropna(subset=["rating_score"]).copy()
    rated["status"] = rated["late"].map({False: "On-time", True: "Late"})
    summary = rated.groupby("status").agg(avg_rating=("rating_score", "mean"), count=("transaction_id", "size")).reset_index()

    fig, ax = plt.subplots(figsize=(8.5, 5.2))
    sns.barplot(
        data=summary,
        x="status",
        y="avg_rating",
        hue="status",
        palette=[ACCENT_2, ACCENT],
        legend=False,
        ax=ax,
    )
    add_header(
        fig,
        "Late delivery is the clearest driver of customer dissatisfaction",
        "Average rating falls from 4.21 on on-time orders to 2.57 on late orders.",
    )
    ax.set_xlabel("")
    ax.set_ylabel("Average Rating")
    ax.set_ylim(0, 5)
    for container in ax.containers:
        ax.bar_label(container, labels=[f"{v:.2f}" for v in container.datavalues], padding=4, color=PRIMARY)
    clean_axes(ax)
    save_fig(fig, "03_late_delivery_rating_impact.png")


def chart_state_late_rates(order: pd.DataFrame) -> None:
    state = (
        order.groupby("buyer_state")
        .agg(
            orders=("transaction_id", "size"),
            late_rate=("late", "mean"),
            avg_rating=("rating_score", "mean"),
        )
        .query("orders >= 1500")
        .sort_values("late_rate", ascending=False)
        .reset_index()
    )

    fig, ax = plt.subplots(figsize=(11, 5.2))
    colors = [ACCENT if s in {"RJ", "BA", "ES", "PE", "SC"} else PRIMARY for s in state["buyer_state"]]
    sns.barplot(
        data=state,
        x="buyer_state",
        y="late_rate",
        hue="buyer_state",
        palette=dict(zip(state["buyer_state"], colors)),
        legend=False,
        ax=ax,
    )
    add_header(
        fig,
        "Late-delivery pain is concentrated in a handful of buyer states",
        "Rio de Janeiro and Bahia stand out as the largest high-friction states among major demand centers.",
    )
    ax.set_xlabel("Buyer State")
    ax.set_ylabel("Late Rate")
    ax.yaxis.set_major_formatter(lambda x, pos: f"{x:.0%}")
    for container in ax.containers:
        ax.bar_label(container, labels=[f"{v:.1%}" for v in container.datavalues], padding=3, color=PRIMARY, fontsize=9)
    clean_axes(ax)
    save_fig(fig, "04_state_late_rate_hotspots.png")


def chart_category_shipping(order: pd.DataFrame) -> None:
    cat = (
        order.groupby("category")
        .agg(
            orders=("transaction_id", "size"),
            ship_share=("order_ship", lambda s: s.sum() / order.loc[s.index, "gross_value"].sum()),
        )
        .sort_values("orders", ascending=False)
        .head(12)
        .sort_values("ship_share", ascending=False)
        .reset_index()
    )

    fig, ax = plt.subplots(figsize=(12, 5.5))
    sns.barplot(data=cat, x="category", y="ship_share", color=PRIMARY, ax=ax)
    add_header(
        fig,
        "Some high-volume categories carry a heavy shipping burden",
        "Electronics and telephony absorb unusually high shipping share, making fulfillment efficiency economically important.",
    )
    ax.set_xlabel("")
    ax.set_ylabel("Shipping Share of Order Value")
    ax.yaxis.set_major_formatter(lambda x, pos: f"{x:.0%}")
    ax.tick_params(axis="x", rotation=70)
    for container in ax.containers:
        ax.bar_label(container, labels=[f"{v:.0%}" for v in container.datavalues], padding=3, color=PRIMARY, fontsize=9)
    clean_axes(ax)
    save_fig(fig, "05_category_shipping_burden.png")


def chart_priority_lanes(order: pd.DataFrame) -> None:
    lane = (
        order.groupby(["artisan_id", "artisan_state", "buyer_state", "category"])
        .agg(
            orders=("transaction_id", "size"),
            gmv=("gross_value", "sum"),
            avg_delivery=("delivery_days", "mean"),
            late_rate=("late", "mean"),
            avg_rating=("rating_score", "mean"),
            ship_share=("ship_share", "mean"),
        )
        .reset_index()
    )
    lane = lane.query("orders >= 10").copy()
    lane["cross_state"] = lane["buyer_state"] != lane["artisan_state"]

    def minmax(series: pd.Series) -> pd.Series:
        series = series.astype(float)
        if series.max() == series.min():
            return pd.Series(0.0, index=series.index)
        return (series - series.min()) / (series.max() - series.min())

    lane["priority_score"] = (
        0.20 * minmax(np.log1p(lane["orders"]))
        + 0.20 * minmax(np.log1p(lane["gmv"]))
        + 0.25 * minmax(lane["late_rate"])
        + 0.15 * minmax(lane["avg_delivery"].fillna(lane["avg_delivery"].median()))
        + 0.10 * minmax(lane["ship_share"].fillna(lane["ship_share"].median()))
        + 0.05 * minmax(5 - lane["avg_rating"].fillna(lane["avg_rating"].median()))
        + 0.05 * lane["cross_state"].astype(int)
    )

    top = lane.sort_values("priority_score", ascending=False).head(10).copy()
    top["lane"] = top["category"] + " | " + top["artisan_state"] + " -> " + top["buyer_state"]
    top = top.sort_values("priority_score")

    fig, ax = plt.subplots(figsize=(12, 6))
    sns.barplot(data=top, y="lane", x="priority_score", color=ACCENT, ax=ax)
    add_header(
        fig,
        "Priority lanes show where intervention should start",
        "SP -> RJ appears repeatedly across categories, indicating a clear first-wave focus for logistics improvement.",
    )
    ax.set_xlabel("Priority Score")
    ax.set_ylabel("")
    for container in ax.containers:
        ax.bar_label(container, labels=[f"{v:.2f}" for v in container.datavalues], padding=4, color=PRIMARY, fontsize=9)
    clean_axes(ax)
    save_fig(fig, "06_priority_lanes.png")


def chart_solution_matrix(order: pd.DataFrame) -> None:
    state = (
        order.groupby("buyer_state")
        .agg(
            orders=("transaction_id", "size"),
            late_rate=("late", "mean"),
            avg_rating=("rating_score", "mean"),
        )
        .query("orders >= 1500")
        .reset_index()
    )
    sp_baseline = float(state.loc[state["buyer_state"] == "SP", "late_rate"].iloc[0])
    state["avoidable_lates"] = ((state["late_rate"] - sp_baseline).clip(lower=0) * state["orders"]).round()
    state["segment"] = np.where(state["avoidable_lates"] >= 100, "Intervene now", "Monitor")

    fig, ax = plt.subplots(figsize=(10.5, 6))
    sns.scatterplot(
        data=state,
        x="orders",
        y="late_rate",
        size="avoidable_lates",
        sizes=(80, 900),
        hue="segment",
        palette={"Intervene now": ACCENT, "Monitor": ACCENT_2},
        ax=ax,
        legend=False,
    )
    add_header(
        fig,
        "State intervention matrix points to a focused solution path",
        "Prioritize high-volume states with above-baseline late rates. Rio de Janeiro is the clearest first target.",
    )
    ax.set_xlabel("Order Volume")
    ax.set_ylabel("Late Rate")
    ax.yaxis.set_major_formatter(lambda x, pos: f"{x:.0%}")
    for _, row in state.iterrows():
        ax.text(row["orders"] + 120, row["late_rate"], row["buyer_state"], color=PRIMARY, fontsize=10)
    ax.axhline(sp_baseline, color=MUTED, linestyle="--", linewidth=1)
    ax.text(
        state["orders"].min(),
        sp_baseline + 0.002,
        "SP late-rate baseline",
        color=MUTED,
        fontsize=9,
    )
    clean_axes(ax)
    save_fig(fig, "07_state_intervention_matrix.png")


def main() -> None:
    configure_style()
    order = build_order_table()
    chart_growth(order)
    chart_cross_state(order)
    chart_late_rating(order)
    chart_state_late_rates(order)
    chart_category_shipping(order)
    chart_priority_lanes(order)
    chart_solution_matrix(order)
    print(f"Wrote visuals to {OUT}")


if __name__ == "__main__":
    main()
