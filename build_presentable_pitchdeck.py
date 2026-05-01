from __future__ import annotations

from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "dataset" / "dataset"
SLIDE_DIR = ROOT / "pitch_deck_slides"
OUT = ROOT / "datafest_fulfillment_reliability_pitchdeck.pptx"
SLIDE_DIR.mkdir(exist_ok=True)

SLIDE_W = 12_192_000
SLIDE_H = 6_858_000

NAVY = "#0B1F33"
MUTED = "#6F7F8F"
LIGHT = "#E9EEF2"
PALE = "#F6F8FA"
ACCENT = "#D77A61"
GREEN = "#5B8E7D"
GOLD = "#E3B23C"
WHITE = "#FFFFFF"


def xml_text(text: str) -> str:
    return escape(text, {'"': "&quot;", "'": "&apos;"})


def emu(inches: float) -> int:
    return int(inches * 914400)


def configure_style() -> None:
    sns.set_theme(style="whitegrid")
    plt.rcParams.update(
        {
            "figure.facecolor": WHITE,
            "axes.facecolor": WHITE,
            "axes.edgecolor": LIGHT,
            "axes.labelcolor": NAVY,
            "xtick.color": NAVY,
            "ytick.color": NAVY,
            "grid.color": LIGHT,
            "grid.linewidth": 0.8,
            "font.family": "DejaVu Sans",
            "axes.spines.top": False,
            "axes.spines.right": False,
        }
    )


def pick_mode(series: pd.Series):
    mode = series.mode(dropna=True)
    if len(mode):
        return mode.iloc[0]
    series = series.dropna()
    return series.iloc[0] if len(series) else np.nan


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
    line = line.merge(artisans[["artisan_id", "artisan_state"]], on="artisan_id", how="left")

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
        )
        .reset_index()
    )

    order = (
        orders.merge(order_lines, on="transaction_id", how="left")
        .merge(buyers[["buyer_id", "buyer_unique_id", "buyer_state"]], on="buyer_id", how="left")
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


def new_slide() -> tuple[plt.Figure, plt.Axes]:
    fig = plt.figure(figsize=(16, 9), dpi=160)
    fig.patch.set_facecolor(WHITE)
    canvas = fig.add_axes([0, 0, 1, 1])
    canvas.axis("off")
    return fig, canvas


def header(fig: plt.Figure, title: str, subtitle: str = "") -> None:
    fig.text(0.055, 0.91, title, ha="left", va="top", fontsize=26, fontweight="bold", color=NAVY)
    if subtitle:
        fig.text(0.057, 0.835, subtitle, ha="left", va="top", fontsize=12.5, color=MUTED)


def metric_card(fig: plt.Figure, x: float, y: float, value: str, label: str) -> None:
    ax = fig.add_axes([x, y, 0.24, 0.17])
    ax.set_facecolor(PALE)
    for spine in ax.spines.values():
        spine.set_color(LIGHT)
    ax.set_xticks([])
    ax.set_yticks([])
    ax.text(0.06, 0.64, value, transform=ax.transAxes, fontsize=25, fontweight="bold", color=ACCENT)
    ax.text(0.06, 0.24, label, transform=ax.transAxes, fontsize=11, color=NAVY)


def save_slide(fig: plt.Figure, number: int) -> Path:
    path = SLIDE_DIR / f"slide_{number:02d}.png"
    fig.savefig(path, dpi=160, facecolor=WHITE)
    plt.close(fig)
    return path


def slide_title(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    fig.patches.append(plt.Rectangle((0, 0.975), 1, 0.025, transform=fig.transFigure, color=ACCENT))
    fig.text(
        0.06,
        0.74,
        "Fulfillment Reliability\nIs the Growth Bottleneck",
        ha="left",
        va="top",
        fontsize=36,
        fontweight="bold",
        color=NAVY,
        linespacing=1.05,
    )
    fig.text(
        0.062,
        0.49,
        "A Brazil-based marketplace can improve customer experience by prioritizing the seller-state-category lanes most responsible for late delivery.",
        ha="left",
        va="top",
        fontsize=15,
        color=MUTED,
        wrap=True,
    )
    metric_card(fig, 0.06, 0.22, "64.1%", "orders cross state lines")
    metric_card(fig, 0.38, 0.22, "2.57", "avg rating when late")
    metric_card(fig, 0.70, 0.22, "SP -> RJ", "first lane focus")
    fig.text(0.06, 0.075, "Cornell MSBA DataFest 2026", fontsize=10.5, color=MUTED)
    return save_slide(fig, 1)


def slide_growth(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    header(
        fig,
        "The problem: growth exposed a logistics reliability gap",
        "Demand scaled quickly, while fulfillment quality remained uneven across geographies.",
    )
    monthly = (
        order.groupby("purchase_month")
        .agg(orders=("transaction_id", "size"))
        .reset_index()
        .query("'2017-01' <= purchase_month <= '2018-08'")
    )
    ax = fig.add_axes([0.07, 0.19, 0.58, 0.56])
    ax.plot(monthly["purchase_month"], monthly["orders"], color=NAVY, lw=3, marker="o", ms=4)
    ax.fill_between(monthly["purchase_month"], monthly["orders"], color=NAVY, alpha=0.08)
    ax.set_ylabel("Orders", color=NAVY)
    ax.set_xlabel("")
    ax.tick_params(axis="x", rotation=65, labelsize=8)
    ax.set_ylim(0, 8500)
    ax.annotate(
        "Holiday spike",
        xy=("2017-11", 7544),
        xytext=("2017-13", 7900),
        arrowprops={"arrowstyle": "->", "color": ACCENT},
        color=ACCENT,
        fontsize=10,
    )
    ax2 = fig.add_axes([0.71, 0.23, 0.24, 0.48])
    ax2.axis("off")
    bullets = [
        "Orders rose from 800 in Jan 2017 to 6,512 in Aug 2018.",
        "90.1% of orders have one item.",
        "98.7% of orders involve one artisan.",
        "The bottleneck is fulfillment reliability, not basket complexity.",
    ]
    for i, text in enumerate(bullets):
        ax2.text(0.0, 1 - i * 0.22, f"• {text}", fontsize=13.5, color=NAVY, va="top", wrap=True)
    return save_slide(fig, 2)


def slide_cross_state(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    header(
        fig,
        "Summary statistic: cross-state delivery carries the penalty",
        "Cross-state orders are common, slower, later, and rated worse.",
    )
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
    charts = [
        ("avg_delivery", "Delivery Days", "{:.1f}", [GREEN, NAVY]),
        ("late_rate", "Late Rate", "{:.1%}", [GREEN, ACCENT]),
        ("avg_rating", "Average Rating", "{:.2f}", [GREEN, GREEN]),
    ]
    for i, (col, title, fmt, colors) in enumerate(charts):
        ax = fig.add_axes([0.07 + i * 0.305, 0.22, 0.24, 0.5])
        sns.barplot(data=perf, x="segment", y=col, hue="segment", palette=colors, legend=False, ax=ax)
        ax.set_title(title, loc="left", fontsize=15, fontweight="bold", color=NAVY, pad=12)
        ax.set_xlabel("")
        ax.set_ylabel("")
        ax.tick_params(axis="x", labelsize=11)
        for container in ax.containers:
            ax.bar_label(container, labels=[fmt.format(v) for v in container.datavalues], padding=4, fontsize=11, color=NAVY)
    fig.text(0.07, 0.105, "Interpretation: target fulfillment lanes, not only aggregate marketplace metrics.", fontsize=12, color=MUTED)
    return save_slide(fig, 3)


def slide_ratings_states(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    header(
        fig,
        "Summary statistic: late delivery is the satisfaction driver",
        "The ratings penalty makes delivery reliability a customer retention issue.",
    )
    rated = order.dropna(subset=["rating_score"]).copy()
    rated["status"] = rated["late"].map({False: "On-time", True: "Late"})
    rating = rated.groupby("status").agg(avg_rating=("rating_score", "mean")).reset_index()
    ax1 = fig.add_axes([0.07, 0.23, 0.36, 0.48])
    sns.barplot(data=rating, x="status", y="avg_rating", hue="status", palette=[GREEN, ACCENT], legend=False, ax=ax1)
    ax1.set_title("Rating impact", loc="left", fontsize=15, fontweight="bold", color=NAVY, pad=12)
    ax1.set_xlabel("")
    ax1.set_ylabel("Average rating")
    ax1.set_ylim(0, 5)
    for container in ax1.containers:
        ax1.bar_label(container, labels=[f"{v:.2f}" for v in container.datavalues], padding=4, fontsize=11, color=NAVY)

    state = (
        order.groupby("buyer_state")
        .agg(orders=("transaction_id", "size"), late_rate=("late", "mean"))
        .query("orders >= 1500")
        .sort_values("late_rate", ascending=False)
        .reset_index()
    )
    ax2 = fig.add_axes([0.51, 0.23, 0.42, 0.48])
    palette = {s: ACCENT if s in {"RJ", "BA", "ES", "PE", "SC"} else NAVY for s in state["buyer_state"]}
    sns.barplot(data=state, x="buyer_state", y="late_rate", hue="buyer_state", palette=palette, legend=False, ax=ax2)
    ax2.set_title("Late-rate hotspots", loc="left", fontsize=15, fontweight="bold", color=NAVY, pad=12)
    ax2.set_xlabel("Buyer state")
    ax2.set_ylabel("Late rate")
    ax2.yaxis.set_major_formatter(lambda x, pos: f"{x:.0%}")
    ax2.tick_params(axis="x", labelsize=10)
    fig.text(0.07, 0.105, "Late orders average 2.57 stars versus 4.21 for on-time orders.", fontsize=12, color=MUTED)
    return save_slide(fig, 4)


def slide_methodology(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    header(
        fig,
        "Methodology: build the decision unit around fulfillment lanes",
        "Raw marketplace tables were converted into order-level and lane-level views.",
    )
    cards = [
        ("1. Join", ["transactions", "line items", "buyers and artisans", "items and categories", "feedback"]),
        ("2. Measure", ["delivery days", "late rate", "rating score", "shipping share", "cross-state status"]),
        ("3. Prioritize", ["seller-state-category lanes", "volume and GMV", "customer pain", "logistics burden"]),
    ]
    for i, (title, bullets) in enumerate(cards):
        ax = fig.add_axes([0.07 + i * 0.31, 0.24, 0.25, 0.45])
        ax.set_facecolor(PALE)
        for spine in ax.spines.values():
            spine.set_color(LIGHT)
        ax.set_xticks([])
        ax.set_yticks([])
        ax.text(0.08, 0.82, title, transform=ax.transAxes, fontsize=19, fontweight="bold", color=ACCENT)
        for j, bullet in enumerate(bullets):
            ax.text(0.08, 0.63 - j * 0.105, f"• {bullet}", transform=ax.transAxes, fontsize=12.5, color=NAVY)
    fig.text(0.07, 0.12, "Output: a ranked list of where limited logistics resources should be deployed first.", fontsize=12.5, color=MUTED)
    return save_slide(fig, 5)


def slide_recommendation(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    header(
        fig,
        "Recommendation: launch a Fulfillment Reliability Program",
        "Use targeted interventions on the worst lanes before broad platform-wide fixes.",
    )
    sections = [
        ("What changes", ["Seller scorecards for late-rate performance", "Premium handling for high-risk lanes", "Local-first ranking where substitutes exist", "Carrier and SLA review for high-friction states", "First-order protection for new buyers"]),
        ("Why this lever", ["Late delivery is the strongest visible ratings driver", "Repeat purchase is only 3.12%", "Pain is concentrated by state and lane", "Top sellers drive a large share of marketplace revenue"]),
    ]
    for i, (title, bullets) in enumerate(sections):
        ax = fig.add_axes([0.08 + i * 0.46, 0.23, 0.38, 0.5])
        ax.set_facecolor(PALE)
        for spine in ax.spines.values():
            spine.set_color(LIGHT)
        ax.set_xticks([])
        ax.set_yticks([])
        ax.text(0.07, 0.84, title, transform=ax.transAxes, fontsize=20, fontweight="bold", color=ACCENT)
        for j, bullet in enumerate(bullets):
            ax.text(0.07, 0.65 - j * 0.12, f"• {bullet}", transform=ax.transAxes, fontsize=12.5, color=NAVY, wrap=True)
    return save_slide(fig, 6)


def priority_lanes(order: pd.DataFrame) -> pd.DataFrame:
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
    return lane.sort_values("priority_score", ascending=False).head(10)


def slide_priorities(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    header(
        fig,
        "Where to start: SP -> RJ is the first-wave focus",
        "The top priority lanes repeatedly point to Rio de Janeiro demand served from Sao Paulo sellers.",
    )
    top = priority_lanes(order).copy()
    short_names = {
        "fashion_bags_accessories": "fashion bags",
        "computers_accessories": "computers",
        "bed_bath_table": "bed/bath",
        "office_furniture": "office furn.",
        "health_beauty": "health/beauty",
        "watches_gifts": "watches",
        "garden_tools": "garden",
    }
    top["category_short"] = top["category"].map(short_names).fillna(top["category"])
    top["lane"] = top["category_short"] + " | " + top["artisan_state"] + " -> " + top["buyer_state"]
    top = top.sort_values("priority_score")
    ax = fig.add_axes([0.17, 0.17, 0.48, 0.58])
    sns.barplot(data=top, y="lane", x="priority_score", color=ACCENT, ax=ax)
    ax.set_xlabel("Priority score")
    ax.set_ylabel("")
    ax.tick_params(axis="y", labelsize=10)
    for container in ax.containers:
        ax.bar_label(container, labels=[f"{v:.2f}" for v in container.datavalues], padding=4, fontsize=9, color=NAVY)
    ax2 = fig.add_axes([0.72, 0.22, 0.22, 0.47])
    ax2.axis("off")
    for i, action in enumerate(["Audit SP -> RJ carriers and seller dispatch times", "Prioritize garden tools, office furniture, bed/bath, electronics, and watches/gifts", "Apply premium handling to the worst lanes", "Measure change in late rate and ratings"]):
        ax2.text(0, 1 - i * 0.22, f"• {action}", fontsize=12.2, color=NAVY, va="top", wrap=True)
    return save_slide(fig, 7)


def slide_takeaways(order: pd.DataFrame) -> Path:
    fig, _ = new_slide()
    fig.patches.append(plt.Rectangle((0, 0.975), 1, 0.025, transform=fig.transFigure, color=ACCENT))
    header(fig, "Key takeaways and Q&A")
    takeaways = [
        "This marketplace's growth challenge is a logistics reliability problem.",
        "Cross-state fulfillment is common, slower, and more likely to be late.",
        "Late delivery sharply reduces ratings.",
        "The solution is lane-level prioritization plus targeted fulfillment intervention.",
        "Start with SP -> RJ, then expand to other high-friction states.",
    ]
    ax = fig.add_axes([0.08, 0.25, 0.75, 0.45])
    ax.axis("off")
    for i, text in enumerate(takeaways):
        ax.text(0, 1 - i * 0.17, f"• {text}", fontsize=17, color=NAVY, va="top", wrap=True)
    fig.text(0.08, 0.12, "Q&A", fontsize=31, fontweight="bold", color=ACCENT)
    return save_slide(fig, 8)


def rels_xml(rels: list[tuple[str, str, str]]) -> str:
    rows = "\n".join(
        f'<Relationship Id="{rid}" Type="{typ}" Target="{xml_text(target)}"/>'
        for rid, typ, target in rels
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rows}
</Relationships>
"""


def content_types(num_slides: int) -> str:
    slides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, num_slides + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {slides}
</Types>
"""


def slide_xml(image_id: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="2" name="Slide Image"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SLIDE_W}" cy="{SLIDE_H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"""


def presentation_xml(num_slides: int) -> str:
    slide_ids = "\n".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, num_slides + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"""


def master_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>
"""


def layout_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
"""


def theme_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="DataFest">
  <a:themeElements>
    <a:clrScheme name="DataFest"><a:dk1><a:srgbClr val="0B1F33"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F2D3D"/></a:dk2><a:lt2><a:srgbClr val="F6F8FA"/></a:lt2><a:accent1><a:srgbClr val="D77A61"/></a:accent1><a:accent2><a:srgbClr val="5B8E7D"/></a:accent2><a:accent3><a:srgbClr val="E3B23C"/></a:accent3><a:accent4><a:srgbClr val="7A8A99"/></a:accent4><a:accent5><a:srgbClr val="E9EEF2"/></a:accent5><a:accent6><a:srgbClr val="0B1F33"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="DataFest"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>
"""


def build_pptx(slides: list[Path]) -> None:
    with ZipFile(OUT, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(slides)))
        z.writestr(
            "_rels/.rels",
            rels_xml(
                [
                    ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "ppt/presentation.xml"),
                    ("rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
                    ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"),
                ]
            ),
        )
        z.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        z.writestr(
            "ppt/_rels/presentation.xml.rels",
            rels_xml(
                [("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "slideMasters/slideMaster1.xml")]
                + [(f"rId{i + 1}", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", f"slides/slide{i}.xml") for i in range(1, len(slides) + 1)]
            ),
        )
        z.writestr("ppt/slideMasters/slideMaster1.xml", master_xml())
        z.writestr(
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            rels_xml(
                [
                    ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
                    ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "../theme/theme1.xml"),
                ]
            ),
        )
        z.writestr("ppt/slideLayouts/slideLayout1.xml", layout_xml())
        z.writestr(
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            rels_xml([("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "../slideMasters/slideMaster1.xml")]),
        )
        z.writestr("ppt/theme/theme1.xml", theme_xml())
        z.writestr(
            "docProps/core.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Fulfillment Reliability Pitch Deck</dc:title><dc:creator>Codex</dc:creator></cp:coreProperties>""",
        )
        z.writestr(
            "docProps/app.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>{len(slides)}</Slides></Properties>""",
        )
        for i, path in enumerate(slides, start=1):
            z.write(path, f"ppt/media/image{i}.png")
            z.writestr(f"ppt/slides/slide{i}.xml", slide_xml(i))
            z.writestr(
                f"ppt/slides/_rels/slide{i}.xml.rels",
                rels_xml(
                    [
                        ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
                        ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", f"../media/image{i}.png"),
                    ]
                ),
            )


def main() -> None:
    configure_style()
    order = build_order_table()
    slide_paths = [
        slide_title(order),
        slide_growth(order),
        slide_cross_state(order),
        slide_ratings_states(order),
        slide_methodology(order),
        slide_recommendation(order),
        slide_priorities(order),
        slide_takeaways(order),
    ]
    build_pptx(slide_paths)
    print(f"Wrote {OUT}")
    print(f"Wrote slide images to {SLIDE_DIR}")


if __name__ == "__main__":
    main()
