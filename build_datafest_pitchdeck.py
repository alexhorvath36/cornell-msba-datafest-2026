from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parent
VISUALS = ROOT / "pitch_deck_visuals"
OUT = ROOT / "datafest_fulfillment_reliability_pitchdeck.pptx"

SLIDE_W = 13_333_333
SLIDE_H = 7_500_000

NAVY = "0B1F33"
MUTED = "6F7F8F"
LIGHT = "E9EEF2"
ACCENT = "D77A61"
GREEN = "5B8E7D"
GOLD = "E3B23C"
WHITE = "FFFFFF"
PALE = "F6F8FA"


def emu(inches: float) -> int:
    return int(inches * 914400)


def xml_text(text: str) -> str:
    return escape(text, {'"': "&quot;", "'": "&apos;"})


@dataclass
class TextBox:
    x: float
    y: float
    w: float
    h: float
    paragraphs: list[str]
    size: int = 18
    color: str = NAVY
    bold: bool = False
    name: str = "Text"


@dataclass
class Rect:
    x: float
    y: float
    w: float
    h: float
    fill: str
    line: str | None = None
    name: str = "Rectangle"


@dataclass
class Picture:
    path: Path
    x: float
    y: float
    w: float
    h: float
    name: str = "Picture"


@dataclass
class Slide:
    items: list[TextBox | Rect | Picture] = field(default_factory=list)


def bullet_lines(lines: Iterable[str]) -> list[str]:
    return [f"- {line}" for line in lines]


def text_shape(shape_id: int, box: TextBox) -> str:
    paragraphs = []
    for text in box.paragraphs:
        is_bullet = text.startswith("- ")
        body = text[2:] if is_bullet else text
        bullet_xml = (
            '<a:buChar char="•"/><a:buFont typeface="Aptos"/>'
            if is_bullet
            else '<a:buNone/>'
        )
        margin = ' marL="274320" indent="-171450"' if is_bullet else ""
        paragraphs.append(
            f"""
            <a:p>
              <a:pPr{margin}>{bullet_xml}</a:pPr>
              <a:r>
                <a:rPr lang="en-US" sz="{box.size * 100}"{' b="1"' if box.bold else ''}>
                  <a:solidFill><a:srgbClr val="{box.color}"/></a:solidFill>
                  <a:latin typeface="Aptos"/>
                </a:rPr>
                <a:t>{xml_text(body)}</a:t>
              </a:r>
              <a:endParaRPr lang="en-US" sz="{box.size * 100}"/>
            </a:p>
            """
        )

    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="{shape_id}" name="{xml_text(box.name)}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{emu(box.x)}" y="{emu(box.y)}"/><a:ext cx="{emu(box.w)}" cy="{emu(box.h)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" anchor="t"/>
        <a:lstStyle/>
        {''.join(paragraphs)}
      </p:txBody>
    </p:sp>
    """


def rect_shape(shape_id: int, rect: Rect) -> str:
    line = rect.line or rect.fill
    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="{shape_id}" name="{xml_text(rect.name)}"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{emu(rect.x)}" y="{emu(rect.y)}"/><a:ext cx="{emu(rect.w)}" cy="{emu(rect.h)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="{rect.fill}"/></a:solidFill>
        <a:ln><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>
      </p:spPr>
    </p:sp>
    """


def pic_shape(shape_id: int, pic: Picture, rel_id: str) -> str:
    return f"""
    <p:pic>
      <p:nvPicPr>
        <p:cNvPr id="{shape_id}" name="{xml_text(pic.name)}"/>
        <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip r:embed="{rel_id}"/>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="{emu(pic.x)}" y="{emu(pic.y)}"/><a:ext cx="{emu(pic.w)}" cy="{emu(pic.h)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>
    """


def slide_xml(slide: Slide, image_rel_ids: dict[Path, str]) -> str:
    shapes = [
        """
        <p:nvGrpSpPr>
          <p:cNvPr id="1" name=""/>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
        </p:grpSpPr>
        """
    ]
    shape_id = 2
    for item in slide.items:
        if isinstance(item, TextBox):
            shapes.append(text_shape(shape_id, item))
        elif isinstance(item, Rect):
            shapes.append(rect_shape(shape_id, item))
        elif isinstance(item, Picture):
            shapes.append(pic_shape(shape_id, item, image_rel_ids[item.path]))
        shape_id += 1

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      {''.join(shapes)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"""


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


def base_rect() -> Rect:
    return Rect(0, 0, 13.333, 7.5, WHITE, WHITE, "Background")


def title_box(text: str, y: float = 0.55, size: int = 34) -> TextBox:
    return TextBox(0.65, y, 12.0, 0.65, [text], size=size, color=NAVY, bold=True, name="Title")


def subtitle_box(text: str, y: float = 1.25) -> TextBox:
    return TextBox(0.68, y, 11.7, 0.45, [text], size=15, color=MUTED, name="Subtitle")


def slides() -> list[Slide]:
    img = lambda name: VISUALS / name

    return [
        Slide(
            [
                base_rect(),
                Rect(0, 0, 13.333, 0.18, ACCENT, ACCENT, "Accent Bar"),
                TextBox(
                    0.75,
                    1.05,
                    11.8,
                    1.1,
                    ["Fulfillment Reliability Is the Growth Bottleneck"],
                    size=38,
                    color=NAVY,
                    bold=True,
                ),
                TextBox(
                    0.78,
                    2.3,
                    10.9,
                    0.55,
                    [
                        "A Brazil-based marketplace can improve customer experience by prioritizing the seller-state-category lanes most responsible for late delivery."
                    ],
                    size=17,
                    color=MUTED,
                ),
                Rect(0.78, 3.45, 3.4, 1.35, PALE, LIGHT, "Metric Card"),
                TextBox(1.02, 3.72, 3.0, 0.5, ["64.1%"], size=30, color=ACCENT, bold=True),
                TextBox(1.04, 4.28, 2.8, 0.35, ["orders cross state lines"], size=13, color=NAVY),
                Rect(4.95, 3.45, 3.4, 1.35, PALE, LIGHT, "Metric Card"),
                TextBox(5.18, 3.72, 3.0, 0.5, ["2.57"], size=30, color=ACCENT, bold=True),
                TextBox(5.2, 4.28, 3.0, 0.35, ["avg rating when late"], size=13, color=NAVY),
                Rect(9.1, 3.45, 3.4, 1.35, PALE, LIGHT, "Metric Card"),
                TextBox(9.34, 3.72, 3.0, 0.5, ["SP -> RJ"], size=27, color=ACCENT, bold=True),
                TextBox(9.36, 4.28, 2.8, 0.35, ["first lane focus"], size=13, color=NAVY),
                TextBox(0.8, 6.55, 6.5, 0.35, ["Cornell MSBA DataFest 2026"], size=12, color=MUTED),
            ]
        ),
        Slide(
            [
                base_rect(),
                title_box("The Problem: Growth Exposed a Logistics Reliability Gap"),
                subtitle_box("Demand scaled quickly, but fulfillment quality is uneven across geography and seller-category lanes."),
                Picture(img("01_growth_in_orders.png"), 0.65, 1.95, 7.1, 4.0, "Growth Chart"),
                TextBox(
                    8.15,
                    2.05,
                    4.45,
                    2.1,
                    bullet_lines(
                        [
                            "Orders rose from 800 in Jan 2017 to 6,512 in Aug 2018.",
                            "Most orders are simple: one item and one artisan.",
                            "The bottleneck is fulfillment reliability, not basket complexity.",
                        ]
                    ),
                    size=15,
                    color=NAVY,
                ),
                Rect(8.15, 4.75, 4.4, 1.0, PALE, LIGHT, "Callout"),
                TextBox(8.4, 5.02, 4.0, 0.4, ["Problem statement"], size=13, color=MUTED, bold=True),
                TextBox(
                    8.4,
                    5.38,
                    3.9,
                    0.45,
                    ["High-value cross-state orders are slower, later, and less satisfying."],
                    size=15,
                    color=NAVY,
                ),
            ]
        ),
        Slide(
            [
                base_rect(),
                title_box("Summary Statistic 1: Cross-State Delivery Carries the Penalty"),
                subtitle_box("Cross-state orders are more valuable, but they take almost twice as long and have worse service outcomes."),
                Picture(img("02_cross_state_penalty.png"), 0.65, 1.75, 12.1, 4.65, "Cross State Penalty"),
                TextBox(
                    0.8,
                    6.65,
                    11.6,
                    0.35,
                    ["Interpretation: the marketplace should target fulfillment lanes, not just aggregate platform metrics."],
                    size=13,
                    color=MUTED,
                ),
            ]
        ),
        Slide(
            [
                base_rect(),
                title_box("Summary Statistic 2: Late Delivery Is the Satisfaction Driver"),
                subtitle_box("The ratings penalty is large enough to make delivery reliability a customer retention issue."),
                Picture(img("03_late_delivery_rating_impact.png"), 0.65, 1.85, 5.7, 4.25, "Rating Impact"),
                Picture(img("04_state_late_rate_hotspots.png"), 6.75, 1.85, 5.9, 4.25, "State Hotspots"),
                TextBox(
                    0.8,
                    6.55,
                    11.7,
                    0.4,
                    ["Late orders average 2.57 stars versus 4.21 for on-time orders; RJ, BA, ES, PE, and SC are the largest high-friction states."],
                    size=13,
                    color=MUTED,
                ),
            ]
        ),
        Slide(
            [
                base_rect(),
                title_box("Methodology: Build the Decision Unit Around Fulfillment Lanes"),
                subtitle_box("The analysis joins the marketplace tables into order and lane-level views, then ranks operational pain by business value."),
                Rect(0.75, 2.0, 3.55, 3.6, PALE, LIGHT, "Step Card"),
                TextBox(1.0, 2.25, 3.0, 0.45, ["1. Join"], size=22, color=ACCENT, bold=True),
                TextBox(
                    1.0,
                    2.9,
                    3.0,
                    1.6,
                    bullet_lines(
                        [
                            "transactions",
                            "line items",
                            "buyers and artisans",
                            "items and categories",
                            "payments and feedback",
                        ]
                    ),
                    size=14,
                    color=NAVY,
                ),
                Rect(4.9, 2.0, 3.55, 3.6, PALE, LIGHT, "Step Card"),
                TextBox(5.15, 2.25, 3.0, 0.45, ["2. Measure"], size=22, color=ACCENT, bold=True),
                TextBox(
                    5.15,
                    2.9,
                    3.0,
                    1.6,
                    bullet_lines(
                        [
                            "delivery days",
                            "late rate",
                            "rating score",
                            "shipping share",
                            "cross-state status",
                        ]
                    ),
                    size=14,
                    color=NAVY,
                ),
                Rect(9.05, 2.0, 3.55, 3.6, PALE, LIGHT, "Step Card"),
                TextBox(9.3, 2.25, 3.0, 0.45, ["3. Prioritize"], size=22, color=ACCENT, bold=True),
                TextBox(
                    9.3,
                    2.9,
                    3.0,
                    1.6,
                    bullet_lines(
                        [
                            "seller-state-category lanes",
                            "volume and GMV",
                            "customer pain",
                            "logistics burden",
                        ]
                    ),
                    size=14,
                    color=NAVY,
                ),
                TextBox(
                    0.82,
                    6.35,
                    11.6,
                    0.45,
                    ["Output: a ranked list of where limited logistics resources should be deployed first."],
                    size=14,
                    color=MUTED,
                ),
            ]
        ),
        Slide(
            [
                base_rect(),
                title_box("Recommendation: Launch a Fulfillment Reliability Program"),
                subtitle_box("Use targeted interventions on the worst lanes before investing in broad platform-wide fixes."),
                Rect(0.75, 1.95, 5.85, 4.55, PALE, LIGHT, "Program"),
                TextBox(1.05, 2.28, 5.2, 0.45, ["What changes"], size=22, color=ACCENT, bold=True),
                TextBox(
                    1.05,
                    3.0,
                    5.15,
                    2.4,
                    bullet_lines(
                        [
                            "Seller scorecards for late-rate performance",
                            "Premium handling for high-risk lanes",
                            "Local-first ranking where substitutes exist",
                            "Carrier and SLA review for high-friction states",
                            "First-order protection for new buyers",
                        ]
                    ),
                    size=15,
                    color=NAVY,
                ),
                Rect(7.0, 1.95, 5.6, 4.55, PALE, LIGHT, "Why"),
                TextBox(7.3, 2.28, 5.0, 0.45, ["Why this is the right lever"], size=22, color=ACCENT, bold=True),
                TextBox(
                    7.3,
                    3.0,
                    4.9,
                    2.4,
                    bullet_lines(
                        [
                            "Late delivery is the strongest visible ratings driver",
                            "Repeat purchase is only 3.12%",
                            "Pain is concentrated by state and lane",
                            "Top sellers drive a large share of marketplace revenue",
                        ]
                    ),
                    size=15,
                    color=NAVY,
                ),
            ]
        ),
        Slide(
            [
                base_rect(),
                title_box("Where to Start: SP -> RJ Is the First-Wave Focus"),
                subtitle_box("The top priority lanes repeatedly point to Rio de Janeiro demand served from Sao Paulo sellers."),
                Picture(img("06_priority_lanes.png"), 0.65, 1.75, 8.0, 4.75, "Priority Lanes"),
                Rect(9.0, 1.95, 3.65, 4.25, PALE, LIGHT, "Action Card"),
                TextBox(9.28, 2.25, 3.1, 0.45, ["First actions"], size=21, color=ACCENT, bold=True),
                TextBox(
                    9.28,
                    2.95,
                    3.05,
                    2.3,
                    bullet_lines(
                        [
                            "Audit SP -> RJ carriers and seller dispatch times",
                            "Prioritize garden tools, office furniture, bed/bath, electronics, and watches/gifts",
                            "Apply premium handling to the worst lanes",
                            "Measure change in late rate and ratings",
                        ]
                    ),
                    size=13,
                    color=NAVY,
                ),
            ]
        ),
        Slide(
            [
                base_rect(),
                Rect(0, 0, 13.333, 0.18, ACCENT, ACCENT, "Accent Bar"),
                title_box("Key Takeaways and Q&A", y=0.85, size=34),
                TextBox(
                    0.85,
                    2.05,
                    11.4,
                    3.1,
                    bullet_lines(
                        [
                            "This marketplace's growth challenge is a logistics reliability problem.",
                            "Cross-state fulfillment is common, slower, and more likely to be late.",
                            "Late delivery sharply reduces ratings, making this a retention risk.",
                            "The solution is a lane-level prioritization system and targeted fulfillment intervention.",
                            "Start with SP -> RJ lanes, then expand to other high-friction states.",
                        ]
                    ),
                    size=19,
                    color=NAVY,
                ),
                TextBox(0.88, 6.15, 4.0, 0.6, ["Q&A"], size=32, color=ACCENT, bold=True),
            ]
        ),
    ]


def content_types(num_slides: int, image_names: list[str]) -> str:
    image_defaults = "\n".join(
        '<Default Extension="png" ContentType="image/png"/>' for _ in sorted(set(image_names[:1]))
    )
    slide_overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, num_slides + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {image_defaults}
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {slide_overrides}
</Types>
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


def simple_master_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>
"""


def simple_layout_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
"""


def theme_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="DataFest">
  <a:themeElements>
    <a:clrScheme name="DataFest">
      <a:dk1><a:srgbClr val="0B1F33"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2D3D"/></a:dk2><a:lt2><a:srgbClr val="F6F8FA"/></a:lt2>
      <a:accent1><a:srgbClr val="D77A61"/></a:accent1><a:accent2><a:srgbClr val="5B8E7D"/></a:accent2>
      <a:accent3><a:srgbClr val="E3B23C"/></a:accent3><a:accent4><a:srgbClr val="7A8A99"/></a:accent4>
      <a:accent5><a:srgbClr val="E9EEF2"/></a:accent5><a:accent6><a:srgbClr val="0B1F33"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="DataFest"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>
"""


def write_deck() -> None:
    deck = slides()
    all_images: list[Path] = []
    for slide in deck:
        for item in slide.items:
            if isinstance(item, Picture) and item.path not in all_images:
                all_images.append(item.path)

    with ZipFile(OUT, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(deck), [p.name for p in all_images]))
        z.writestr(
            "_rels/.rels",
            rels_xml(
                [
                    (
                        "rId1",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                        "ppt/presentation.xml",
                    ),
                    (
                        "rId2",
                        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                        "docProps/core.xml",
                    ),
                    (
                        "rId3",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                        "docProps/app.xml",
                    ),
                ]
            ),
        )
        z.writestr(
            "docProps/core.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Fulfillment Reliability Pitch Deck</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
</cp:coreProperties>
""",
        )
        z.writestr(
            "docProps/app.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>{len(deck)}</Slides>
</Properties>
""",
        )
        z.writestr("ppt/presentation.xml", presentation_xml(len(deck)))
        z.writestr(
            "ppt/_rels/presentation.xml.rels",
            rels_xml(
                [("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "slideMasters/slideMaster1.xml")]
                + [
                    (
                        f"rId{i + 1}",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                        f"slides/slide{i}.xml",
                    )
                    for i in range(1, len(deck) + 1)
                ]
            ),
        )
        z.writestr("ppt/slideMasters/slideMaster1.xml", simple_master_xml())
        z.writestr(
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            rels_xml(
                [
                    ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", "../slideLayouts/slideLayout1.xml"),
                    ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "../theme/theme1.xml"),
                ]
            ),
        )
        z.writestr("ppt/slideLayouts/slideLayout1.xml", simple_layout_xml())
        z.writestr(
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            rels_xml([("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", "../slideMasters/slideMaster1.xml")]),
        )
        z.writestr("ppt/theme/theme1.xml", theme_xml())

        media_ids = {path: idx + 1 for idx, path in enumerate(all_images)}
        for path, idx in media_ids.items():
            z.write(path, f"ppt/media/image{idx}.png")

        for i, slide in enumerate(deck, start=1):
            slide_images = [item.path for item in slide.items if isinstance(item, Picture)]
            image_rels = {path: f"rId{idx + 2}" for idx, path in enumerate(slide_images)}
            z.writestr(f"ppt/slides/slide{i}.xml", slide_xml(slide, image_rels))
            rels = [
                (
                    "rId1",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                    "../slideLayouts/slideLayout1.xml",
                )
            ]
            rels.extend(
                (
                    image_rels[path],
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    f"../media/image{media_ids[path]}.png",
                )
                for path in slide_images
            )
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", rels_xml(rels))

    print(f"Wrote {OUT}")


if __name__ == "__main__":
    write_deck()
