"""Generates ppt-toolkit/templates/white_template.pptx and
wb_branded_template.pptx from scratch as valid, minimal OOXML packages.

Both templates are built from the exact same code path and the shared
`LAYOUTS` spec in `src/layouts.py`, so they are guaranteed to expose the same
5 named slide layouts with identical placeholder types/positions -- they
differ only in theme (colors extracted from the two source decks the user
provided: `templates/_source_white.pptx` = Slideworks consulting template,
`templates/_source_wb.pptx` = World Bank "CLIMATE WAREHOUSE 2" theme) and a
handful of per-layout decorative shapes (accent bars / full-bleed brand
backgrounds).

Run: python tools/build_templates.py
"""

from __future__ import annotations

import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from layouts import LAYOUTS, LAYOUT_ORDER, SLIDE_WIDTH, SLIDE_HEIGHT  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.normpath(os.path.join(HERE, "..", "templates"))

NS = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
)

THEMES = {
    "white": {
        "clr_scheme_name": "Slideworks",
        "colors": {
            "dk1": "061F32",
            "lt1": "FFFFFF",
            "dk2": "1D4769",
            "lt2": "F2F2F2",
            "accent1": "22A3DF",
            "accent2": "3478AB",
            "accent3": "1D4769",
            "accent4": "061F32",
            "accent5": "2222F6",
            "accent6": "F8BC3B",
            "hlink": "3478AB",
            "folHlink": "8B78AB",
        },
        "major_font": "Arial",
        "minor_font": "Arial",
        "brand_dark": "061F32",  # Section Divider / Closing full-bleed bg
        "accent_bar": "22A3DF",  # top accent bar on light layouts
    },
    "wb": {
        "clr_scheme_name": "CLIMATE WAREHOUSE 2",
        "colors": {
            "dk1": "494949",
            "lt1": "FBFCFF",
            "dk2": "183153",
            "lt2": "F8F9F8",
            "accent1": "183153",
            "accent2": "004B87",
            "accent3": "049AD7",
            "accent4": "00955E",
            "accent5": "FCC402",
            "accent6": "E7485F",
            "hlink": "611F41",
            "folHlink": "00955E",
        },
        "major_font": "Calibri",
        "minor_font": "Calibri",
        "brand_dark": "E7485F",  # WBG Prosperity template's pink/coral accent
        "accent_bar": "183153",  # top accent bar on light layouts (navy)
    },
}


def content_types_xml() -> str:
    layout_overrides = "".join(
        f'<Override PartName="/ppt/slideLayouts/slideLayout{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        for i in range(1, len(LAYOUT_ORDER) + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
{layout_overrides}
<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""


def root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def core_xml(title: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>{title}</dc:title>
<dc:creator>ppt-toolkit</dc:creator>
</cp:coreProperties>"""


def app_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>ppt-toolkit</Application>
</Properties>"""


def theme_xml(theme: dict) -> str:
    c = theme["colors"]

    def clr(tag):
        return f'<a:{tag}><a:srgbClr val="{c[tag]}"/></a:{tag}>'

    scheme = "".join(clr(k) for k in ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"])
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{theme['clr_scheme_name']}">
<a:themeElements>
<a:clrScheme name="{theme['clr_scheme_name']}">{scheme}</a:clrScheme>
<a:fontScheme name="{theme['clr_scheme_name']}">
<a:majorFont><a:latin typeface="{theme['major_font']}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="{theme['minor_font']}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>"""


def _text_style_block() -> str:
    return """<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>"""


def default_text_style_xml() -> str:
    return f"<p:defaultTextStyle>{_text_style_block()}</p:defaultTextStyle>"


def presentation_xml() -> str:
    layout_rels = "".join(f'<p:sldMasterId id="{2147483648 + 1}" r:id="rId1"/>' for _ in [0])
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation {NS}>
<p:sldMasterIdLst>{layout_rels}</p:sldMasterIdLst>
<p:sldIdLst/>
<p:sldSz cx="{SLIDE_WIDTH}" cy="{SLIDE_HEIGHT}"/>
<p:notesSz cx="6858000" cy="9144000"/>
{default_text_style_xml()}
</p:presentation>"""


def presentation_rels_xml() -> str:
    # No slide relationships: templates ship with zero pre-existing slides so
    # render.py's Presentation.slides.add_slide() can freely allocate
    # "slide1.xml", "slide2.xml", ... without colliding with an unused part.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
</Relationships>"""


def pres_props_xml() -> str:
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentationPr {NS}/>'


def view_props_xml() -> str:
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:viewPr {NS}/>'


def table_styles_xml() -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'


def _ph_xml(ph, extra_style: str = "") -> str:
    idx_attr = f' idx="{ph.idx}"' if ph.idx is not None else ""
    ph_el = f'<p:ph type="{ph.type}"{idx_attr}/>'
    body_style = ""
    if extra_style:
        body_style = f"<a:lstStyle>{extra_style}</a:lstStyle>"
    if ph.type == "pic":
        return f"""<p:sp><p:nvSpPr><p:cNvPr id="0" name="{ph.name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>{ph_el}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{ph.left}" y="{ph.top}"/><a:ext cx="{ph.width}" cy="{ph.height}"/></a:xfrm></p:spPr></p:sp>"""
    return f"""<p:sp><p:nvSpPr><p:cNvPr id="0" name="{ph.name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>{ph_el}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{ph.left}" y="{ph.top}"/><a:ext cx="{ph.width}" cy="{ph.height}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/>{body_style}<a:p><a:r><a:rPr lang="en-US"/><a:t> </a:t></a:r></a:p></p:txBody></p:sp>"""


def slide_master_xml(theme: dict) -> str:
    c = theme["colors"]
    title_ph = _ph_xml_master("title", None, 685800, 274638, 10820400, 1143000, "Title Placeholder", c["dk1"])
    body_ph = _ph_xml_master("body", 1, 685800, 1600200, 10820400, 4525963, "Text Placeholder", c["dk1"])
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster {NS}>
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{title_ph}
{body_ph}
</p:spTree>
</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst>
{"".join(f'<p:sldLayoutId id="{2147483649 + i}" r:id="rId{i + 1}"/>' for i in range(len(LAYOUT_ORDER)))}
</p:sldLayoutIdLst>
{default_text_style_xml()}
</p:sldMaster>"""


def _ph_xml_master(ph_type, idx, left, top, width, height, name, color_hex) -> str:
    idx_attr = f' idx="{idx}"' if idx is not None else ""
    ph_el = f'<p:ph type="{ph_type}"{idx_attr}/>'
    return f"""<p:sp><p:nvSpPr><p:cNvPr id="0" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>{ph_el}</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t> </a:t></a:r></a:p></p:txBody></p:sp>"""


def slide_master_rels_xml() -> str:
    layout_rels = "".join(
        f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{i + 1}.xml"/>'
        for i in range(len(LAYOUT_ORDER))
    )
    theme_rel = f'<Relationship Id="rId{len(LAYOUT_ORDER) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{layout_rels}{theme_rel}
</Relationships>"""


def _white_rect(x, y, cx, cy, color_hex) -> str:
    return f"""<p:sp><p:nvSpPr><p:cNvPr id="0" name="Decoration"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"""


def slide_layout_xml(theme: dict, layout_name: str) -> str:
    spec = LAYOUTS[layout_name]
    c = theme["colors"]
    decoration = ""
    white_title_style = ""

    if layout_name == "Title Slide":
        decoration = _white_rect(0, 0, SLIDE_WIDTH, 137160, theme["accent_bar"])
    elif layout_name == "Section Divider":
        decoration = f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{theme["brand_dark"]}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
        white_title_style = '<a:lvl1pPr><a:defRPr sz="4000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:lvl1pPr>'
    elif layout_name == "Half Half":
        decoration = _white_rect(6092952, 0, 6099048, SLIDE_HEIGHT, c["lt2"]) + _white_rect(0, 0, SLIDE_WIDTH, 137160, theme["accent_bar"])
    elif layout_name == "Title and Bullets":
        decoration = _white_rect(0, 0, SLIDE_WIDTH, 137160, theme["accent_bar"])
    elif layout_name == "Closing":
        decoration = f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{theme["brand_dark"]}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
        white_title_style = '<a:lvl1pPr><a:defRPr sz="3600" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:lvl1pPr>'

    bg_xml = decoration if decoration.startswith("<p:bg>") else ""
    other_decoration = decoration if not decoration.startswith("<p:bg>") else ""

    ph_shapes = []
    for ph in spec.placeholders:
        style = white_title_style if (ph.type == "title" and white_title_style) else ""
        if ph.type == "body" and white_title_style:
            style = '<a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:lvl1pPr>'
        ph_shapes.append(_ph_xml(ph, style))

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout {NS}>
<p:cSld name="{spec.name}">
{bg_xml}
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{other_decoration}
{"".join(ph_shapes)}
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""


def slide_layout_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""


def example_slide_xml(brand_label: str) -> str:
    title_spec = LAYOUTS["Title Slide"].placeholder("title")
    sub_spec = LAYOUTS["Title Slide"].placeholder("subTitle")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld {NS}>
<p:cSld>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{title_spec.left}" y="{title_spec.top}"/><a:ext cx="{title_spec.width}" cy="{title_spec.height}"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{brand_label} Template</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="{sub_spec.left}" y="{sub_spec.top}"/><a:ext cx="{sub_spec.width}" cy="{sub_spec.height}"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>ppt-toolkit example slide -- see README for the 5 available layouts</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree>
</p:cSld>
</p:sld>"""


def slide_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""


def build_template(theme_key: str, brand_label: str, out_path: str) -> None:
    theme = THEMES[theme_key]
    parts: dict[str, str] = {
        "[Content_Types].xml": content_types_xml(),
        "_rels/.rels": root_rels_xml(),
        "docProps/core.xml": core_xml(f"{brand_label} Template"),
        "docProps/app.xml": app_xml(),
        "ppt/presentation.xml": presentation_xml(),
        "ppt/_rels/presentation.xml.rels": presentation_rels_xml(),
        "ppt/theme/theme1.xml": theme_xml(theme),
        "ppt/presProps.xml": pres_props_xml(),
        "ppt/viewProps.xml": view_props_xml(),
        "ppt/tableStyles.xml": table_styles_xml(),
        "ppt/slideMasters/slideMaster1.xml": slide_master_xml(theme),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": slide_master_rels_xml(),
    }
    for i, layout_name in enumerate(LAYOUT_ORDER, start=1):
        parts[f"ppt/slideLayouts/slideLayout{i}.xml"] = slide_layout_xml(theme, layout_name)
        parts[f"ppt/slideLayouts/_rels/slideLayout{i}.xml.rels"] = slide_layout_rels_xml()

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in parts.items():
            zf.writestr(name, content)
    print(f"wrote {out_path} ({len(parts)} parts)")


def main() -> None:
    build_template("white", "White", os.path.join(TEMPLATES_DIR, "white_template.pptx"))
    build_template("wb", "WB-Branded", os.path.join(TEMPLATES_DIR, "wb_branded_template.pptx"))


if __name__ == "__main__":
    main()
