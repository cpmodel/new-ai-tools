"""Single source of truth for the 5 canonical slide layouts shared by both
the white and WB-branded templates.

Both `templates/white_template.pptx` and `templates/wb_branded_template.pptx`
expose exactly these 5 named slide layouts, in this order, with identical
placeholder types/positions. `tools/build_templates.py` uses this module to
generate both template files; `render.py` and `harvest.py` use it to find the
right layout/placeholder by name instead of hard-coding indices.

To add a new layout to BOTH templates in sync: add an entry here (name +
placeholders + geometry), add matching per-theme decoration in
`tools/build_templates.py`'s `THEMES[...]["layout_decoration"]`, then re-run
`python tools/build_templates.py`.
"""

from __future__ import annotations

from dataclasses import dataclass, field

# Standard 16:9 widescreen slide size in EMU (matches both source decks'
# aspect ratio; python-pptx / PowerPoint default).
SLIDE_WIDTH = 12192000
SLIDE_HEIGHT = 6858000

# Placeholder type strings match python-pptx / OOXML `<p:ph type="...">`
# values: "title", "subTitle", "body", "pic". idx=None means no idx attribute
# (used for the single title/ctrTitle placeholder on a layout).


@dataclass(frozen=True)
class PlaceholderSpec:
    type: str  # "title" | "subTitle" | "body" | "pic"
    idx: int | None
    left: int
    top: int
    width: int
    height: int
    name: str  # human-readable shape name


@dataclass(frozen=True)
class LayoutSpec:
    name: str
    description: str
    placeholders: list[PlaceholderSpec] = field(default_factory=list)

    def placeholder(self, ph_type: str) -> PlaceholderSpec | None:
        for p in self.placeholders:
            if p.type == ph_type:
                return p
        return None


LAYOUTS: dict[str, LayoutSpec] = {
    "Title Slide": LayoutSpec(
        name="Title Slide",
        description="Deck cover: title, subtitle/byline.",
        placeholders=[
            PlaceholderSpec("title", None, 685800, 2286000, 10820400, 1143000, "Title"),
            PlaceholderSpec("subTitle", 1, 685800, 3600450, 10820400, 914400, "Subtitle"),
        ],
    ),
    "Section Divider": LayoutSpec(
        name="Section Divider",
        description="Full-bleed section break with a large title.",
        placeholders=[
            PlaceholderSpec("title", None, 685800, 2971800, 10820400, 914400, "Title"),
        ],
    ),
    "Half Half": LayoutSpec(
        name="Half Half",
        description="Left half: title + bullets. Right half: image.",
        placeholders=[
            PlaceholderSpec("title", None, 685800, 457200, 5486400, 685800, "Title"),
            PlaceholderSpec("body", 1, 685800, 1200150, 5486400, 5200650, "Bullets"),
            PlaceholderSpec("pic", 2, 6400800, 0, 5791200, 6858000, "Image"),
        ],
    ),
    "Title and Bullets": LayoutSpec(
        name="Title and Bullets",
        description="Standard content slide: title + bullet body.",
        placeholders=[
            PlaceholderSpec("title", None, 685800, 457200, 10820400, 685800, "Title"),
            PlaceholderSpec("body", 1, 685800, 1200150, 10820400, 5200650, "Bullets"),
        ],
    ),
    "Closing": LayoutSpec(
        name="Closing",
        description="Full-bleed closing/thank-you slide with a message body.",
        placeholders=[
            PlaceholderSpec("title", None, 685800, 2286000, 10820400, 914400, "Title"),
            PlaceholderSpec("body", 1, 685800, 3352800, 10820400, 1600200, "Message"),
        ],
    ),
}

LAYOUT_ORDER = ["Title Slide", "Section Divider", "Half Half", "Title and Bullets", "Closing"]
