# ppt-toolkit

Generate and convert PowerPoint decks between a clean **White/McKinsey-style**
template and a **WB-branded** template, from a single template-agnostic JSON
description. Existing hand-built decks can be "harvested" into that JSON and
re-rendered in either style.

```
ppt-toolkit/
  schema/deck_schema.json      # JSON Schema an LLM (or harvest.py) targets
  templates/
    white_template.pptx        # clean/white template
    wb_branded_template.pptx   # WB-branded template
    _source_white.pptx         # reference source deck (not used by the generator directly)
    _source_wb.pptx             # reference source deck (not used by the generator directly)
  src/
    layouts.py                  # single source of truth: the 5 layouts + placeholder geometry
    render.py                   # JSON deck -> .pptx
    harvest.py                  # existing .pptx -> JSON deck (best-effort)
    cli.py                      # `ppt render|harvest|convert`
  tools/build_templates.py      # generates both .pptx templates from layouts.py + THEMES
  examples/
    sample_deck.json            # 4-slide example deck
    output/                     # rendered/harvested outputs (generated, not hand-edited)
  pyproject.toml, requirements.txt
```

## Why two templates generated from one source

Both templates must expose the **same 5 named slide layouts** with identical
placeholder types/positions, differing only in theme (colors, fonts,
background decoration). Rather than hand-maintaining two separate `.pptx`
files and risking them drifting apart, `tools/build_templates.py` generates
**both** template files from raw OOXML, driven by:

- `src/layouts.py` — the layout/placeholder spec (names, placeholder types,
  `idx`, geometry in EMU). This is the single source of truth for structure.
- The `THEMES` dict in `build_templates.py` — per-template colors, fonts, and
  background decoration (accent bars, full-bleed brand-dark backgrounds).

This guarantees the two templates stay in lock-step by construction: a slide
built against layout `"Half Half"` in one template has the exact same
placeholder shapes/positions in the other, so a deck JSON renders identically
laid out in both styles.

### The 5 canonical layouts

| Layout | Purpose | Placeholders |
|---|---|---|
| `Title Slide` | Deck cover | title, subtitle |
| `Section Divider` | Full-bleed section break | title only |
| `Half Half` | Left text / right image split | title, body (bullets), picture |
| `Title and Bullets` | Standard content slide | title, body (bullets) |
| `Closing` | Full-bleed thank-you/closing slide | title, body (message) |

Layout choices were derived from two user-supplied "cutdown" reference decks
(`templates/_source_white.pptx`, a Slideworks proposal template excerpt, and
`templates/_source_wb.pptx`, a WB Prosperity/CPAT template excerpt) — kept in
`templates/` for reference only; the generator does not read them at build
time. Theme colors/fonts (white: Slideworks navy/blue/gold, Arial; WB: the
deck's actual "Climate Warehouse 2" theme — navy/blue/green/yellow, with
coral/pink `E7485F` as the WB brand-dark background) were extracted from
those files' `theme1.xml`.

### Adding a new layout to both templates in sync

1. Add a new entry to `LAYOUTS` (and `LAYOUT_ORDER`) in `src/layouts.py` with
   its placeholders (type, `idx`, geometry).
2. In `tools/build_templates.py`, add a branch for the new layout name in the
   per-layout decoration logic (background, accent bar, text color
   overrides) for each theme — reuse an existing pattern where possible.
3. Add the new layout name to the `layout` enum in
   `schema/deck_schema.json`, plus any new fields it needs.
4. Re-run `python tools/build_templates.py` (from `ppt-toolkit/`) to
   regenerate both `.pptx` files.
5. Update `render.py`/`harvest.py` only if the new layout needs special
   handling beyond the existing title/subtitle/bullets/message/chart/table/
   image/notes fields.

Because both templates are generated from the same layout spec, there's no
way for them to drift out of sync as long as you always regenerate both after
editing `layouts.py`.

## The content schema

`schema/deck_schema.json` describes a deck as `{title, author?, slides: [...]}`
where each slide has a `layout` (one of the 5 names above) plus whatever
fields that layout uses: `title`, `subtitle`, `bullets` (array of strings or
`{text, level}`), `message`, `image` (`{path, alt_text?}`), `chart`
(`{type, categories, series}`), `table` (`{headers, rows}`), `source`
(footnote text), `notes` (speaker notes). Unused fields for a given layout are
ignored by the renderer. `chart`/`table` take priority over `bullets` if more
than one is supplied on the same slide (they share the body placeholder).

### Example prompt for an LLM to generate a deck

```
You produce PowerPoint deck content as JSON that strictly matches the schema
at schema/deck_schema.json. Rules:
- Every slide has a "layout" field, one of:
  "Title Slide", "Section Divider", "Half Half", "Title and Bullets", "Closing".
- Only include fields that layout uses (see schema/deck_schema.json).
- Keep bullets concise (<= 15 words each); use "level" for sub-bullets.
- Output ONLY the JSON document, no commentary.

Task: Create a 4-slide deck about <topic>, covering: a title slide, one
content slide with 4-6 bullets, one slide with a column chart of <data>,
and a closing slide with contact info.
```

Validate the LLM's output against the schema (`jsonschema`) before rendering
— `render.py` does this automatically unless `--skip-validate` is passed.

## CLI usage

```bash
pip install -r requirements.txt   # python-pptx, jsonschema

# Render a JSON deck into a .pptx with the chosen template
python -m src.cli render examples/sample_deck.json --template white -o out_white.pptx
python -m src.cli render examples/sample_deck.json --template wb    -o out_wb.pptx

# Harvest an existing .pptx into schema-shaped JSON (best-effort layout mapping)
python -m src.cli harvest old_deck.pptx -o extracted.json

# Convert an existing .pptx to the other template (harvest + render chained)
python -m src.cli convert old_deck.pptx --template white -o new_deck.pptx
```

If installed as a package (`pip install -e .`), the same commands are
available as `ppt render|harvest|convert ...`.

## Example deck

`examples/sample_deck.json` is a 4-slide CPAT-style deck (Title Slide, Title
and Bullets, Half Half with an embedded chart, Closing). It's rendered to
`examples/output/sample_white.pptx` and `sample_wb.pptx` so you can compare
both styles side by side, and `examples/output/converted_white_to_wb.pptx`
demonstrates round-tripping white → harvest → wb with content preserved.

## Real-world example: converting a hand-built deck

`examples/real_decks/` demonstrates converting an actual 21-slide hand-built
deck (`CPAT_Final_Summary.pptx`, a dense consulting-style deck with no
slide-layout placeholders at all — just freeform text boxes, tables, and one
image, arranged in multi-column "card" layouts):

- `CPAT_Final_Summary_extracted.json` — output of `ppt harvest`
- `CPAT_Final_Summary_white.pptx` / `_wb.pptx` — rendered in both templates
- `images/` — pictures extracted from the source deck during harvest

Because this deck has no placeholders to key off, `harvest.py` falls back to
a positional/heuristic extractor (see "Harvesting freeform decks" below) that
reconstructs a title, section-label, bullet hierarchy, tables, and images
from raw text-box positions. All 21 slides converted successfully with
content preserved; see "Known limitations" for what this approach can't
capture (the original's precise multi-column visual layout is flattened into
one of the 5 canonical layouts, with cards becoming indented bullet groups).

### Harvesting freeform decks (no placeholders)

Many hand-built decks — especially ones with custom card/grid layouts —
don't use PowerPoint's placeholder mechanism at all; every "slide" is just a
pile of positioned text boxes, tables, and pictures. `harvest.py` detects
this (`slide.placeholders` is empty) and falls back to:

1. Sort shapes **column-major** (left, then top) — this matches how
   multi-column card/agenda layouts are meant to be read, and degenerates
   gracefully to top-to-bottom reading order for single-column slides.
2. Drop running page footers (any text shape in the bottom ~15% of the
   slide) unless it contains a "Source:" citation, in which case the
   trailing "<deck name> <page number>" tail is stripped and the citation
   text is kept as the slide's `source` field.
3. Treat the first short, punchy text shape as an optional "eyebrow"/section
   label and fold it into the slide `title` (e.g. `"PART 2 -- TWO PRODUCTS
   COMPARED -- One set of equations, two forms"`), since the schema has no
   separate eyebrow field.
4. Classify remaining text shapes as a "header" (short, no trailing
   punctuation, or ALL CAPS) at bullet level 0, or a "detail" line at level 1
   — approximating card layouts (header + description) as an indented list.
5. The first table shape becomes the slide's `table` field; the first
   picture is saved to an `images/` folder next to the output JSON and
   referenced via the slide's `image.path`.

This is inherently approximate — always skim the extracted JSON (and the
rendered output) before treating it as final, especially for dense grid
layouts where the header/detail split or reading order may need a manual
tweak.

## Validation performed (no PowerPoint installed)

There is no PowerPoint installation available in this environment, so output
could not be visually rendered/verified. Validation instead covers:

- **Zip/XML structural integrity**: every generated `.pptx` is a valid zip
  (`zipfile.testzip()` reports no errors) and every XML part parses.
- **python-pptx round-trip load**: both templates and all example outputs
  load cleanly via `python-pptx` with no duplicate-part warnings, correct
  layout names, and correct placeholder types/positions.
- **Content round-trip**: `harvest.py` run on a rendered deck reproduces the
  original layout choice and text content (title/subtitle/bullets/message),
  confirmed on `examples/output/sample_white.pptx` and via the chained
  `convert` command.

**You should still open the outputs in PowerPoint/LibreOffice to visually
confirm colors, fonts, and layout look right before relying on this for real
decks** — structural validation catches corruption but not visual/design
issues.

## Known limitations

- `harvest.py`'s layout guessing is heuristic for *foreign* decks (not built
  by this toolkit): it inspects placeholder types and text length, since
  arbitrary source decks won't use our 5 layout names. Decks previously
  rendered by this toolkit round-trip exactly (their layout name is trusted
  directly).
- Chart/table content is not currently reconstructed by `harvest.py` (it only
  extracts text); harvesting a chart/table slide preserves the slide's title
  and layout guess but drops the data. Re-populate `chart`/`table` fields by
  hand if you need to convert a chart-bearing legacy slide.
- The `source` footnote field is only rendered when a slide uses plain
  `bullets` (it's appended as a trailing bullet). It is not shown alongside
  `chart`/`table` content; a dedicated footnote placeholder is a possible
  future addition.
- Images referenced via the `image` field are inserted only into the `Half
  Half` layout's picture placeholder, and only if the path exists on disk.
- The freeform-deck header/detail heuristic (see above) is imperfect on
  atypical card shapes (e.g. a standalone short fact with no following
  detail line may be mis-leveled); review/adjust levels in the extracted
  JSON for anything presentation-critical.

## Do not modify

`templates/_source_white.pptx` and `templates/_source_wb.pptx` are read-only
copies of user-supplied reference decks, kept only so the layout/theme
extraction is reproducible. They are not read by `build_templates.py` at
runtime and should not be edited. Original files under the user's OneDrive
folder were never modified — only read for inspection during design.
