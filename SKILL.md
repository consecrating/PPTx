---
name: pptx-advanced
description: "Advanced PPTX skill for creating, editing, and analyzing PowerPoint presentations. Extends the base skill with intelligent layout engine, theme management, accessibility-checked palettes, chart builder with dual-axis safety, animation/transition support, batch processing, and image optimization. Use this skill any time a .pptx or .potx file is involved — creating slide decks, pitch decks, reading/parsing .pptx files, editing presentations, working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions 'deck,' 'slides,' 'presentation,' or references a .pptx/.potx filename."
version: "2.0.0"
---

# Advanced PPTX Skill — Creation, Editing, Analysis & Automation

A `.pptx` is a ZIP archive of XML files. This skill provides a complete professional pipeline.

## Quick Reference — Choose Your Approach

| Task | Approach |
|---|---|
| **Create** a new deck | Write a `pptxgenjs` script — see gotchas below |
| **Edit** an existing deck, or build from a template | unzip → edit `ppt/slides/slideN.xml` → zip |
| **Read** content | `markitdown deck.pptx` or `python scripts/advanced/batch_processor.py extract deck.pptx` |
| **Generate themed deck** | Use `scripts/advanced/theme_engine.py` + `layout_engine.py` |
| **Batch operations** | `python scripts/advanced/batch_processor.py` |
| **Visual QA** | `python scripts/office/soffice.py --headless --convert-to pdf deck.pptx` → `pdftoppm` |

---

## Scripts Reference

### Core Scripts (Base)

| Script | What it does |
|---|---|
| `scripts/thumbnail.py deck.pptx [prefix]` | Labeled grid of every slide for picking template layouts |
| `scripts/add_slide.py unpacked/ slide2.xml [--after slideN.xml]` | Duplicate a slide with all package bookkeeping |
| `scripts/clean.py unpacked/` | Delete unreferenced slides, media, and rels |
| `scripts/office/validate.py deck.pptx [--original src.pptx]` | Schema, relationship, content-type, chart and slide checks |
| `scripts/office/soffice.py --headless --convert-to pdf deck.pptx` | LibreOffice wrapper for sandboxed environments |

### Advanced Scripts (Enhanced)

| Script | What it does |
|---|---|
| `scripts/advanced/theme_engine.py` | 12+ professional palettes, WCAG accessibility checking, custom palette derivation from brand colors |
| `scripts/advanced/layout_engine.py` | Intelligent layout calculator: two-column, card grids, timelines, stat callouts, comparisons, icon rows |
| `scripts/advanced/chart_builder.py` | Safe chart configuration generator avoiding pptxgenjs pitfalls (dual-axis, stacked labels) |
| `scripts/advanced/content_formatter.py` | Text formatting helpers, bullet lists, stat numbers, font sizing, text fitting |
| `scripts/advanced/animation_helper.py` | Slide transitions, animation sequences, staggered animations |
| `scripts/advanced/batch_processor.py` | Extract text, metadata, compare decks, batch operations |
| `scripts/advanced/image_optimizer.py` | Gradient generation, image resizing, circle icons, overlays for backgrounds |

---

## Creating with pptxgenjs — Gotchas

`pptxgenjs` is preinstalled — do not run `npm install` first; write the script and `require('pptxgenjs')` directly. Only if that require fails: `npm install pptxgenjs`.

### Critical Rules

- **Set `pres.layout` before adding slides.** Default canvas is `LAYOUT_16x9` = **10" × 5.625"**. (`LAYOUT_WIDE` is 13.3" × 7.5".)
- **Hex colors: never `#`, never 8 digits.** `color: "FF0000"`. Both `"#FF0000"` and alpha (`"00000020"`) **corrupt the file**.
- **pptxgenjs mutates option objects in place** — build a fresh object each time.
- **Shadow `offset` must be ≥ 0** — negative corrupts. Use `angle: 270` for upward shadow.
- **`letterSpacing` is silently ignored** — use `charSpacing`.
- **Lists:** `bullet: true` on each item, never literal `•`. Set `breakLine: true` on every item except last. Space with `paraSpaceAfter`.
- **One `new pptxgen()` per output file** — never reuse an instance.
- **`rectRadius` only works on `ROUNDED_RECTANGLE`**, not `RECTANGLE`.
- **Gradient fills aren't supported** — use `scripts/advanced/image_optimizer.py` to generate gradient images instead.
- **Text boxes have built-in padding** — set `margin: 0` for precise alignment.
- **Speaker notes: `slide.addNotes("...")`** (plain text, once per slide).
- **Keep charts native.** Use `addChart()` for everything PowerPoint can chart.
- **Default charts render bare** — always set `showTitle`, `showValue: true`, `chartColors`, and quiet the frame.
- **Stacked bar/column: `dataLabelPosition` must be `ctr`, `inEnd`, or `inBase`.** `outEnd` **corrupts the file**.
- **Combo charts with `secondaryValAxis` need both `valAxes` and `catAxes` (two entries each).** Without them PowerPoint discards the chart.
- **After `writeFile()`, run `python scripts/office/validate.py deck.pptx`.**
- **Never reorder the children of `<p:presentation>`.**
- **Icons:** render `react-icons` to SVG, rasterize with `sharp` at ≥256px, insert via `addImage({ data: "image/png;base64," + buf.toString("base64") })`.

### Using Advanced Features in pptxgenjs Scripts

```javascript
// Example: Using theme + layout engines together
const pptxgen = require('pptxgenjs');
// Theme and layout are configured in the script logic:
// - Use theme_engine.py palettes for chartColors and element colors
// - Use layout_engine.py regions for x, y, w, h positioning
// - Use content_formatter.py text options for fontSize, fontFace, etc.

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

// Add transition to slide
const slide = pres.addSlide();
slide.addText("Title", {
    x: 0.5, y: 0.5, w: 9, h: 1,
    fontSize: 40, color: "1E2761", bold: true, fontFace: "Cambria"
});

// Transitions (from animation_helper.py)
// pptxgenjs supports: fade, push, wipe, zoom, cover, uncover, cut
slide.transition = { type: "fade", speed: 0.5 };
```

---

## Advanced Theme System

### Using Predefined Palettes

```python
from advanced.theme_engine import get_palette, ThemeEngine, list_palettes

# Available palettes:
# midnight_executive, forest_moss, coral_energy, warm_terracotta,
# ocean_gradient, charcoal_minimal, teal_trust, berry_cream,
# sage_calm, cherry_bold, sunset_warmth, deep_purple

palette = get_palette("midnight_executive")
print(palette.primary)      # "1E2761"
print(palette.chart_colors) # ["1E2761", "4169E1", ...]

# Derive a palette from a brand color
engine = ThemeEngine()
custom = engine.derive_palette("FF6B35", style="vibrant")

# Check accessibility (WCAG 2.1)
ratio = engine.contrast_ratio("FFFFFF", "1E2761")  # 14.7:1 ✓
passes = engine.passes_wcag_aa("FFFFFF", "1E2761") # True
```

### Color Palette Quick Reference

| Theme | Primary | Secondary | Accent | Best For |
|-------|---------|-----------|--------|----------|
| Midnight Executive | `1E2761` | `CADCFC` | `7B68EE` | Corporate, Finance |
| Forest & Moss | `2C5F2D` | `97BC62` | `E8E8E8` | Sustainability, Nature |
| Coral Energy | `F96167` | `F9E795` | `2F3C7E` | Marketing, Creative |
| Warm Terracotta | `B85042` | `E7E8D1` | `A7BEAE` | Architecture, Design |
| Ocean Gradient | `065A82` | `1C7293` | `02C39A` | Tech, SaaS |
| Charcoal Minimal | `36454F` | `F2F2F2` | `E63946` | Minimalist, Editorial |
| Teal Trust | `028090` | `00A896` | `02C39A` | Healthcare, Trust |
| Berry & Cream | `6D2E46` | `A26769` | `ECE2D0` | Luxury, Beauty |
| Sage Calm | `84B59F` | `69A297` | `50808E` | Wellness, Calm |
| Cherry Bold | `990011` | `FCF6F5` | `2F3C7E` | Bold, Impact |
| Sunset Warmth | `FF6B35` | `F7C59F` | `004E89` | Startups, Energy |
| Deep Purple | `4A148C` | `CE93D8` | `FF6F00` | Creative, Innovation |

---

## Advanced Layout Engine

```python
from advanced.layout_engine import LayoutEngine

engine = LayoutEngine(slide_width=10, slide_height=5.625)

# Available layouts:
layout = engine.title_only()
layout = engine.title_subtitle()
layout = engine.title_content()
layout = engine.two_column(left_ratio=0.6)
layout = engine.three_column()
layout = engine.card_grid(rows=2, cols=3)
layout = engine.stat_callouts(count=3)
layout = engine.timeline(steps=4)
layout = engine.comparison()           # Before/After, Pros/Cons
layout = engine.half_image("right")    # Half-bleed image layouts
layout = engine.icon_rows(count=4)     # Feature list with icons

# Access regions for positioning
region = layout.get("title")
print(f"x={region.x}, y={region.y}, w={region.w}, h={region.h}")
```

---

## Slide Transitions & Animations

### Transitions (via pptxgenjs)

```python
from advanced.animation_helper import TransitionHelper

helper = TransitionHelper()
transition = helper.fade(duration=0.5)
transition = helper.push(duration=0.7)
transition = helper.wipe(duration=0.5)
transition = helper.zoom(duration=0.6)

# Get recommended transition for slide type
transition = TransitionHelper.recommended_for("title")     # Slow fade
transition = TransitionHelper.recommended_for("content")   # Fast fade
transition = TransitionHelper.recommended_for("section")   # Push
```

### Element Animations (via XML editing)

```python
from advanced.animation_helper import AnimationSequenceBuilder

builder = AnimationSequenceBuilder()
builder.add("title_1", "fade", duration=500)
builder.add("bullet_1", "fly_in", duration=400)
builder.add("bullet_2", "fly_in", duration=400)
sequence = builder.get_sequence()

# Staggered animations for lists
animations = builder.build_stagger(
    ["item_1", "item_2", "item_3"],
    effect="fade", stagger_delay=150
)
```

---

## Image & Gradient Generation

```python
from advanced.image_optimizer import ImageOptimizer

optimizer = ImageOptimizer()

# Create gradient backgrounds (since pptxgenjs doesn't support native gradients)
gradient = optimizer.create_gradient(1920, 1080, "1E2761", "7B68EE", direction="horizontal")
# Directions: horizontal, vertical, diagonal, radial

# Create circle icon backgrounds
icon_bg = optimizer.create_circle_icon_bg(size=64, color="1E2761")

# Create semi-transparent overlays
overlay = optimizer.create_overlay(1920, 1080, color="000000", opacity=0.4)

# Prepare images for embedding
base64_img = optimizer.prepare_for_slide("photo.jpg", max_width=1920)

# Calculate image fit in container
x, y, w, h = optimizer.calculate_fit(1920, 1080, container_width=5.0, container_height=4.0)
```

---

## Chart Builder (Safe Configuration)

```python
from advanced.chart_builder import ChartBuilder

builder = ChartBuilder(palette=["1E2761", "4169E1", "7B68EE"])

# Bar chart (handles stacked label position safety)
config = builder.bar_chart(data, title="Revenue", stacked=True)

# Line chart with markers
config = builder.line_chart(data, title="Growth", smooth=True)

# Pie / Doughnut
config = builder.pie_chart(data, title="Market Share", doughnut=True)

# COMBO CHART (automatically handles dual-axis safety)
config = builder.combo_chart(bar_data, line_data, title="Revenue vs Growth")
# ^ This properly defines valAxes + catAxes to avoid PowerPoint corruption

# Scatter plot
config = builder.scatter_chart(data, title="Correlation")

# Area chart
config = builder.area_chart(data, title="Trend", stacked=True)
```

---

## Editing Existing Decks and Templates

Pick layouts first: `python scripts/thumbnail.py template.pptx template-thumbs` writes a labeled grid.

```bash
python3 -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall('unpacked')" deck.pptx
python scripts/add_slide.py unpacked/ slide2.xml --after slide2.xml
# reorder / delete slides = edit <p:sldIdLst> in ppt/presentation.xml
python scripts/clean.py unpacked/
# edit slide content in ppt/slides/slideN.xml
(cd unpacked && rm -f ../out.pptx && zip -Xr ../out.pptx .)
python scripts/office/validate.py out.pptx --original deck.pptx
```

### Key Rules for Template Editing

- **Do all structural work (add, delete, reorder) before editing content.**
- **Never copy a slide file by hand** — use `add_slide.py`.
- Parse with `defusedxml.minidom` — `xml.etree.ElementTree` rewrites namespace prefixes and corrupts.
- One `<a:p>` per list item — never concatenate.
- Let bullets inherit from layout; only add `<a:buChar>` or `<a:buNone>` to override.
- Text with leading/trailing spaces needs `xml:space="preserve"` on its `<a:t>`.

---

## Design Guidelines

### Typography (Safe Fonts)

**Safe fonts** (render correctly in QA AND ship with Office):
- **Sans-serif:** Arial, Calibri
- **Serif:** Cambria, Times New Roman, Bookman Old Style, Century Schoolbook
- **Monospace:** Courier New

**Never default to Aptos** — no compatible substitute exists.

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | 14-16pt |
| Captions | 10-12pt muted |

### Layout Rules

- 0.5" minimum margins
- 0.3-0.5" between content blocks
- Leave breathing room — don't fill every inch
- Vary layouts across slides (use layout_engine.py)

### Design Anti-Patterns (AVOID)

- Text-only slides without any visual element
- Same layout repeated on every slide
- Centered body text (left-align paragraphs)
- Low-contrast text (check with theme_engine.py)
- Accent lines/bars under titles (AI hallmark)
- Decorative color bars or accent stripes
- Default beige/cream backgrounds (use white or brand)
- Text overflow beyond shape bounds

---

## Batch Processing

```bash
# Extract text content
python scripts/advanced/batch_processor.py extract deck.pptx --format markdown

# Get file metadata
python scripts/advanced/batch_processor.py info deck.pptx

# Compare two decks
python scripts/advanced/batch_processor.py compare deck1.pptx deck2.pptx
```

---

## QA Pipeline (Required)

### 1. Content QA
```bash
markitdown output.pptx
# Check for placeholder text:
markitdown output.pptx | grep -iE "\bx{3,}\b|lorem|ipsum|\bTODO|\[insert"
```

### 2. File Validation
```bash
python scripts/office/validate.py output.pptx                      # built from scratch
python scripts/office/validate.py output.pptx --original src.pptx  # from template
```

### 3. Visual QA
```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 output.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

Check for: text overflow, overlapping elements, low contrast, uneven spacing, misalignment.

---

## Dependencies

**npm:** `pptxgenjs` (preinstalled), `react-icons`, `react`, `react-dom`, `sharp`
**pip:** `markitdown[pptx]`, `Pillow`, `defusedxml`, `lxml`
**system:** LibreOffice (`soffice`), `pdftoppm` (Poppler), `gcc` (for socket shim)
