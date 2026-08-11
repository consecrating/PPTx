# PPTx — Advanced PowerPoint Presentation Skill

> A comprehensive, production-ready skill for creating, editing, and analyzing professional PowerPoint presentations (.pptx/.potx) using AI assistants.

Built on top of the [Anthropic PPTX Skill](https://github.com/anthropics/skills/tree/main/skills/pptx) with significant enhancements including an intelligent layout engine, WCAG-compliant theme system, safe chart configuration, animation support, and batch processing capabilities.

---

## Features

### Core Capabilities
- **Create** professional presentations from scratch using `pptxgenjs`
- **Edit** existing decks via XML manipulation with full package bookkeeping
- **Read** and extract content from any `.pptx` or `.potx` file
- **Validate** generated files against OOXML schemas before delivery
- **Visual QA** pipeline with PDF conversion and slide-by-slide inspection

### Advanced Enhancements (v2.0)

| Feature | Description |
|---------|-------------|
| **Theme Engine** | 12 professional palettes + custom derivation from brand colors + WCAG accessibility checking |
| **Layout Engine** | 11 intelligent layout types with automatic positioning (grids, timelines, comparisons, etc.) |
| **Chart Builder** | Safe chart configuration that avoids pptxgenjs corruption bugs (dual-axis, stacked labels) |
| **Content Formatter** | Smart text formatting, bullet generation, font sizing, and text-fit calculations |
| **Animation Helper** | Slide transitions, element animation sequences, staggered reveal effects |
| **Image Optimizer** | Gradient generation, image resizing, circle icons, overlays (workaround for no native gradients) |
| **Batch Processor** | Text extraction, metadata analysis, deck comparison, bulk operations |

---

## Project Structure

```
PPTx/
├── SKILL.md                          # Complete skill documentation & instructions
├── README.md                         # This file
├── package.json                      # Node.js dependencies & scripts
├── requirements.txt                  # Python dependencies
├── examples/
│   └── demo_presentation.js          # Example presentation generator
├── scripts/
│   ├── __init__.py
│   ├── thumbnail.py                  # Slide thumbnail grid generator
│   ├── add_slide.py                  # Slide duplication with bookkeeping
│   ├── clean.py                      # Remove orphaned files
│   ├── office/
│   │   ├── __init__.py
│   │   ├── soffice.py                # LibreOffice sandbox wrapper
│   │   ├── validate.py               # OOXML schema validator
│   │   ├── helpers/
│   │   │   ├── __init__.py           # OPC utilities, safe_extract, rezip
│   │   │   ├── pptx_chart.py         # Chart validation
│   │   │   ├── pptx_slide.py         # Slide schema checks
│   │   │   └── pptx_theme.py         # Theme/master validation
│   │   ├── schemas/                  # XSD schema files
│   │   └── validators/
│   │       ├── __init__.py
│   │       ├── base.py               # Base validator class
│   │       ├── docx.py               # DOCX validator
│   │       ├── pptx.py               # PPTX validator
│   │       └── redlining.py          # Tracked changes validator
│   └── advanced/
│       ├── __init__.py
│       ├── theme_engine.py           # Color palettes & accessibility
│       ├── layout_engine.py          # Intelligent layout calculator
│       ├── chart_builder.py          # Safe chart configurations
│       ├── content_formatter.py      # Text formatting utilities
│       ├── animation_helper.py       # Transitions & animations
│       ├── batch_processor.py        # Bulk processing tools
│       └── image_optimizer.py        # Image & gradient generation
```

---

## Quick Start

### Prerequisites

```bash
# Node.js (pptxgenjs for creation)
npm install

# Python (validation, thumbnails, advanced features)
pip install -r requirements.txt

# System dependencies
# LibreOffice (soffice) — for PDF conversion & visual QA
# Poppler (pdftoppm) — for PDF to image conversion
# GCC — for LibreOffice socket shim in sandboxes
```

### Create a Presentation

```javascript
const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";  // Always set before adding slides!

const slide = pres.addSlide();
slide.addText("Hello World", {
    x: 0.5, y: 0.5, w: 9, h: 1.2,
    fontSize: 44, color: "1E2761", bold: true, fontFace: "Cambria",
    align: "left"
});
slide.transition = { type: "fade", speed: 0.5 };

pres.writeFile({ fileName: "output.pptx" });
```

### Validate Output

```bash
python scripts/office/validate.py output.pptx
```

### Visual QA

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
# Inspect slide-*.jpg files
```

---

## Advanced Usage

### Theme Engine — Professional Color Palettes

```python
from scripts.advanced.theme_engine import get_palette, ThemeEngine

# 12 predefined professional palettes
palette = get_palette("midnight_executive")
print(palette.primary)       # "1E2761"
print(palette.secondary)     # "CADCFC"
print(palette.chart_colors)  # ["1E2761", "4169E1", "7B68EE", ...]

# Generate palette from brand color
engine = ThemeEngine()
custom = engine.derive_palette("FF6B35", style="vibrant")

# WCAG accessibility checking
ratio = engine.contrast_ratio("FFFFFF", "1E2761")  # 14.7:1
engine.passes_wcag_aa("FFFFFF", "1E2761")          # True
engine.passes_wcag_aaa("FFFFFF", "1E2761")         # True
```

### Layout Engine — Intelligent Positioning

```python
from scripts.advanced.layout_engine import LayoutEngine

engine = LayoutEngine(slide_width=10, slide_height=5.625)

# Two-column with 60/40 split
layout = engine.two_column(left_ratio=0.6)
title = layout.get("title")   # Rect(x=0.5, y=0.5, w=9.0, h=0.8)
left = layout.get("left")     # Rect(x=0.5, y=1.6, w=5.16, h=3.825)
right = layout.get("right")   # Rect(x=6.06, y=1.6, w=3.44, h=3.825)

# Card grid for feature showcases
layout = engine.card_grid(rows=2, cols=3, gap=0.25)

# Timeline for process flows
layout = engine.timeline(steps=5)

# Stat callouts for KPIs
layout = engine.stat_callouts(count=4)
```

### Chart Builder — Corruption-Safe Charts

```python
from scripts.advanced.chart_builder import ChartBuilder

builder = ChartBuilder(palette=["1E2761", "4169E1", "7B68EE"])

# Stacked bar (auto-sets safe dataLabelPosition)
config = builder.bar_chart(data, title="Revenue", stacked=True)

# Combo chart (auto-defines valAxes + catAxes to prevent corruption)
config = builder.combo_chart(bar_data, line_data, title="Revenue vs Growth")
```

### Image Optimizer — Gradient Backgrounds

```python
from scripts.advanced.image_optimizer import ImageOptimizer

optimizer = ImageOptimizer()

# Generate gradient (pptxgenjs workaround)
gradient = optimizer.create_gradient(1920, 1080, "1E2761", "7B68EE", "diagonal")

# Circle icon backgrounds
icon_bg = optimizer.create_circle_icon_bg(size=64, color="1E2761")

# Semi-transparent overlay
overlay = optimizer.create_overlay(1920, 1080, "000000", opacity=0.4)
```

---

## Available Palettes

| Name | Primary | Style | Best For |
|------|---------|-------|----------|
| `midnight_executive` | Navy | Dark & Professional | Corporate, Finance |
| `forest_moss` | Green | Natural & Fresh | Sustainability, Nature |
| `coral_energy` | Coral | Vibrant & Bold | Marketing, Creative |
| `warm_terracotta` | Terracotta | Warm & Earthy | Architecture, Design |
| `ocean_gradient` | Deep Blue | Cool & Tech | SaaS, Technology |
| `charcoal_minimal` | Charcoal | Clean & Sharp | Editorial, Minimalist |
| `teal_trust` | Teal | Trustworthy | Healthcare, Finance |
| `berry_cream` | Berry | Luxurious | Beauty, Luxury |
| `sage_calm` | Sage | Calming | Wellness, Mindfulness |
| `cherry_bold` | Cherry | High-Impact | Sales, Bold Pitches |
| `sunset_warmth` | Orange | Energetic | Startups, Innovation |
| `deep_purple` | Purple | Creative | Tech, AI, Innovation |

---

## Key Differences from Base Skill

| Feature | Base Skill | This Repo (v2.0) |
|---------|-----------|-------------------|
| Color Palettes | 10 static palettes | 12 palettes + dynamic derivation + WCAG checking |
| Layout | Manual positioning | Intelligent layout engine with 11 layout types |
| Charts | Manual config | Safe builder preventing common corruption bugs |
| Animations | Not covered | Transition helper + animation sequence builder |
| Images | Not covered | Gradient generation + optimization + overlays |
| Batch Operations | Not covered | Text extraction, metadata, deck comparison |
| Content Formatting | Not covered | Smart text fitting, bullet generation, stat formatting |
| Accessibility | Not covered | WCAG AA/AAA contrast checking built-in |

---

## Common Pitfalls Avoided

This skill automatically handles or documents workarounds for:

1. **Hex color corruption** — enforces 6-digit hex without `#`
2. **Stacked chart label crash** — only uses `ctr`/`inEnd`/`inBase`
3. **Combo chart axis crash** — always defines both `valAxes` and `catAxes`
4. **Options object mutation** — documents fresh-object pattern
5. **Gradient fill limitation** — provides image-based gradient generation
6. **Font substitution in QA** — safe-font list with known-good metrics
7. **Namespace corruption** — uses `defusedxml.minidom` not `ElementTree`
8. **Shadow negative offset** — validates offset ≥ 0

---

## License

The base skill files are derived from [Anthropic's Skills Repository](https://github.com/anthropics/skills) under their proprietary license. Advanced enhancements in `scripts/advanced/` are original work.

See [LICENSE.txt](LICENSE.txt) for full terms.
