"""Advanced Layout Engine for PPTX presentations.

Provides intelligent layout calculation, automatic content placement,
and responsive grid systems for professional slides.

Usage:
    from advanced.layout_engine import LayoutEngine, SlideLayout

    engine = LayoutEngine(slide_width=10, slide_height=5.625)
    layout = engine.two_column(left_ratio=0.6)
    grid = engine.card_grid(rows=2, cols=3, gap=0.3)
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Rect:
    """A rectangle on the slide (in inches)."""
    x: float
    y: float
    w: float
    h: float

    @property
    def right(self) -> float:
        return self.x + self.w

    @property
    def bottom(self) -> float:
        return self.y + self.h

    @property
    def center_x(self) -> float:
        return self.x + self.w / 2

    @property
    def center_y(self) -> float:
        return self.y + self.h / 2


@dataclass
class SlideLayout:
    """A calculated slide layout with named regions."""
    regions: dict = field(default_factory=dict)
    description: str = ""

    def get(self, name: str) -> Rect:
        if name not in self.regions:
            raise KeyError(f"Region '{name}' not in layout. Available: {list(self.regions.keys())}")
        return self.regions[name]

    def __repr__(self) -> str:
        return f"SlideLayout({self.description}, regions={list(self.regions.keys())})"


class LayoutEngine:
    """Intelligent layout calculator for PPTX slides."""

    # Standard slide sizes
    LAYOUT_16x9 = (10.0, 5.625)
    LAYOUT_WIDE = (13.333, 7.5)
    LAYOUT_4x3 = (10.0, 7.5)

    def __init__(self, slide_width: float = 10.0, slide_height: float = 5.625,
                 margin: float = 0.5, padding: float = 0.3):
        self.width = slide_width
        self.height = slide_height
        self.margin = margin
        self.padding = padding
        # Usable content area
        self.content_x = margin
        self.content_y = margin
        self.content_w = slide_width - 2 * margin
        self.content_h = slide_height - 2 * margin

    def title_only(self, title_height: float = 1.0) -> SlideLayout:
        """Single centered title layout."""
        return SlideLayout(
            description="Title Only",
            regions={
                "title": Rect(self.content_x, self.height / 2 - title_height / 2,
                             self.content_w, title_height),
            }
        )

    def title_subtitle(self, title_height: float = 1.2, subtitle_height: float = 0.8,
                       gap: float = 0.2) -> SlideLayout:
        """Title + subtitle centered layout."""
        total_h = title_height + gap + subtitle_height
        start_y = (self.height - total_h) / 2
        return SlideLayout(
            description="Title + Subtitle",
            regions={
                "title": Rect(self.content_x, start_y, self.content_w, title_height),
                "subtitle": Rect(self.content_x, start_y + title_height + gap,
                                self.content_w, subtitle_height),
            }
        )

    def title_content(self, title_height: float = 0.8, gap: float = 0.3) -> SlideLayout:
        """Title at top + content area below."""
        content_y = self.content_y + title_height + gap
        content_h = self.content_h - title_height - gap
        return SlideLayout(
            description="Title + Content",
            regions={
                "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
                "content": Rect(self.content_x, content_y, self.content_w, content_h),
            }
        )

    def two_column(self, left_ratio: float = 0.5, gap: float = 0.4,
                   title_height: float = 0.8) -> SlideLayout:
        """Two-column layout with optional title."""
        top_y = self.content_y + title_height + 0.3
        col_h = self.content_h - title_height - 0.3
        left_w = (self.content_w - gap) * left_ratio
        right_w = self.content_w - gap - left_w
        return SlideLayout(
            description=f"Two Column ({left_ratio:.0%}/{1-left_ratio:.0%})",
            regions={
                "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
                "left": Rect(self.content_x, top_y, left_w, col_h),
                "right": Rect(self.content_x + left_w + gap, top_y, right_w, col_h),
            }
        )

    def three_column(self, gap: float = 0.3, title_height: float = 0.8) -> SlideLayout:
        """Three equal columns with title."""
        top_y = self.content_y + title_height + 0.3
        col_h = self.content_h - title_height - 0.3
        col_w = (self.content_w - 2 * gap) / 3
        return SlideLayout(
            description="Three Column",
            regions={
                "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
                "col1": Rect(self.content_x, top_y, col_w, col_h),
                "col2": Rect(self.content_x + col_w + gap, top_y, col_w, col_h),
                "col3": Rect(self.content_x + 2 * (col_w + gap), top_y, col_w, col_h),
            }
        )

    def card_grid(self, rows: int = 2, cols: int = 3, gap: float = 0.25,
                  title_height: float = 0.8) -> SlideLayout:
        """Grid of equal cards."""
        top_y = self.content_y + title_height + 0.3
        grid_h = self.content_h - title_height - 0.3
        card_w = (self.content_w - (cols - 1) * gap) / cols
        card_h = (grid_h - (rows - 1) * gap) / rows

        regions = {
            "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
        }
        for r in range(rows):
            for c in range(cols):
                x = self.content_x + c * (card_w + gap)
                y = top_y + r * (card_h + gap)
                regions[f"card_{r}_{c}"] = Rect(x, y, card_w, card_h)

        return SlideLayout(description=f"Card Grid ({rows}x{cols})", regions=regions)

    def stat_callouts(self, count: int = 3, title_height: float = 0.8) -> SlideLayout:
        """Large stat numbers with labels below."""
        top_y = self.content_y + title_height + 0.5
        gap = 0.4
        stat_w = (self.content_w - (count - 1) * gap) / count
        stat_h = 1.5
        label_h = 0.8

        regions = {
            "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
        }
        for i in range(count):
            x = self.content_x + i * (stat_w + gap)
            regions[f"stat_{i}"] = Rect(x, top_y, stat_w, stat_h)
            regions[f"label_{i}"] = Rect(x, top_y + stat_h + 0.1, stat_w, label_h)

        return SlideLayout(description=f"Stat Callouts ({count})", regions=regions)

    def timeline(self, steps: int = 4, title_height: float = 0.8) -> SlideLayout:
        """Horizontal timeline/process flow."""
        top_y = self.content_y + title_height + 0.5
        gap = 0.2
        step_w = (self.content_w - (steps - 1) * gap) / steps
        circle_size = 0.6
        desc_h = 1.5

        regions = {
            "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
            "line": Rect(self.content_x, top_y + circle_size / 2 - 0.02,
                        self.content_w, 0.04),
        }
        for i in range(steps):
            x = self.content_x + i * (step_w + gap)
            center_x = x + step_w / 2
            regions[f"circle_{i}"] = Rect(center_x - circle_size/2, top_y,
                                         circle_size, circle_size)
            regions[f"step_{i}"] = Rect(x, top_y + circle_size + 0.3, step_w, desc_h)

        return SlideLayout(description=f"Timeline ({steps} steps)", regions=regions)

    def comparison(self, title_height: float = 0.8) -> SlideLayout:
        """Before/After or Pros/Cons comparison layout."""
        gap = 0.5
        top_y = self.content_y + title_height + 0.3
        col_h = self.content_h - title_height - 0.3
        col_w = (self.content_w - gap) / 2
        header_h = 0.6
        return SlideLayout(
            description="Comparison (A vs B)",
            regions={
                "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
                "left_header": Rect(self.content_x, top_y, col_w, header_h),
                "left_content": Rect(self.content_x, top_y + header_h + 0.1,
                                    col_w, col_h - header_h - 0.1),
                "right_header": Rect(self.content_x + col_w + gap, top_y, col_w, header_h),
                "right_content": Rect(self.content_x + col_w + gap, top_y + header_h + 0.1,
                                     col_w, col_h - header_h - 0.1),
                "divider": Rect(self.content_x + col_w + gap/2 - 0.01, top_y,
                               0.02, col_h),
            }
        )

    def half_image(self, image_side: str = "right", image_ratio: float = 0.45,
                   title_height: float = 0.8) -> SlideLayout:
        """Half-bleed image with content on the other side."""
        img_w = self.width * image_ratio
        content_w = self.width - img_w - self.margin
        top_y = self.content_y + title_height + 0.3
        content_h = self.content_h - title_height - 0.3

        if image_side == "right":
            return SlideLayout(
                description="Half Image (Right)",
                regions={
                    "title": Rect(self.content_x, self.content_y, content_w - self.margin, title_height),
                    "content": Rect(self.content_x, top_y, content_w - self.margin, content_h),
                    "image": Rect(self.width - img_w, 0, img_w, self.height),
                }
            )
        else:
            return SlideLayout(
                description="Half Image (Left)",
                regions={
                    "title": Rect(img_w + self.margin, self.content_y, content_w - self.margin, title_height),
                    "content": Rect(img_w + self.margin, top_y, content_w - self.margin, content_h),
                    "image": Rect(0, 0, img_w, self.height),
                }
            )

    def icon_rows(self, count: int = 3, title_height: float = 0.8) -> SlideLayout:
        """Icon + text rows for feature lists."""
        top_y = self.content_y + title_height + 0.4
        available_h = self.content_h - title_height - 0.4
        row_h = available_h / count
        icon_size = min(0.6, row_h * 0.7)
        icon_margin = 0.3

        regions = {
            "title": Rect(self.content_x, self.content_y, self.content_w, title_height),
        }
        for i in range(count):
            y = top_y + i * row_h
            regions[f"icon_{i}"] = Rect(self.content_x, y + (row_h - icon_size)/2,
                                       icon_size, icon_size)
            regions[f"text_{i}"] = Rect(self.content_x + icon_size + icon_margin, y,
                                       self.content_w - icon_size - icon_margin, row_h)

        return SlideLayout(description=f"Icon Rows ({count})", regions=regions)


if __name__ == "__main__":
    engine = LayoutEngine()
    print("Available layouts:")
    layouts = [
        engine.title_only(),
        engine.title_subtitle(),
        engine.title_content(),
        engine.two_column(),
        engine.three_column(),
        engine.card_grid(2, 3),
        engine.stat_callouts(3),
        engine.timeline(4),
        engine.comparison(),
        engine.half_image("right"),
        engine.icon_rows(4),
    ]
    for layout in layouts:
        print(f"  {layout}")
