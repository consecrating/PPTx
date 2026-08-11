"""Advanced Content Formatter for PPTX presentations.

Handles text formatting, smart text fitting, bullet generation,
and content structure optimization.

Usage:
    from advanced.content_formatter import ContentFormatter

    formatter = ContentFormatter()
    text_opts = formatter.title_text("Hello World", color="1E2761")
    bullets = formatter.bullet_list(["Item 1", "Item 2", "Item 3"])
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class TextStyle:
    """Text styling options for pptxgenjs."""
    fontSize: int = 14
    color: str = "333333"
    bold: bool = False
    italic: bool = False
    fontFace: str = "Calibri"
    align: str = "left"
    valign: str = "top"
    margin: float = 0
    lineSpacing: Optional[int] = None
    charSpacing: Optional[float] = None  # Note: NOT letterSpacing (that's ignored)

    def to_dict(self) -> dict:
        d = {
            "fontSize": self.fontSize,
            "color": self.color,
            "bold": self.bold,
            "italic": self.italic,
            "fontFace": self.fontFace,
            "align": self.align,
            "valign": self.valign,
            "margin": self.margin,
        }
        if self.lineSpacing:
            d["lineSpacing"] = self.lineSpacing
        if self.charSpacing:
            d["charSpacing"] = self.charSpacing
        return {k: v for k, v in d.items() if v is not None}


class ContentFormatter:
    """Format content for professional PPTX presentations."""

    # Safe fonts (true-to-width in QA AND ship with Office)
    SAFE_FONTS = {
        "sans": ["Arial", "Calibri"],
        "serif": ["Cambria", "Times New Roman", "Bookman Old Style", "Century Schoolbook"],
        "mono": ["Courier New"],
    }

    def __init__(self, primary_font: str = "Calibri", heading_font: str = "Cambria",
                 primary_color: str = "333333", accent_color: str = "1E2761"):
        self.primary_font = primary_font
        self.heading_font = heading_font
        self.primary_color = primary_color
        self.accent_color = accent_color

    def title_text(self, text: str, color: Optional[str] = None,
                   size: int = 40, align: str = "left") -> list[dict]:
        """Format text as a slide title."""
        return [{
            "text": text,
            "options": {
                "fontSize": size,
                "color": color or self.accent_color,
                "bold": True,
                "fontFace": self.heading_font,
                "align": align,
            }
        }]

    def subtitle_text(self, text: str, color: Optional[str] = None,
                      size: int = 20) -> list[dict]:
        """Format text as a subtitle."""
        return [{
            "text": text,
            "options": {
                "fontSize": size,
                "color": color or "666666",
                "fontFace": self.primary_font,
                "italic": True,
            }
        }]

    def body_text(self, text: str, size: int = 14, bold: bool = False) -> list[dict]:
        """Format body text."""
        return [{
            "text": text,
            "options": {
                "fontSize": size,
                "color": self.primary_color,
                "fontFace": self.primary_font,
                "bold": bold,
            }
        }]

    def bullet_list(self, items: list[str], size: int = 14,
                    numbered: bool = False, indent_level: int = 0,
                    color: Optional[str] = None) -> list[dict]:
        """Format a bullet or numbered list.

        IMPORTANT: Use bullet:true, never literal bullet chars.
        Set breakLine:true on every item except the last.
        """
        result = []
        for i, item in enumerate(items):
            opts = {
                "fontSize": size,
                "color": color or self.primary_color,
                "fontFace": self.primary_font,
                "bullet": {"type": "number"} if numbered else True,
                "indentLevel": indent_level,
                "paraSpaceAfter": 6,
            }
            if i < len(items) - 1:
                opts["breakLine"] = True
            result.append({"text": item, "options": opts})
        return result

    def stat_number(self, number: str, label: str,
                    number_size: int = 60, label_size: int = 12,
                    number_color: Optional[str] = None) -> list[dict]:
        """Format a large statistic with label."""
        return [
            {
                "text": number,
                "options": {
                    "fontSize": number_size,
                    "color": number_color or self.accent_color,
                    "bold": True,
                    "fontFace": self.heading_font,
                    "align": "center",
                    "breakLine": True,
                }
            },
            {
                "text": label,
                "options": {
                    "fontSize": label_size,
                    "color": "888888",
                    "fontFace": self.primary_font,
                    "align": "center",
                }
            }
        ]

    def section_header(self, text: str, size: int = 22,
                       color: Optional[str] = None) -> list[dict]:
        """Format a section header within a slide."""
        return [{
            "text": text,
            "options": {
                "fontSize": size,
                "color": color or self.accent_color,
                "bold": True,
                "fontFace": self.heading_font,
                "paraSpaceAfter": 8,
            }
        }]

    def caption_text(self, text: str, size: int = 10) -> list[dict]:
        """Format caption/footnote text."""
        return [{
            "text": text,
            "options": {
                "fontSize": size,
                "color": "999999",
                "fontFace": self.primary_font,
                "italic": True,
            }
        }]

    def two_tone_text(self, highlight: str, normal: str,
                      highlight_color: Optional[str] = None,
                      size: int = 14) -> list[dict]:
        """Format text with a highlighted portion."""
        return [
            {
                "text": highlight,
                "options": {
                    "fontSize": size,
                    "color": highlight_color or self.accent_color,
                    "bold": True,
                    "fontFace": self.primary_font,
                }
            },
            {
                "text": f" {normal}",
                "options": {
                    "fontSize": size,
                    "color": self.primary_color,
                    "fontFace": self.primary_font,
                }
            }
        ]

    def estimate_text_width(self, text: str, font_size: int, bold: bool = False) -> float:
        """Estimate text width in inches (approximate).

        Uses average character widths for common fonts.
        Calibri: ~0.055 inches per pt per char
        """
        avg_char_width = font_size * 0.006  # Approximate inches per character
        if bold:
            avg_char_width *= 1.05
        return len(text) * avg_char_width

    def fit_text_to_width(self, text: str, max_width: float,
                          max_size: int = 40, min_size: int = 10) -> int:
        """Calculate the largest font size that fits text in width."""
        for size in range(max_size, min_size - 1, -1):
            width = self.estimate_text_width(text, size)
            if width <= max_width:
                return size
        return min_size

    def wrap_text(self, text: str, max_chars_per_line: int = 60) -> str:
        """Wrap text at word boundaries."""
        words = text.split()
        lines = []
        current_line = ""
        for word in words:
            if len(current_line) + len(word) + 1 <= max_chars_per_line:
                current_line += (" " if current_line else "") + word
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        return "\n".join(lines)


if __name__ == "__main__":
    formatter = ContentFormatter()
    print("Title:", formatter.title_text("Sample Presentation"))
    print("Bullets:", formatter.bullet_list(["Feature 1", "Feature 2", "Feature 3"]))
    print("Stat:", formatter.stat_number("98%", "Customer Satisfaction"))
    print("Fit size for 'Hello World' in 4 inches:", formatter.fit_text_to_width("Hello World", 4.0))
