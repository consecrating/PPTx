"""Advanced Theme Engine for PPTX presentations.

Provides professional color palettes, gradient generation, and theme management.
Supports automatic palette derivation from brand colors and accessibility checking.

Usage:
    from advanced.theme_engine import ThemeEngine, get_palette

    # Get a predefined palette
    palette = get_palette("midnight_executive")

    # Generate a custom palette from a brand color
    engine = ThemeEngine()
    palette = engine.derive_palette("#1E2761")

    # Check color accessibility
    ratio = engine.contrast_ratio("#FFFFFF", "#1E2761")
"""

import colorsys
import math
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ColorPalette:
    """A professional presentation color palette."""
    name: str
    primary: str          # Main dominant color (60-70% visual weight)
    secondary: str        # Supporting color
    accent: str           # Sharp accent color
    text_dark: str        # Dark text color
    text_light: str       # Light text on dark backgrounds
    background_light: str # Light slide backgrounds
    background_dark: str  # Dark slide backgrounds
    success: str = "2ECC71"
    warning: str = "F39C12"
    danger: str = "E74C3C"
    info: str = "3498DB"
    muted: str = "95A5A6"
    chart_colors: list = field(default_factory=list)

    def __post_init__(self):
        if not self.chart_colors:
            self.chart_colors = [
                self.primary, self.secondary, self.accent,
                self.success, self.warning, self.info
            ]


# Predefined professional palettes
PALETTES = {
    "midnight_executive": ColorPalette(
        name="Midnight Executive",
        primary="1E2761", secondary="CADCFC", accent="7B68EE",
        text_dark="1A1A2E", text_light="FFFFFF",
        background_light="FFFFFF", background_dark="0F0F1A",
        chart_colors=["1E2761", "4169E1", "7B68EE", "CADCFC", "6C63FF", "9B59B6"]
    ),
    "forest_moss": ColorPalette(
        name="Forest & Moss",
        primary="2C5F2D", secondary="97BC62", accent="E8E8E8",
        text_dark="1B3A1C", text_light="F5F5F5",
        background_light="FFFFFF", background_dark="162517",
        chart_colors=["2C5F2D", "97BC62", "5D9B4A", "B8D98C", "3E8E41", "6DBE45"]
    ),
    "coral_energy": ColorPalette(
        name="Coral Energy",
        primary="F96167", secondary="F9E795", accent="2F3C7E",
        text_dark="2F3C7E", text_light="FFFFFF",
        background_light="FFFFFF", background_dark="1A2040",
        chart_colors=["F96167", "F9E795", "2F3C7E", "FF8A80", "FFD54F", "5C6BC0"]
    ),
    "warm_terracotta": ColorPalette(
        name="Warm Terracotta",
        primary="B85042", secondary="E7E8D1", accent="A7BEAE",
        text_dark="3D1F17", text_light="F5F5F0",
        background_light="FFFFFF", background_dark="2D1810",
        chart_colors=["B85042", "A7BEAE", "E7C8A0", "D4726A", "8FB39A", "C9A96E"]
    ),
    "ocean_gradient": ColorPalette(
        name="Ocean Gradient",
        primary="065A82", secondary="1C7293", accent="02C39A",
        text_dark="21295C", text_light="FFFFFF",
        background_light="FFFFFF", background_dark="0A1628",
        chart_colors=["065A82", "1C7293", "02C39A", "21295C", "00A896", "028090"]
    ),
    "charcoal_minimal": ColorPalette(
        name="Charcoal Minimal",
        primary="36454F", secondary="F2F2F2", accent="E63946",
        text_dark="212121", text_light="FFFFFF",
        background_light="FFFFFF", background_dark="1A1A1A",
        chart_colors=["36454F", "E63946", "457B9D", "A8DADC", "1D3557", "F4A261"]
    ),
    "teal_trust": ColorPalette(
        name="Teal Trust",
        primary="028090", secondary="00A896", accent="02C39A",
        text_dark="023047", text_light="FFFFFF",
        background_light="FFFFFF", background_dark="011627",
        chart_colors=["028090", "00A896", "02C39A", "05668D", "007F5F", "80B918"]
    ),
    "berry_cream": ColorPalette(
        name="Berry & Cream",
        primary="6D2E46", secondary="A26769", accent="ECE2D0",
        text_dark="3D1A2B", text_light="FFF8F0",
        background_light="FFFFFF", background_dark="2A1020",
        chart_colors=["6D2E46", "A26769", "D4A5A5", "8B4367", "C97B84", "E8B4B8"]
    ),
    "sage_calm": ColorPalette(
        name="Sage Calm",
        primary="84B59F", secondary="69A297", accent="50808E",
        text_dark="2D4A3E", text_light="F0F5F3",
        background_light="FFFFFF", background_dark="1A2F28",
        chart_colors=["84B59F", "69A297", "50808E", "A3D9C5", "4A7C6F", "7FB3A4"]
    ),
    "cherry_bold": ColorPalette(
        name="Cherry Bold",
        primary="990011", secondary="FCF6F5", accent="2F3C7E",
        text_dark="1A0003", text_light="FCF6F5",
        background_light="FFFFFF", background_dark="1A0003",
        chart_colors=["990011", "2F3C7E", "CC0033", "5C6BC0", "FF4444", "7986CB"]
    ),
    "sunset_warmth": ColorPalette(
        name="Sunset Warmth",
        primary="FF6B35", secondary="F7C59F", accent="004E89",
        text_dark="1A1A2E", text_light="FFFCF5",
        background_light="FFFFFF", background_dark="1A0E06",
        chart_colors=["FF6B35", "004E89", "F7C59F", "1A659E", "FF9F1C", "2EC4B6"]
    ),
    "deep_purple": ColorPalette(
        name="Deep Purple",
        primary="4A148C", secondary="CE93D8", accent="FF6F00",
        text_dark="1A0033", text_light="F3E5F5",
        background_light="FFFFFF", background_dark="12002B",
        chart_colors=["4A148C", "7B1FA2", "CE93D8", "FF6F00", "AB47BC", "E040FB"]
    ),
}


def get_palette(name: str) -> ColorPalette:
    """Get a predefined color palette by name."""
    if name not in PALETTES:
        available = ", ".join(sorted(PALETTES.keys()))
        raise ValueError(f"Unknown palette '{name}'. Available: {available}")
    return PALETTES[name]


def list_palettes() -> list[str]:
    """List all available palette names."""
    return sorted(PALETTES.keys())


class ThemeEngine:
    """Advanced theme engine for generating and manipulating color palettes."""

    @staticmethod
    def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
        """Convert hex color (without #) to RGB tuple."""
        h = hex_color.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

    @staticmethod
    def rgb_to_hex(r: int, g: int, b: int) -> str:
        """Convert RGB to hex string (without #)."""
        return f"{r:02X}{g:02X}{b:02X}"

    @staticmethod
    def hex_to_hsl(hex_color: str) -> tuple[float, float, float]:
        """Convert hex to HSL."""
        r, g, b = ThemeEngine.hex_to_rgb(hex_color)
        h, l, s = colorsys.rgb_to_hls(r/255, g/255, b/255)
        return h * 360, s * 100, l * 100

    @staticmethod
    def hsl_to_hex(h: float, s: float, l: float) -> str:
        """Convert HSL to hex."""
        r, g, b = colorsys.hls_to_rgb(h/360, l/100, s/100)
        return ThemeEngine.rgb_to_hex(int(r*255), int(g*255), int(b*255))

    @staticmethod
    def relative_luminance(hex_color: str) -> float:
        """Calculate relative luminance for WCAG contrast calculations."""
        r, g, b = ThemeEngine.hex_to_rgb(hex_color)
        rs, gs, bs = r/255, g/255, b/255
        r_lin = rs/12.92 if rs <= 0.03928 else ((rs + 0.055)/1.055) ** 2.4
        g_lin = gs/12.92 if gs <= 0.03928 else ((gs + 0.055)/1.055) ** 2.4
        b_lin = bs/12.92 if bs <= 0.03928 else ((bs + 0.055)/1.055) ** 2.4
        return 0.2126 * r_lin + 0.7152 * g_lin + 0.0722 * b_lin

    @staticmethod
    def contrast_ratio(color1: str, color2: str) -> float:
        """Calculate WCAG 2.1 contrast ratio between two colors."""
        l1 = ThemeEngine.relative_luminance(color1)
        l2 = ThemeEngine.relative_luminance(color2)
        lighter = max(l1, l2)
        darker = min(l1, l2)
        return (lighter + 0.05) / (darker + 0.05)

    @staticmethod
    def passes_wcag_aa(foreground: str, background: str, large_text: bool = False) -> bool:
        """Check if color combination passes WCAG AA."""
        ratio = ThemeEngine.contrast_ratio(foreground, background)
        threshold = 3.0 if large_text else 4.5
        return ratio >= threshold

    @staticmethod
    def passes_wcag_aaa(foreground: str, background: str, large_text: bool = False) -> bool:
        """Check if color combination passes WCAG AAA."""
        ratio = ThemeEngine.contrast_ratio(foreground, background)
        threshold = 4.5 if large_text else 7.0
        return ratio >= threshold

    def derive_palette(self, brand_color: str, style: str = "professional") -> ColorPalette:
        """Generate a full palette from a single brand color."""
        brand = brand_color.lstrip("#")
        h, s, l = self.hex_to_hsl(brand)

        if style == "professional":
            secondary = self.hsl_to_hex(h, max(s - 30, 10), min(l + 30, 90))
            accent = self.hsl_to_hex((h + 180) % 360, s, l)
            bg_dark = self.hsl_to_hex(h, min(s, 30), 8)
        elif style == "vibrant":
            secondary = self.hsl_to_hex((h + 30) % 360, s, l)
            accent = self.hsl_to_hex((h + 150) % 360, min(s + 20, 100), l)
            bg_dark = self.hsl_to_hex(h, 40, 10)
        elif style == "muted":
            secondary = self.hsl_to_hex(h, max(s - 40, 5), min(l + 40, 92))
            accent = self.hsl_to_hex((h + 120) % 360, max(s - 20, 10), l)
            bg_dark = self.hsl_to_hex(h, 15, 12)
        else:
            raise ValueError(f"Unknown style: {style}")

        return ColorPalette(
            name=f"Custom ({brand})",
            primary=brand,
            secondary=secondary,
            accent=accent,
            text_dark=self.hsl_to_hex(h, min(s, 20), 15),
            text_light="FFFFFF",
            background_light="FFFFFF",
            background_dark=bg_dark,
        )

    def generate_gradient_stops(self, color1: str, color2: str, steps: int = 5) -> list[str]:
        """Generate gradient color stops between two colors."""
        r1, g1, b1 = self.hex_to_rgb(color1)
        r2, g2, b2 = self.hex_to_rgb(color2)
        stops = []
        for i in range(steps):
            t = i / (steps - 1)
            r = int(r1 + (r2 - r1) * t)
            g = int(g1 + (g2 - g1) * t)
            b = int(b1 + (b2 - b1) * t)
            stops.append(self.rgb_to_hex(r, g, b))
        return stops

    def lighten(self, hex_color: str, amount: float = 20) -> str:
        """Lighten a color by a percentage."""
        h, s, l = self.hex_to_hsl(hex_color)
        return self.hsl_to_hex(h, s, min(l + amount, 100))

    def darken(self, hex_color: str, amount: float = 20) -> str:
        """Darken a color by a percentage."""
        h, s, l = self.hex_to_hsl(hex_color)
        return self.hsl_to_hex(h, s, max(l - amount, 0))

    def desaturate(self, hex_color: str, amount: float = 30) -> str:
        """Desaturate a color."""
        h, s, l = self.hex_to_hsl(hex_color)
        return self.hsl_to_hex(h, max(s - amount, 0), l)


if __name__ == "__main__":
    print("Available palettes:")
    for name in list_palettes():
        p = get_palette(name)
        print(f"  {name}: {p.name} (primary: #{p.primary})")

    print("\nDerived palette from brand blue #1E90FF:")
    engine = ThemeEngine()
    custom = engine.derive_palette("1E90FF")
    print(f"  Primary: #{custom.primary}")
    print(f"  Secondary: #{custom.secondary}")
    print(f"  Accent: #{custom.accent}")
    print(f"  Contrast (white on primary): {engine.contrast_ratio('FFFFFF', custom.primary):.2f}:1")
