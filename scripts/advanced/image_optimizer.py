"""Image Optimizer for PPTX presentations.

Handles image preparation, resizing, format conversion, and optimization
for embedding in presentations.

Usage:
    from advanced.image_optimizer import ImageOptimizer

    optimizer = ImageOptimizer()
    optimized = optimizer.prepare_for_slide(image_path, max_width=1920)
    gradient_img = optimizer.create_gradient(1920, 1080, "1E2761", "7B68EE")
"""

import base64
import io
import math
from pathlib import Path
from typing import Optional, Tuple

try:
    from PIL import Image, ImageDraw, ImageFilter, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


class ImageOptimizer:
    """Optimize and prepare images for PPTX embedding."""

    # Recommended max dimensions for various use cases
    FULL_SLIDE_16x9 = (1920, 1080)
    HALF_SLIDE = (960, 1080)
    THUMBNAIL = (400, 300)
    ICON_LARGE = (256, 256)
    ICON_SMALL = (64, 64)

    def __init__(self, quality: int = 85, max_file_size_kb: int = 500):
        self.quality = quality
        self.max_file_size_kb = max_file_size_kb
        if not HAS_PIL:
            print("Warning: Pillow not available. Image operations will be limited.")

    def prepare_for_slide(self, image_path: str, max_width: int = 1920,
                         max_height: int = 1080) -> Optional[str]:
        """Resize and optimize an image for slide embedding.
        Returns base64-encoded PNG/JPEG string for pptxgenjs.
        """
        if not HAS_PIL:
            return None

        img = Image.open(image_path)

        # Convert RGBA to RGB with white background if needed
        if img.mode == "RGBA":
            background = Image.new("RGB", img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            img = background

        # Resize maintaining aspect ratio
        img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)

        # Save to buffer
        buffer = io.BytesIO()
        img.save(buffer, format="PNG", optimize=True)
        buffer.seek(0)

        return "image/png;base64," + base64.b64encode(buffer.read()).decode()

    def create_gradient(self, width: int, height: int,
                       color1: str, color2: str,
                       direction: str = "horizontal") -> str:
        """Create a gradient image for slide backgrounds.

        Since pptxgenjs doesn't support gradient fills natively,
        use this to generate a gradient image as background.

        Returns base64-encoded PNG for pptxgenjs addImage().
        """
        if not HAS_PIL:
            raise RuntimeError("Pillow required for gradient generation")

        img = Image.new("RGB", (width, height))
        draw = ImageDraw.Draw(img)

        r1, g1, b1 = self._hex_to_rgb(color1)
        r2, g2, b2 = self._hex_to_rgb(color2)

        if direction == "horizontal":
            for x in range(width):
                t = x / (width - 1)
                r = int(r1 + (r2 - r1) * t)
                g = int(g1 + (g2 - g1) * t)
                b = int(b1 + (b2 - b1) * t)
                draw.line([(x, 0), (x, height)], fill=(r, g, b))
        elif direction == "vertical":
            for y in range(height):
                t = y / (height - 1)
                r = int(r1 + (r2 - r1) * t)
                g = int(g1 + (g2 - g1) * t)
                b = int(b1 + (b2 - b1) * t)
                draw.line([(0, y), (width, y)], fill=(r, g, b))
        elif direction == "diagonal":
            for y in range(height):
                for x in range(width):
                    t = (x / width + y / height) / 2
                    r = int(r1 + (r2 - r1) * t)
                    g = int(g1 + (g2 - g1) * t)
                    b = int(b1 + (b2 - b1) * t)
                    img.putpixel((x, y), (r, g, b))
        elif direction == "radial":
            cx, cy = width / 2, height / 2
            max_dist = math.sqrt(cx**2 + cy**2)
            for y in range(height):
                for x in range(width):
                    dist = math.sqrt((x - cx)**2 + (y - cy)**2)
                    t = min(dist / max_dist, 1.0)
                    r = int(r1 + (r2 - r1) * t)
                    g = int(g1 + (g2 - g1) * t)
                    b = int(b1 + (b2 - b1) * t)
                    img.putpixel((x, y), (r, g, b))

        buffer = io.BytesIO()
        img.save(buffer, format="PNG", optimize=True)
        buffer.seek(0)
        return "image/png;base64," + base64.b64encode(buffer.read()).decode()

    def create_solid_color(self, width: int, height: int, color: str) -> str:
        """Create a solid color image (useful for backgrounds)."""
        if not HAS_PIL:
            raise RuntimeError("Pillow required")

        img = Image.new("RGB", (width, height), self._hex_to_rgb(color))
        buffer = io.BytesIO()
        img.save(buffer, format="PNG", optimize=True)
        buffer.seek(0)
        return "image/png;base64," + base64.b64encode(buffer.read()).decode()

    def create_overlay(self, width: int, height: int,
                      color: str = "000000", opacity: float = 0.5) -> str:
        """Create a semi-transparent overlay image."""
        if not HAS_PIL:
            raise RuntimeError("Pillow required")

        r, g, b = self._hex_to_rgb(color)
        a = int(opacity * 255)
        img = Image.new("RGBA", (width, height), (r, g, b, a))
        buffer = io.BytesIO()
        img.save(buffer, format="PNG")
        buffer.seek(0)
        return "image/png;base64," + base64.b64encode(buffer.read()).decode()

    def create_circle_icon_bg(self, size: int = 64, color: str = "1E2761") -> str:
        """Create a circular background for icons."""
        if not HAS_PIL:
            raise RuntimeError("Pillow required")

        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        r, g, b = self._hex_to_rgb(color)
        draw.ellipse([0, 0, size-1, size-1], fill=(r, g, b, 255))

        buffer = io.BytesIO()
        img.save(buffer, format="PNG")
        buffer.seek(0)
        return "image/png;base64," + base64.b64encode(buffer.read()).decode()

    def get_image_dimensions(self, image_path: str) -> Tuple[int, int]:
        """Get image dimensions without loading full image."""
        if not HAS_PIL:
            return (0, 0)
        with Image.open(image_path) as img:
            return img.size

    def calculate_fit(self, img_width: int, img_height: int,
                     container_width: float, container_height: float) -> Tuple[float, float, float, float]:
        """Calculate position and size to fit image in container (in inches).
        Returns (x_offset, y_offset, display_width, display_height).
        """
        img_ratio = img_width / img_height
        container_ratio = container_width / container_height

        if img_ratio > container_ratio:
            # Image is wider - fit to width
            display_width = container_width
            display_height = container_width / img_ratio
        else:
            # Image is taller - fit to height
            display_height = container_height
            display_width = container_height * img_ratio

        x_offset = (container_width - display_width) / 2
        y_offset = (container_height - display_height) / 2

        return x_offset, y_offset, display_width, display_height

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
        h = hex_color.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


if __name__ == "__main__":
    optimizer = ImageOptimizer()
    print("Image Optimizer initialized")
    print(f"  Quality: {optimizer.quality}")
    print(f"  Max file size: {optimizer.max_file_size_kb}KB")
    print(f"  Full slide dimensions: {optimizer.FULL_SLIDE_16x9}")

    if HAS_PIL:
        # Create sample gradient
        gradient = optimizer.create_gradient(200, 100, "1E2761", "7B68EE")
        print(f"\n  Generated gradient: {len(gradient)} chars (base64)")

        # Test fit calculation
        x, y, w, h = optimizer.calculate_fit(1920, 1080, 10.0, 5.625)
        print(f"  Fit 1920x1080 in 10x5.625: ({x:.2f}, {y:.2f}, {w:.2f}, {h:.2f})")
