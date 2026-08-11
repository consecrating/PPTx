"""Animation & Transition Helper for PPTX presentations.

Provides utilities for adding slide transitions and generating
animation-ready element configurations.

Note: pptxgenjs supports slide-level transitions but not per-element animations.
For per-element animations, use the XML editing approach documented in SKILL.md.

Usage:
    from advanced.animation_helper import TransitionHelper

    helper = TransitionHelper()
    transition = helper.fade(duration=0.5)
    transition = helper.push(direction="left", duration=0.7)
"""

from dataclasses import dataclass
from typing import Optional


@dataclass
class SlideTransition:
    """Slide transition configuration for pptxgenjs."""
    type: str
    speed: Optional[float] = None  # Duration in seconds
    advance_time: Optional[float] = None  # Auto-advance after N seconds

    def to_options(self) -> dict:
        """Convert to pptxgenjs slide options."""
        opts = {"transition": {"type": self.type}}
        if self.speed:
            opts["transition"]["speed"] = self.speed
        if self.advance_time:
            opts["advanceAfter"] = int(self.advance_time * 1000)
        return opts


@dataclass
class AnimationConfig:
    """Element animation XML configuration (for XML editing approach)."""
    element_id: str
    effect: str
    trigger: str = "onClick"  # onClick, withPrev, afterPrev
    duration: int = 500  # milliseconds
    delay: int = 0  # milliseconds

    def to_xml(self) -> str:
        """Generate OOXML animation XML snippet."""
        return f"""<p:par>
  <p:cTn id="{{auto}}" fill="hold">
    <p:stCondLst>
      <p:cond delay="{self.delay}"/>
    </p:stCondLst>
    <p:childTnLst>
      <p:par>
        <p:cTn id="{{auto}}" presetID="{self._preset_id()}" presetClass="entr"
               presetSubtype="0" fill="hold" dur="{self.duration}">
          <p:stCondLst>
            <p:cond delay="0"/>
          </p:stCondLst>
        </p:cTn>
      </p:par>
    </p:childTnLst>
  </p:cTn>
</p:par>"""

    def _preset_id(self) -> int:
        """Map effect names to OOXML preset IDs."""
        presets = {
            "appear": 1,
            "fade": 10,
            "fly_in": 2,
            "wipe": 22,
            "split": 16,
            "wheel": 21,
            "zoom": 53,
            "bounce": 26,
            "float_up": 42,
            "grow": 52,
        }
        return presets.get(self.effect, 1)


class TransitionHelper:
    """Helper for slide transitions in pptxgenjs."""

    # Available transition types in pptxgenjs
    TRANSITIONS = {
        "fade": "fade",
        "push": "push",
        "wipe": "wipe",
        "zoom": "zoom",
        "none": "none",
        "cover": "cover",
        "uncover": "uncover",
        "cut": "cut",
    }

    def fade(self, duration: float = 0.5, advance_time: Optional[float] = None) -> SlideTransition:
        """Fade transition between slides."""
        return SlideTransition(type="fade", speed=duration, advance_time=advance_time)

    def push(self, direction: str = "left", duration: float = 0.7,
             advance_time: Optional[float] = None) -> SlideTransition:
        """Push transition (one slide pushes another off)."""
        return SlideTransition(type="push", speed=duration, advance_time=advance_time)

    def wipe(self, duration: float = 0.5, advance_time: Optional[float] = None) -> SlideTransition:
        """Wipe transition."""
        return SlideTransition(type="wipe", speed=duration, advance_time=advance_time)

    def zoom(self, duration: float = 0.6, advance_time: Optional[float] = None) -> SlideTransition:
        """Zoom transition."""
        return SlideTransition(type="zoom", speed=duration, advance_time=advance_time)

    def cover(self, duration: float = 0.5, advance_time: Optional[float] = None) -> SlideTransition:
        """Cover transition."""
        return SlideTransition(type="cover", speed=duration, advance_time=advance_time)

    def none(self) -> SlideTransition:
        """No transition (instant cut)."""
        return SlideTransition(type="none")

    def auto_advance(self, seconds: float = 5.0) -> dict:
        """Auto-advance slide after N seconds (for kiosk mode)."""
        return {"advanceAfter": int(seconds * 1000)}

    @staticmethod
    def recommended_for(slide_type: str) -> SlideTransition:
        """Get recommended transition for a slide type."""
        recommendations = {
            "title": SlideTransition(type="fade", speed=0.8),
            "content": SlideTransition(type="fade", speed=0.4),
            "section": SlideTransition(type="push", speed=0.6),
            "conclusion": SlideTransition(type="fade", speed=1.0),
            "data": SlideTransition(type="wipe", speed=0.5),
            "comparison": SlideTransition(type="push", speed=0.5),
        }
        return recommendations.get(slide_type, SlideTransition(type="fade", speed=0.5))


class AnimationSequenceBuilder:
    """Build sequential element animations for XML editing."""

    def __init__(self):
        self.animations: list[AnimationConfig] = []
        self._delay_accumulator = 0

    def add(self, element_id: str, effect: str = "fade",
            duration: int = 500, delay: int = 200,
            trigger: str = "afterPrev") -> "AnimationSequenceBuilder":
        """Add an animation to the sequence."""
        self.animations.append(AnimationConfig(
            element_id=element_id,
            effect=effect,
            trigger=trigger,
            duration=duration,
            delay=self._delay_accumulator + delay,
        ))
        self._delay_accumulator += duration + delay
        return self

    def build_stagger(self, element_ids: list[str], effect: str = "fade",
                      duration: int = 400, stagger_delay: int = 150) -> list[AnimationConfig]:
        """Build a staggered animation for multiple elements."""
        animations = []
        for i, eid in enumerate(element_ids):
            animations.append(AnimationConfig(
                element_id=eid,
                effect=effect,
                trigger="afterPrev" if i > 0 else "onClick",
                duration=duration,
                delay=stagger_delay * i,
            ))
        return animations

    def get_sequence(self) -> list[AnimationConfig]:
        """Get the built animation sequence."""
        return self.animations.copy()


if __name__ == "__main__":
    helper = TransitionHelper()
    print("Available transitions:", list(TransitionHelper.TRANSITIONS.keys()))
    print("\nFade transition:", helper.fade().to_options())
    print("Push transition:", helper.push().to_options())
    print("Recommended for title:", TransitionHelper.recommended_for("title").to_options())

    builder = AnimationSequenceBuilder()
    builder.add("title_1", "fade", 500)
    builder.add("bullet_1", "fly_in", 400)
    builder.add("bullet_2", "fly_in", 400)
    print("\nAnimation sequence:")
    for anim in builder.get_sequence():
        print(f"  {anim.element_id}: {anim.effect} (delay={anim.delay}ms)")
