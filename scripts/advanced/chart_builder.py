"""Advanced Chart Builder for pptxgenjs presentations.

Generates properly configured chart options that avoid common pitfalls.
Supports combo charts, dual-axis, and proper styling.

Usage:
    from advanced.chart_builder import ChartBuilder

    builder = ChartBuilder(palette=["1E2761", "4169E1", "7B68EE"])
    config = builder.bar_chart(data, title="Revenue by Quarter")
    config = builder.combo_chart(bar_data, line_data, title="Revenue vs Growth")
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ChartConfig:
    """Configuration for a pptxgenjs chart."""
    chart_type: str
    data: list
    options: dict = field(default_factory=dict)
    combo: bool = False

    def to_js(self) -> str:
        """Generate JavaScript code for pptxgenjs."""
        lines = []
        if self.combo:
            lines.append(f"slide.addChart({self.chart_type}, {self._format_data()}, {self._format_options()});")
        else:
            lines.append(f"slide.addChart(pptx.charts.{self.chart_type}, {self._format_data()}, {self._format_options()});")
        return "\n".join(lines)

    def _format_data(self) -> str:
        import json
        return json.dumps(self.data, indent=2)

    def _format_options(self) -> str:
        import json
        return json.dumps(self.options, indent=2)


class ChartBuilder:
    """Build professional chart configurations for pptxgenjs."""

    CHART_TYPES = {
        "bar": "BAR",
        "bar3d": "BAR3D",
        "line": "LINE",
        "area": "AREA",
        "pie": "PIE",
        "doughnut": "DOUGHNUT",
        "scatter": "SCATTER",
        "bubble": "BUBBLE",
        "radar": "RADAR",
    }

    def __init__(self, palette: Optional[list[str]] = None):
        self.palette = palette or [
            "1E2761", "4169E1", "7B68EE", "CADCFC", "6C63FF", "9B59B6"
        ]

    def _base_options(self, title: str = "", position: Optional[dict] = None) -> dict:
        """Base chart options with professional defaults."""
        opts = {
            "showTitle": bool(title),
            "chartColors": self.palette[:6],
            "catAxisLabelColor": "666666",
            "valAxisLabelColor": "666666",
            "catAxisLabelFontSize": 10,
            "valAxisLabelFontSize": 10,
            "valGridLine": {"color": "E8E8E8", "size": 0.5},
            "catGridLine": {"style": "none"},
            "showLegend": False,
            "legendPos": "b",
            "legendFontSize": 9,
            "legendColor": "666666",
        }
        if title:
            opts["title"] = title
            opts["titleColor"] = "333333"
            opts["titleFontSize"] = 14
        if position:
            opts.update(position)
        return opts

    def bar_chart(self, data: list[dict], title: str = "",
                  stacked: bool = False, horizontal: bool = False,
                  show_values: bool = True, position: Optional[dict] = None) -> ChartConfig:
        """Create a bar/column chart configuration."""
        opts = self._base_options(title, position)

        if stacked:
            opts["barGrouping"] = "stacked"
            if show_values:
                opts["showValue"] = True
                opts["dataLabelPosition"] = "ctr"  # Safe for stacked
        else:
            if show_values:
                opts["showValue"] = True
                opts["dataLabelPosition"] = "outEnd"

        opts["dataLabelColor"] = "555555"
        opts["dataLabelFontSize"] = 9

        if horizontal:
            opts["barDir"] = "bar"

        if len(data) > 1:
            opts["showLegend"] = True

        return ChartConfig(
            chart_type="BAR",
            data=data,
            options=opts,
        )

    def line_chart(self, data: list[dict], title: str = "",
                   smooth: bool = False, show_markers: bool = True,
                   show_values: bool = False, position: Optional[dict] = None) -> ChartConfig:
        """Create a line chart configuration."""
        opts = self._base_options(title, position)

        if smooth:
            opts["lineSmooth"] = True
        if show_markers:
            opts["lineDataSymbol"] = "circle"
            opts["lineDataSymbolSize"] = 6
        if show_values:
            opts["showValue"] = True
            opts["dataLabelPosition"] = "t"

        opts["lineSize"] = 2.5

        if len(data) > 1:
            opts["showLegend"] = True

        return ChartConfig(
            chart_type="LINE",
            data=data,
            options=opts,
        )

    def pie_chart(self, data: list[dict], title: str = "",
                  doughnut: bool = False, show_percent: bool = True,
                  position: Optional[dict] = None) -> ChartConfig:
        """Create a pie/doughnut chart configuration."""
        opts = self._base_options(title, position)
        opts["showLegend"] = True
        opts["legendPos"] = "r"

        if show_percent:
            opts["showPercent"] = True
            opts["dataLabelColor"] = "FFFFFF"
            opts["dataLabelFontSize"] = 11

        chart_type = "DOUGHNUT" if doughnut else "PIE"
        if doughnut:
            opts["holeSize"] = 50

        return ChartConfig(
            chart_type=chart_type,
            data=data,
            options=opts,
        )

    def combo_chart(self, bar_data: list[dict], line_data: list[dict],
                    title: str = "", position: Optional[dict] = None) -> ChartConfig:
        """Create a combo (bar + line) chart with dual axes.

        IMPORTANT: Combo charts with secondaryValAxis need both valAxes and catAxes
        defined, two entries each, or PowerPoint rejects the file.
        """
        opts = self._base_options(title, position)
        opts["showLegend"] = True

        # CRITICAL: Define both axes pairs for combo charts
        opts["valAxes"] = [
            {
                "showValAxisTitle": True,
                "valAxisTitle": "Values",
                "valAxisTitleColor": "666666",
                "valAxisLabelColor": "666666",
                "valGridLine": {"color": "E8E8E8", "size": 0.5},
            },
            {
                "showValAxisTitle": True,
                "valAxisTitle": "Percentage",
                "valAxisTitleColor": "666666",
                "valAxisLabelColor": "666666",
                "valAxisHidden": False,
            }
        ]
        opts["catAxes"] = [
            {"catAxisLabelColor": "666666"},
            {"catAxisHidden": True},
        ]

        # Build combo data array
        combo_data = []
        for d in bar_data:
            combo_data.append({
                "type": "pptx.charts.BAR",
                "data": [d],
                "options": {"chartColors": [self.palette[0]]},
            })
        for d in line_data:
            combo_data.append({
                "type": "pptx.charts.LINE",
                "data": [d],
                "options": {
                    "chartColors": [self.palette[2]],
                    "secondaryValAxis": True,
                    "secondaryCatAxis": True,
                    "lineSize": 2.5,
                    "lineDataSymbol": "circle",
                },
            })

        return ChartConfig(
            chart_type="combo_data",
            data=combo_data,
            options=opts,
            combo=True,
        )

    def scatter_chart(self, data: list[dict], title: str = "",
                      show_trend: bool = False, position: Optional[dict] = None) -> ChartConfig:
        """Create a scatter plot configuration."""
        opts = self._base_options(title, position)
        opts["lineDataSymbol"] = "circle"
        opts["lineDataSymbolSize"] = 8
        opts["lineSize"] = 0

        if len(data) > 1:
            opts["showLegend"] = True

        return ChartConfig(
            chart_type="SCATTER",
            data=data,
            options=opts,
        )

    def area_chart(self, data: list[dict], title: str = "",
                   stacked: bool = False, position: Optional[dict] = None) -> ChartConfig:
        """Create an area chart configuration."""
        opts = self._base_options(title, position)
        opts["opacity"] = 50

        if stacked:
            opts["barGrouping"] = "stacked"

        if len(data) > 1:
            opts["showLegend"] = True

        return ChartConfig(
            chart_type="AREA",
            data=data,
            options=opts,
        )


if __name__ == "__main__":
    builder = ChartBuilder()

    sample_data = [
        {"name": "Revenue", "labels": ["Q1", "Q2", "Q3", "Q4"], "values": [100, 120, 140, 180]},
    ]
    config = builder.bar_chart(sample_data, title="Quarterly Revenue")
    print(f"Bar chart config: {config.chart_type}")
    print(f"  Options: showValue={config.options.get('showValue')}")

    config = builder.pie_chart(sample_data, title="Revenue Split", doughnut=True)
    print(f"\nDoughnut chart config: {config.chart_type}")
