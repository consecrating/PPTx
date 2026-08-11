/**
 * Demo Presentation Generator
 * 
 * Demonstrates all advanced features of the PPTx skill:
 * - Professional color palette usage
 * - Multiple layout types
 * - Charts with safe configuration
 * - Transitions between slides
 * - Proper typography and spacing
 * 
 * Usage: node examples/demo_presentation.js
 * Output: examples/demo_output.pptx
 */

const pptxgen = require('pptxgenjs');

// ============================================================
// THEME CONFIGURATION (from theme_engine.py: midnight_executive)
// ============================================================
const THEME = {
    primary: "1E2761",
    secondary: "CADCFC",
    accent: "7B68EE",
    textDark: "1A1A2E",
    textLight: "FFFFFF",
    bgLight: "FFFFFF",
    bgDark: "0F0F1A",
    muted: "95A5A6",
    success: "2ECC71",
    warning: "F39C12",
    danger: "E74C3C",
    chartColors: ["1E2761", "4169E1", "7B68EE", "CADCFC", "6C63FF", "9B59B6"],
};

// ============================================================
// TYPOGRAPHY (safe fonts from SKILL.md)
// ============================================================
const FONTS = {
    heading: "Cambria",
    body: "Calibri",
};

// ============================================================
// CREATE PRESENTATION
// ============================================================
function createPresentation() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9"; // 10" x 5.625"

    // Define reusable master slides
    pres.defineSlideMaster({
        title: "TITLE_SLIDE",
        background: { color: THEME.bgDark },
    });

    pres.defineSlideMaster({
        title: "CONTENT_SLIDE",
        background: { color: THEME.bgLight },
    });

    // ----------------------------------------------------------
    // SLIDE 1: Title Slide (Dark background)
    // ----------------------------------------------------------
    const slide1 = pres.addSlide({ masterName: "TITLE_SLIDE" });
    slide1.transition = { type: "fade", speed: 0.8 };

    slide1.addText("Advanced PPTx Skill", {
        x: 0.5, y: 1.5, w: 9.0, h: 1.5,
        fontSize: 44, color: THEME.textLight, bold: true,
        fontFace: FONTS.heading, align: "left", margin: 0,
    });

    slide1.addText("Professional Presentations with AI", {
        x: 0.5, y: 3.0, w: 9.0, h: 0.8,
        fontSize: 20, color: THEME.secondary, italic: true,
        fontFace: FONTS.body, align: "left", margin: 0,
    });

    slide1.addText("Powered by pptxgenjs + Advanced Layout Engine", {
        x: 0.5, y: 4.8, w: 9.0, h: 0.4,
        fontSize: 12, color: THEME.muted,
        fontFace: FONTS.body, align: "left", margin: 0,
    });

    // ----------------------------------------------------------
    // SLIDE 2: Features Overview (Icon Rows Layout)
    // ----------------------------------------------------------
    const slide2 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
    slide2.transition = { type: "fade", speed: 0.4 };

    slide2.addText("Key Features", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: THEME.primary, bold: true,
        fontFace: FONTS.heading, align: "left", margin: 0,
    });

    const features = [
        { title: "Theme Engine", desc: "12 professional palettes with WCAG accessibility checking" },
        { title: "Layout Engine", desc: "11 intelligent layout types with automatic positioning" },
        { title: "Chart Builder", desc: "Safe configurations preventing PowerPoint corruption bugs" },
        { title: "Animation System", desc: "Slide transitions and staggered element animations" },
    ];

    features.forEach((feature, i) => {
        const y = 1.6 + i * 1.0;

        // Icon circle
        slide2.addShape(pres.shapes.OVAL, {
            x: 0.5, y: y + 0.1, w: 0.5, h: 0.5,
            fill: { color: THEME.primary },
        });

        // Feature number in circle
        slide2.addText(String(i + 1), {
            x: 0.5, y: y + 0.1, w: 0.5, h: 0.5,
            fontSize: 14, color: THEME.textLight, bold: true,
            fontFace: FONTS.body, align: "center", valign: "middle", margin: 0,
        });

        // Title
        slide2.addText(feature.title, {
            x: 1.2, y: y, w: 8.0, h: 0.4,
            fontSize: 16, color: THEME.primary, bold: true,
            fontFace: FONTS.heading, align: "left", margin: 0,
        });

        // Description
        slide2.addText(feature.desc, {
            x: 1.2, y: y + 0.4, w: 8.0, h: 0.4,
            fontSize: 13, color: "666666",
            fontFace: FONTS.body, align: "left", margin: 0,
        });
    });

    // ----------------------------------------------------------
    // SLIDE 3: Statistics (Stat Callouts Layout)
    // ----------------------------------------------------------
    const slide3 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
    slide3.transition = { type: "fade", speed: 0.4 };

    slide3.addText("By the Numbers", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: THEME.primary, bold: true,
        fontFace: FONTS.heading, align: "left", margin: 0,
    });

    const stats = [
        { number: "12", label: "Professional\nPalettes" },
        { number: "11", label: "Layout\nTypes" },
        { number: "6", label: "Chart\nFormats" },
        { number: "100%", label: "WCAG AA\nCompliant" },
    ];

    const statWidth = (9.0 - 0.6) / 4; // 4 stats with gaps
    stats.forEach((stat, i) => {
        const x = 0.5 + i * (statWidth + 0.2);
        const y = 2.0;

        // Number
        slide3.addText(stat.number, {
            x: x, y: y, w: statWidth, h: 1.2,
            fontSize: 52, color: THEME.accent, bold: true,
            fontFace: FONTS.heading, align: "center", valign: "bottom", margin: 0,
        });

        // Label
        slide3.addText(stat.label, {
            x: x, y: y + 1.4, w: statWidth, h: 0.8,
            fontSize: 12, color: "888888",
            fontFace: FONTS.body, align: "center", valign: "top", margin: 0,
        });
    });

    // ----------------------------------------------------------
    // SLIDE 4: Chart Demo (Bar Chart)
    // ----------------------------------------------------------
    const slide4 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
    slide4.transition = { type: "fade", speed: 0.4 };

    slide4.addText("Chart Performance", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: THEME.primary, bold: true,
        fontFace: FONTS.heading, align: "left", margin: 0,
    });

    // Bar chart with safe configuration
    slide4.addChart(pres.charts.BAR, [
        {
            name: "Q1 2024",
            labels: ["Theme", "Layout", "Charts", "Animation", "Images"],
            values: [95, 88, 92, 78, 85],
        },
        {
            name: "Q2 2024",
            labels: ["Theme", "Layout", "Charts", "Animation", "Images"],
            values: [98, 94, 96, 88, 92],
        },
    ], {
        x: 0.5, y: 1.5, w: 9.0, h: 3.8,
        showTitle: false,
        chartColors: [THEME.primary, THEME.accent],
        showValue: true,
        dataLabelPosition: "outEnd",  // Safe for non-stacked
        dataLabelColor: "555555",
        dataLabelFontSize: 9,
        catAxisLabelColor: "666666",
        valAxisLabelColor: "666666",
        catAxisLabelFontSize: 10,
        valAxisLabelFontSize: 10,
        valGridLine: { color: "E8E8E8", size: 0.5 },
        catGridLine: { style: "none" },
        showLegend: true,
        legendPos: "b",
        legendFontSize: 9,
        legendColor: "666666",
    });

    // ----------------------------------------------------------
    // SLIDE 5: Two-Column Comparison
    // ----------------------------------------------------------
    const slide5 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
    slide5.transition = { type: "push", speed: 0.6 };

    slide5.addText("Base vs Advanced", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: THEME.primary, bold: true,
        fontFace: FONTS.heading, align: "left", margin: 0,
    });

    // Left column header
    slide5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.5, y: 1.5, w: 4.2, h: 0.6,
        fill: { color: THEME.muted },
        rectRadius: 0.05,
    });
    slide5.addText("Base Skill", {
        x: 0.5, y: 1.5, w: 4.2, h: 0.6,
        fontSize: 16, color: THEME.textLight, bold: true,
        fontFace: FONTS.heading, align: "center", valign: "middle", margin: 0,
    });

    // Right column header
    slide5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 5.3, y: 1.5, w: 4.2, h: 0.6,
        fill: { color: THEME.primary },
        rectRadius: 0.05,
    });
    slide5.addText("Advanced (v2.0)", {
        x: 5.3, y: 1.5, w: 4.2, h: 0.6,
        fontSize: 16, color: THEME.textLight, bold: true,
        fontFace: FONTS.heading, align: "center", valign: "middle", margin: 0,
    });

    // Left items
    const baseFeatures = [
        "10 static palettes",
        "Manual positioning",
        "Manual chart config",
        "No animation support",
        "No image tools",
    ];
    baseFeatures.forEach((item, i) => {
        slide5.addText([{ text: item, options: { bullet: true, breakLine: i < baseFeatures.length - 1 } }], {
            x: 0.5, y: 2.3 + i * 0.5, w: 4.2, h: 0.45,
            fontSize: 13, color: "666666",
            fontFace: FONTS.body, align: "left", margin: 0,
            paraSpaceAfter: 4,
        });
    });

    // Right items
    const advFeatures = [
        "12 palettes + WCAG + derivation",
        "11 intelligent layout types",
        "Safe chart builder (no corruption)",
        "Transitions + animation sequences",
        "Gradients, overlays, optimization",
    ];
    advFeatures.forEach((item, i) => {
        slide5.addText([{ text: item, options: { bullet: true, breakLine: i < advFeatures.length - 1 } }], {
            x: 5.3, y: 2.3 + i * 0.5, w: 4.2, h: 0.45,
            fontSize: 13, color: THEME.primary,
            fontFace: FONTS.body, align: "left", margin: 0,
            paraSpaceAfter: 4,
        });
    });

    // ----------------------------------------------------------
    // SLIDE 6: Color Palettes Showcase
    // ----------------------------------------------------------
    const slide6 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
    slide6.transition = { type: "fade", speed: 0.4 };

    slide6.addText("Color Palettes", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: THEME.primary, bold: true,
        fontFace: FONTS.heading, align: "left", margin: 0,
    });

    const palettes = [
        { name: "Midnight Executive", colors: ["1E2761", "CADCFC", "7B68EE"] },
        { name: "Forest & Moss", colors: ["2C5F2D", "97BC62", "E8E8E8"] },
        { name: "Coral Energy", colors: ["F96167", "F9E795", "2F3C7E"] },
        { name: "Ocean Gradient", colors: ["065A82", "1C7293", "02C39A"] },
        { name: "Charcoal Minimal", colors: ["36454F", "F2F2F2", "E63946"] },
        { name: "Deep Purple", colors: ["4A148C", "CE93D8", "FF6F00"] },
    ];

    palettes.forEach((palette, i) => {
        const row = Math.floor(i / 3);
        const col = i % 3;
        const x = 0.5 + col * 3.2;
        const y = 1.5 + row * 2.0;

        // Palette name
        slide6.addText(palette.name, {
            x: x, y: y, w: 2.8, h: 0.4,
            fontSize: 11, color: "555555", bold: true,
            fontFace: FONTS.body, align: "left", margin: 0,
        });

        // Color swatches
        palette.colors.forEach((color, ci) => {
            slide6.addShape(pres.shapes.ROUNDED_RECTANGLE, {
                x: x + ci * 0.9, y: y + 0.5, w: 0.8, h: 0.8,
                fill: { color: color },
                rectRadius: 0.05,
            });
            slide6.addText(`#${color}`, {
                x: x + ci * 0.9, y: y + 1.4, w: 0.8, h: 0.3,
                fontSize: 7, color: "999999",
                fontFace: FONTS.body, align: "center", margin: 0,
            });
        });
    });

    // ----------------------------------------------------------
    // SLIDE 7: Conclusion (Dark background)
    // ----------------------------------------------------------
    const slide7 = pres.addSlide({ masterName: "TITLE_SLIDE" });
    slide7.transition = { type: "fade", speed: 1.0 };

    slide7.addText("Ready to Create\nStunning Presentations", {
        x: 0.5, y: 1.2, w: 9.0, h: 2.0,
        fontSize: 40, color: THEME.textLight, bold: true,
        fontFace: FONTS.heading, align: "left", margin: 0,
        lineSpacingMultiple: 1.3,
    });

    slide7.addText("github.com/consecrating/PPTx", {
        x: 0.5, y: 3.5, w: 9.0, h: 0.5,
        fontSize: 16, color: THEME.accent,
        fontFace: FONTS.body, align: "left", margin: 0,
    });

    slide7.addText("MIT License | Built with pptxgenjs", {
        x: 0.5, y: 4.8, w: 9.0, h: 0.4,
        fontSize: 11, color: THEME.muted,
        fontFace: FONTS.body, align: "left", margin: 0,
    });

    // Speaker notes on conclusion slide
    slide7.addNotes("Thank the audience. Mention that the repo is open source and contributions are welcome. Point them to the README for setup instructions.");

    // ----------------------------------------------------------
    // SAVE
    // ----------------------------------------------------------
    const outputPath = __dirname + "/demo_output.pptx";
    pres.writeFile({ fileName: outputPath })
        .then(() => {
            console.log(`Presentation created: ${outputPath}`);
            console.log(`Slides: 7`);
            console.log(`Theme: Midnight Executive`);
            console.log(`Layout: 16:9 (10" x 5.625")`);
            console.log(`\nRun validation:`);
            console.log(`  python scripts/office/validate.py ${outputPath}`);
        })
        .catch(err => {
            console.error("Error creating presentation:", err);
            process.exit(1);
        });
}

// Run
createPresentation();
