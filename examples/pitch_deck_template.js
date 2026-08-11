/**
 * Pitch Deck Template
 * 
 * A 10-slide startup pitch deck template following the standard VC pitch structure.
 * Demonstrates advanced layout usage, varied slide types, and professional design.
 * 
 * Slides: Problem → Solution → Market → Product → Business Model → 
 *         Traction → Team → Competition → Financials → Ask
 * 
 * Usage: node examples/pitch_deck_template.js
 */

const pptxgen = require('pptxgenjs');

// Theme: Coral Energy (startup-friendly, bold)
const T = {
    primary: "F96167",
    secondary: "F9E795",
    accent: "2F3C7E",
    textDark: "2F3C7E",
    textLight: "FFFFFF",
    bgDark: "1A2040",
    muted: "8E99A4",
};

const HEADING = "Cambria";
const BODY = "Calibri";

function createPitchDeck() {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";

    // ---- SLIDE 1: Title ----
    const s1 = pres.addSlide();
    s1.background = { color: T.bgDark };
    s1.transition = { type: "fade", speed: 0.8 };
    s1.addText("[Company Name]", {
        x: 0.5, y: 1.8, w: 9.0, h: 1.5,
        fontSize: 48, color: T.textLight, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s1.addText("[One-line description of what you do]", {
        x: 0.5, y: 3.3, w: 9.0, h: 0.6,
        fontSize: 20, color: T.primary, italic: true,
        fontFace: BODY, align: "left", margin: 0,
    });
    s1.addText("[Presenter Name] | [Date]", {
        x: 0.5, y: 4.8, w: 9.0, h: 0.4,
        fontSize: 12, color: T.muted,
        fontFace: BODY, align: "left", margin: 0,
    });

    // ---- SLIDE 2: Problem ----
    const s2 = pres.addSlide();
    s2.transition = { type: "fade", speed: 0.4 };
    s2.addText("The Problem", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: T.accent, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s2.addText("[Describe the pain point your target customers face.\nMake it relatable and quantifiable.]", {
        x: 0.5, y: 1.6, w: 5.5, h: 3.5,
        fontSize: 16, color: "444444",
        fontFace: BODY, align: "left", valign: "top", margin: 0,
    });
    // Placeholder for image/stat
    s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 6.5, y: 1.6, w: 3.0, h: 3.0,
        fill: { color: "F5F5F5" },
        rectRadius: 0.1,
    });
    s2.addText("[Image or\nKey Stat]", {
        x: 6.5, y: 1.6, w: 3.0, h: 3.0,
        fontSize: 14, color: T.muted,
        fontFace: BODY, align: "center", valign: "middle", margin: 0,
    });

    // ---- SLIDE 3: Solution ----
    const s3 = pres.addSlide();
    s3.transition = { type: "fade", speed: 0.4 };
    s3.addText("Our Solution", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: T.accent, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s3.addText("[How your product solves the problem.\nKeep it clear and benefit-focused.]", {
        x: 0.5, y: 1.6, w: 9.0, h: 1.0,
        fontSize: 16, color: "444444",
        fontFace: BODY, align: "left", margin: 0,
    });
    // 3 feature cards
    for (let i = 0; i < 3; i++) {
        const x = 0.5 + i * 3.2;
        s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: x, y: 3.0, w: 2.8, h: 2.0,
            fill: { color: "F8F9FA" },
            rectRadius: 0.08,
            shadow: { type: "outer", color: "000000", blur: 4, offset: 2, opacity: 0.1, angle: 270 },
        });
        s3.addText(`[Feature ${i + 1}]`, {
            x: x, y: 3.2, w: 2.8, h: 0.5,
            fontSize: 14, color: T.accent, bold: true,
            fontFace: BODY, align: "center", margin: 0,
        });
        s3.addText("[Brief description]", {
            x: x + 0.2, y: 3.8, w: 2.4, h: 1.0,
            fontSize: 11, color: "666666",
            fontFace: BODY, align: "center", margin: 0,
        });
    }

    // ---- SLIDE 4: Market Size ----
    const s4 = pres.addSlide();
    s4.transition = { type: "fade", speed: 0.4 };
    s4.addText("Market Opportunity", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: T.accent, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    // TAM / SAM / SOM
    const markets = [
        { label: "TAM", value: "$[X]B", desc: "Total Addressable Market" },
        { label: "SAM", value: "$[X]B", desc: "Serviceable Addressable Market" },
        { label: "SOM", value: "$[X]M", desc: "Serviceable Obtainable Market" },
    ];
    markets.forEach((m, i) => {
        const x = 0.5 + i * 3.2;
        s4.addText(m.value, {
            x: x, y: 1.8, w: 2.8, h: 1.2,
            fontSize: 44, color: T.primary, bold: true,
            fontFace: HEADING, align: "center", valign: "bottom", margin: 0,
        });
        s4.addText(m.label, {
            x: x, y: 3.1, w: 2.8, h: 0.5,
            fontSize: 18, color: T.accent, bold: true,
            fontFace: BODY, align: "center", margin: 0,
        });
        s4.addText(m.desc, {
            x: x, y: 3.6, w: 2.8, h: 0.5,
            fontSize: 11, color: T.muted,
            fontFace: BODY, align: "center", margin: 0,
        });
    });

    // ---- SLIDE 5: Traction ----
    const s5 = pres.addSlide();
    s5.transition = { type: "fade", speed: 0.4 };
    s5.addText("Traction", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: T.accent, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s5.addChart(pres.charts.LINE, [
        {
            name: "MRR ($K)",
            labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            values: [10, 18, 32, 48, 72, 105],
        },
    ], {
        x: 0.5, y: 1.5, w: 9.0, h: 3.8,
        showTitle: false,
        chartColors: [T.primary],
        lineSize: 3,
        lineDataSymbol: "circle",
        lineDataSymbolSize: 8,
        showValue: true,
        dataLabelPosition: "t",
        dataLabelColor: T.accent,
        dataLabelFontSize: 10,
        catAxisLabelColor: "666666",
        valAxisLabelColor: "666666",
        valGridLine: { color: "EEEEEE", size: 0.5 },
        catGridLine: { style: "none" },
        showLegend: false,
    });

    // ---- SLIDE 6: Business Model ----
    const s6 = pres.addSlide();
    s6.transition = { type: "fade", speed: 0.4 };
    s6.addText("Business Model", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: T.accent, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s6.addText([
        { text: "Revenue Streams", options: { fontSize: 18, bold: true, color: T.accent, breakLine: true } },
        { text: "\n", options: { fontSize: 8, breakLine: true } },
        { text: "[Stream 1]: [Description and pricing]", options: { bullet: true, breakLine: true, fontSize: 14 } },
        { text: "[Stream 2]: [Description and pricing]", options: { bullet: true, breakLine: true, fontSize: 14 } },
        { text: "[Stream 3]: [Description and pricing]", options: { bullet: true, fontSize: 14 } },
    ], {
        x: 0.5, y: 1.5, w: 9.0, h: 3.5,
        color: "444444", fontFace: BODY, align: "left", margin: 0,
        paraSpaceAfter: 6,
    });

    // ---- SLIDE 7: Team ----
    const s7 = pres.addSlide();
    s7.transition = { type: "push", speed: 0.5 };
    s7.addText("The Team", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: T.accent, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    const team = [
        { name: "[CEO Name]", role: "CEO & Co-Founder", bg: "[Experience]" },
        { name: "[CTO Name]", role: "CTO & Co-Founder", bg: "[Experience]" },
        { name: "[COO Name]", role: "COO", bg: "[Experience]" },
    ];
    team.forEach((member, i) => {
        const x = 0.5 + i * 3.2;
        // Avatar placeholder
        s7.addShape(pres.shapes.OVAL, {
            x: x + 0.9, y: 1.6, w: 1.0, h: 1.0,
            fill: { color: "E8E8E8" },
        });
        s7.addText(member.name, {
            x: x, y: 2.8, w: 2.8, h: 0.5,
            fontSize: 14, color: T.accent, bold: true,
            fontFace: BODY, align: "center", margin: 0,
        });
        s7.addText(member.role, {
            x: x, y: 3.3, w: 2.8, h: 0.4,
            fontSize: 12, color: T.primary,
            fontFace: BODY, align: "center", margin: 0,
        });
        s7.addText(member.bg, {
            x: x, y: 3.7, w: 2.8, h: 0.5,
            fontSize: 10, color: T.muted,
            fontFace: BODY, align: "center", margin: 0,
        });
    });

    // ---- SLIDE 8: Financials ----
    const s8 = pres.addSlide();
    s8.transition = { type: "fade", speed: 0.4 };
    s8.addText("Financial Projections", {
        x: 0.5, y: 0.5, w: 9.0, h: 0.8,
        fontSize: 36, color: T.accent, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s8.addChart(pres.charts.BAR, [
        {
            name: "Revenue ($M)",
            labels: ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"],
            values: [0.5, 2.1, 6.8, 15.2, 32.0],
        },
    ], {
        x: 0.5, y: 1.5, w: 9.0, h: 3.8,
        showTitle: false,
        chartColors: [T.accent],
        showValue: true,
        dataLabelPosition: "outEnd",
        dataLabelColor: T.accent,
        dataLabelFontSize: 11,
        catAxisLabelColor: "666666",
        valAxisLabelColor: "666666",
        valGridLine: { color: "EEEEEE", size: 0.5 },
        catGridLine: { style: "none" },
        showLegend: false,
    });

    // ---- SLIDE 9: The Ask ----
    const s9 = pres.addSlide();
    s9.background = { color: T.bgDark };
    s9.transition = { type: "fade", speed: 0.8 };
    s9.addText("The Ask", {
        x: 0.5, y: 0.8, w: 9.0, h: 0.8,
        fontSize: 36, color: T.textLight, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s9.addText("$[X]M Series [A/Seed]", {
        x: 0.5, y: 1.8, w: 9.0, h: 1.0,
        fontSize: 48, color: T.primary, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s9.addText([
        { text: "Use of Funds:", options: { bold: true, fontSize: 16, color: T.secondary, breakLine: true } },
        { text: "\n", options: { fontSize: 6, breakLine: true } },
        { text: "[X]% — Product Development", options: { bullet: true, breakLine: true, fontSize: 14, color: T.textLight } },
        { text: "[X]% — Sales & Marketing", options: { bullet: true, breakLine: true, fontSize: 14, color: T.textLight } },
        { text: "[X]% — Team Growth", options: { bullet: true, breakLine: true, fontSize: 14, color: T.textLight } },
        { text: "[X]% — Operations", options: { bullet: true, fontSize: 14, color: T.textLight } },
    ], {
        x: 0.5, y: 3.0, w: 9.0, h: 2.3,
        fontFace: BODY, align: "left", margin: 0,
        paraSpaceAfter: 6,
    });

    // ---- SLIDE 10: Contact ----
    const s10 = pres.addSlide();
    s10.background = { color: T.bgDark };
    s10.transition = { type: "fade", speed: 1.0 };
    s10.addText("Let's Talk", {
        x: 0.5, y: 1.8, w: 9.0, h: 1.2,
        fontSize: 44, color: T.textLight, bold: true,
        fontFace: HEADING, align: "left", margin: 0,
    });
    s10.addText("[name@company.com]\n[+1 (555) 000-0000]\n[company.com]", {
        x: 0.5, y: 3.2, w: 9.0, h: 1.5,
        fontSize: 18, color: T.primary,
        fontFace: BODY, align: "left", margin: 0,
        lineSpacingMultiple: 1.5,
    });

    // Save
    const outputPath = __dirname + "/pitch_deck_output.pptx";
    pres.writeFile({ fileName: outputPath }).then(() => {
        console.log(`Pitch Deck created: ${outputPath}`);
        console.log("10 slides: Title → Problem → Solution → Market → Traction → ");
        console.log("           Business Model → Team → Financials → Ask → Contact");
        console.log(`\nValidate: python scripts/office/validate.py ${outputPath}`);
    });
}

createPitchDeck();
