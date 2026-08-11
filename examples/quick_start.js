/**
 * Quick Start Example
 * 
 * Minimal example showing how to create a professional 3-slide presentation
 * using the advanced skill's recommended patterns.
 * 
 * Usage: node examples/quick_start.js
 */

const pptxgen = require('pptxgenjs');

// Step 1: Choose a palette (see theme_engine.py for all 12)
const palette = {
    primary: "065A82",    // Ocean Gradient theme
    secondary: "1C7293",
    accent: "02C39A",
    dark: "0A1628",
};

// Step 2: Create presentation
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";  // ALWAYS set before adding slides

// Step 3: Title slide
const titleSlide = pres.addSlide();
titleSlide.background = { color: palette.dark };
titleSlide.transition = { type: "fade", speed: 0.8 };
titleSlide.addText("Your Presentation Title", {
    x: 0.5, y: 2.0, w: 9.0, h: 1.2,
    fontSize: 42, color: "FFFFFF", bold: true,
    fontFace: "Cambria", align: "left", margin: 0,
});
titleSlide.addText("Subtitle or tagline goes here", {
    x: 0.5, y: 3.3, w: 9.0, h: 0.6,
    fontSize: 18, color: palette.accent, italic: true,
    fontFace: "Calibri", align: "left", margin: 0,
});

// Step 4: Content slide with bullets
const contentSlide = pres.addSlide();
contentSlide.transition = { type: "fade", speed: 0.4 };
contentSlide.addText("Key Points", {
    x: 0.5, y: 0.5, w: 9.0, h: 0.8,
    fontSize: 36, color: palette.primary, bold: true,
    fontFace: "Cambria", align: "left", margin: 0,
});
contentSlide.addText([
    { text: "First important point with details", options: { bullet: true, breakLine: true } },
    { text: "Second point explaining your value", options: { bullet: true, breakLine: true } },
    { text: "Third point with supporting evidence", options: { bullet: true, breakLine: true } },
    { text: "Final point driving action", options: { bullet: true } },
], {
    x: 0.5, y: 1.5, w: 9.0, h: 3.5,
    fontSize: 16, color: "333333",
    fontFace: "Calibri", align: "left", margin: 0,
    paraSpaceAfter: 8,
});

// Step 5: Closing slide
const closeSlide = pres.addSlide();
closeSlide.background = { color: palette.dark };
closeSlide.transition = { type: "fade", speed: 1.0 };
closeSlide.addText("Thank You", {
    x: 0.5, y: 2.2, w: 9.0, h: 1.0,
    fontSize: 44, color: "FFFFFF", bold: true,
    fontFace: "Cambria", align: "left", margin: 0,
});
closeSlide.addText("your@email.com", {
    x: 0.5, y: 3.5, w: 9.0, h: 0.5,
    fontSize: 16, color: palette.accent,
    fontFace: "Calibri", align: "left", margin: 0,
});
closeSlide.addNotes("Thank the audience and open for questions.");

// Step 6: Save
const outputPath = __dirname + "/quick_start_output.pptx";
pres.writeFile({ fileName: outputPath }).then(() => {
    console.log(`Created: ${outputPath}`);
    console.log("Next: python scripts/office/validate.py " + outputPath);
});
