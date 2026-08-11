const pptxgen = require("pptxgenjs");

// ============================================================
// THEME: Deep Navy Corporate (trust, premium, professional)
// ============================================================
const T = {
  navy: "0F1B3D",
  navyLight: "1A2A5E",
  gold: "C4A35A",
  white: "FFFFFF",
  offWhite: "F5F6FA",
  textDark: "1A1A2E",
  textMuted: "6B7280",
  cardBg: "F8F9FC",
  accent: "2563EB",
  green: "059669",
  greenLight: "D1FAE5",
};

const H = "Cambria";
const B = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";


// ============================================================
// SLIDE 1 — THE CHALLENGE
// ============================================================
const s1 = pres.addSlide();
s1.background = { color: T.navy };
s1.transition = { type: "fade", speed: 0.8 };

// Top label
s1.addText("THE CHALLENGE", {
  x: 0.5, y: 0.4, w: 9.0, h: 0.4,
  fontSize: 11, color: T.gold, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 2,
});

// Main title
s1.addText("Creating an Alternative Digital\nPresence for Payment Gateway\nEvaluation", {
  x: 0.5, y: 0.9, w: 9.0, h: 1.6,
  fontSize: 32, color: T.white, bold: true,
  fontFace: H, align: "left", margin: 0,
  lineSpacingMultiple: 1.2,
});

// Context paragraph 1
s1.addText("Casino Pride operates as an offshore casino in Goa. However, casino and gaming-related businesses fall under restricted/high-risk categories for many banks and payment gateway providers, which can result in additional scrutiny or rejection during merchant onboarding.", {
  x: 0.5, y: 2.7, w: 5.8, h: 1.0,
  fontSize: 11, color: T.offWhite,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.4,
});


// Context paragraph 2
s1.addText("The current website, www.cpofficial.in, is primarily positioned around the Casino Pride casino/gaming business. This may be one of the factors considered during the payment gateway's website and merchant assessment.", {
  x: 0.5, y: 3.75, w: 5.8, h: 0.85,
  fontSize: 11, color: T.offWhite,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.4,
});

// Proposed Approach Box
s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.7, y: 2.7, w: 2.9, h: 2.6,
  fill: { color: T.navyLight },
  rectRadius: 0.08,
});

s1.addText("Proposed Approach", {
  x: 6.9, y: 2.85, w: 2.5, h: 0.4,
  fontSize: 12, color: T.gold, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s1.addText("Create a separate, professionally designed website with a broader focus on:", {
  x: 6.9, y: 3.3, w: 2.5, h: 0.6,
  fontSize: 10, color: T.offWhite,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});

s1.addText("Entertainment\nNightlife\nHospitality\nRecreation\nEvents", {
  x: 6.9, y: 3.95, w: 2.5, h: 1.1,
  fontSize: 11, color: T.gold, bold: true,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.5,
});


// Disclaimer footer
s1.addText("Note: Website development cannot guarantee payment gateway approval. Final approval depends on the bank/payment gateway's KYC, merchant category, business activity, compliance and risk policies.", {
  x: 0.5, y: 5.0, w: 9.0, h: 0.4,
  fontSize: 8, color: T.textMuted, italic: true,
  fontFace: B, align: "left", margin: 0,
});

// ============================================================
// SLIDE 2 — TWO POSSIBLE SOLUTIONS
// ============================================================
const s2 = pres.addSlide();
s2.background = { color: T.white };
s2.transition = { type: "fade", speed: 0.4 };

// Section label
s2.addText("TWO POSSIBLE SOLUTIONS", {
  x: 0.5, y: 0.4, w: 9.0, h: 0.35,
  fontSize: 10, color: T.gold, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 2,
});

// OPTION 1 Header
s2.addText("OPTION 1", {
  x: 0.5, y: 0.85, w: 4.3, h: 0.35,
  fontSize: 10, color: T.accent, bold: true,
  fontFace: B, align: "left", margin: 0,
});

s2.addText("CP Goa", {
  x: 0.5, y: 1.15, w: 4.3, h: 0.5,
  fontSize: 24, color: T.navy, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s2.addText("Entertainment & Nightlife Website", {
  x: 0.5, y: 1.6, w: 4.3, h: 0.35,
  fontSize: 12, color: T.textMuted, italic: true,
  fontFace: B, align: "left", margin: 0,
});


// Option 1 description
s2.addText("Create a dedicated sub-domain website positioned around CP Goa rather than making the website primarily casino-focused.", {
  x: 0.5, y: 2.05, w: 4.3, h: 0.55,
  fontSize: 11, color: T.textDark,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});

// Option 1 - Website Focus
s2.addText("Website Focus", {
  x: 0.5, y: 2.7, w: 4.3, h: 0.3,
  fontSize: 10, color: T.navy, bold: true,
  fontFace: B, align: "left", margin: 0,
});

s2.addText([
  { text: "Entertainment", options: { bullet: true, breakLine: true } },
  { text: "Nightlife", options: { bullet: true, breakLine: true } },
  { text: "Live events & performances", options: { bullet: true, breakLine: true } },
  { text: "Dining experiences", options: { bullet: true, breakLine: true } },
  { text: "Premium ambience & recreation", options: { bullet: true, breakLine: true } },
  { text: "Guest experiences & event enquiries", options: { bullet: true } },
], {
  x: 0.5, y: 3.0, w: 4.3, h: 1.4,
  fontSize: 10, color: T.textDark,
  fontFace: B, align: "left", margin: 0,
  paraSpaceAfter: 3,
});

// Option 1 - Pages card
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.45, w: 4.3, h: 0.7,
  fill: { color: T.cardBg },
  rectRadius: 0.05,
});

s2.addText("Proposed 5 Pages:  Home  |  About CP Goa  |  Entertainment & Nightlife  |  Experiences & Events  |  Contact", {
  x: 0.7, y: 4.5, w: 4.0, h: 0.6,
  fontSize: 9, color: T.navy,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});


// Vertical divider
s2.addShape(pres.shapes.RECTANGLE, {
  x: 4.95, y: 0.85, w: 0.02, h: 4.4,
  fill: { color: "E5E7EB" },
});

// OPTION 2 Header
s2.addText("OPTION 2", {
  x: 5.2, y: 0.85, w: 4.3, h: 0.35,
  fontSize: 10, color: T.green, bold: true,
  fontFace: B, align: "left", margin: 0,
});

s2.addText("Goa Coastal Resorts &\nRecreation Pvt. Ltd.", {
  x: 5.2, y: 1.15, w: 4.3, h: 0.65,
  fontSize: 20, color: T.navy, bold: true,
  fontFace: H, align: "left", margin: 0,
  lineSpacingMultiple: 1.1,
});

s2.addText("Corporate Business Website", {
  x: 5.2, y: 1.8, w: 4.3, h: 0.3,
  fontSize: 12, color: T.textMuted, italic: true,
  fontFace: B, align: "left", margin: 0,
});

// Option 2 description
s2.addText("Develop a new 5-page website under the domain of the company associated with Casino Pride. The website presents the company's broader business identity.", {
  x: 5.2, y: 2.15, w: 4.3, h: 0.6,
  fontSize: 11, color: T.textDark,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});

// Option 2 - Website Focus
s2.addText("Website Focus", {
  x: 5.2, y: 2.8, w: 4.3, h: 0.3,
  fontSize: 10, color: T.navy, bold: true,
  fontFace: B, align: "left", margin: 0,
});

s2.addText([
  { text: "Hospitality", options: { bullet: true, breakLine: true } },
  { text: "Entertainment", options: { bullet: true, breakLine: true } },
  { text: "Recreation", options: { bullet: true, breakLine: true } },
  { text: "Dining", options: { bullet: true, breakLine: true } },
  { text: "Events", options: { bullet: true } },
], {
  x: 5.2, y: 3.1, w: 4.3, h: 1.2,
  fontSize: 10, color: T.textDark,
  fontFace: B, align: "left", margin: 0,
  paraSpaceAfter: 3,
});


// Option 2 - Pages card
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 5.2, y: 4.35, w: 4.3, h: 0.55,
  fill: { color: T.cardBg },
  rectRadius: 0.05,
});

s2.addText("Proposed 5 Pages:  Home  |  About the Company  |  Our Businesses  |  Brands  |  Contact", {
  x: 5.4, y: 4.4, w: 3.9, h: 0.5,
  fontSize: 9, color: T.navy,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});

// Key advantage note at bottom
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 5.2, y: 4.95, w: 4.3, h: 0.4,
  fill: { color: T.greenLight },
  rectRadius: 0.04,
});
s2.addText("Key: Creates stronger corporate identity for current & future businesses", {
  x: 5.4, y: 4.98, w: 4.0, h: 0.35,
  fontSize: 8.5, color: T.green, bold: true,
  fontFace: B, align: "left", margin: 0,
});

// ============================================================
// SLIDE 3 — OPTION 2 DETAILED
// ============================================================
const s3 = pres.addSlide();
s3.background = { color: T.white };
s3.transition = { type: "fade", speed: 0.4 };

// Section label
s3.addText("OPTION 2 — RECOMMENDED", {
  x: 0.5, y: 0.4, w: 9.0, h: 0.35,
  fontSize: 10, color: T.green, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 2,
});


s3.addText("Goa Coastal Resorts & Recreation Pvt. Ltd.", {
  x: 0.5, y: 0.85, w: 9.0, h: 0.55,
  fontSize: 26, color: T.navy, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s3.addText("Corporate Business Website — Broader Business Identity", {
  x: 0.5, y: 1.4, w: 9.0, h: 0.35,
  fontSize: 13, color: T.textMuted, italic: true,
  fontFace: B, align: "left", margin: 0,
});

// Description
s3.addText("Develop a new 5-page website under the domain of Goa Coastal Resorts & Recreation Pvt. Ltd., the company associated with Casino Pride. Instead of positioning the entire website around casino/gaming, the website will present the company's broader business identity and legitimate activities.", {
  x: 0.5, y: 1.9, w: 9.0, h: 0.75,
  fontSize: 11, color: T.textDark,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.4,
});

// Website Focus pills
const focusItems = ["Hospitality", "Entertainment", "Recreation", "Dining", "Events"];
focusItems.forEach((item, i) => {
  const pillX = 0.5 + i * 1.85;
  s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: pillX, y: 2.8, w: 1.7, h: 0.4,
    fill: { color: T.navy },
    rectRadius: 0.2,
  });
  s3.addText(item, {
    x: pillX, y: 2.8, w: 1.7, h: 0.4,
    fontSize: 10, color: T.white, bold: true,
    fontFace: B, align: "center", valign: "middle", margin: 0,
  });
});

// Pages structure card
s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.4, w: 9.0, h: 0.6,
  fill: { color: T.cardBg },
  rectRadius: 0.06,
});

s3.addText("Proposed 5 Pages:   Home   |   About the Company   |   Our Businesses & Experiences   |   Brands / Business Interests   |   Contact", {
  x: 0.7, y: 3.45, w: 8.6, h: 0.5,
  fontSize: 11, color: T.navy, bold: true,
  fontFace: B, align: "left", valign: "middle", margin: 0,
});


// Casino Pride positioning note
s3.addText("Casino Pride can be presented accurately as one of the company's brands/business interests where appropriate.", {
  x: 0.5, y: 4.15, w: 9.0, h: 0.35,
  fontSize: 11, color: T.textDark, italic: true,
  fontFace: B, align: "left", margin: 0,
});

// Key Advantage box
s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.6, w: 9.0, h: 0.8,
  fill: { color: T.greenLight },
  rectRadius: 0.06,
});

s3.addText("KEY ADVANTAGE", {
  x: 0.7, y: 4.65, w: 8.6, h: 0.25,
  fontSize: 9, color: T.green, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 1.5,
});

s3.addText("This creates a stronger corporate digital identity that can accommodate the company's existing and future business activities, rather than creating a website solely around one entertainment brand.", {
  x: 0.7, y: 4.9, w: 8.6, h: 0.45,
  fontSize: 10.5, color: T.green,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});

// ============================================================
// SLIDE 4 — OUR RECOMMENDATION
// ============================================================
const s4 = pres.addSlide();
s4.background = { color: T.navy };
s4.transition = { type: "fade", speed: 0.5 };

// Section label
s4.addText("OUR RECOMMENDATION", {
  x: 0.5, y: 0.4, w: 9.0, h: 0.35,
  fontSize: 10, color: T.gold, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 2,
});

s4.addText("OPTION 2 — Corporate Website", {
  x: 0.5, y: 0.85, w: 9.0, h: 0.55,
  fontSize: 28, color: T.white, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s4.addText("Goa Coastal Resorts & Recreation Pvt. Ltd.", {
  x: 0.5, y: 1.4, w: 9.0, h: 0.35,
  fontSize: 14, color: T.gold, italic: true,
  fontFace: B, align: "left", margin: 0,
});


// Why cards - 4 reasons
const reasons = [
  { title: "More Credible", desc: "Represents the actual corporate/business entity" },
  { title: "More Flexible", desc: "Showcases hospitality, entertainment, recreation, dining & other activities" },
  { title: "More Scalable", desc: "Additional brands, properties or business interests can be added in future" },
  { title: "Better Long-Term Value", desc: "Remains useful as a corporate digital asset beyond gateway requirement" },
];

reasons.forEach((reason, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.7;
  const y = 2.0 + row * 1.3;

  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.4, h: 1.1,
    fill: { color: T.navyLight },
    rectRadius: 0.06,
  });

  s4.addText(reason.title, {
    x: x + 0.2, y: y + 0.1, w: 4.0, h: 0.4,
    fontSize: 13, color: T.gold, bold: true,
    fontFace: H, align: "left", margin: 0,
  });

  s4.addText(reason.desc, {
    x: x + 0.2, y: y + 0.5, w: 4.0, h: 0.5,
    fontSize: 10.5, color: T.offWhite,
    fontFace: B, align: "left", margin: 0,
    lineSpacingMultiple: 1.3,
  });
});

// Our Approach section
s4.addText("Our Approach", {
  x: 0.5, y: 4.75, w: 9.0, h: 0.3,
  fontSize: 11, color: T.gold, bold: true,
  fontFace: B, align: "left", margin: 0,
});

s4.addText("The website will be designed as a modern, premium corporate website, with clear business positioning and accurate information rather than simply modifying or hiding the existing Casino Pride identity.", {
  x: 0.5, y: 5.05, w: 9.0, h: 0.4,
  fontSize: 10, color: T.offWhite,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});


// ============================================================
// SLIDE 5 — WEBSITE & COMMERCIAL
// ============================================================
const s5 = pres.addSlide();
s5.background = { color: T.white };
s5.transition = { type: "fade", speed: 0.4 };

// Section label
s5.addText("WEBSITE & COMMERCIAL", {
  x: 0.5, y: 0.4, w: 9.0, h: 0.35,
  fontSize: 10, color: T.navy, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 2,
});

s5.addText("5-Page Minimal Professional Website", {
  x: 0.5, y: 0.8, w: 5.5, h: 0.5,
  fontSize: 22, color: T.navy, bold: true,
  fontFace: H, align: "left", margin: 0,
});

// Included features - left column
s5.addText("Included", {
  x: 0.5, y: 1.45, w: 5.0, h: 0.3,
  fontSize: 11, color: T.navy, bold: true,
  fontFace: B, align: "left", margin: 0,
});

s5.addText([
  { text: "Modern UI/UX design", options: { bullet: true, breakLine: true } },
  { text: "5-page responsive website", options: { bullet: true, breakLine: true } },
  { text: "Desktop & mobile optimisation", options: { bullet: true, breakLine: true } },
  { text: "Professional content structuring", options: { bullet: true, breakLine: true } },
  { text: "Image integration", options: { bullet: true, breakLine: true } },
  { text: "Contact/enquiry form", options: { bullet: true, breakLine: true } },
  { text: "Google Maps integration", options: { bullet: true, breakLine: true } },
  { text: "Social media integration", options: { bullet: true, breakLine: true } },
  { text: "Testing & deployment", options: { bullet: true } },
], {
  x: 0.5, y: 1.75, w: 5.0, h: 2.6,
  fontSize: 11, color: T.textDark,
  fontFace: B, align: "left", margin: 0,
  paraSpaceAfter: 4,
});


// Commercial card - right side
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.0, y: 1.35, w: 3.6, h: 3.4,
  fill: { color: T.navy },
  rectRadius: 0.1,
});

s5.addText("Commercial", {
  x: 6.2, y: 1.5, w: 3.2, h: 0.35,
  fontSize: 12, color: T.gold, bold: true,
  fontFace: B, align: "center", margin: 0,
});

// Strikethrough price
s5.addText("Standard: \u20B921,000", {
  x: 6.2, y: 1.95, w: 3.2, h: 0.3,
  fontSize: 11, color: T.textMuted,
  fontFace: B, align: "center", margin: 0,
  strike: true,
});

// Special price label
s5.addText("Special Discounted Price", {
  x: 6.2, y: 2.3, w: 3.2, h: 0.25,
  fontSize: 9, color: T.offWhite,
  fontFace: B, align: "center", margin: 0,
});

// Big price
s5.addText("\u20B916,000", {
  x: 6.2, y: 2.55, w: 3.2, h: 0.6,
  fontSize: 36, color: T.white, bold: true,
  fontFace: H, align: "center", margin: 0,
});

s5.addText("+ GST", {
  x: 6.2, y: 3.1, w: 3.2, h: 0.25,
  fontSize: 11, color: T.offWhite,
  fontFace: B, align: "center", margin: 0,
});

// Breakdown
s5.addText("GST @ 18%: \u20B92,880", {
  x: 6.2, y: 3.5, w: 3.2, h: 0.25,
  fontSize: 10, color: T.offWhite,
  fontFace: B, align: "center", margin: 0,
});

// Total
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.4, y: 3.85, w: 2.8, h: 0.45,
  fill: { color: T.gold },
  rectRadius: 0.04,
});

s5.addText("Total: \u20B918,880/-", {
  x: 6.4, y: 3.85, w: 2.8, h: 0.45,
  fontSize: 14, color: T.navy, bold: true,
  fontFace: H, align: "center", valign: "middle", margin: 0,
});

// Saving note
s5.addText("Client Saving: \u20B95,000 discount", {
  x: 6.2, y: 4.4, w: 3.2, h: 0.25,
  fontSize: 10, color: T.green, bold: true,
  fontFace: B, align: "center", margin: 0,
});


// NEXT STEP section at bottom
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.65, w: 9.0, h: 0.8,
  fill: { color: T.cardBg },
  rectRadius: 0.06,
});

s5.addText("NEXT STEP", {
  x: 0.7, y: 4.7, w: 8.6, h: 0.25,
  fontSize: 9, color: T.navy, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 1.5,
});

s5.addText("Select preferred approach  \u2192  Finalise website structure  \u2192  Design  \u2192  Development  \u2192  Testing  \u2192  Deployment", {
  x: 0.7, y: 4.95, w: 8.6, h: 0.25,
  fontSize: 10, color: T.navy, bold: true,
  fontFace: B, align: "left", margin: 0,
});

s5.addText("The website will be developed as a professional, accurate and scalable digital presence that can subsequently be submitted to the concerned bank/payment provider for its independent review. Final payment gateway approval remains subject to the bank/payment provider's policies and compliance assessment.", {
  x: 0.7, y: 5.2, w: 8.6, h: 0.35,
  fontSize: 8, color: T.textMuted, italic: true,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});

// ============================================================
// SAVE FILE
// ============================================================
const outputPath = __dirname + "/Casino_Pride_Digital_Strategy.pptx";
pres.writeFile({ fileName: outputPath })
  .then(() => {
    console.log("Presentation created successfully!");
    console.log("File: " + outputPath);
    console.log("Slides: 5");
    console.log("Theme: Deep Navy Corporate");
  })
  .catch(err => {
    console.error("Error:", err);
    process.exit(1);
  });
