const pptxgen = require("pptxgenjs");

// ============================================================
// MODERN PREMIUM THEME — Dark + Gold + Clean White
// ============================================================
const T = {
  // Primary
  dark: "0A0E1A",
  darkCard: "141B2D",
  darkSurface: "1C2438",
  // Accents
  gold: "D4A853",
  goldLight: "F0D98D",
  goldMuted: "8B7535",
  blue: "3B82F6",
  blueLight: "60A5FA",
  emerald: "10B981",
  emeraldBg: "064E3B",
  // Neutrals
  white: "FFFFFF",
  offWhite: "F8FAFC",
  gray50: "F9FAFB",
  gray100: "F3F4F6",
  gray200: "E5E7EB",
  gray400: "9CA3AF",
  gray500: "6B7280",
  gray600: "4B5563",
  gray700: "374151",
  gray800: "1F2937",
  gray900: "111827",
};

const H = "Cambria";
const B = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10" x 5.625"


// ============================================================
// SLIDE 1 — THE CHALLENGE (Full dark immersive)
// ============================================================
const s1 = pres.addSlide();
s1.background = { color: T.dark };
s1.transition = { type: "fade", speed: 0.8 };

// Subtle top-left geometric accent (large faded circle)
s1.addShape(pres.shapes.OVAL, {
  x: -1.5, y: -1.5, w: 4.5, h: 4.5,
  fill: { color: T.darkCard },
});

// Gold chip label
s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.6, y: 0.5, w: 1.6, h: 0.32,
  fill: { color: T.goldMuted },
  rectRadius: 0.16,
});
s1.addText("THE CHALLENGE", {
  x: 0.6, y: 0.5, w: 1.6, h: 0.32,
  fontSize: 8, color: T.white, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
  charSpacing: 1.5,
});

// Main headline
s1.addText("Creating an Alternative\nDigital Presence for\nPayment Gateway Evaluation", {
  x: 0.6, y: 1.1, w: 5.5, h: 1.8,
  fontSize: 30, color: T.white, bold: true,
  fontFace: H, align: "left", margin: 0,
  lineSpacingMultiple: 1.15,
});

// Two context paragraphs
s1.addText("Casino Pride operates as an offshore casino in Goa. Casino and gaming-related businesses fall under restricted/high-risk categories for many banks and payment gateway providers, resulting in additional scrutiny or rejection during merchant onboarding.", {
  x: 0.6, y: 3.1, w: 5.2, h: 0.85,
  fontSize: 10.5, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.5,
});

s1.addText("The current website (cpofficial.in) is primarily positioned around the Casino Pride casino/gaming business \u2014 a factor considered during payment gateway merchant assessment.", {
  x: 0.6, y: 4.05, w: 5.2, h: 0.7,
  fontSize: 10.5, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.5,
});


// Right side — Proposed Approach Card (floating glass card)
s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.2, y: 0.8, w: 3.4, h: 4.2,
  fill: { color: T.darkCard },
  rectRadius: 0.15,
  shadow: { type: "outer", color: "000000", blur: 12, offset: 4, opacity: 0.3, angle: 270 },
});

// Card gold top accent line
s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.5, y: 1.05, w: 2.0, h: 0.04,
  fill: { color: T.gold },
  rectRadius: 0.02,
});

s1.addText("Proposed Approach", {
  x: 6.5, y: 1.25, w: 2.8, h: 0.4,
  fontSize: 14, color: T.white, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s1.addText("Create a separate, professionally designed website with a broader focus:", {
  x: 6.5, y: 1.7, w: 2.8, h: 0.55,
  fontSize: 10, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.4,
});

// Focus items with gold bullets
const focusAreas = ["Entertainment", "Nightlife", "Hospitality", "Recreation", "Events"];
focusAreas.forEach((item, i) => {
  const y = 2.4 + i * 0.42;
  // Gold dot
  s1.addShape(pres.shapes.OVAL, {
    x: 6.5, y: y + 0.08, w: 0.12, h: 0.12,
    fill: { color: T.gold },
  });
  s1.addText(item, {
    x: 6.8, y: y, w: 2.5, h: 0.3,
    fontSize: 12, color: T.white, bold: true,
    fontFace: B, align: "left", margin: 0,
  });
});

// Disclaimer at very bottom
s1.addText("Note: Website development cannot guarantee payment gateway approval. Final approval depends on the bank/payment gateway\u2019s KYC, compliance and risk policies.", {
  x: 0.6, y: 5.15, w: 8.8, h: 0.3,
  fontSize: 7.5, color: T.gray600,
  fontFace: B, align: "left", margin: 0, italic: true,
});


// ============================================================
// SLIDE 2 — TWO POSSIBLE SOLUTIONS (Modern split cards)
// ============================================================
const s2 = pres.addSlide();
s2.background = { color: T.offWhite };
s2.transition = { type: "fade", speed: 0.4 };

// Section label
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 0.4, w: 2.4, h: 0.3,
  fill: { color: T.dark },
  rectRadius: 0.15,
});
s2.addText("TWO POSSIBLE SOLUTIONS", {
  x: 0.5, y: 0.4, w: 2.4, h: 0.3,
  fontSize: 7.5, color: T.white, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
  charSpacing: 1.2,
});

// ---- OPTION 1 CARD ----
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 0.95, w: 4.4, h: 4.4,
  fill: { color: T.white },
  rectRadius: 0.12,
  shadow: { type: "outer", color: "000000", blur: 8, offset: 2, opacity: 0.08, angle: 270 },
});

// Option 1 badge
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 1.2, w: 1.1, h: 0.28,
  fill: { color: T.blue },
  rectRadius: 0.14,
});
s2.addText("OPTION 1", {
  x: 0.8, y: 1.2, w: 1.1, h: 0.28,
  fontSize: 7.5, color: T.white, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
});

s2.addText("CP Goa", {
  x: 0.8, y: 1.65, w: 3.8, h: 0.45,
  fontSize: 22, color: T.gray900, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s2.addText("Entertainment & Nightlife Website", {
  x: 0.8, y: 2.05, w: 3.8, h: 0.3,
  fontSize: 10, color: T.gray500,
  fontFace: B, align: "left", margin: 0,
});

// Thin separator
s2.addShape(pres.shapes.RECTANGLE, {
  x: 0.8, y: 2.45, w: 3.8, h: 0.015,
  fill: { color: T.gray200 },
});

// Option 1 focus items
const opt1Items = ["Entertainment & nightlife", "Live events & performances", "Dining experiences", "Premium ambience", "Guest experiences & enquiries"];
opt1Items.forEach((item, i) => {
  const y = 2.6 + i * 0.36;
  s2.addShape(pres.shapes.OVAL, {
    x: 0.8, y: y + 0.1, w: 0.09, h: 0.09,
    fill: { color: T.blue },
  });
  s2.addText(item, {
    x: 1.05, y: y, w: 3.5, h: 0.28,
    fontSize: 10, color: T.gray700,
    fontFace: B, align: "left", margin: 0,
  });
});

// Option 1 pages
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 4.55, w: 3.8, h: 0.6,
  fill: { color: T.gray50 },
  rectRadius: 0.06,
});
s2.addText("5 Pages", {
  x: 0.95, y: 4.58, w: 3.5, h: 0.2,
  fontSize: 7.5, color: T.gray500, bold: true,
  fontFace: B, align: "left", margin: 0,
});
s2.addText("Home \u2022 About CP Goa \u2022 Entertainment \u2022 Experiences \u2022 Contact", {
  x: 0.95, y: 4.8, w: 3.5, h: 0.3,
  fontSize: 8.5, color: T.gray700,
  fontFace: B, align: "left", margin: 0,
});


// ---- OPTION 2 CARD (recommended — highlighted) ----
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 5.1, y: 0.95, w: 4.4, h: 4.4,
  fill: { color: T.white },
  rectRadius: 0.12,
  shadow: { type: "outer", color: "000000", blur: 10, offset: 3, opacity: 0.1, angle: 270 },
  line: { color: T.emerald, width: 1.5 },
});

// Option 2 badge
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 5.4, y: 1.2, w: 1.7, h: 0.28,
  fill: { color: T.emerald },
  rectRadius: 0.14,
});
s2.addText("OPTION 2 \u2014 RECOMMENDED", {
  x: 5.4, y: 1.2, w: 1.7, h: 0.28,
  fontSize: 6.5, color: T.white, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
});

s2.addText("Goa Coastal Resorts\n& Recreation Pvt. Ltd.", {
  x: 5.4, y: 1.6, w: 3.8, h: 0.7,
  fontSize: 18, color: T.gray900, bold: true,
  fontFace: H, align: "left", margin: 0,
  lineSpacingMultiple: 1.1,
});

s2.addText("Corporate Business Website", {
  x: 5.4, y: 2.3, w: 3.8, h: 0.3,
  fontSize: 10, color: T.gray500,
  fontFace: B, align: "left", margin: 0,
});

// Thin separator
s2.addShape(pres.shapes.RECTANGLE, {
  x: 5.4, y: 2.65, w: 3.8, h: 0.015,
  fill: { color: T.gray200 },
});

// Option 2 focus items
const opt2Items = ["Hospitality & recreation", "Entertainment & events", "Dining & lifestyle", "Corporate identity", "Multi-brand scalability"];
opt2Items.forEach((item, i) => {
  const y = 2.8 + i * 0.36;
  s2.addShape(pres.shapes.OVAL, {
    x: 5.4, y: y + 0.1, w: 0.09, h: 0.09,
    fill: { color: T.emerald },
  });
  s2.addText(item, {
    x: 5.65, y: y, w: 3.5, h: 0.28,
    fontSize: 10, color: T.gray700,
    fontFace: B, align: "left", margin: 0,
  });
});

// Option 2 pages
s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 5.4, y: 4.55, w: 3.8, h: 0.6,
  fill: { color: "ECFDF5" },
  rectRadius: 0.06,
});
s2.addText("5 Pages", {
  x: 5.55, y: 4.58, w: 3.5, h: 0.2,
  fontSize: 7.5, color: T.emerald, bold: true,
  fontFace: B, align: "left", margin: 0,
});
s2.addText("Home \u2022 About \u2022 Businesses & Experiences \u2022 Brands \u2022 Contact", {
  x: 5.55, y: 4.8, w: 3.5, h: 0.3,
  fontSize: 8.5, color: T.gray700,
  fontFace: B, align: "left", margin: 0,
});


// ============================================================
// SLIDE 3 — OPTION 2 DEEP DIVE (Modern light layout)
// ============================================================
const s3 = pres.addSlide();
s3.background = { color: T.white };
s3.transition = { type: "fade", speed: 0.4 };

// Full-width dark header band
s3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1.8,
  fill: { color: T.dark },
});

// Badge
s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 0.35, w: 1.6, h: 0.26,
  fill: { color: T.emerald },
  rectRadius: 0.13,
});
s3.addText("RECOMMENDED", {
  x: 0.5, y: 0.35, w: 1.6, h: 0.26,
  fontSize: 7, color: T.white, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
});

s3.addText("Goa Coastal Resorts & Recreation Pvt. Ltd.", {
  x: 0.5, y: 0.75, w: 9.0, h: 0.5,
  fontSize: 24, color: T.white, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s3.addText("Corporate Business Website \u2014 Representing the Company\u2019s Broader Identity", {
  x: 0.5, y: 1.3, w: 9.0, h: 0.3,
  fontSize: 11, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
});

// Description paragraph
s3.addText("Develop a new 5-page website under the domain of Goa Coastal Resorts & Recreation Pvt. Ltd., the company associated with Casino Pride. Instead of positioning the website around casino/gaming, it presents the company\u2019s broader business identity and legitimate activities.", {
  x: 0.5, y: 2.1, w: 9.0, h: 0.65,
  fontSize: 11, color: T.gray700,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.5,
});

// Focus area pills row
const pills = ["Hospitality", "Entertainment", "Recreation", "Dining", "Events"];
pills.forEach((pill, i) => {
  const px = 0.5 + i * 1.9;
  s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: px, y: 2.95, w: 1.75, h: 0.38,
    fill: { color: T.dark },
    rectRadius: 0.19,
  });
  s3.addText(pill, {
    x: px, y: 2.95, w: 1.75, h: 0.38,
    fontSize: 9.5, color: T.white, bold: true,
    fontFace: B, align: "center", valign: "middle", margin: 0,
  });
});

// Pages structure card
s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.55, w: 9.0, h: 0.55,
  fill: { color: T.gray50 },
  rectRadius: 0.08,
});
s3.addText("PROPOSED PAGES:   Home   \u2022   About the Company   \u2022   Our Businesses & Experiences   \u2022   Brands / Business Interests   \u2022   Contact", {
  x: 0.7, y: 3.55, w: 8.6, h: 0.55,
  fontSize: 10, color: T.gray800, bold: true,
  fontFace: B, align: "left", valign: "middle", margin: 0,
});

// Casino Pride note
s3.addText("Casino Pride can be presented accurately as one of the company\u2019s brands/business interests where appropriate.", {
  x: 0.5, y: 4.25, w: 9.0, h: 0.3,
  fontSize: 10.5, color: T.gray600, italic: true,
  fontFace: B, align: "left", margin: 0,
});

// Key Advantage box
s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.7, w: 9.0, h: 0.75,
  fill: { color: "ECFDF5" },
  rectRadius: 0.08,
  line: { color: T.emerald, width: 0.75 },
});

s3.addShape(pres.shapes.OVAL, {
  x: 0.75, y: 4.88, w: 0.35, h: 0.35,
  fill: { color: T.emerald },
});
s3.addText("\u2713", {
  x: 0.75, y: 4.88, w: 0.35, h: 0.35,
  fontSize: 14, color: T.white, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
});

s3.addText("KEY ADVANTAGE", {
  x: 1.25, y: 4.78, w: 8.0, h: 0.2,
  fontSize: 8, color: T.emerald, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 1,
});
s3.addText("Creates a stronger corporate digital identity that accommodates existing and future business activities, rather than a website solely around one entertainment brand.", {
  x: 1.25, y: 4.98, w: 8.0, h: 0.4,
  fontSize: 10, color: T.emeraldBg,
  fontFace: B, align: "left", margin: 0,
  lineSpacingMultiple: 1.3,
});


// ============================================================
// SLIDE 4 — OUR RECOMMENDATION (Dark premium)
// ============================================================
const s4 = pres.addSlide();
s4.background = { color: T.dark };
s4.transition = { type: "fade", speed: 0.5 };

// Subtle background shape
s4.addShape(pres.shapes.OVAL, {
  x: 7.0, y: -2.0, w: 6.0, h: 6.0,
  fill: { color: T.darkCard },
});

// Badge
s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 0.4, w: 2.2, h: 0.28,
  fill: { color: T.gold },
  rectRadius: 0.14,
});
s4.addText("OUR RECOMMENDATION", {
  x: 0.5, y: 0.4, w: 2.2, h: 0.28,
  fontSize: 7.5, color: T.dark, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
  charSpacing: 1,
});

s4.addText("Option 2 \u2014 Corporate Website", {
  x: 0.5, y: 0.9, w: 9.0, h: 0.55,
  fontSize: 26, color: T.white, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s4.addText("Goa Coastal Resorts & Recreation Pvt. Ltd.", {
  x: 0.5, y: 1.45, w: 9.0, h: 0.3,
  fontSize: 12, color: T.gold, italic: true,
  fontFace: B, align: "left", margin: 0,
});

// 4 reason cards in 2x2 grid
const reasons = [
  { icon: "\u2726", title: "More Credible", desc: "Represents the actual corporate/business entity" },
  { icon: "\u2194", title: "More Flexible", desc: "Showcases hospitality, entertainment, recreation, dining & other activities" },
  { icon: "\u2197", title: "More Scalable", desc: "Additional brands, properties or business interests can be added in future" },
  { icon: "\u2605", title: "Better Long-Term Value", desc: "Remains useful as a corporate digital asset beyond gateway requirement" },
];

reasons.forEach((r, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.7;
  const y = 2.0 + row * 1.4;

  // Card
  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.4, h: 1.2,
    fill: { color: T.darkSurface },
    rectRadius: 0.1,
  });

  // Gold left accent
  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.15, y: y + 0.25, w: 0.04, h: 0.7,
    fill: { color: T.gold },
    rectRadius: 0.02,
  });

  // Title
  s4.addText(r.title, {
    x: x + 0.4, y: y + 0.15, w: 3.8, h: 0.4,
    fontSize: 13, color: T.white, bold: true,
    fontFace: H, align: "left", margin: 0,
  });

  // Description
  s4.addText(r.desc, {
    x: x + 0.4, y: y + 0.55, w: 3.8, h: 0.5,
    fontSize: 10, color: T.gray400,
    fontFace: B, align: "left", margin: 0,
    lineSpacingMultiple: 1.3,
  });
});

// Our Approach section
s4.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 4.95, w: 3.0, h: 0.02,
  fill: { color: T.gold },
});

s4.addText("OUR APPROACH", {
  x: 0.5, y: 5.05, w: 9.0, h: 0.2,
  fontSize: 8, color: T.gold, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 1.5,
});

s4.addText("A modern, premium corporate website with clear business positioning and accurate information \u2014 not simply modifying or hiding the existing Casino Pride identity.", {
  x: 0.5, y: 5.25, w: 9.0, h: 0.3,
  fontSize: 10, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
});


// ============================================================
// SLIDE 5 — WEBSITE & COMMERCIAL (Clean white + dark card)
// ============================================================
const s5 = pres.addSlide();
s5.background = { color: T.offWhite };
s5.transition = { type: "fade", speed: 0.4 };

// Section badge
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 0.4, w: 2.2, h: 0.28,
  fill: { color: T.dark },
  rectRadius: 0.14,
});
s5.addText("WEBSITE & COMMERCIAL", {
  x: 0.5, y: 0.4, w: 2.2, h: 0.28,
  fontSize: 7, color: T.white, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
  charSpacing: 1,
});

s5.addText("5-Page Minimal Professional Website", {
  x: 0.5, y: 0.85, w: 5.5, h: 0.5,
  fontSize: 22, color: T.gray900, bold: true,
  fontFace: H, align: "left", margin: 0,
});

// Features list (left side — modern check marks)
const features = [
  "Modern UI/UX design",
  "5-page responsive website",
  "Desktop & mobile optimisation",
  "Professional content structuring",
  "Image integration",
  "Contact/enquiry form",
  "Google Maps integration",
  "Social media integration",
  "Testing & deployment",
];

features.forEach((feat, i) => {
  const col = i < 5 ? 0 : 1;
  const row = i < 5 ? i : i - 5;
  const x = 0.5 + col * 2.8;
  const y = 1.55 + row * 0.42;

  // Check circle
  s5.addShape(pres.shapes.OVAL, {
    x: x, y: y + 0.05, w: 0.2, h: 0.2,
    fill: { color: T.emerald },
  });
  s5.addText("\u2713", {
    x: x, y: y + 0.03, w: 0.2, h: 0.22,
    fontSize: 8, color: T.white, bold: true,
    fontFace: B, align: "center", valign: "middle", margin: 0,
  });

  s5.addText(feat, {
    x: x + 0.3, y: y, w: 2.4, h: 0.3,
    fontSize: 10.5, color: T.gray700,
    fontFace: B, align: "left", margin: 0,
  });
});

// ---- PRICING CARD (dark premium) ----
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.2, y: 0.85, w: 3.4, h: 3.2,
  fill: { color: T.dark },
  rectRadius: 0.15,
  shadow: { type: "outer", color: "000000", blur: 12, offset: 4, opacity: 0.2, angle: 270 },
});

// Gold accent at top of card
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.5, y: 1.1, w: 1.5, h: 0.035,
  fill: { color: T.gold },
  rectRadius: 0.02,
});

s5.addText("Investment", {
  x: 6.5, y: 1.25, w: 2.8, h: 0.3,
  fontSize: 11, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
});

// Strikethrough original price
s5.addText("\u20B921,000", {
  x: 6.5, y: 1.55, w: 2.8, h: 0.3,
  fontSize: 13, color: T.gray500,
  fontFace: B, align: "left", margin: 0,
  strike: true,
});

// Big discounted price
s5.addText("\u20B916,000", {
  x: 6.5, y: 1.85, w: 2.8, h: 0.65,
  fontSize: 38, color: T.white, bold: true,
  fontFace: H, align: "left", margin: 0,
});

s5.addText("+ GST", {
  x: 8.55, y: 2.15, w: 0.8, h: 0.3,
  fontSize: 10, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
});

// Breakdown
s5.addShape(pres.shapes.RECTANGLE, {
  x: 6.5, y: 2.6, w: 2.8, h: 0.015,
  fill: { color: T.darkSurface },
});

s5.addText("GST @ 18%", {
  x: 6.5, y: 2.7, w: 1.8, h: 0.25,
  fontSize: 9, color: T.gray400,
  fontFace: B, align: "left", margin: 0,
});
s5.addText("\u20B92,880", {
  x: 8.1, y: 2.7, w: 1.2, h: 0.25,
  fontSize: 9, color: T.gray400,
  fontFace: B, align: "right", margin: 0,
});

// Total box
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.5, y: 3.05, w: 2.8, h: 0.5,
  fill: { color: T.gold },
  rectRadius: 0.06,
});
s5.addText("Total: \u20B918,880/-", {
  x: 6.5, y: 3.05, w: 2.8, h: 0.5,
  fontSize: 16, color: T.dark, bold: true,
  fontFace: H, align: "center", valign: "middle", margin: 0,
});

// Saving tag
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.5, y: 3.7, w: 1.8, h: 0.28,
  fill: { color: T.emeraldBg },
  rectRadius: 0.14,
});
s5.addText("\u20B95,000 SAVED", {
  x: 6.5, y: 3.7, w: 1.8, h: 0.28,
  fontSize: 8, color: T.emerald, bold: true,
  fontFace: B, align: "center", valign: "middle", margin: 0,
});


// ---- NEXT STEP (bottom section) ----
s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.35, w: 9.0, h: 1.1,
  fill: { color: T.white },
  rectRadius: 0.1,
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, opacity: 0.06, angle: 270 },
});

s5.addText("NEXT STEP", {
  x: 0.75, y: 4.45, w: 8.5, h: 0.22,
  fontSize: 8, color: T.dark, bold: true,
  fontFace: B, align: "left", margin: 0,
  charSpacing: 1.5,
});

// Process steps
const steps = ["Select approach", "Finalise structure", "Design", "Development", "Testing", "Deployment"];
steps.forEach((step, i) => {
  const sx = 0.75 + i * 1.48;
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: sx, y: 4.75, w: 1.3, h: 0.3,
    fill: { color: i === 0 ? T.dark : T.gray100 },
    rectRadius: 0.15,
  });
  s5.addText(step, {
    x: sx, y: 4.75, w: 1.3, h: 0.3,
    fontSize: 7.5, color: i === 0 ? T.white : T.gray700, bold: i === 0,
    fontFace: B, align: "center", valign: "middle", margin: 0,
  });
  // Arrow between steps
  if (i < steps.length - 1) {
    s5.addText("\u203A", {
      x: sx + 1.3, y: 4.73, w: 0.18, h: 0.3,
      fontSize: 12, color: T.gray400,
      fontFace: B, align: "center", valign: "middle", margin: 0,
    });
  }
});

// Footer disclaimer
s5.addText("The website will be developed as a professional, accurate and scalable digital presence. Final payment gateway approval remains subject to the bank/payment provider\u2019s policies and compliance assessment.", {
  x: 0.75, y: 5.12, w: 8.5, h: 0.25,
  fontSize: 7.5, color: T.gray500, italic: true,
  fontFace: B, align: "left", margin: 0,
});

// ============================================================
// SAVE
// ============================================================
const outputPath = __dirname + "/Casino_Pride_Digital_Strategy.pptx";
pres.writeFile({ fileName: outputPath })
  .then(() => {
    console.log("\u2713 Presentation created successfully!");
    console.log("  File: " + outputPath);
    console.log("  Slides: 5");
    console.log("  Theme: Modern Dark Premium");
    console.log("  Layout: 16:9 (10\" x 5.625\")");
  })
  .catch(err => {
    console.error("Error:", err);
    process.exit(1);
  });
