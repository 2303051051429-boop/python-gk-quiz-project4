/**
 * Python GK Quiz Training Deck — 20-slide .pptx generator
 * DecodeLabs | Batch 2026
 * Built with pptxgenjs
 */

const PptxGenJS = require("pptxgenjs");
const pptx = new PptxGenJS();

// ─── Design Tokens ───────────────────────────────────────────────────────────
const BG        = "0B0B0F";
const ACCENT1   = "5B8CFF";   // blue
const ACCENT2   = "A05BFF";   // purple
const WHITE     = "FFFFFF";
const LIGHT     = "C8D6FF";   // soft lavender-white for body
const DIM       = "6B7280";   // muted grey for secondary labels
const CARD_FILL = "1A1E2E";   // dark navy card
const CODE_BG   = "111827";   // deeper card for code blocks
const GREEN     = "34D399";
const RED       = "F87171";
const YELLOW    = "FBBF24";

const TITLE_FONT   = "Segoe UI";
const BODY_FONT    = "Segoe UI";
const CODE_FONT    = "Courier New";

// slide dimensions (inches) — widescreen 13.33 × 7.5
const W = 13.33;
const H = 7.5;

pptx.layout = "LAYOUT_WIDE";
pptx.author = "DecodeLabs";
pptx.subject = "Python GK Quiz Training Deck";
pptx.title   = "Python Programming – Industrial Training Kit – Project 4";

// ─── Helper: add slide bg ────────────────────────────────────────────────────
function addBg(slide) {
  slide.background = { color: BG };
}

// ─── Helper: heading text ────────────────────────────────────────────────────
function addTitle(slide, text, y = 0.35, size = 34, color = WHITE) {
  slide.addText(text, {
    x: 0.5, y, w: W - 1, h: 0.7,
    fontSize: size, bold: true, color,
    fontFace: TITLE_FONT, valign: "middle",
  });
}

// ─── Helper: section sub-label ───────────────────────────────────────────────
function addLabel(slide, text, y = 0.28) {
  slide.addText(text, {
    x: 0.5, y, w: 4, h: 0.28,
    fontSize: 11, color: ACCENT1, bold: true,
    fontFace: BODY_FONT, charSpacing: 2,
  });
}

// ─── Helper: body paragraph ──────────────────────────────────────────────────
function addBody(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: 16, color: LIGHT, fontFace: BODY_FONT,
    valign: "top", wrap: true,
    ...opts,
  });
}

// ─── Helper: glass card (rounded rect) ───────────────────────────────────────
function addCard(slide, x, y, w, h, fillColor = CARD_FILL, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: fillColor, transparency: 20 },
    line: { color: ACCENT1, width: 0.5, transparency: 60 },
    rectRadius: 0.12,
    shadow: { type: "outer", blur: 8, offset: 3, angle: 45, color: "000000", opacity: 0.4 },
    ...opts,
  });
}

// ─── Helper: code block ──────────────────────────────────────────────────────
function addCode(slide, lines, x, y, w, h) {
  addCard(slide, x, y, w, h, CODE_BG);
  slide.addText(lines, {
    x: x + 0.18, y: y + 0.15, w: w - 0.36, h: h - 0.3,
    fontSize: 13, fontFace: CODE_FONT, color: "D4E4FF",
    valign: "top", wrap: true, lineSpacing: 20,
  });
}

// ─── Helper: divider line ────────────────────────────────────────────────────
function addDivider(slide, y = 1.08) {
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5, y, w: W - 1, h: 0.018,
    fill: { color: ACCENT1, transparency: 65 },
    line: { type: "none" },
  });
}

// ─── Helper: bullet list ─────────────────────────────────────────────────────
function addBullets(slide, items, x, y, w, h, color = LIGHT) {
  const bulletItems = items.map(t => ({
    text: t,
    options: { bullet: { type: "bullet", characterCode: "25B8" }, indent: 14 },
  }));
  slide.addText(bulletItems, {
    x, y, w, h,
    fontSize: 16, color, fontFace: BODY_FONT,
    lineSpacing: 26, valign: "top", wrap: true,
  });
}

// ─── Helper: small badge ─────────────────────────────────────────────────────
function addBadge(slide, text, x, y, color = ACCENT1) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w: 1.6, h: 0.32,
    fill: { color }, line: { type: "none" }, rectRadius: 0.08,
  });
  slide.addText(text, {
    x: x + 0.04, y: y + 0.04, w: 1.52, h: 0.24,
    fontSize: 11, bold: true, color: WHITE, fontFace: BODY_FONT,
    align: "center",
  });
}

// ─── Helper: box with label ──────────────────────────────────────────────────
function addBoxLabel(slide, label, sub, x, y, w = 2.6, h = 0.95, fillC = CARD_FILL, accentC = ACCENT1) {
  addCard(slide, x, y, w, h, fillC);
  slide.addText(label, {
    x: x + 0.1, y: y + 0.1, w: w - 0.2, h: 0.4,
    fontSize: 17, bold: true, color: accentC, fontFace: TITLE_FONT, align: "center",
  });
  if (sub) {
    slide.addText(sub, {
      x: x + 0.1, y: y + 0.52, w: w - 0.2, h: 0.35,
      fontSize: 13, color: LIGHT, fontFace: BODY_FONT, align: "center",
    });
  }
}

// ─── Helper: arrow right ─────────────────────────────────────────────────────
function addArrow(slide, x, y) {
  slide.addShape(pptx.ShapeType.rightArrow, {
    x, y: y + 0.3, w: 0.55, h: 0.35,
    fill: { color: ACCENT1, transparency: 40 },
    line: { type: "none" },
  });
}

// ─── Helper: analogy line ────────────────────────────────────────────────────
function addAnalogy(slide, text, y = 6.65) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.5, y, w: W - 1, h: 0.48,
    fill: { color: "16213E", transparency: 20 },
    line: { color: ACCENT2, width: 0.8, transparency: 50 },
    rectRadius: 0.08,
  });
  slide.addText("💡 " + text, {
    x: 0.72, y: y + 0.06, w: W - 1.44, h: 0.36,
    fontSize: 14, color: "C4B5FD", fontFace: BODY_FONT,
    italic: true, valign: "middle",
  });
}

// ─── Helper: footer tag ──────────────────────────────────────────────────────
function addFooter(slide, num) {
  slide.addText("DecodeLabs · Batch 2026 · Project 4", {
    x: 0.5, y: H - 0.28, w: W - 1.5, h: 0.22,
    fontSize: 9, color: DIM, fontFace: BODY_FONT,
  });
  slide.addText(num + " / 20", {
    x: W - 1.0, y: H - 0.28, w: 0.7, h: 0.22,
    fontSize: 9, color: DIM, fontFace: BODY_FONT, align: "right",
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 1 — Title
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);

  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: W, h: 0.06,
    fill: { color: ACCENT1 }, line: { type: "none" },
  });

  addCard(s, 1.8, 1.7, 9.7, 1.6, "131929");
  s.addText("Python Programming", {
    x: 1.9, y: 1.8, w: 9.5, h: 0.65,
    fontSize: 38, bold: true, color: WHITE,
    fontFace: TITLE_FONT, align: "center",
  });
  s.addText("Industrial Training Kit — Project 4", {
    x: 1.9, y: 2.48, w: 9.5, h: 0.38,
    fontSize: 18, color: ACCENT1, fontFace: BODY_FONT,
    align: "center",
  });

  addCard(s, 3.5, 3.6, 6.3, 0.82, "1E1B4B");
  s.addText("The General Knowledge Quiz", {
    x: 3.6, y: 3.72, w: 6.1, h: 0.5,
    fontSize: 24, bold: true, color: ACCENT2,
    fontFace: TITLE_FONT, align: "center",
  });

  s.addText("DecodeLabs  ·  Batch 2026", {
    x: 0, y: 5.1, w: W, h: 0.35,
    fontSize: 15, color: DIM, fontFace: BODY_FONT, align: "center",
  });

  const badges = ["Variables & Input", "Sanitization", "Score Tracking"];
  badges.forEach(function(b, i) { addBadge(s, b, 3.5 + i * 2.1, 5.7); });

  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: H - 0.06, w: W, h: 0.06,
    fill: { color: ACCENT2 }, line: { type: "none" },
  });

  s.addNotes("Title slide. Welcome students and introduce the project.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 2 — What You'll Build
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "OVERVIEW");
  addTitle(s, "What You'll Build", 0.55);
  addDivider(s, 1.08);

  addCard(s, 0.5, 1.25, W - 1, 1.05, "131929");
  s.addText("A 3-question General Knowledge quiz that accepts user input, sanitizes it, checks answers, and outputs a final score — all in under 20 lines of Python.", {
    x: 0.75, y: 1.35, w: W - 1.5, h: 0.85,
    fontSize: 18, color: WHITE, fontFace: BODY_FONT,
    valign: "middle", align: "center", bold: true,
  });

  const features = [
    { icon: "❓", title: "3 GK Questions", sub: "Capital, Planet, Language" },
    { icon: "🧹", title: "Input Sanitization", sub: "strip() + lower()" },
    { icon: "🏆", title: "Final Score", sub: "Printed with f-string" },
  ];
  features.forEach(function(f, i) {
    const x = 0.5 + i * 4.28;
    addCard(s, x, 2.55, 3.9, 2.1, CARD_FILL);
    s.addText(f.icon, { x: x, y: 2.68, w: 3.9, h: 0.6, fontSize: 30, align: "center" });
    s.addText(f.title, {
      x: x + 0.1, y: 3.3, w: 3.7, h: 0.45,
      fontSize: 17, bold: true, color: ACCENT1, fontFace: TITLE_FONT, align: "center",
    });
    s.addText(f.sub, {
      x: x + 0.1, y: 3.76, w: 3.7, h: 0.6,
      fontSize: 14, color: LIGHT, fontFace: BODY_FONT, align: "center",
    });
  });

  addBody(s, "The same logic powers every online quiz, form validator, and authentication check you've ever used.", 0.5, 4.9, W - 1, 0.5);
  addAnalogy(s, "Think of this as the skeleton key for all interactive programs — once you build one quiz, you understand every quiz.");
  addFooter(s, 2);
  s.addNotes("Walk through each feature box. Emphasise that the whole program fits in 20 lines, yet covers the core of interactive Python.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 3 — Learning Objectives
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "OBJECTIVES");
  addTitle(s, "Learning Objectives", 0.55);
  addDivider(s, 1.08);

  const objectives = [
    ["Variables",     "Store and update quiz state (score, answers)"],
    ["input()",       "Capture user responses from the terminal"],
    ["print()",       "Display questions, results, and final score"],
    ["Sanitization",  ".strip() + .lower() to normalise raw input"],
    ["if / else",     "Branch on correct vs. wrong answer"],
    ["== operator",   "Compare cleaned input to stored answer"],
    ["Score State",   "Accumulate score across multiple iterations"],
  ];

  // first 6 in 2-col layout
  for (var i = 0; i < 6; i++) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = col === 0 ? 0.5 : 7.0;
    var y = 1.28 + row * 1.08;
    var w = col === 0 ? 6.0 : 5.83;
    addCard(s, x, y, w, 0.88, CARD_FILL);
    addBadge(s, objectives[i][0], x + 0.12, y + 0.28, ACCENT1);
    s.addText(objectives[i][1], {
      x: x + 1.85, y: y + 0.2, w: w - 1.97, h: 0.5,
      fontSize: 15, color: LIGHT, fontFace: BODY_FONT, valign: "middle",
    });
  }

  // 7th item full-width
  addCard(s, 0.5, 1.28 + 3 * 1.08, W - 1, 0.88, CARD_FILL);
  addBadge(s, objectives[6][0], 0.62, 1.28 + 3 * 1.08 + 0.28, ACCENT2);
  s.addText(objectives[6][1], {
    x: 2.35, y: 1.28 + 3 * 1.08 + 0.2, w: W - 2.97, h: 0.5,
    fontSize: 15, color: LIGHT, fontFace: BODY_FONT, valign: "middle",
  });

  addFooter(s, 3);
  s.addNotes("Seven Python concepts the student will have working knowledge of by the end of the project.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 4 — From Static Script to Decision Engine
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "MOTIVATION");
  addTitle(s, "From Static Script to Decision Engine", 0.55);
  addDivider(s, 1.08);

  addCard(s, 0.5, 1.25, 5.7, 2.6, CODE_BG);
  s.addText("❌  Hardcoded Output", {
    x: 0.65, y: 1.35, w: 5.4, h: 0.38, fontSize: 14, bold: true, color: RED, fontFace: BODY_FONT,
  });
  s.addText('print("What is the capital of France?")\nprint("Answer: paris")\nprint("Correct!")   # always correct!', {
    x: 0.68, y: 1.76, w: 5.3, h: 1.0,
    fontSize: 13, fontFace: CODE_FONT, color: "D4E4FF", valign: "top", wrap: true, lineSpacing: 20,
  });
  s.addText("Problem: the program can never tell if the user is right or wrong. It just prints a pre-determined result.", {
    x: 0.65, y: 2.82, w: 5.4, h: 0.72, fontSize: 14, color: LIGHT, fontFace: BODY_FONT, wrap: true,
  });

  s.addShape(pptx.ShapeType.rightArrow, {
    x: 6.38, y: 2.35, w: 0.72, h: 0.5,
    fill: { color: ACCENT1 }, line: { type: "none" },
  });

  addCard(s, 7.2, 1.25, 5.6, 2.6, CODE_BG);
  s.addText("✅  Decision Engine", {
    x: 7.35, y: 1.35, w: 5.3, h: 0.38, fontSize: 14, bold: true, color: GREEN, fontFace: BODY_FONT,
  });
  s.addText('answer = input("Capital of France? ")\nif answer.strip().lower() == "paris":\n    print("Correct!")\nelse:\n    print("Wrong!")', {
    x: 7.38, y: 1.76, w: 5.28, h: 1.0,
    fontSize: 13, fontFace: CODE_FONT, color: "D4E4FF", valign: "top", wrap: true, lineSpacing: 20,
  });
  s.addText("Now the program evaluates the user's actual response and branches based on correctness.", {
    x: 7.35, y: 2.82, w: 5.3, h: 0.72, fontSize: 14, color: LIGHT, fontFace: BODY_FONT, wrap: true,
  });

  addBody(s, "The shift from print-only scripts to interactive, branching programs is the single most important step in learning Python.", 0.5, 4.1, W - 1, 0.5);
  addAnalogy(s, "A thermostat that always prints '22 C' is useless. One that reads the room and reacts — that's engineering.");
  addFooter(s, 4);
  s.addNotes("Contrast the two approaches side by side. The left is what most beginners write first. The right is what we're aiming for.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 5 — IPO Architecture
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "ARCHITECTURE");
  addTitle(s, "IPO Architecture: Input — Process — Output", 0.55);
  addDivider(s, 1.08);

  const boxes = [
    { label: "INPUT",   sub: "input(question)\nUser types answer",            x: 0.55,  color: "3B82F6" },
    { label: "PROCESS", sub: ".strip().lower()\nif cleaned == answer",         x: 4.97,  color: ACCENT2 },
    { label: "OUTPUT",  sub: 'print("Correct!")\nprint(f"Score: {score}")',    x: 9.38,  color: GREEN },
  ];
  boxes.forEach(function(b) {
    addCard(s, b.x, 1.5, 3.55, 2.3, CARD_FILL);
    s.addText(b.label, {
      x: b.x + 0.1, y: 1.62, w: 3.35, h: 0.52,
      fontSize: 22, bold: true, color: b.color, fontFace: TITLE_FONT, align: "center",
    });
    s.addShape(pptx.ShapeType.rect, {
      x: b.x + 0.3, y: 2.18, w: 2.95, h: 0.018,
      fill: { color: b.color, transparency: 60 }, line: { type: "none" },
    });
    s.addText(b.sub, {
      x: b.x + 0.1, y: 2.26, w: 3.35, h: 1.3,
      fontSize: 14, color: LIGHT, fontFace: CODE_FONT, valign: "top", wrap: true, lineSpacing: 20,
    });
  });

  [4.26, 8.67].forEach(function(ax) {
    s.addShape(pptx.ShapeType.rightArrow, {
      x: ax, y: 2.32, w: 0.55, h: 0.42,
      fill: { color: ACCENT1 }, line: { type: "none" },
    });
  });

  addCard(s, 0.5, 4.1, W - 1, 1.55, "131929");
  s.addText("Quiz Mapping", {
    x: 0.7, y: 4.18, w: 3, h: 0.35, fontSize: 14, bold: true, color: ACCENT1, fontFace: BODY_FONT,
  });
  var mapping = [
    ["INPUT",   "User types 'PARIS' or ' paris ' in terminal"],
    ["PROCESS", "Sanitize to 'paris', compare to stored answer 'paris'"],
    ["OUTPUT",  "Print 'Correct!' and increment score by 1"],
  ];
  mapping.forEach(function(row, i) {
    s.addText(row[0] + ":", { x: 0.7, y: 4.6 + i * 0.32, w: 1.1, h: 0.28, fontSize: 13, bold: true, color: ACCENT1, fontFace: BODY_FONT });
    s.addText(row[1],       { x: 1.85, y: 4.6 + i * 0.32, w: W - 2.55, h: 0.28, fontSize: 13, color: LIGHT, fontFace: BODY_FONT });
  });

  addFooter(s, 5);
  s.addNotes("IPO is the universal model for all programs. Map each phase to a concrete line of Python in the quiz.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 6 — Program Flow Overview
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "FLOW");
  addTitle(s, "Program Flow Overview", 0.55);
  addDivider(s, 1.08);

  var nodes = [
    { label: "START",              y: 1.25, color: GREEN },
    { label: "score = 0",          y: 2.05, color: ACCENT1 },
    { label: "for each question",  y: 2.85, color: ACCENT2 },
    { label: "Ask question",       y: 3.65, color: ACCENT1 },
    { label: "Sanitize input",     y: 4.45, color: ACCENT1 },
    { label: "if cleaned == answer", y: 5.25, color: YELLOW },
    { label: "END",                y: 6.05, color: RED },
  ];

  nodes.forEach(function(n) {
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.5, y: n.y, w: 3.8, h: 0.62,
      fill: { color: CARD_FILL, transparency: 20 },
      line: { color: n.color, width: 1.5 }, rectRadius: 0.08,
    });
    s.addText(n.label, {
      x: 0.55, y: n.y + 0.1, w: 3.7, h: 0.42,
      fontSize: 14, bold: true, color: n.color,
      fontFace: CODE_FONT, align: "center", valign: "middle",
    });
    if (n.y < 6.05) {
      s.addShape(pptx.ShapeType.line, {
        x: 2.4, y: n.y + 0.62, w: 0, h: 0.43,
        line: { color: DIM, width: 1.2 },
      });
    }
  });

  addCard(s, 4.6, 5.08, 3.6, 0.95, CARD_FILL);
  s.addText("True (Correct)", { x: 4.72, y: 5.15, w: 3.36, h: 0.35, fontSize: 13, bold: true, color: GREEN, fontFace: BODY_FONT });
  s.addText('print("Correct!") + score += 1', { x: 4.72, y: 5.5, w: 3.36, h: 0.35, fontSize: 12, color: LIGHT, fontFace: CODE_FONT });

  addCard(s, 8.55, 5.08, 4.28, 0.95, CARD_FILL);
  s.addText("False (Wrong)", { x: 8.67, y: 5.15, w: 4.04, h: 0.35, fontSize: 13, bold: true, color: RED, fontFace: BODY_FONT });
  s.addText("print(f\"Wrong. Answer: '{answer}'\")", { x: 8.67, y: 5.5, w: 4.04, h: 0.35, fontSize: 12, color: LIGHT, fontFace: CODE_FONT });

  s.addShape(pptx.ShapeType.line, { x: 2.4, y: 5.57, w: 2.2, h: 0, line: { color: GREEN, width: 1 } });
  s.addShape(pptx.ShapeType.line, { x: 2.4, y: 5.57, w: 6.15, h: 0, line: { color: RED, width: 1 } });

  addCard(s, 4.6, 6.25, 8.23, 0.62, "131929");
  s.addText('After loop: print(f"Final score: {score}/{len(questions)}")', {
    x: 4.75, y: 6.35, w: 7.93, h: 0.35, fontSize: 13, color: LIGHT, fontFace: CODE_FONT,
  });

  addFooter(s, 6);
  s.addNotes("Walk the flowchart step by step. Score accumulates but is never reset inside the loop.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 7 — Input Chaos
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "PROBLEM");
  addTitle(s, "Input Chaos: Why Naive == Fails", 0.55);
  addDivider(s, 1.08);

  addBody(s, 'Stored answer: "paris"   |   Comparison: user_input == "paris"', 0.5, 1.22, W - 1, 0.4, { bold: true, color: ACCENT1 });

  var variants = [
    { raw: '"Paris"',   result: "FAIL  X", why: "Capital P — different character",    color: RED },
    { raw: '" paris "', result: "FAIL  X", why: "Leading + trailing spaces",           color: RED },
    { raw: '"PARIS "',  result: "FAIL  X", why: "All-caps + trailing space",           color: RED },
    { raw: '"paris"',   result: "PASS  OK", why: "Exact match — only this passes",    color: GREEN },
  ];

  variants.forEach(function(v, i) {
    var x = (i % 2 === 0) ? 0.5 : 6.92;
    var y = i < 2 ? 1.8 : 3.55;
    addCard(s, x, y, 6.08, 1.52, CARD_FILL);
    s.addText(v.raw, { x: x + 0.15, y: y + 0.12, w: 3.8, h: 0.45, fontSize: 20, bold: true, color: WHITE, fontFace: CODE_FONT });
    s.addText(v.result, { x: x + 4.0, y: y + 0.12, w: 1.9, h: 0.45, fontSize: 18, bold: true, color: v.color, fontFace: BODY_FONT, align: "right" });
    s.addText(v.why, { x: x + 0.15, y: y + 0.65, w: 5.78, h: 0.72, fontSize: 14, color: LIGHT, fontFace: BODY_FONT, wrap: true });
  });

  addAnalogy(s, "A vending machine that only accepts exactly '1.00' rejects the same dollar if it enters slightly crooked. That's your naive == check.");
  addFooter(s, 7);
  s.addNotes("This is the motivating problem for sanitization. Real users don't type precisely — your program must handle that.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 8 — Sanitization Pipeline
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "SOLUTION");
  addTitle(s, "Sanitization Pipeline: .strip() then .lower()", 0.55);
  addDivider(s, 1.08);

  var steps = [
    { stage: "Raw input",         value: '" PARIS "',                    note: "Straight from input()",                 color: RED    },
    { stage: ".strip()",          value: '"PARIS"',                      note: "Removes leading/trailing whitespace",    color: YELLOW },
    { stage: ".lower()",          value: '"paris"',                      note: "Converts all chars to lowercase",        color: GREEN  },
    { stage: "== answer",         value: '"paris" == "paris"  ->  True', note: "Safe, reliable comparison",             color: GREEN  },
  ];

  steps.forEach(function(st, i) {
    var y = 1.28 + i * 1.15;
    addCard(s, 0.5, y, 2.6, 0.9, CARD_FILL);
    s.addText(st.stage, { x: 0.6, y: y + 0.1, w: 2.4, h: 0.7, fontSize: 13, bold: true, color: st.color, fontFace: BODY_FONT, valign: "middle" });
    addCard(s, 3.35, y, 4.0, 0.9, CODE_BG);
    s.addText(st.value, { x: 3.48, y: y + 0.18, w: 3.74, h: 0.55, fontSize: 16, color: WHITE, fontFace: CODE_FONT, valign: "middle" });
    s.addText(st.note, { x: 7.55, y: y + 0.18, w: 5.3, h: 0.55, fontSize: 14, color: LIGHT, fontFace: BODY_FONT, valign: "middle" });
    if (i < steps.length - 1) {
      s.addShape(pptx.ShapeType.line, { x: 1.8, y: y + 0.9, w: 0, h: 0.25, line: { color: DIM, width: 1 } });
    }
  });

  addCode(s, "cleaned = user_input.strip().lower()\nif cleaned == answer:", 0.5, 5.65, 8.5, 0.82);
  addAnalogy(s, "strip() removes the wrapper; lower() puts everyone in the same uniform before comparison.");
  addFooter(s, 8);
  s.addNotes("Order matters: strip first, then lower. Both are chained on one line. This pattern repeats for every question.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 9 — Comparison Operators
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "OPERATORS");
  addTitle(s, "Comparison Operators", 0.55);
  addDivider(s, 1.08);

  var ops = [
    { op: "==", desc: "Equal to",             ex: '"paris" == "paris"  ->  True',        color: GREEN },
    { op: "!=", desc: "Not equal to",          ex: '"paris" != "london"  ->  True',       color: YELLOW },
    { op: ">",  desc: "Greater than",          ex: "score > 0   ->  True",                color: ACCENT1 },
    { op: "<",  desc: "Less than",             ex: "score < 3   ->  True",                color: ACCENT2 },
    { op: ">=", desc: "Greater than or equal", ex: "score >= 2  ->  True (pass threshold)", color: GREEN },
  ];

  ops.forEach(function(o, i) {
    var y = 1.28 + i * 0.98;
    addCard(s, 0.5, y, 0.9, 0.75, "1E1B4B");
    s.addText(o.op, { x: 0.52, y: y + 0.1, w: 0.86, h: 0.55, fontSize: 22, bold: true, color: o.color, fontFace: CODE_FONT, align: "center" });
    s.addText(o.desc, { x: 1.52, y: y + 0.1, w: 2.2, h: 0.55, fontSize: 15, color: WHITE, fontFace: BODY_FONT, valign: "middle" });
    addCard(s, 3.9, y, 9.0, 0.75, CODE_BG);
    s.addText(o.ex, { x: 4.05, y: y + 0.16, w: 8.72, h: 0.44, fontSize: 13, color: LIGHT, fontFace: CODE_FONT, valign: "middle" });
  });

  addBody(s, "The == operator is the workhorse of this project. Every correct-answer check is a string equality comparison on sanitized input.", 0.5, 6.15, W - 1, 0.42);
  addFooter(s, 9);
  s.addNotes("Clarify that == tests value equality. Don't introduce 'is' yet — it's a distraction at this stage.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 10 — State Management: Score Vault
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "STATE");
  addTitle(s, "State Management — The Score Vault", 0.55);
  addDivider(s, 1.08);

  var moments = [
    { step: "Initialise", code: "score = 0",      note: "Before the loop. Set once.",    color: ACCENT1 },
    { step: "Accumulate", code: "score += 1",     note: "Inside the loop, on correct.",  color: GREEN },
    { step: "Read",       code: 'f"{score}/3"',   note: "After the loop. Final output.", color: YELLOW },
  ];

  moments.forEach(function(m, i) {
    var x = 0.5 + i * 4.28;
    addCard(s, x, 1.28, 3.85, 2.6, CARD_FILL);
    addBadge(s, m.step, x + 0.15, 1.4, m.color === GREEN ? "059669" : (m.color === YELLOW ? "B45309" : "2563EB"));
    addCard(s, x + 0.15, 1.85, 3.55, 0.72, CODE_BG);
    s.addText(m.code, { x: x + 0.28, y: 1.93, w: 3.29, h: 0.56, fontSize: 18, bold: true, color: m.color, fontFace: CODE_FONT, valign: "middle" });
    s.addText(m.note, { x: x + 0.15, y: 2.65, w: 3.55, h: 1.0, fontSize: 14, color: LIGHT, fontFace: BODY_FONT, wrap: true });
  });

  addCard(s, 0.5, 4.1, W - 1, 1.85, "131929");
  s.addText("Score trace across 3 iterations", { x: 0.7, y: 4.18, w: 6, h: 0.35, fontSize: 14, bold: true, color: ACCENT1, fontFace: BODY_FONT });

  ["Iteration", "Question", "Correct?", "score after"].forEach(function(h, i) {
    s.addText(h, { x: 0.7 + i * 3.0, y: 4.6, w: 2.8, h: 0.32, fontSize: 13, bold: true, color: WHITE, fontFace: BODY_FONT });
  });
  [
    ["1", "Capital of France?",     "Yes", "1"],
    ["2", "Largest planet?",        "No",  "1"],
    ["3", "Language named snake?",  "Yes", "2"],
  ].forEach(function(r, ri) {
    r.forEach(function(cell, ci) {
      s.addText(cell, {
        x: 0.7 + ci * 3.0, y: 5.0 + ri * 0.32, w: 2.8, h: 0.28,
        fontSize: 13, color: ci === 2 ? (cell === "Yes" ? GREEN : RED) : LIGHT,
        fontFace: ci === 3 ? CODE_FONT : BODY_FONT,
      });
    });
  });

  addFooter(s, 10);
  s.addNotes("Key insight: score lives OUTSIDE the loop so it survives across iterations. If it were inside, it would reset to 0 every time.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 11 — Normalization Matrix
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "DEEP DIVE");
  addTitle(s, "Normalization Matrix", 0.55);
  addDivider(s, 1.08);

  var colW = [2.2, 2.9, 4.3, 3.4];
  var colX = [0.5, 2.8, 5.8, 10.2];
  ["Method", "What it removes", "Example in -> out", "Use in this project"].forEach(function(h, i) {
    addCard(s, colX[i], 1.28, colW[i] - 0.1, 0.5, "1E1B4B");
    s.addText(h, { x: colX[i] + 0.1, y: 1.31, w: colW[i] - 0.22, h: 0.44, fontSize: 13, bold: true, color: ACCENT1, fontFace: BODY_FONT, valign: "middle" });
  });

  [
    [".strip()",    "Leading & trailing whitespace", '" PARIS "  ->  "PARIS"', "Always — user input has invisible spaces"],
    [".lower()",    "Mixed-case differences",        '"PARIS"  ->  "paris"',   "Always — answers stored in lowercase"],
    [".casefold()", "Additional Unicode cases",      '"ss"  for German sz',    "Optional — for international input"],
  ].forEach(function(row, ri) {
    row.forEach(function(cell, ci) {
      addCard(s, colX[ci], 1.92 + ri * 1.1, colW[ci] - 0.1, 0.95, CARD_FILL);
      s.addText(cell, {
        x: colX[ci] + 0.1, y: 1.97 + ri * 1.1, w: colW[ci] - 0.22, h: 0.82,
        fontSize: ci === 0 ? 15 : 13, bold: ci === 0, color: ci === 0 ? ACCENT2 : LIGHT,
        fontFace: ci <= 1 ? CODE_FONT : BODY_FONT, valign: "middle", wrap: true,
      });
    });
  });

  addBody(s, "For this project, .strip().lower() is sufficient. .casefold() is a good extension exercise for advanced students.", 0.5, 5.35, W - 1, 0.42);
  addAnalogy(s, "strip() sweeps the doorstep; lower() makes everyone wear the same uniform before the check.");
  addFooter(s, 11);
  s.addNotes("The matrix gives students a reference card. .strip().lower() covers 99% of English-language input scenarios.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 12 — Logic Gate: if/else
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "BRANCHING");
  addTitle(s, "Logic Gate: if / else Branching", 0.55);
  addDivider(s, 1.08);

  addCard(s, 0.5, 1.28, 7.5, 3.2, CODE_BG);
  [
    { text: "if cleaned == answer:",                               color: YELLOW, y: 1.42 },
    { text: '    print("Correct!")',                               color: GREEN,  y: 1.82 },
    { text: "    score += 1",                                      color: GREEN,  y: 2.22 },
    { text: "else:",                                               color: RED,    y: 2.62 },
    { text: "    print(f\"Wrong. The correct answer was '{answer}'.\")", color: RED, y: 3.02 },
  ].forEach(function(line) {
    s.addText(line.text, { x: 0.68, y: line.y, w: 7.1, h: 0.38, fontSize: 15, fontFace: CODE_FONT, color: line.color });
  });

  // callout annotations
  [
    { text: "Condition: equality check on sanitized string", y: 1.52, color: YELLOW },
    { text: "True branch: reward correct answer",            y: 2.0,  color: GREEN },
    { text: "False branch: inform of correct answer",        y: 2.82, color: RED },
  ].forEach(function(a) {
    s.addShape(pptx.ShapeType.line, { x: 8.1, y: a.y, w: 0.5, h: 0, line: { color: a.color, width: 0.8 } });
    s.addText(a.text, { x: 8.68, y: a.y - 0.12, w: 4.2, h: 0.35, fontSize: 12, color: a.color, fontFace: BODY_FONT, wrap: true });
  });

  addBody(s, "The if block runs only when the condition is True. The else block handles every other case. Together they form a complete decision with no gaps.", 0.5, 4.65, W - 1, 0.6);
  addAnalogy(s, "A traffic light either shows green (go) or not. if/else enforces the same binary certainty in your code.");
  addFooter(s, 12);
  s.addNotes("Point out the indentation — Python uses 4 spaces to define which code belongs to which branch. Missing indentation is the #1 syntax error beginners make.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 13 — Question Block Pattern
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "PATTERN");
  addTitle(s, "The Question Block Pattern", 0.55);
  addDivider(s, 1.08);

  var steps6 = [
    { n: "1", label: "ASK",          detail: "Print the question to terminal",          color: ACCENT1 },
    { n: "2", label: "CAPTURE",      detail: "user_input = input(question + ' ')",      color: ACCENT1 },
    { n: "3", label: "SANITIZE",     detail: "cleaned = user_input.strip().lower()",    color: YELLOW  },
    { n: "4", label: "EVALUATE",     detail: "if cleaned == answer:",                   color: YELLOW  },
    { n: "5", label: "EXECUTE",      detail: 'print("Correct!") or print("Wrong")',     color: GREEN   },
    { n: "6", label: "UPDATE SCORE", detail: "score += 1  (only on correct)",           color: GREEN   },
  ];

  steps6.forEach(function(st, i) {
    var row = Math.floor(i / 3);
    var col = i % 3;
    var x = 0.5 + col * 4.28;
    var y = 1.28 + row * 1.85;
    addCard(s, x, y, 3.9, 1.65, CARD_FILL);
    addBadge(s, st.n + ". " + st.label, x + 0.15, y + 0.12, st.color === GREEN ? "059669" : (st.color === YELLOW ? "B45309" : "2563EB"));
    s.addText(st.detail, { x: x + 0.15, y: y + 0.58, w: 3.6, h: 0.95, fontSize: 13, color: LIGHT, fontFace: CODE_FONT, wrap: true, valign: "top" });
  });

  addBody(s, "Every question in the quiz follows this exact six-step pattern. Extract it into a function and you have a reusable quiz engine.", 0.5, 5.15, W - 1, 0.5);
  addAnalogy(s, "Like a recipe card — same steps, different ingredients (questions). The pattern doesn't change; only the data does.");
  addFooter(s, 13);
  s.addNotes("This pattern is the bridge to functions. Once students see steps 1-6 repeat identically, they naturally ask 'can we write this once?' — that's the function concept.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 14 — f-strings
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "SYNTAX");
  addTitle(s, "f-strings: Clean String Formatting", 0.55);
  addDivider(s, 1.08);

  addCard(s, 0.5, 1.28, 5.8, 2.35, CODE_BG);
  s.addText("X  String Concatenation", { x: 0.65, y: 1.36, w: 5.5, h: 0.38, fontSize: 14, bold: true, color: RED, fontFace: BODY_FONT });
  s.addText('"Score: " + str(score) + "/" + str(len(questions))', { x: 0.65, y: 1.78, w: 5.5, h: 0.45, fontSize: 13, color: LIGHT, fontFace: CODE_FONT });
  s.addText("Requires manual str() conversion; hard to read; breaks easily.", { x: 0.65, y: 2.3, w: 5.5, h: 0.65, fontSize: 13, color: DIM, fontFace: BODY_FONT, wrap: true });
  s.addText("-> TypeError if you forget str()", { x: 0.65, y: 2.96, w: 5.5, h: 0.42, fontSize: 13, color: RED, fontFace: CODE_FONT });

  addCard(s, 6.62, 1.28, 6.2, 2.35, CODE_BG);
  s.addText("OK  f-string", { x: 6.77, y: 1.36, w: 5.9, h: 0.38, fontSize: 14, bold: true, color: GREEN, fontFace: BODY_FONT });
  s.addText('f"Score: {score}/{len(questions)}"', { x: 6.77, y: 1.78, w: 5.9, h: 0.45, fontSize: 13, color: LIGHT, fontFace: CODE_FONT });
  s.addText("Variables embedded directly. Auto-converts to string. Reads like plain English.", { x: 6.77, y: 2.3, w: 5.9, h: 0.65, fontSize: 13, color: LIGHT, fontFace: BODY_FONT, wrap: true });
  s.addText("-> Always works — Python handles types", { x: 6.77, y: 2.96, w: 5.9, h: 0.42, fontSize: 13, color: GREEN, fontFace: CODE_FONT });

  addCode(s, 'print(f"\\nFinal score: {score}/{len(questions)}")\nprint(f"Wrong. The correct answer was \'{answer}\'.")', 0.5, 3.82, W - 1, 1.0);
  addBody(s, "f-strings (formatted string literals) are the modern Python way to embed variables in text. Introduced in Python 3.6 — now the standard.", 0.5, 5.05, W - 1, 0.5);
  addAnalogy(s, "An f-string is a Mad Libs template — you write the sentence with blanks, Python fills them in at runtime.");
  addFooter(s, 14);
  s.addNotes("Common mistake: students write {} without the f prefix. The string looks right but variables are never substituted.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 15 — Full Working Code
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "REFERENCE IMPLEMENTATION");
  addTitle(s, "Full Working Code", 0.55);
  addDivider(s, 1.08);

  addCard(s, 0.5, 1.22, 8.3, 5.55, CODE_BG);

  var codeLines = [
    { n: " 1", text: "questions = [",                                                                 color: LIGHT  },
    { n: " 2", text: '    ("What is the capital of France?", "paris"),',                              color: LIGHT  },
    { n: " 3", text: '    ("Which planet is the largest in our solar system?", "jupiter"),',          color: LIGHT  },
    { n: " 4", text: '    ("Which programming language is named after a snake?", "python"),',        color: LIGHT  },
    { n: " 5", text: "]",                                                                             color: LIGHT  },
    { n: " 6", text: "",                                                                              color: LIGHT  },
    { n: " 7", text: "score = 0",                                                                     color: GREEN  },
    { n: " 8", text: "",                                                                              color: LIGHT  },
    { n: " 9", text: "for question, answer in questions:",                                            color: ACCENT1 },
    { n: "10", text: "    user_input = input(question + ' ')",                                        color: LIGHT  },
    { n: "11", text: "    cleaned = user_input.strip().lower()",                                      color: YELLOW },
    { n: "12", text: "    if cleaned == answer:",                                                     color: LIGHT  },
    { n: "13", text: '        print("Correct!")',                                                     color: GREEN  },
    { n: "14", text: "        score += 1",                                                            color: GREEN  },
    { n: "15", text: "    else:",                                                                     color: LIGHT  },
    { n: "16", text: "        print(f\"Wrong. The correct answer was '{answer}'.\")",                 color: RED    },
    { n: "17", text: "",                                                                              color: LIGHT  },
    { n: "18", text: 'print(f"\\nFinal score: {score}/{len(questions)}")',                            color: ACCENT2 },
  ];

  codeLines.forEach(function(line, i) {
    s.addText(line.n, { x: 0.6, y: 1.3 + i * 0.28, w: 0.45, h: 0.26, fontSize: 11, color: DIM, fontFace: CODE_FONT, align: "right" });
    if (line.text) {
      s.addText(line.text, { x: 1.12, y: 1.3 + i * 0.28, w: 7.55, h: 0.26, fontSize: 11.5, color: line.color, fontFace: CODE_FONT });
    }
  });

  // callout annotations
  var callouts = [
    { label: "State Init",   note: "score starts at 0 before the loop",        lineY: 1.3 + 6  * 0.28 + 0.12, color: GREEN   },
    { label: "Sanitization", note: ".strip().lower() normalises every answer",  lineY: 1.3 + 10 * 0.28 + 0.12, color: YELLOW  },
    { label: "Score Update", note: "score += 1 only when cleaned == answer",    lineY: 1.3 + 13 * 0.28 + 0.12, color: ACCENT1 },
  ];

  callouts.forEach(function(c) {
    addCard(s, 9.0, c.lineY - 0.1, 3.85, 0.78, CARD_FILL);
    s.addText(c.label, { x: 9.12, y: c.lineY - 0.04, w: 3.62, h: 0.3, fontSize: 12, bold: true, color: c.color, fontFace: BODY_FONT });
    s.addText(c.note, { x: 9.12, y: c.lineY + 0.26, w: 3.62, h: 0.35, fontSize: 11, color: LIGHT, fontFace: BODY_FONT, wrap: true });
    s.addShape(pptx.ShapeType.line, { x: 8.82, y: c.lineY + 0.12, w: 0.18, h: 0, line: { color: c.color, width: 1 } });
  });

  addFooter(s, 15);
  s.addNotes("Complete, runnable script — 18 lines. Students should type this out, not copy-paste, to build muscle memory. Test in IDLE or the terminal before the session.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 16 — Handling Bad Input
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "DEFENSIVE CODING");
  addTitle(s, "Handling Bad Input", 0.55);
  addDivider(s, 1.08);

  addBody(s, "What happens when a student hits Enter without typing anything? Defensive code handles it gracefully without crashing.", 0.5, 1.22, W - 1, 0.55);

  [
    { input: "''  (empty — just Enter)",         behavior: "strip() -> '' -> does not match answer -> Wrong (safe)", color: YELLOW },
    { input: "'   '  (only spaces)",             behavior: "strip() -> '' -> same as empty -> Wrong (safe)",         color: YELLOW },
    { input: "None  (not possible via input())", behavior: "input() always returns a string — no TypeError risk",    color: GREEN  },
  ].forEach(function(c, i) {
    addCard(s, 0.5, 1.95 + i * 0.95, W - 1, 0.78, CARD_FILL);
    s.addText("Input:", { x: 0.65, y: 2.02 + i * 0.95, w: 0.7, h: 0.28, fontSize: 12, bold: true, color: ACCENT1, fontFace: BODY_FONT });
    s.addText(c.input, { x: 1.4, y: 2.02 + i * 0.95, w: 3.8, h: 0.28, fontSize: 12, color: WHITE, fontFace: CODE_FONT });
    s.addText(c.behavior, { x: 5.4, y: 2.02 + i * 0.95, w: 7.5, h: 0.55, fontSize: 13, color: c.color, fontFace: BODY_FONT, valign: "middle", wrap: true });
  });

  addCode(s,
    'user_input = input(question + " ")\nif user_input.strip() == "":\n    print("No answer entered — marked as incorrect.")\nelse:\n    cleaned = user_input.strip().lower()\n    if cleaned == answer:\n        print("Correct!")\n        score += 1\n    else:\n        print(f"Wrong. The correct answer was \'{answer}\'.")',
    0.5, 4.85, W - 1, 1.62);

  addFooter(s, 16);
  s.addNotes("input() ALWAYS returns a string — no TypeError risk. The only edge case worth handling at this level is empty string. Don't over-engineer with try/except yet.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 17 — Common Mistakes & Debug Tips
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "DEBUG");
  addTitle(s, "Common Mistakes & Debug Tips", 0.55);
  addDivider(s, 1.08);

  var mistakes = [
    {
      bug:  "Forgetting .strip()",
      code: 'if user_input == "paris":  # fails on " paris "',
      fix:  'Use: if user_input.strip().lower() == "paris":',
      why:  "Invisible whitespace is the silent killer. Always strip before compare.",
      color: RED,
    },
    {
      bug:  "Comparing before .lower()",
      code: 'if user_input.strip() == "paris":  # fails on "Paris"',
      fix:  "Apply .lower() after .strip(), then compare.",
      why:  "Case differences cause false negatives. Lower everything, then check.",
      color: YELLOW,
    },
    {
      bug:  "score = 0 inside the loop",
      code: "for q, a in questions:\n    score = 0  # resets every iteration!",
      fix:  "Move score = 0 to BEFORE the for loop.",
      why:  "Score must survive loop iterations. Initialise outside, update inside.",
      color: RED,
    },
    {
      bug:  "Wrong f-string syntax",
      code: '"Score: " + score  # TypeError: can only concatenate str',
      fix:  'Use: f"Score: {score}"',
      why:  "Don't concatenate int + str. Use f-strings for clean embedding.",
      color: YELLOW,
    },
  ];

  mistakes.forEach(function(m, i) {
    var row = Math.floor(i / 2);
    var col = i % 2;
    var x = col === 0 ? 0.5 : 6.78;
    var y = 1.28 + row * 2.6;
    addCard(s, x, y, 6.08, 2.42, CARD_FILL);
    s.addText("! " + m.bug, { x: x + 0.15, y: y + 0.1, w: 5.78, h: 0.35, fontSize: 14, bold: true, color: m.color, fontFace: BODY_FONT });
    addCard(s, x + 0.12, y + 0.52, 5.84, 0.55, CODE_BG);
    s.addText(m.code, { x: x + 0.22, y: y + 0.58, w: 5.64, h: 0.42, fontSize: 11, color: RED, fontFace: CODE_FONT });
    s.addText("Fix: " + m.fix, { x: x + 0.15, y: y + 1.15, w: 5.78, h: 0.4, fontSize: 12, color: GREEN, fontFace: CODE_FONT, wrap: true });
    s.addText(m.why, { x: x + 0.15, y: y + 1.62, w: 5.78, h: 0.62, fontSize: 12, color: DIM, fontFace: BODY_FONT, wrap: true });
  });

  addFooter(s, 17);
  s.addNotes("These four bugs account for ~80% of student errors. Spend extra time on 'score inside the loop' — it's the most confusing because the code looks almost right.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 18 — Where This Logic Actually Runs
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "REAL WORLD");
  addTitle(s, "Where This Logic Actually Runs", 0.55);
  addDivider(s, 1.08);

  var apps = [
    { name: "Google Forms",     desc: "Reads answers, checks against correct responses, scores instantly",           icon: "FORMS" },
    { name: "Kahoot",           desc: "Compares player's choice with stored answer key, tracks per-player score",   icon: "QUIZ"  },
    { name: "HackerRank",       desc: "Reads code output, strips and compares against expected string",             icon: "CODE"  },
    { name: "Google Classroom", desc: "Short-answer grading uses exact-match or keyword-match on student input",    icon: "CLASS" },
    { name: "ATM PIN Check",    desc: "Compares entered PIN (normalised) against stored hash — same if/else gate",  icon: "ATM"   },
  ];

  // emoji replacements for safety
  var icons = ["📋", "🎯", "💻", "🎓", "🏧"];

  apps.forEach(function(a, i) {
    var col = i % 3;
    var row = Math.floor(i / 3);
    var x = 0.5 + col * 4.28;
    var y = 1.3 + row * 2.1;
    addCard(s, x, y, 3.9, 1.88, CARD_FILL);
    s.addText(icons[i], { x: x, y: y + 0.12, w: 3.9, h: 0.52, fontSize: 26, align: "center" });
    s.addText(a.name, { x: x + 0.1, y: y + 0.65, w: 3.7, h: 0.35, fontSize: 15, bold: true, color: ACCENT1, fontFace: TITLE_FONT, align: "center" });
    s.addText(a.desc, { x: x + 0.1, y: y + 1.02, w: 3.7, h: 0.72, fontSize: 12, color: LIGHT, fontFace: BODY_FONT, align: "center", wrap: true });
  });

  addBody(s, "Every input-compare-score loop you've seen in an app uses a variation of exactly what you just built.", 0.5, 5.58, W - 1, 0.42);
  addAnalogy(s, "Your 20-line quiz is the same engine that powers million-user platforms — scaled with databases and APIs, but the logic is identical.");
  addFooter(s, 18);
  s.addNotes("Connecting the homework to real products they use every day creates genuine engagement and retention.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 19 — Practice Exercises + Mini Quiz
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);
  addLabel(s, "EXERCISES");
  addTitle(s, "Practice Exercises & Mini Quiz", 0.55);
  addDivider(s, 1.08);

  var exercises = [
    {
      num: "01", title: "Add a 4th Question",
      desc: 'Add a 4th tuple to questions:\n("What is H2O?", "water")\nExpected: score shows x/4.',
      diff: "Beginner", color: GREEN,
    },
    {
      num: "02", title: "Play Again Loop",
      desc: 'Wrap the quiz in a while loop.\nAfter score, ask:\n"Play again? (yes/no)"\nLoop if yes, exit if no.',
      diff: "Intermediate", color: YELLOW,
    },
    {
      num: "03", title: "Hint System",
      desc: "On wrong answer, offer a hint:\n\"Hint: starts with 'p'. Try again?\"\nAllow one re-attempt per question.",
      diff: "Advanced", color: ACCENT2,
    },
  ];

  exercises.forEach(function(ex, i) {
    var x = 0.5 + i * 4.28;
    addCard(s, x, 1.28, 3.9, 3.6, CARD_FILL);
    s.addText(ex.num, { x: x + 0.15, y: 1.38, w: 3.6, h: 0.52, fontSize: 28, bold: true, color: ex.color, fontFace: TITLE_FONT, align: "center" });
    s.addText(ex.title, { x: x + 0.15, y: 1.95, w: 3.6, h: 0.38, fontSize: 16, bold: true, color: WHITE, fontFace: TITLE_FONT, align: "center" });
    addBadge(s, ex.diff, x + 1.15, 2.42, ex.color === GREEN ? "059669" : (ex.color === YELLOW ? "B45309" : "6D28D9"));
    s.addText(ex.desc, { x: x + 0.15, y: 2.85, w: 3.6, h: 1.85, fontSize: 13, color: LIGHT, fontFace: CODE_FONT, valign: "top", wrap: true, lineSpacing: 20 });
  });

  addCard(s, 0.5, 5.08, W - 1, 1.08, "131929");
  s.addText("Quick Check:", { x: 0.7, y: 5.15, w: 1.8, h: 0.3, fontSize: 13, bold: true, color: ACCENT1, fontFace: BODY_FONT });
  s.addText("1. What does .strip() remove?   2. Where must score = 0 appear?   3. What does the f prefix do?", { x: 2.6, y: 5.15, w: W - 3.3, h: 0.3, fontSize: 13, color: WHITE, fontFace: BODY_FONT });
  s.addText("Answers: 1) leading/trailing whitespace  2) before the for loop  3) enables variable embedding in strings", { x: 0.7, y: 5.5, w: W - 1.4, h: 0.45, fontSize: 12, color: DIM, fontFace: BODY_FONT, wrap: true });

  addFooter(s, 19);
  s.addNotes("Exercise 01 is achievable in 5 minutes. Exercise 02 requires while loops. Exercise 03 introduces nested if — suitable as home assignment.");
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 20 — Recap + Next Steps
// ═══════════════════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  addBg(s);

  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.06, fill: { color: ACCENT1 }, line: { type: "none" } });

  addLabel(s, "WRAP-UP");
  addTitle(s, "Recap & Next Steps", 0.55);
  addDivider(s, 1.08);

  addCard(s, 0.5, 1.22, 6.55, 4.15, CARD_FILL);
  s.addText("Key Takeaways — Project 4", { x: 0.65, y: 1.32, w: 6.25, h: 0.38, fontSize: 16, bold: true, color: ACCENT1, fontFace: TITLE_FONT });
  addBullets(s, [
    "input() captures raw text — always sanitize before comparing",
    ".strip().lower() is your defensive pair — apply on every string input",
    "score = 0 outside the loop; score += 1 inside, only on correct",
    "if/else gives programs the ability to make decisions",
    "f-strings embed variables cleanly — ditch string concatenation",
    "The Question Block pattern scales: 3 questions or 300, same logic",
  ], 0.65, 1.78, 6.2, 3.35);

  addCard(s, 7.35, 1.22, 5.5, 4.15, "1E1B4B");
  s.addText("Project 5 Preview", { x: 7.5, y: 1.32, w: 5.2, h: 0.38, fontSize: 16, bold: true, color: ACCENT2, fontFace: TITLE_FONT });
  addBullets(s, [
    "Functions — wrap Question Block into def ask_question()",
    "Lists & loops — dynamic question sets from a file",
    "try / except — handle non-string inputs properly",
    "Dictionaries — store question, answer, category together",
    "Randomisation — shuffle with random.shuffle()",
  ], 7.5, 1.78, 5.15, 3.35, LIGHT);

  addCard(s, 0.5, 5.55, W - 1, 0.75, "131929");
  s.addText("You now have a working, production-pattern quiz engine in Python. Every concept here is reused directly in Project 5.", {
    x: 0.7, y: 5.63, w: W - 1.4, h: 0.55,
    fontSize: 15, color: WHITE, fontFace: BODY_FONT, valign: "middle", align: "center",
  });

  s.addShape(pptx.ShapeType.rect, { x: 0, y: H - 0.06, w: W, h: 0.06, fill: { color: ACCENT2 }, line: { type: "none" } });
  s.addText("DecodeLabs  ·  Batch 2026  ·  Project 4", { x: 0, y: H - 0.28, w: W, h: 0.22, fontSize: 10, color: DIM, fontFace: BODY_FONT, align: "center" });

  s.addNotes("Final slide. Summarise the six core skills. Set expectations for Project 5 — functions and dictionaries — so students know where today's knowledge leads.");
}

// ─── Write Output ────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: "Python-GK-Quiz-Training-Deck.pptx" })
  .then(function() { console.log("Deck saved: Python-GK-Quiz-Training-Deck.pptx"); })
  .catch(function(err) { console.error("Error:", err); process.exit(1); });
