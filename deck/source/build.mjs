import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, Presentation, PresentationFile } from "@oai/artifact-tool";

const SOURCE_DIR = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(SOURCE_DIR, "../..");
const RENDERED_DIR = path.join(ROOT, "tmp/deck-v2/rendered/artifact-tool");
const PPTX_RENDERED_DIR = path.join(ROOT, "tmp/deck-v2/rendered/pptx");
const OUT = path.join(ROOT, "deck/archive/v2/Nana-Pitch-Deck-v2.pptx");
const BRAND_BUBBLE = path.join(ROOT, "assets/bubbles/nana-glass-bubble-deck.png");
const brandBubbleBuffer = await fs.readFile(BRAND_BUBBLE);
const BRAND_BUBBLE_BYTES = new Uint8Array(
  brandBubbleBuffer.buffer,
  brandBubbleBuffer.byteOffset,
  brandBubbleBuffer.byteLength,
);

const C = {
  cream: "#F6F1E8",
  paper: "#FFFCF7",
  purple: "#5639DD",
  purpleDark: "#3C239F",
  action: "#684CF6",
  ink: "#17151B",
  muted: "#5A5860",
  line: "#D8D0E8",
  grid: "#E8E0F3",
  lilac: "#EDE8FF",
  lilacDeep: "#D8CBFF",
  lavender: "#C5B3FF",
  white: "#FFFFFF",
  mint: "#DDF5E8",
  green: "#2F7D5A",
  blush: "#F8E2E7",
  red: "#B44959",
  gold: "#F4E2A5",
};

const DISPLAY = "Fredoka";
const BODY = "Nunito";
const W = 1280;
const H = 720;
const M = 64;

const presentation = Presentation.create({ slideSize: { width: W, height: H } });

function addShape(slide, geometry, position, fill, options = {}) {
  return slide.shapes.add({
    geometry,
    name: options.name,
    position,
    fill,
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
    ...(options.borderRadius ? { borderRadius: options.borderRadius } : {}),
    ...(options.shadow ? { shadow: options.shadow } : {}),
  });
}

function addText(slide, text, position, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: options.name,
    position,
    fill: options.fill ?? "none",
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
    ...(options.borderRadius ? { borderRadius: options.borderRadius } : {}),
  });
  shape.text = text;
  shape.text.style = {
    fontSize: options.fontSize ?? 24,
    bold: options.bold ?? false,
    color: options.color ?? C.ink,
    typeface: options.typeface ?? BODY,
    alignment: options.align ?? "left",
    verticalAlignment: options.valign ?? "top",
    autoFit: options.autoFit ?? "shrinkText",
    insets: options.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return shape;
}

function addRule(slide, x, y, width, color = C.line, weight = 1) {
  return slide.shapes.add({
    geometry: "line",
    position: { left: x, top: y, width, height: 0 },
    fill: "none",
    line: { style: "solid", fill: color, width: weight },
  });
}

function addGrid(slide, options = {}) {
  const left = options.left ?? 0;
  const top = options.top ?? 0;
  const width = options.width ?? W;
  const height = options.height ?? H;
  const step = options.step ?? 76;
  const color = options.color ?? C.grid;
  for (let x = left; x <= left + width; x += step) {
    slide.shapes.add({
      geometry: "line",
      position: { left: x, top, width: 0, height },
      fill: "none",
      line: { style: "solid", fill: color, width: 1 },
    });
  }
  for (let y = top; y <= top + height; y += step) {
    slide.shapes.add({
      geometry: "line",
      position: { left, top: y, width, height: 0 },
      fill: "none",
      line: { style: "solid", fill: color, width: 1 },
    });
  }
}

function addGlossOrb(slide, x, y, size) {
  return slide.images.add({
    blob: BRAND_BUBBLE_BYTES,
    contentType: "image/png",
    alt: "Nana translucent lavender glass bubble",
    prompt: "A soft translucent lavender glass sphere with a milky upper-left highlight and subtle purple inset shadow, matching Nana's landing-page bubble texture.",
    fit: "contain",
    position: { left: x, top: y, width: size, height: size },
  });
}

function addBase(slide, page, options = {}) {
  slide.background.fill = options.background ?? C.paper;
  if (options.grid !== false) addGrid(slide, { color: options.gridColor ?? C.grid });
  if (options.orbs !== false) {
    addGlossOrb(slide, 1090, 0, 190);
    addGlossOrb(slide, 0, 580, 140);
  }
  addText(slide, "nana", { left: M, top: 30, width: 128, height: 38 }, {
    fontSize: 29,
    bold: true,
    typeface: DISPLAY,
    color: options.brandColor ?? C.ink,
  });
  addText(slide, String(page).padStart(2, "0"), { left: 1180, top: 36, width: 36, height: 20 }, {
    fontSize: 12,
    bold: true,
    color: options.pageColor ?? C.muted,
    align: "right",
  });
}

function addHeader(slide, kicker, title, page, options = {}) {
  addBase(slide, page, options);
  addText(slide, kicker.toUpperCase(), { left: M, top: 90, width: 520, height: 21 }, {
    fontSize: 12,
    bold: true,
    color: options.kickerColor ?? C.purple,
  });
  addText(slide, title, { left: M, top: 122, width: options.titleWidth ?? 1100, height: options.titleHeight ?? 74 }, {
    fontSize: options.titleSize ?? 44,
    bold: true,
    color: options.titleColor ?? C.ink,
    typeface: DISPLAY,
  });
}

function addFooter(slide, label = "NANA · CONFIDENTIAL") {
  addRule(slide, M, 678, W - M * 2, C.line, 1);
  addText(slide, label, { left: M, top: 689, width: 600, height: 14 }, {
    fontSize: 9,
    bold: true,
    color: C.muted,
  });
}

async function bytes(filePath) {
  const buffer = await fs.readFile(filePath);
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

function imageType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".jpg" || ext === ".jpeg" || filePath.includes("/team/")) return "image/jpeg";
  if (ext === ".webp") return "image/webp";
  return "image/png";
}

async function addImage(slide, relativePath, position, options = {}) {
  const filePath = path.join(ROOT, relativePath);
  return slide.images.add({
    blob: await bytes(filePath),
    contentType: options.contentType ?? imageType(filePath),
    alt: options.alt ?? path.basename(filePath),
    fit: options.fit ?? "contain",
    position,
    ...(options.geometry ? { geometry: options.geometry } : {}),
    ...(options.borderRadius ? { borderRadius: options.borderRadius } : {}),
    ...(options.crop ? { crop: options.crop } : {}),
  });
}

function addNotes(slide, lines) {
  slide.speakerNotes.textFrame.setText(`[Sources]\n${lines.map((line) => `- ${line}`).join("\n")}`);
}

function addPill(slide, text, x, y, width, options = {}) {
  addShape(slide, "roundRect", { left: x, top: y, width, height: options.height ?? 38 }, options.fill ?? C.lilac, {
    line: { style: "solid", fill: options.line ?? C.line, width: 1 },
    borderRadius: 18,
  });
  addText(slide, text, { left: x + 12, top: y + 9, width: width - 24, height: 18 }, {
    fontSize: options.fontSize ?? 12,
    bold: true,
    color: options.color ?? C.purpleDark,
    align: "center",
  });
}

function addCard(slide, x, y, width, height, options = {}) {
  return addShape(slide, "roundRect", { left: x, top: y, width, height }, options.fill ?? C.white, {
    line: { style: "solid", fill: options.line ?? C.line, width: options.lineWidth ?? 1 },
    borderRadius: options.radius ?? 26,
    ...(options.shadow === false ? {} : { shadow: options.shadow ?? "shadow-sm" }),
  });
}

function addIconBadge(slide, text, x, y, options = {}) {
  addShape(slide, "ellipse", { left: x, top: y, width: 48, height: 48 }, options.fill ?? C.lilac, {
    line: { style: "solid", fill: options.line ?? C.line, width: 1 },
  });
  addText(slide, text, { left: x, top: y + 8, width: 48, height: 30 }, {
    fontSize: options.fontSize ?? 22,
    bold: true,
    color: options.color ?? C.purple,
    align: "center",
    typeface: DISPLAY,
  });
}

function addArrow(slide, x, y, width, color = C.purple) {
  addRule(slide, x, y, width, color, 3);
  addText(slide, "›", { left: x + width - 13, top: y - 18, width: 22, height: 34 }, {
    fontSize: 32,
    bold: true,
    color,
    typeface: DISPLAY,
    align: "center",
  });
}

// 01: Cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addGrid(slide, { left: 540, top: 0, width: 740, height: H });
  addGlossOrb(slide, 950, 0, 330);
  addGlossOrb(slide, 0, 570, 150);
  addGlossOrb(slide, 640, 72, 570);
  addText(slide, "nana", { left: M, top: 42, width: 170, height: 55 }, {
    fontSize: 43,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addPill(slide, "AGENTIC WALLET", M, 153, 166, { fill: C.lilac });
  addText(slide, "Independence\nthrough voice.", { left: M, top: 220, width: 610, height: 152 }, {
    fontSize: 66,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, "Protection when it matters.", { left: M, top: 390, width: 550, height: 58 }, {
    fontSize: 34,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
  });
  addText(slide, "The agentic wallet for seniors and people with limited mobility.", {
    left: M,
    top: 473,
    width: 535,
    height: 70,
  }, { fontSize: 23, color: C.muted, bold: true });
  addText(slide, "PITCH DECK · SEPTEMBER 2026", { left: M, top: 644, width: 280, height: 18 }, {
    fontSize: 10,
    bold: true,
    color: C.muted,
  });
  await addImage(slide, "assets/sprites/nani-01-lista.png", { left: 728, top: 112, width: 420, height: 540 }, {
    alt: "Nani, the Nana wallet assistant",
  });
  addNotes(slide, [
    "app/apps/nana-wallet/src/components/HybridLandingIntro.tsx",
    "assets/sprites/nani-01-lista.png",
    "docs/business-model/nana-business-model.pdf",
  ]);
}

// 02: Problem
{
  const slide = presentation.slides.add();
  addHeader(slide, "The problem", "Protection today often means taking over.", 2, { titleWidth: 950 });
  addText(slide, "When financial interfaces become difficult, families usually choose between two bad outcomes:", {
    left: M,
    top: 205,
    width: 610,
    height: 62,
  }, { fontSize: 23, color: C.muted });

  const cards = [
    ["01", "Do it for them", "The person loses agency and privacy."],
    ["02", "Leave them alone", "The risk of error, fraud or lockout grows."],
    ["03", "Add another app", "More screens and passwords do not solve the real problem."],
  ];
  for (let i = 0; i < cards.length; i += 1) {
    const y = 305 + i * 102;
    addCard(slide, M, y, 620, 82, { fill: i === 2 ? C.lilac : C.white, shadow: false });
    addText(slide, cards[i][0], { left: 84, top: y + 22, width: 45, height: 24 }, {
      fontSize: 16,
      bold: true,
      color: C.purple,
      typeface: DISPLAY,
    });
    addText(slide, cards[i][1], { left: 145, top: y + 14, width: 190, height: 28 }, {
      fontSize: 21,
      bold: true,
      color: C.ink,
      typeface: DISPLAY,
    });
    addText(slide, cards[i][2], { left: 345, top: y + 15, width: 315, height: 48 }, {
      fontSize: 17,
      color: C.muted,
    });
  }

  addGlossOrb(slide, 750, 210, 410);
  addGlossOrb(slide, 1010, 193, 115);
  addText(slide, "The missing option", { left: 838, top: 300, width: 240, height: 30 }, {
    fontSize: 18,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
    align: "center",
  });
  addText(slide, "Act independently.\nEscalate only\nwhen needed.", { left: 810, top: 349, width: 295, height: 135 }, {
    fontSize: 36,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
    align: "center",
  });
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 2 and 4",
    "docs/business-model/nana-business-model.pdf · problem and product thesis",
  ]);
}

// 03: Why now
{
  const slide = presentation.slides.add();
  addHeader(slide, "Why now", "The need is structural and growing.", 3);
  addText(slide, "A massive population is being asked to navigate an increasingly screen-first financial life.", {
    left: M,
    top: 194,
    width: 810,
    height: 56,
  }, { fontSize: 23, color: C.muted });

  const metrics = [
    ["1 in 6", "people worldwide will be 60+ by 2030", C.lilac],
    ["1.4B", "people aged 60+ by 2030", C.mint],
    ["1.3B", "people live with significant disability today", C.blush],
  ];
  for (let i = 0; i < metrics.length; i += 1) {
    const x = M + i * 382;
    addCard(slide, x, 300, 350, 245, { fill: metrics[i][2], shadow: false, radius: 32 });
    addText(slide, metrics[i][0], { left: x + 28, top: 330, width: 290, height: 74 }, {
      fontSize: 58,
      bold: true,
      color: i === 1 ? C.green : i === 2 ? C.red : C.purple,
      typeface: DISPLAY,
    });
    addText(slide, metrics[i][1], { left: x + 28, top: 422, width: 286, height: 78 }, {
      fontSize: 22,
      bold: true,
      color: C.ink,
    });
  }
  addPill(slide, "OUR WEDGE: SENIORS + LIMITED MOBILITY", 380, 584, 520, {
    fill: C.purple,
    line: C.purple,
    color: C.white,
    height: 44,
    fontSize: 14,
  });
  addFooter(slide);
  addNotes(slide, [
    "https://www.who.int/news-room/fact-sheets/detail/ageing-and-health",
    "https://www.who.int/news-room/fact-sheets/detail/disability-and-health",
    "WHO figures are global and describe population context, not Nana's addressable market.",
  ]);
}

// 04: Product
{
  const slide = presentation.slides.add();
  addHeader(slide, "The product", "Ask. Prepare. Confirm. Settle.", 4);
  addText(slide, "A financial action starts in natural language and ends with explicit user consent.", {
    left: M,
    top: 190,
    width: 760,
    height: 50,
  }, { fontSize: 22, color: C.muted });

  const screens = [
    ["assets/landing/agent-command.jpg", "1", "ASK"],
    ["assets/landing/agent-confirmation.jpg", "2", "PREPARE"],
    ["assets/landing/agent-confirmed.jpg", "3", "CONFIRM"],
    ["assets/landing/agent-home.jpg", "4", "SETTLE"],
  ];
  for (let i = 0; i < screens.length; i += 1) {
    const x = 96 + i * 288;
    addShape(slide, "roundRect", { left: x - 8, top: 259, width: 190, height: 350 }, C.ink, {
      line: { style: "solid", fill: C.ink, width: 0 },
      borderRadius: 28,
      shadow: "shadow-md",
    });
    await addImage(slide, screens[i][0], { left: x, top: 267, width: 174, height: 334 }, {
      fit: "cover",
      geometry: "roundRect",
      borderRadius: 22,
      alt: `Nana product screen: ${screens[i][2].toLowerCase()}`,
    });
    addText(slide, screens[i][1], { left: x - 2, top: 624, width: 24, height: 20 }, {
      fontSize: 13,
      bold: true,
      color: C.purple,
      typeface: DISPLAY,
    });
    addText(slide, screens[i][2], { left: x + 28, top: 624, width: 130, height: 20 }, {
      fontSize: 12,
      bold: true,
      color: C.ink,
    });
    if (i < 3) addArrow(slide, x + 195, 438, 56, C.lavender);
  }
  addFooter(slide);
  addNotes(slide, [
    "assets/landing/agent-command.jpg",
    "assets/landing/agent-confirmation.jpg",
    "assets/landing/agent-confirmed.jpg",
    "assets/landing/agent-home.jpg",
  ]);
}

// 05: Dual user
{
  const slide = presentation.slides.add();
  addHeader(slide, "Dual-user product", "One account. Two distinct roles.", 5);
  addText(slide, "Nana separates the person who acts from the person who can help when something looks wrong.", {
    left: M,
    top: 190,
    width: 860,
    height: 52,
  }, { fontSize: 22, color: C.muted });

  addCard(slide, 74, 278, 470, 306, { fill: C.white, radius: 34 });
  addPill(slide, "END USER", 105, 309, 130, { fill: C.lilac });
  addText(slide, "Acts independently", { left: 105, top: 370, width: 350, height: 38 }, {
    fontSize: 29,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "• Speaks the intent\n• Reviews every consequence\n• Keeps final authority", {
    left: 105,
    top: 431,
    width: 345,
    height: 105,
  }, { fontSize: 21, color: C.muted });

  addCard(slide, 736, 278, 470, 306, { fill: C.lilac, radius: 34 });
  addPill(slide, "RESPONSIBLE", 767, 309, 150, { fill: C.white, line: C.white });
  addText(slide, "Helps by exception", { left: 767, top: 370, width: 360, height: 38 }, {
    fontSize: 29,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "• Receives unusual-event alerts\n• Can co-approve if policy requires\n• Never silently takes control", {
    left: 767,
    top: 431,
    width: 370,
    height: 105,
  }, { fontSize: 21, color: C.muted });

  addGlossOrb(slide, 566, 332, 148);
  addText(slide, "trusted\nrelationship", { left: 580, top: 388, width: 120, height: 50 }, {
    fontSize: 14,
    bold: true,
    color: C.white,
    align: "center",
  });
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slide 6",
    "docs/business-model/nana-business-model.pdf · dual-user model",
    "docs/mission-board/Nana-Mission-Board.pptx · product principles",
  ]);
}

// 06: Exception principle
{
  const slide = presentation.slides.add();
  addHeader(slide, "The product principle", "Intervention happens only by exception.", 6, { titleWidth: 1000 });
  addText(slide, "The user acts independently. Nana pauses only when an action breaks policy or context.", {
    left: M,
    top: 195,
    width: 860,
    height: 55,
  }, { fontSize: 22, color: C.muted });

  addShape(slide, "ellipse", { left: 145, top: 292, width: 690, height: 280 }, C.mint, {
    line: { style: "solid", fill: "#A9E0C3", width: 2 },
  });
  addShape(slide, "ellipse", { left: 632, top: 331, width: 380, height: 205 }, C.blush, {
    line: { style: "solid", fill: "#E7A8B5", width: 2 },
  });
  addText(slide, "DEFAULT", { left: 193, top: 331, width: 120, height: 22 }, {
    fontSize: 12,
    bold: true,
    color: C.green,
  });
  addText(slide, "User acts", { left: 193, top: 372, width: 260, height: 46 }, {
    fontSize: 37,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "Nana prepares, explains and executes\nafter explicit confirmation.", {
    left: 193,
    top: 437,
    width: 360,
    height: 68,
  }, { fontSize: 21, color: C.muted });
  addText(slide, "EXCEPTION", { left: 750, top: 362, width: 130, height: 22 }, {
    fontSize: 12,
    bold: true,
    color: C.red,
    align: "center",
  });
  addText(slide, "Pause + alert", { left: 690, top: 407, width: 250, height: 42 }, {
    fontSize: 29,
    bold: true,
    typeface: DISPLAY,
    align: "center",
  });
  addText(slide, "Nana holds the action and alerts\nthe responsible person.", {
    left: 702,
    top: 468,
    width: 225,
    height: 48,
  }, { fontSize: 18, color: C.muted, align: "center" });
  addGlossOrb(slide, 1025, 424, 105);
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 7 and 8",
    "docs/business-model/nana-business-model.pdf · intervention-by-exception rule",
  ]);
}

// 07: Protection engine
{
  const slide = presentation.slides.add();
  addHeader(slide, "Protection engine", "Normal intent flows. Unusual intent pauses.", 7);
  addText(slide, "The same simple interaction can produce two different safety outcomes.", {
    left: M,
    top: 190,
    width: 760,
    height: 48,
  }, { fontSize: 22, color: C.muted });

  const rows = [
    { y: 285, label: "ORDINARY", fill: C.mint, color: C.green, steps: ["Intent", "Explain", "Confirm", "Execute"] },
    { y: 457, label: "UNUSUAL", fill: C.blush, color: C.red, steps: ["Intent", "Detect", "Pause", "Escalate"] },
  ];
  for (const row of rows) {
    addCard(slide, M, row.y, 1148, 126, { fill: row.fill, shadow: false, radius: 28, line: row.fill });
    addText(slide, row.label, { left: 88, top: row.y + 18, width: 120, height: 20 }, {
      fontSize: 11,
      bold: true,
      color: row.color,
    });
    for (let i = 0; i < row.steps.length; i += 1) {
      const x = 98 + i * 272;
      addIconBadge(slide, String(i + 1), x, row.y + 52, { fill: C.white, color: row.color, line: C.white });
      addText(slide, row.steps[i], { left: x + 62, top: row.y + 64, width: 140, height: 26 }, {
        fontSize: 20,
        bold: true,
        typeface: DISPLAY,
      });
      if (i < row.steps.length - 1) addArrow(slide, x + 192, row.y + 77, 48, row.color);
    }
  }
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slide 8",
    "docs/business-model/nana-business-model.pdf · Protection Engine and policy orchestration",
    "Flow is a product model; concrete thresholds remain subject to user testing and compliance review.",
  ]);
}

// 08: Delegated execution
{
  const slide = presentation.slides.add();
  addHeader(slide, "Delegated execution", "The agent can act inside boundaries the user controls.", 8, {
    titleWidth: 1080,
    titleHeight: 96,
    titleSize: 39,
  });
  addText(slide, "Authority stays with the person. Agent permissions are limited, transparent and revocable.", {
    left: M,
    top: 225,
    width: 880,
    height: 52,
  }, { fontSize: 22, color: C.muted });

  const nodes = [
    ["1", "USER AUTHORITY", "Defines who can do what", C.lilac],
    ["2", "POLICY", "Limits, recipients and escalation", C.white],
    ["3", "NANA AGENT", "Prepares and explains the action", C.lilac],
    ["4", "EXECUTION", "Execute or hold and escalate", C.white],
  ];
  for (let i = 0; i < nodes.length; i += 1) {
    const x = 68 + i * 296;
    addCard(slide, x, 315, 248, 218, { fill: nodes[i][3], radius: 30 });
    addIconBadge(slide, nodes[i][0], x + 24, 341, { fill: i % 2 ? C.lilac : C.white });
    addText(slide, nodes[i][1], { left: x + 24, top: 411, width: 200, height: 26 }, {
      fontSize: 16,
      bold: true,
      color: C.purple,
    });
    addText(slide, nodes[i][2], { left: x + 24, top: 455, width: 200, height: 58 }, {
      fontSize: 20,
      bold: true,
      color: C.ink,
      typeface: DISPLAY,
    });
    if (i < nodes.length - 1) addArrow(slide, x + 250, 424, 43, C.purple);
  }
  addPill(slide, "BOUNDED", 262, 574, 150, { fill: C.ink, line: C.ink, color: C.white });
  addPill(slide, "TRANSPARENT", 445, 574, 170, { fill: C.ink, line: C.ink, color: C.white });
  addPill(slide, "REVOCABLE", 648, 574, 150, { fill: C.ink, line: C.ink, color: C.white });
  addPill(slide, "USER-CONTROLLED", 831, 574, 194, { fill: C.ink, line: C.ink, color: C.white });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · delegated execution architecture",
    "Implementation primitives under evaluation include smart accounts, programmable permissions, session keys and multisig escalation.",
  ]);
}

// 09: Proof
{
  const slide = presentation.slides.add();
  addHeader(slide, "Proof", "The first agentic transfer loop already works.", 9);
  addText(slide, "We built the core interaction, shipped a safety floor and validated the direction in public.", {
    left: M,
    top: 190,
    width: 860,
    height: 50,
  }, { fontSize: 22, color: C.muted });

  const stats = [
    ["7", "AGENTIC TRANSFERS", C.purple, C.lilac],
    ["2nd", "TETHER WDK HACKATHON", C.green, C.mint],
    ["1", "WORKING DUAL-USER LOOP", C.red, C.blush],
  ];
  for (let i = 0; i < stats.length; i += 1) {
    const x = M + i * 275;
    addCard(slide, x, 286, 250, 180, { fill: stats[i][3], shadow: false, radius: 28, line: stats[i][3] });
    addText(slide, stats[i][0], { left: x + 24, top: 308, width: 190, height: 67 }, {
      fontSize: 52,
      bold: true,
      color: stats[i][2],
      typeface: DISPLAY,
    });
    addText(slide, stats[i][1], { left: x + 24, top: 394, width: 195, height: 48 }, {
      fontSize: 13,
      bold: true,
      color: C.ink,
    });
  }

  addShape(slide, "roundRect", { left: 914, top: 252, width: 226, height: 390 }, C.ink, {
    line: { style: "solid", fill: C.ink, width: 0 },
    borderRadius: 30,
    shadow: "shadow-md",
  });
  await addImage(slide, "assets/landing/agent-confirmed.jpg", {
    left: 922,
    top: 260,
    width: 210,
    height: 374,
  }, { fit: "cover", geometry: "roundRect", borderRadius: 24, alt: "Confirmed Nana transaction" });
  addText(slide, "Built, not rendered", { left: M, top: 520, width: 690, height: 38 }, {
    fontSize: 28,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "Voice input · intent parsing · explicit confirmation · receipt · unusual-event controls", {
    left: M,
    top: 572,
    width: 780,
    height: 48,
  }, { fontSize: 19, color: C.muted });
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slide 9",
    "https://github.com/theboyplunger0x/nana-wallet",
    "assets/landing/agent-confirmed.jpg",
    "Hackathon placement and transfer count are team-reported proof points carried forward from the existing deck.",
  ]);
}

// 10: Ownership boundary
{
  const slide = presentation.slides.add();
  addHeader(slide, "Our boundary", "Nana owns the experience and the trust layer.", 10);
  addText(slide, "Infrastructure can evolve without changing the user relationship or product promise.", {
    left: M,
    top: 190,
    width: 780,
    height: 50,
  }, { fontSize: 22, color: C.muted });

  addCard(slide, M, 280, 565, 320, { fill: C.purple, line: C.purple, radius: 36 });
  addText(slide, "NANA OWNS", { left: 100, top: 313, width: 200, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.lilacDeep,
  });
  addText(slide, "The product\npeople trust", { left: 100, top: 357, width: 420, height: 86 }, {
    fontSize: 39,
    bold: true,
    color: C.white,
    typeface: DISPLAY,
  });
  addText(slide, "Agent · voice UX · Family Graph · Protection Engine\npolicies · orchestration · user relationship · brand", {
    left: 100,
    top: 473,
    width: 440,
    height: 82,
  }, { fontSize: 19, color: C.lilac });

  addCard(slide, 674, 280, 542, 320, { fill: C.white, radius: 36 });
  addText(slide, "PARTNERS PROVIDE", { left: 710, top: 313, width: 220, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.purple,
  });
  addText(slide, "The rails beneath it", { left: 710, top: 357, width: 430, height: 49 }, {
    fontSize: 35,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, "Wallet infrastructure · stablecoins · settlement\nfiat rails · KYC · bill pay · liquidity", {
    left: 710,
    top: 446,
    width: 430,
    height: 72,
  }, { fontSize: 20, color: C.muted });
  addPill(slide, "WDK / USDT · PROVEN PATH", 710, 543, 218, { fill: C.mint, color: C.green, line: C.mint });
  addPill(slide, "ARC / USDC · EVALUATING", 946, 543, 225, { fill: C.lilac, color: C.purple, line: C.lilac });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · ownership boundary and infrastructure paths",
    "https://github.com/theboyplunger0x/nana-wallet",
    "Tether/WDK is the proven implementation path; Circle/Arc is an evaluation path, not a committed dependency.",
  ]);
}

// 11: Business model
{
  const slide = presentation.slides.add();
  addHeader(slide, "Business model", "Three hypotheses. One non-negotiable.", 11);
  addText(slide, "We will validate monetization without putting the safety floor behind a paywall.", {
    left: M,
    top: 190,
    width: 820,
    height: 50,
  }, { fontSize: 22, color: C.muted });

  const models = [
    ["01", "Financial activity", "Revenue from transactions and embedded financial flows.", "HYPOTHESIS"],
    ["02", "Nana Family", "Paid coordination, visibility and family-level controls. Pricing unit: TBD.", "HYPOTHESIS"],
    ["03", "B2B2C", "Distribution through institutions serving seniors, families and care networks.", "HYPOTHESIS"],
  ];
  for (let i = 0; i < models.length; i += 1) {
    const x = M + i * 382;
    addCard(slide, x, 286, 350, 264, { fill: i === 1 ? C.lilac : C.white, radius: 30 });
    addText(slide, models[i][0], { left: x + 26, top: 310, width: 48, height: 25 }, {
      fontSize: 18,
      bold: true,
      color: C.purple,
      typeface: DISPLAY,
    });
    addPill(slide, models[i][3], x + 194, 306, 128, { fill: C.paper, line: C.line, fontSize: 10 });
    addText(slide, models[i][1], { left: x + 26, top: 370, width: 295, height: 38 }, {
      fontSize: 27,
      bold: true,
      typeface: DISPLAY,
    });
    addText(slide, models[i][2], { left: x + 26, top: 433, width: 294, height: 90 }, {
      fontSize: 19,
      color: C.muted,
    });
  }
  addShape(slide, "roundRect", { left: 225, top: 582, width: 830, height: 58 }, C.ink, {
    line: { style: "solid", fill: C.ink, width: 0 },
    borderRadius: 26,
  });
  addText(slide, "THE NON-NEGOTIABLE · CORE SAFETY REMAINS FREE", {
    left: 255,
    top: 600,
    width: 770,
    height: 24,
  }, { fontSize: 15, bold: true, color: C.white, align: "center" });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · revenue hypotheses and Nana Free / Nana Family",
    "Pricing and revenue units are intentionally shown as hypotheses pending validation.",
  ]);
}

// 12: Wedge to platform
{
  const slide = presentation.slides.add();
  addHeader(slide, "Expansion", "A focused wedge into a broader financial primitive.", 12, { titleWidth: 1030 });
  addText(slide, "Start where the cost of inaccessible finance is highest, then expand through trusted relationships.", {
    left: M,
    top: 193,
    width: 910,
    height: 52,
  }, { fontSize: 22, color: C.muted });

  addShape(slide, "ellipse", { left: 86, top: 278, width: 1090, height: 340 }, C.lilac, {
    line: { style: "solid", fill: C.lavender, width: 1 },
  });
  addShape(slide, "ellipse", { left: 347, top: 319, width: 790, height: 258 }, C.white, {
    line: { style: "solid", fill: C.line, width: 1 },
  });
  addShape(slide, "ellipse", { left: 692, top: 357, width: 405, height: 184 }, C.purple, {
    line: { style: "solid", fill: C.purple, width: 1 },
  });

  addText(slide, "01", { left: 145, top: 342, width: 50, height: 30 }, {
    fontSize: 21,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
  });
  addText(slide, "Seniors +\nlimited mobility", { left: 145, top: 386, width: 220, height: 70 }, {
    fontSize: 28,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "Accessible action", { left: 145, top: 480, width: 190, height: 28 }, {
    fontSize: 17,
    color: C.muted,
  });

  addText(slide, "02", { left: 445, top: 361, width: 50, height: 30 }, {
    fontSize: 21,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
  });
  addText(slide, "Family finance", { left: 445, top: 407, width: 240, height: 40 }, {
    fontSize: 30,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "Trusted coordination", { left: 445, top: 471, width: 220, height: 28 }, {
    fontSize: 17,
    color: C.muted,
  });

  addText(slide, "03", { left: 778, top: 389, width: 50, height: 30 }, {
    fontSize: 21,
    bold: true,
    color: C.lilacDeep,
    typeface: DISPLAY,
  });
  addText(slide, "Delegated\nfinance", { left: 778, top: 429, width: 220, height: 68 }, {
    fontSize: 30,
    bold: true,
    color: C.white,
    typeface: DISPLAY,
  });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · wedge and expansion sequence",
    "This slide replaces the previous unsupported TAM arithmetic with an explicit product expansion thesis.",
  ]);
}

// 13: GTM
{
  const slide = presentation.slides.add();
  addHeader(slide, "Go to market", "The responsible person opens the door.", 13);
  addText(slide, "Distribution starts with the family member who already helps, without turning them into the account owner.", {
    left: M,
    top: 190,
    width: 900,
    height: 52,
  }, { fontSize: 22, color: C.muted });

  const steps = [
    ["1", "DISCOVER", "A daughter, son or trusted person finds Nana."],
    ["2", "ASSIST", "They help with setup, trusted contacts and preferences."],
    ["3", "STEP BACK", "The end user speaks, confirms and keeps authority."],
  ];
  for (let i = 0; i < steps.length; i += 1) {
    const x = 78 + i * 386;
    addCard(slide, x, 295, 350, 258, { fill: i === 2 ? C.purple : i === 1 ? C.lilac : C.white, radius: 34 });
    addIconBadge(slide, steps[i][0], x + 28, 325, {
      fill: i === 2 ? C.white : C.lilac,
      line: i === 2 ? C.white : C.line,
      color: C.purple,
    });
    addText(slide, steps[i][1], { left: x + 28, top: 398, width: 170, height: 25 }, {
      fontSize: 13,
      bold: true,
      color: i === 2 ? C.lilacDeep : C.purple,
    });
    addText(slide, steps[i][2], { left: x + 28, top: 443, width: 292, height: 82 }, {
      fontSize: 22,
      bold: true,
      color: i === 2 ? C.white : C.ink,
      typeface: DISPLAY,
    });
    if (i < 2) addArrow(slide, x + 351, 426, 34, C.purple);
  }
  addText(slide, "Acquisition by trust. Retention by independence.", { left: 295, top: 599, width: 690, height: 34 }, {
    fontSize: 25,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
    align: "center",
  });
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slide 15",
    "docs/business-model/nana-business-model.pdf · initial distribution thesis",
    "GTM sequence is a hypothesis to validate, not a reported funnel.",
  ]);
}

// 14: Built / next
{
  const slide = presentation.slides.add();
  addHeader(slide, "Execution", "A working core with a focused validation plan.", 14);
  addText(slide, "What exists today is separated from what we need to learn next.", {
    left: M,
    top: 190,
    width: 780,
    height: 48,
  }, { fontSize: 22, color: C.muted });

  addCard(slide, M, 276, 550, 340, { fill: C.mint, line: C.mint, radius: 34, shadow: false });
  addPill(slide, "BUILT", 96, 307, 104, { fill: C.white, line: C.white, color: C.green });
  addText(slide, "Working product core", { left: 96, top: 369, width: 390, height: 40 }, {
    fontSize: 31,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "✓ Voice-led intent\n✓ Agent prepares a transfer\n✓ Explicit confirmation\n✓ Transaction receipt\n✓ Dual-user safety concept", {
    left: 96,
    top: 434,
    width: 410,
    height: 155,
  }, { fontSize: 21, color: C.ink });

  addCard(slide, 666, 276, 550, 340, { fill: C.lilac, line: C.lilac, radius: 34, shadow: false });
  addPill(slide, "NEXT TO VALIDATE", 698, 307, 180, { fill: C.white, line: C.white });
  addText(slide, "The riskiest assumptions", { left: 698, top: 369, width: 420, height: 40 }, {
    fontSize: 31,
    bold: true,
    typeface: DISPLAY,
  });
  addText(slide, "→ Responsible-person acquisition\n→ Family Graph configuration\n→ Escalation thresholds\n→ Nana Family willingness to pay\n→ Infrastructure path by market", {
    left: 698,
    top: 434,
    width: 440,
    height: 155,
  }, { fontSize: 21, color: C.ink });
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 9 and 16",
    "docs/business-model/nana-business-model.pdf · open questions and validation priorities",
    "The previous claim that two experiments were already running has been removed because it was not independently verified.",
  ]);
}

// 15: Team
{
  const slide = presentation.slides.add();
  addHeader(slide, "Team", "Four builders, one product thesis.", 15);
  addText(slide, "Product, mobile, agent flows, wallet infrastructure and voice. Built together in the open.", {
    left: M,
    top: 190,
    width: 850,
    height: 50,
  }, { fontSize: 22, color: C.muted });

  const people = [
    ["assets/team/robertino.png", "Robertino Barbuto", "@rober8b"],
    ["assets/team/ramiro.png", "Ramiro Carnicer Souble", "@ram4-dev"],
    ["assets/team/ignacio.png", "Ignacio Becerra", "@BecerraIgnacio"],
    ["assets/team/marcos-professional.png", "Marcos Lanzani", "@theboyplunger0x"],
  ];
  for (let i = 0; i < people.length; i += 1) {
    const x = 66 + i * 299;
    addCard(slide, x, 276, 266, 331, { fill: i % 2 ? C.lilac : C.white, radius: 30 });
    await addImage(slide, people[i][0], { left: x + 25, top: 299, width: 216, height: 185 }, {
      fit: "cover",
      geometry: "roundRect",
      borderRadius: 24,
      alt: `${people[i][1]} portrait`,
      contentType: i === 3 ? "image/png" : "image/jpeg",
      ...(i === 3 ? { crop: { left: 0, top: 0, right: 0, bottom: 0.35 } } : {}),
    });
    addText(slide, people[i][1], { left: x + 24, top: 515, width: 218, height: 47 }, {
      fontSize: 21,
      bold: true,
      typeface: DISPLAY,
      align: "center",
    });
    addText(slide, people[i][2], { left: x + 24, top: 568, width: 218, height: 22 }, {
      fontSize: 13,
      bold: true,
      color: C.purple,
      align: "center",
    });
  }
  addFooter(slide);
  addNotes(slide, [
    "https://github.com/theboyplunger0x/nana-wallet",
    "https://github.com/rober8b",
    "https://github.com/ram4-dev",
    "https://github.com/BecerraIgnacio",
    "https://github.com/theboyplunger0x",
    "Role labels were intentionally omitted rather than inferred from limited public evidence.",
  ]);
}

// 16: Close
{
  const slide = presentation.slides.add();
  slide.background.fill = C.purple;
  addGrid(slide, { color: "#6E58E8" });
  addGlossOrb(slide, 1000, 0, 280);
  addGlossOrb(slide, 700, 95, 465);
  addText(slide, "nana", { left: M, top: 39, width: 155, height: 50 }, {
    fontSize: 40,
    bold: true,
    color: C.white,
    typeface: DISPLAY,
  });
  addText(slide, "The next financial\naccount is agentic.", { left: M, top: 148, width: 650, height: 134 }, {
    fontSize: 59,
    bold: true,
    color: C.white,
    typeface: DISPLAY,
  });
  addText(slide, "It lets people act independently inside a trusted, programmable network of relationships and permissions.", {
    left: M,
    top: 322,
    width: 600,
    height: 95,
  }, { fontSize: 25, bold: true, color: C.lilac });
  addShape(slide, "roundRect", { left: M, top: 495, width: 535, height: 94 }, C.white, {
    line: { style: "solid", fill: C.white, width: 0 },
    borderRadius: 30,
  });
  addText(slide, "Independence through voice.\nProtection when it matters.", {
    left: 93,
    top: 518,
    width: 475,
    height: 54,
  }, { fontSize: 24, bold: true, color: C.purple, typeface: DISPLAY, align: "center" });
  addText(slide, "nana-wallet-hybrid.vercel.app", { left: M, top: 653, width: 280, height: 22 }, {
    fontSize: 14,
    bold: true,
    color: C.white,
  });
  addText(slide, "github.com/theboyplunger0x/nana-wallet", { left: 340, top: 653, width: 470, height: 22 }, {
    fontSize: 14,
    bold: true,
    color: C.lilac,
  });
  await addImage(slide, "assets/sprites/nani-02-escuchando.png", { left: 733, top: 117, width: 402, height: 525 }, {
    alt: "Nani listening",
  });
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · final product primitive",
    "assets/sprites/nani-02-escuchando.png",
    "https://nana-wallet-hybrid.vercel.app",
    "https://github.com/theboyplunger0x/nana-wallet",
  ]);
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(path.dirname(OUT), { recursive: true });
await fs.mkdir(RENDERED_DIR, { recursive: true });

for (const [index, slide] of presentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(path.join(RENDERED_DIR, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(RENDERED_DIR, `${stem}.layout.json`), await layout.text());
}

await writeBlob(path.join(RENDERED_DIR, "montage.webp"), await presentation.export({ format: "webp", montage: true, scale: 1 }));

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(OUT);

const regenerated = await PresentationFile.importPptx(await FileBlob.load(OUT));
await fs.mkdir(PPTX_RENDERED_DIR, { recursive: true });
for (const [index, slide] of regenerated.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(path.join(PPTX_RENDERED_DIR, `${stem}.png`), await regenerated.export({ slide, format: "png", scale: 1 }));
}

const snapshot = await presentation.inspect({ kind: "slide,textbox,shape,image,notes", maxChars: 24000 });
await fs.writeFile(path.join(RENDERED_DIR, "inspect.ndjson"), snapshot.ndjson);

console.log(OUT);
