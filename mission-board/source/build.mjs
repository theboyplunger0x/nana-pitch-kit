import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, Presentation, PresentationFile } from "@oai/artifact-tool";

const SOURCE_DIR = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(SOURCE_DIR, "../../..");
const RENDERED_DIR = path.join(ROOT, "docs/mission-board/preview/artifact-tool");
const PPTX_RENDERED_DIR = path.join(ROOT, "docs/mission-board/preview/slides");
const OUT = path.join(ROOT, "docs/mission-board/Nana-Mission-Board.pptx");

const C = {
  cream: "#F6F1E8",
  paper: "#FFFCF7",
  purple: "#5639DD",
  action: "#684CF6",
  purpleDark: "#3C239F",
  purpleInk: "#4631BA",
  lilac: "#EDE8FF",
  lilacDeep: "#D8CBFF",
  glow: "#A996FF",
  ink: "#17151B",
  muted: "#5A5860",
  border: "#D8D0E8",
  grid: "#E8E0F3",
  success: "#2F7D5A",
  successSoft: "#E5F6EC",
  alert: "#B44959",
  alertSoft: "#FBE9EC",
  white: "#FFFFFF",
};

const DISPLAY = "Fredoka";
const BODY = "Nunito";
const SLIDE_W = 1280;
const SLIDE_H = 720;
const M = 68;

const presentation = Presentation.create({
  slideSize: { width: SLIDE_W, height: SLIDE_H },
});

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

function addRule(slide, x, y, w, color = C.border, weight = 1) {
  return slide.shapes.add({
    geometry: "line",
    position: { left: x, top: y, width: w, height: 0 },
    fill: "none",
    line: { style: "solid", fill: color, width: weight },
  });
}

function addGrid(slide, options = {}) {
  const left = options.left ?? 0;
  const top = options.top ?? 0;
  const width = options.width ?? SLIDE_W;
  const height = options.height ?? SLIDE_H;
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

function addOrb(slide, position, fill = C.lilacDeep, line = C.border) {
  return addShape(slide, "ellipse", position, fill, {
    line: { style: "solid", fill: line, width: 1 },
  });
}

function addHeader(slide, kicker, title, page) {
  addText(slide, kicker.toUpperCase(), { left: M, top: 38, width: 440, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.purple,
    typeface: BODY,
    name: `kicker-${page}`,
  });
  addText(slide, title, { left: M, top: 70, width: 1080, height: 76 }, {
    fontSize: 47,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
    name: `title-${page}`,
  });
  addText(slide, String(page).padStart(2, "0"), { left: 1170, top: 42, width: 42, height: 20 }, {
    fontSize: 13,
    bold: true,
    color: C.muted,
    align: "right",
  });
}

function addFooter(slide, text = "NANA MISSION BOARD · INTERNAL WORKING SYSTEM") {
  addRule(slide, M, 680, SLIDE_W - M * 2, C.border, 1);
  addText(slide, text, { left: M, top: 690, width: 760, height: 18 }, {
    fontSize: 10,
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
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".webp") return "image/webp";
  return "image/png";
}

async function addImage(slide, relativePath, position, options = {}) {
  const filePath = path.join(ROOT, relativePath);
  return slide.images.add({
    blob: await bytes(filePath),
    contentType: imageType(filePath),
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

function addNumberedPrinciple(slide, number, title, body, x, y, width) {
  addText(slide, String(number).padStart(2, "0"), { left: x, top: y, width: 54, height: 28 }, {
    fontSize: 22,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
  });
  addText(slide, title, { left: x + 66, top: y - 2, width: width - 66, height: 30 }, {
    fontSize: 23,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, body, { left: x + 66, top: y + 32, width: width - 66, height: 58 }, {
    fontSize: 20,
    color: C.muted,
  });
  addRule(slide, x + 66, y + 98, width - 66, C.border, 1);
}

// 01: Cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addGrid(slide, { left: 520, top: 0, width: 760, height: 720 });
  addOrb(slide, { left: 910, top: 0, width: 370, height: 370 }, C.lilacDeep, C.glow);
  addOrb(slide, { left: 0, top: 570, width: 150, height: 150 }, C.lilac, C.border);

  addText(slide, "nana", { left: M, top: 54, width: 260, height: 74 }, {
    fontSize: 54,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, "Mission Board", { left: M, top: 210, width: 650, height: 96 }, {
    fontSize: 72,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, "The internal system for mission, vision, product and brand.", {
    left: M,
    top: 326,
    width: 600,
    height: 78,
  }, { fontSize: 25, bold: true, color: C.muted });
  addText(slide, "Independence through voice.\nProtection when it matters.", {
    left: M,
    top: 492,
    width: 610,
    height: 92,
  }, { fontSize: 32, bold: true, color: C.purple, typeface: DISPLAY });
  addText(slide, "V0.1 · AUGUST 2026", { left: M, top: 650, width: 240, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.muted,
  });
  await addImage(slide, "assets/sprites/nani-01-lista.png", {
    left: 765,
    top: 155,
    width: 430,
    height: 545,
  }, { alt: "Nani, Nana's wallet assistant" });
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 1 and 15",
    "app/apps/nana-wallet/src/components/HybridLandingIntro.tsx",
    "assets/sprites/nani-01-lista.png",
  ]);
}

// 02: Mission board
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addHeader(slide, "The board", "The system, on one page.", 2);

  addRule(slide, M, 166, 1144, C.border, 1);
  addRule(slide, 444, 185, 0, C.border, 1);
  addRule(slide, 840, 185, 0, C.border, 1);

  const columns = [M, 470, 865];
  const widths = [340, 330, 345];
  const labels = ["MISSION", "VISION", "PRODUCT"];
  const statements = [
    "Return financial agency to people excluded by digital interfaces, without asking families to surrender peace of mind.",
    "Financial independence should not expire with age, mobility or interface complexity.",
    "A voice-first agentic financial account with a dual-user model: independence for the person using it, protection for the person who cares for them.",
  ];
  for (let i = 0; i < 3; i += 1) {
    addText(slide, labels[i], { left: columns[i], top: 194, width: widths[i], height: 22 }, {
      fontSize: 12,
      bold: true,
      color: C.purple,
    });
    addText(slide, statements[i], { left: columns[i], top: 234, width: widths[i], height: 230 }, {
      fontSize: i === 2 ? 25 : 28,
      bold: true,
      color: C.ink,
      typeface: DISPLAY,
    });
  }

  addShape(slide, "roundRect", { left: M, top: 505, width: 1144, height: 122 }, C.lilac, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 28,
  });
  addText(slide, "PROMISE", { left: 96, top: 527, width: 150, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Independence through voice. Protection when it matters.", {
    left: 96,
    top: 553,
    width: 660,
    height: 44,
  }, { fontSize: 27, bold: true, color: C.purpleDark, typeface: DISPLAY });
  addText(slide, "NORTH-STAR RULE", { left: 815, top: 527, width: 180, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Family intervention happens by exception only.", {
    left: 815,
    top: 553,
    width: 355,
    height: 48,
  }, { fontSize: 22, bold: true, color: C.ink, typeface: DISPLAY });
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 1, 6 and 7",
    "docs/business-model/nana-business-model.pdf · pages 1-2",
    "Mission and vision are a working synthesis of the team's existing thesis.",
  ]);
}

// 03: Product architecture
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addGrid(slide, { left: 575, top: 0, width: 705, height: 720 });
  addOrb(slide, { left: 765, top: 120, width: 470, height: 470 }, C.lilacDeep, C.glow);
  addHeader(slide, "Product architecture", "Two users. One account. One owner.", 3);

  addText(slide, "THE END USER", { left: M, top: 204, width: 260, height: 20 }, {
    fontSize: 12,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Speaks. Reviews. Confirms.\nPays independently.", {
    left: M,
    top: 238,
    width: 430,
    height: 92,
  }, { fontSize: 30, bold: true, color: C.ink, typeface: DISPLAY });
  addText(slide, "The account stays in their name. Daily use is a conversation, not a dashboard.", {
    left: M,
    top: 342,
    width: 430,
    height: 72,
  }, { fontSize: 21, color: C.muted });

  addText(slide, "THE RESPONSIBLE", { left: M, top: 450, width: 260, height: 20 }, {
    fontSize: 12,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Sets up once.\nIntervenes only by exception.", {
    left: M,
    top: 484,
    width: 430,
    height: 88,
  }, { fontSize: 28, bold: true, color: C.ink, typeface: DISPLAY });
  addText(slide, "KYC, funding, family graph and limits happen once, before daily use.", {
    left: M,
    top: 584,
    width: 440,
    height: 56,
  }, { fontSize: 21, color: C.muted });

  addShape(slide, "roundRect", { left: 665, top: 206, width: 470, height: 370 }, C.paper, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 36,
    shadow: "shadow-lg",
  });
  addText(slide, "nana", { left: 770, top: 256, width: 260, height: 70 }, {
    fontSize: 60,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
    align: "center",
  });
  addText(slide, "VOICE AGENT\nWALLET\nPROTECTION\nFAMILY GRAPH", {
    left: 760,
    top: 354,
    width: 280,
    height: 132,
  }, { fontSize: 18, bold: true, color: C.ink, align: "center" });
  addText(slide, "Rails stay invisible", { left: 760, top: 516, width: 280, height: 26 }, {
    fontSize: 15,
    bold: true,
    color: C.muted,
    align: "center",
  });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · pages 2-4",
    "docs/deck/nana-deck.pdf · slide 6",
  ]);
}

// 04: Interaction model
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addHeader(slide, "Interaction model", "The interface becomes a conversation.", 4);

  addShape(slide, "roundRect", { left: M, top: 172, width: 1144, height: 372 }, C.cream, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 30,
  });
  await addImage(slide, "assets/social/tweet5-flow-355KB.png", {
    left: 90,
    top: 190,
    width: 1100,
    height: 334,
  }, { alt: "Nana's ask and confirm product flow" });
  addText(slide, "The agent does the work. The user makes the one decision that carries consequence.", {
    left: 160,
    top: 578,
    width: 960,
    height: 58,
  }, { fontSize: 28, bold: true, color: C.ink, typeface: DISPLAY, align: "center" });
  addFooter(slide, "PRODUCT LOOP · ASK → PREPARE → CONFIRM → SETTLE");
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slide 5",
    "docs/business-model/nana-business-model.pdf · page 4",
    "assets/social/tweet5-flow-355KB.png",
  ]);
}

// 05: Protection doctrine
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addShape(slide, "roundRect", { left: 28, top: 28, width: 1224, height: 664 }, C.purple, {
    line: { style: "solid", fill: C.purple, width: 0 },
    borderRadius: 38,
  });
  addOrb(slide, { left: 1000, top: 28, width: 252, height: 252 }, C.action, C.glow);
  addText(slide, "THE PROTECTION DOCTRINE", { left: 78, top: 70, width: 360, height: 22 }, {
    fontSize: 12,
    bold: true,
    color: C.lilacDeep,
  });
  addText(slide, "Protection intervenes\nonly when it matters.", {
    left: 78,
    top: 116,
    width: 620,
    height: 118,
  }, { fontSize: 52, bold: true, color: C.white, typeface: DISPLAY });

  addText(slide, "The family must not approve every payment.", {
    left: 78,
    top: 286,
    width: 540,
    height: 68,
  }, { fontSize: 31, bold: true, color: C.white, typeface: DISPLAY });
  addText(slide, "The moment they do, independence was not returned. It was transferred.", {
    left: 78,
    top: 382,
    width: 540,
    height: 102,
  }, { fontSize: 31, bold: true, color: C.lilacDeep, typeface: DISPLAY });
  addText(slide, "SAFETY IS NEVER THE UPSELL.", { left: 78, top: 610, width: 520, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.white,
  });

  addShape(slide, "roundRect", { left: 730, top: 150, width: 440, height: 430 }, C.paper, {
    line: { style: "solid", fill: C.lilacDeep, width: 1 },
    borderRadius: 34,
    shadow: "shadow-xl",
  });
  await addImage(slide, "assets/screenshots-english/03-agent-confirmation.jpg", {
    left: 852,
    top: 174,
    width: 196,
    height: 380,
  }, { alt: "Nana confirmation screen" });
  addText(slide, "Ordinary", { left: 754, top: 194, width: 88, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.success,
  });
  addText(slide, "User confirms.\nFamily stays out.", { left: 754, top: 236, width: 100, height: 100 }, {
    fontSize: 20,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, "Unusual", { left: 1062, top: 194, width: 88, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.alert,
    align: "right",
  });
  addText(slide, "Held first.\nFamily alerted.", { left: 1052, top: 236, width: 100, height: 80 }, {
    fontSize: 20,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
    align: "right",
  });
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 7-8",
    "docs/business-model/nana-business-model.pdf · page 5",
    "assets/screenshots-english/03-agent-confirmation.jpg",
  ]);
}

// 06: Product principles
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addHeader(slide, "Product doctrine", "Six rules that survive every roadmap.", 6);

  const leftX = M;
  const rightX = 656;
  const colW = 540;
  addNumberedPrinciple(slide, 1, "Agency by default", "The user owns the account, initiates actions and confirms consequence.", leftX, 188, colW);
  addNumberedPrinciple(slide, 2, "Intervention by exception", "Family enters only when activity is unusual or crosses a policy.", leftX, 336, colW);
  addNumberedPrinciple(slide, 3, "Complexity stays invisible", "Addresses, rails, gas and settlement never become daily-user tasks.", leftX, 484, colW);
  addNumberedPrinciple(slide, 4, "Plain language first", "One action, one consequence, no forms and no crypto vocabulary.", rightX, 188, colW);
  addNumberedPrinciple(slide, 5, "Safety is universal", "Core confirmation, limits, anomaly detection and protection stay free.", rightX, 336, colW);
  addNumberedPrinciple(slide, 6, "Fail closed and explain", "If nobody responds, the held action expires safely and Nana explains why.", rightX, 484, colW);
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · pages 2-6",
    "docs/business-model/nana-modelo.pdf · pages 4-7",
    "app/apps/nana-wallet/src/routes/index.tsx",
  ]);
}

// 07: Audience and message system
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addHeader(slide, "Audience and message", "Design for the user. Acquire the responsible.", 7);

  await addImage(slide, "assets/social/tweet6-users-331KB.png", {
    left: M,
    top: 174,
    width: 600,
    height: 286,
  }, { alt: "Nana's dual-user model" });
  addShape(slide, "roundRect", { left: 705, top: 174, width: 507, height: 424 }, C.cream, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 30,
  });
  const messages = [
    ["WEDGE", "Start with seniors and people with limited mobility."],
    ["PROMISE", "Independence through voice. Protection when it matters."],
    ["PRODUCT", "Say what you want. Confirm what matters."],
    ["FAMILY", "Give your parents independence without giving up peace of mind."],
    ["EXPANSION", "Families · minors · delegated finance"],
  ];
  for (let i = 0; i < messages.length; i += 1) {
    const y = 198 + i * 76;
    addText(slide, messages[i][0], { left: 736, top: y, width: 110, height: 18 }, {
      fontSize: 10,
      bold: true,
      color: C.purple,
    });
    addText(slide, messages[i][1], { left: 736, top: y + 22, width: 438, height: 46 }, {
      fontSize: 20,
      bold: true,
      color: C.ink,
      typeface: DISPLAY,
    });
    if (i < messages.length - 1) addRule(slide, 736, y + 68, 438, C.border, 1);
  }
  addText(slide, "DO NOT FRAME NANA AS", { left: M, top: 508, width: 220, height: 18 }, {
    fontSize: 10,
    bold: true,
    color: C.alert,
  });
  addText(slide, "a crypto wallet · parental controls · pay-to-be-safe · technology simplified for old people", {
    left: M,
    top: 540,
    width: 600,
    height: 62,
  }, { fontSize: 21, bold: true, color: C.muted, typeface: DISPLAY });
  addFooter(slide);
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 1, 6 and 12",
    "docs/business-model/nana-business-model.pdf · page 8",
    "assets/social/tweet6-users-331KB.png",
  ]);
}

// 08: Visual design system
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addGrid(slide, { left: 780, top: 0, width: 500, height: 720 });
  addOrb(slide, { left: 970, top: 0, width: 310, height: 310 }, C.lilacDeep, C.glow);
  addHeader(slide, "Visual system", "Calm for money. Warm for family.", 8);

  addText(slide, "COLOR", { left: M, top: 190, width: 120, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.purple,
  });
  const swatches = [
    [C.purple, "Nana Purple", "#5639DD"],
    [C.action, "Action Purple", "#684CF6"],
    [C.cream, "Cream", "#F6F1E8"],
    [C.ink, "Ink", "#17151B"],
    [C.lilac, "Lilac", "#EDE8FF"],
  ];
  for (let i = 0; i < swatches.length; i += 1) {
    const x = M + i * 112;
    addShape(slide, "roundRect", { left: x, top: 226, width: 92, height: 68 }, swatches[i][0], {
      line: { style: "solid", fill: C.border, width: 1 },
      borderRadius: 18,
    });
    addText(slide, swatches[i][1], { left: x, top: 306, width: 100, height: 20 }, {
      fontSize: 11,
      bold: true,
      color: C.ink,
    });
    addText(slide, swatches[i][2], { left: x, top: 328, width: 100, height: 18 }, {
      fontSize: 10,
      color: C.muted,
    });
  }

  addText(slide, "TYPE", { left: M, top: 394, width: 120, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Fredoka", { left: M, top: 428, width: 240, height: 52 }, {
    fontSize: 40,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, "Rounded display voice for the brand, claims and decisions.", {
    left: M,
    top: 482,
    width: 280,
    height: 48,
  }, { fontSize: 20, color: C.muted });
  addText(slide, "Nunito", { left: 370, top: 428, width: 210, height: 52 }, {
    fontSize: 38,
    bold: true,
    color: C.ink,
    typeface: BODY,
  });
  addText(slide, "Plain, generous body copy for product and explanation.", {
    left: 370,
    top: 482,
    width: 280,
    height: 48,
  }, { fontSize: 20, color: C.muted });

  addText(slide, "MOTIFS", { left: M, top: 566, width: 120, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.purple,
  });
  addText(slide, "soft spheres · faint grid · glass rails · large rounded surfaces · one strong action", {
    left: M,
    top: 596,
    width: 620,
    height: 44,
  }, { fontSize: 20, bold: true, color: C.ink, typeface: DISPLAY });

  await addImage(slide, "assets/sprites/nani-02-escuchando.png", {
    left: 770,
    top: 238,
    width: 215,
    height: 388,
  }, { alt: "Nani listening" });
  await addImage(slide, "assets/sprites/nani-03-pensando.png", {
    left: 968,
    top: 254,
    width: 220,
    height: 370,
  }, { alt: "Nani thinking" });
  addText(slide, "Nani is a calm guide, never a childish mascot.", {
    left: 770,
    top: 622,
    width: 430,
    height: 30,
  }, { fontSize: 17, bold: true, color: C.ink, align: "center" });
  addFooter(slide);
  addNotes(slide, [
    "app/apps/nana-wallet/src/styles.css",
    "video/src/theme.ts",
    "assets/sprites/nani-02-escuchando.png",
    "assets/sprites/nani-03-pensando.png",
  ]);
}

// 09: What is true vs. hypothesis
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addHeader(slide, "Company truth", "Protect the thesis. Test the business model.", 9);

  addShape(slide, "roundRect", { left: M, top: 180, width: 535, height: 345 }, C.successSoft, {
    line: { style: "solid", fill: "#B8DDCB", width: 1 },
    borderRadius: 30,
  });
  addShape(slide, "roundRect", { left: 677, top: 180, width: 535, height: 365 }, C.lilac, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 30,
  });
  addText(slide, "DEFINED", { left: 98, top: 208, width: 170, height: 22 }, {
    fontSize: 12,
    bold: true,
    color: C.success,
  });
  addText(slide, "The product cannot change these without becoming something else.", {
    left: 98,
    top: 246,
    width: 450,
    height: 58,
  }, { fontSize: 22, bold: true, color: C.ink, typeface: DISPLAY });
  addText(slide, "• Dual-user architecture\n• User owns and confirms\n• Intervention by exception\n• Safety stays free\n• GTM through the responsible\n• Voice-first intent-to-action loop", {
    left: 98,
    top: 326,
    width: 450,
    height: 170,
  }, { fontSize: 22, bold: true, color: C.ink });

  addText(slide, "TESTING", { left: 707, top: 208, width: 170, height: 22 }, {
    fontSize: 12,
    bold: true,
    color: C.purple,
  });
  addText(slide, "LOW / MEDIUM = current confidence today, not priority or opportunity size.", {
    left: 707,
    top: 230,
    width: 450,
    height: 18,
  }, { fontSize: 11, bold: true, color: C.muted });
  addText(slide, "These are hypotheses, not identity statements or promises.", {
    left: 707,
    top: 260,
    width: 450,
    height: 52,
  }, { fontSize: 22, bold: true, color: C.ink, typeface: DISPLAY });
  const testingItems = [
    ["Transaction economics", "LOW"],
    ["Nana Family willingness to pay", "MEDIUM"],
    ["B2B2C demand", "LOW"],
    ["Production custody + compliance", "LOW"],
    ["Risk thresholds + slow scams", "LOW"],
    ["Family Graph growth loop", "LOW"],
  ];
  testingItems.forEach(([hypothesis, confidence], index) => {
    const y = 330 + index * 31;
    addText(slide, `• ${hypothesis}`, {
      left: 707,
      top: y,
      width: 348,
      height: 27,
    }, { fontSize: 19, bold: true, color: C.ink });
    // MEDIUM no entra a 19pt en Fredoka/Nunito reales: se parte en dos lineas
    addText(slide, confidence, {
      left: 1055,
      top: y + 2,
      width: 102,
      height: 25,
    }, { fontSize: 16, bold: true, color: C.ink, align: "right" });
  });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · pages 6-9",
    "docs/deck/nana-deck.pdf · slides 9, 11 and 13",
  ]);
}

// 10: Problem cascade
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addHeader(slide, "The protection trap", "Protection became control.", 10);
  addText(slide, "Financial products are built around interfaces, not people.", {
    left: M,
    top: 160,
    width: 980,
    height: 44,
  }, { fontSize: 24, bold: true, color: C.muted });

  const steps = [
    ["01", "Can't navigate\nbanking"],
    ["02", "Family\nsteps in"],
    ["03", "Family pays\nand moves money"],
    ["04", "Family controls\nthe account"],
    ["05", "Less\nrisk"],
    ["06", "Less\nindependence"],
  ];
  for (let i = 0; i < steps.length; i += 1) {
    const x = M + i * 189;
    const isLast = i === steps.length - 1;
    addShape(slide, "roundRect", { left: x, top: 246, width: 160, height: 142 }, isLast ? C.alertSoft : C.paper, {
      line: { style: "solid", fill: isLast ? C.alert : C.border, width: 1 },
      borderRadius: 24,
    });
    addText(slide, steps[i][0], { left: x + 18, top: 266, width: 45, height: 20 }, {
      fontSize: 12,
      bold: true,
      color: isLast ? C.alert : C.purple,
    });
    addText(slide, steps[i][1], { left: x + 18, top: 304, width: 124, height: 64 }, {
      fontSize: 21,
      bold: true,
      color: C.ink,
      typeface: DISPLAY,
      align: "center",
    });
    if (!isLast) {
      addText(slide, "→", { left: x + 163, top: 292, width: 26, height: 38 }, {
        fontSize: 24,
        bold: true,
        color: C.purple,
        typeface: DISPLAY,
        align: "center",
      });
    }
  }

  addShape(slide, "roundRect", { left: 226, top: 476, width: 828, height: 118 }, C.purple, {
    line: { style: "solid", fill: C.purple, width: 0 },
    borderRadius: 30,
    shadow: "shadow-lg",
  });
  addText(slide, "NANA BREAKS THE CHAIN", { left: 270, top: 500, width: 230, height: 20 }, {
    fontSize: 12,
    bold: true,
    color: C.lilacDeep,
  });
  addText(slide, "Protection without taking control.", {
    left: 270,
    top: 532,
    width: 740,
    height: 42,
  }, { fontSize: 29, bold: true, color: C.white, typeface: DISPLAY });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · pages 1-3",
    "/Users/lanzanimarcos7/.codex/attachments/0c94102c-c9bf-41e2-9d0d-65a4ef170350/pasted-text.txt",
    "Problem cascade phrasing is a synthesis of the team's existing problem statement.",
  ]);
}

// 11: Family graph
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addGrid(slide, { left: 590, top: 0, width: 690, height: 720 });
  addHeader(slide, "Strategic system", "The family is a graph, not a permission queue.", 11);

  addText(slide, "BANKS HAVE ACCOUNTS. NANA UNDERSTANDS FAMILIES.", {
    left: M,
    top: 190,
    width: 430,
    height: 42,
  }, { fontSize: 15, bold: true, color: C.purple });
  addText(slide, "Relationships carry meaning.", {
    left: M,
    top: 246,
    width: 430,
    height: 48,
  }, { fontSize: 32, bold: true, color: C.ink, typeface: DISPLAY });
  addText(slide, "Responsible, alerts, trusted recipient and protected person are distinct roles, not a chain of approvals.", {
    left: M,
    top: 316,
    width: 430,
    height: 100,
  }, { fontSize: 22, color: C.muted });

  const centerX = 895;
  const centerY = 368;
  const nodes = [
    { x: 648, y: 205, w: 196, h: 104, name: "LAURA", role: "responsible · admin" },
    { x: 1024, y: 205, w: 196, h: 104, name: "MARCOS", role: "alerts" },
    { x: 648, y: 492, w: 196, h: 104, name: "TOMI", role: "trusted recipient" },
    { x: 1024, y: 492, w: 196, h: 104, name: "PAPÁ", role: "protected account" },
  ];
  for (const node of nodes) {
    addShape(slide, "roundRect", { left: node.x, top: node.y, width: node.w, height: node.h }, C.cream, {
      line: { style: "solid", fill: C.border, width: 1 },
      borderRadius: 24,
    });
    addText(slide, node.name, { left: node.x + 18, top: node.y + 18, width: node.w - 36, height: 22 }, {
      fontSize: 13,
      bold: true,
      color: C.purple,
      align: "center",
    });
    addText(slide, node.role, { left: node.x + 16, top: node.y + 52, width: node.w - 32, height: 30 }, {
      fontSize: 19,
      bold: true,
      color: C.ink,
      typeface: DISPLAY,
      align: "center",
    });
  }
  for (const connector of [
    { left: 746, top: 309, width: 0, height: 183 },
    { left: 746, top: 368, width: 62, height: 0 },
    { left: 982, top: 368, width: 140, height: 0 },
    { left: 1122, top: 309, width: 0, height: 183 },
  ]) {
    slide.shapes.add({
      geometry: "line",
      position: connector,
      fill: "none",
      line: { style: "solid", fill: C.glow, width: 3 },
    });
  }
  addOrb(slide, { left: 808, top: 282, width: 174, height: 174 }, C.purple, C.purpleDark);
  addText(slide, "MAMÁ", { left: 830, top: 320, width: 130, height: 32 }, {
    fontSize: 24,
    bold: true,
    color: C.white,
    typeface: DISPLAY,
    align: "center",
  });
  addText(slide, "account owner", { left: 830, top: 366, width: 130, height: 30 }, {
    fontSize: 18,
    bold: true,
    color: C.lilacDeep,
    align: "center",
  });

  addShape(slide, "roundRect", { left: M, top: 474, width: 438, height: 134 }, C.lilac, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 26,
  });
  addText(slide, "The graph powers", { left: 94, top: 500, width: 170, height: 22 }, {
    fontSize: 12,
    bold: true,
    color: C.purple,
  });
  addText(slide, "permissions · protection routing\ncoordination · future expansion", {
    left: 94,
    top: 534,
    width: 370,
    height: 56,
  }, { fontSize: 21, bold: true, color: C.ink, typeface: DISPLAY });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · pages 5-7",
    "docs/business-model/nana-modelo.pdf · pages 6-8",
    "/Users/lanzanimarcos7/.codex/attachments/0c94102c-c9bf-41e2-9d0d-65a4ef170350/pasted-text.txt",
  ]);
}

// 12: Ownership boundary
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addHeader(slide, "Operating boundary", "Own the relationship. Integrate the rails.", 12);

  addShape(slide, "roundRect", { left: M, top: 180, width: 535, height: 315 }, C.purple, {
    line: { style: "solid", fill: C.purple, width: 0 },
    borderRadius: 32,
  });
  addShape(slide, "roundRect", { left: 677, top: 180, width: 535, height: 315 }, C.paper, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 32,
  });
  addText(slide, "NANA MUST OWN", { left: 100, top: 214, width: 210, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.lilacDeep,
  });
  addText(slide, "The experience and trust layer", { left: 100, top: 252, width: 450, height: 48 }, {
    fontSize: 28,
    bold: true,
    color: C.white,
    typeface: DISPLAY,
  });
  addText(slide, "• Product experience\n• Voice agent + intent orchestration\n• Family Graph + permissions\n• Protection and risk experience\n• Brand and user relationship", {
    left: 100,
    top: 330,
    width: 450,
    height: 150,
  }, { fontSize: 22, bold: true, color: C.white });

  addText(slide, "PARTNERS CAN PROVIDE", { left: 710, top: 214, width: 240, height: 24 }, {
    fontSize: 13,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Regulated financial infrastructure", { left: 710, top: 252, width: 450, height: 48 }, {
    fontSize: 28,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });
  addText(slide, "• Wallet infrastructure\n• Stablecoin and local rails\n• ARS on/off-ramp\n• KYC and compliance infrastructure\n• QR, bill payment and settlement", {
    left: 710,
    top: 330,
    width: 450,
    height: 150,
  }, { fontSize: 22, bold: true, color: C.ink });
  addText(slide, "Nana should own the user relationship, not reinvent every financial rail.", {
    left: 170,
    top: 535,
    width: 940,
    height: 38,
  }, { fontSize: 25, bold: true, color: C.purpleDark, typeface: DISPLAY, align: "center" });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · pages 3 and 7-9",
    "/Users/lanzanimarcos7/.codex/attachments/0c94102c-c9bf-41e2-9d0d-65a4ef170350/pasted-text.txt",
  ]);
}

// 13: North star and category boundary
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addHeader(slide, "Measurement and scope", "Measure independence, not activity alone.", 13);

  addShape(slide, "roundRect", { left: M, top: 180, width: 650, height: 424 }, C.lilac, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 32,
  });
  addText(slide, "NORTH STAR", { left: 100, top: 214, width: 180, height: 22 }, {
    fontSize: 13,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Financial actions completed independently that previously required assistance.", {
    left: 100,
    top: 258,
    width: 585,
    height: 112,
  }, { fontSize: 32, bold: true, color: C.ink, typeface: DISPLAY });
  addText(slide, "BEFORE NANA", { left: 100, top: 414, width: 140, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.alert,
  });
  addText(slide, "“Send me the bill. I will pay it for you.”", { left: 100, top: 444, width: 540, height: 40 }, {
    fontSize: 21,
    bold: true,
    color: C.muted,
  });
  addText(slide, "WITH NANA", { left: 100, top: 510, width: 140, height: 20 }, {
    fontSize: 11,
    bold: true,
    color: C.success,
  });
  addText(slide, "“Nana, pay the electricity bill.”  ✓", { left: 100, top: 540, width: 540, height: 40 }, {
    fontSize: 22,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });

  addShape(slide, "roundRect", { left: 752, top: 180, width: 460, height: 424 }, C.cream, {
    line: { style: "solid", fill: C.border, width: 1 },
    borderRadius: 32,
  });
  addText(slide, "NANA IS NOT", { left: 786, top: 214, width: 170, height: 22 }, {
    fontSize: 13,
    bold: true,
    color: C.alert,
  });
  addText(slide, "• a crypto wallet for crypto users\n• a chatbot inside a bank app\n• parental-control software\n• approval for every transaction\n• an interface exposing rails and gas", {
    left: 786,
    top: 252,
    width: 390,
    height: 210,
  }, { fontSize: 21, bold: true, color: C.ink });
  addRule(slide, 786, 478, 390, C.border, 1);
  addText(slide, "NANA IS", { left: 786, top: 500, width: 120, height: 20 }, {
    fontSize: 12,
    bold: true,
    color: C.purple,
  });
  addText(slide, "Intent turned into action with bounded human control.", {
    left: 786,
    top: 532,
    width: 390,
    height: 50,
  }, { fontSize: 22, bold: true, color: C.purpleDark, typeface: DISPLAY });
  addFooter(slide);
  addNotes(slide, [
    "docs/business-model/nana-business-model.pdf · pages 2-6",
    "/Users/lanzanimarcos7/.codex/attachments/0c94102c-c9bf-41e2-9d0d-65a4ef170350/pasted-text.txt",
    "The North Star is a proposed mission metric, not a validated KPI definition.",
  ]);
}

// 14: Decision filter
{
  const slide = presentation.slides.add();
  slide.background.fill = C.cream;
  addGrid(slide, { left: 0, top: 0, width: SLIDE_W, height: SLIDE_H });
  addOrb(slide, { left: 870, top: 0, width: 410, height: 410 }, C.lilacDeep, C.glow);
  addText(slide, "THE DECISION FILTER", { left: M, top: 52, width: 300, height: 22 }, {
    fontSize: 12,
    bold: true,
    color: C.purple,
  });
  addText(slide, "14", { left: 1170, top: 42, width: 42, height: 20 }, {
    fontSize: 13,
    bold: true,
    color: C.purpleDark,
    align: "right",
  });
  addText(slide, "Before Nana ships anything, ask:", { left: M, top: 94, width: 790, height: 68 }, {
    fontSize: 48,
    bold: true,
    color: C.ink,
    typeface: DISPLAY,
  });

  const questions = [
    "Does it return agency to the person using it?",
    "Does family intervene only when necessary?",
    "Does complexity stay invisible?",
    "Is safety available to everyone?",
    "Can the user understand the consequence in plain language?",
  ];
  for (let i = 0; i < questions.length; i += 1) {
    const y = 200 + i * 72;
    addText(slide, "✓", { left: M, top: y, width: 38, height: 34 }, {
      fontSize: 26,
      bold: true,
      color: C.purple,
      typeface: DISPLAY,
    });
    addText(slide, questions[i], { left: 118, top: y, width: 630, height: 38 }, {
      fontSize: 24,
      bold: true,
      color: C.ink,
      typeface: DISPLAY,
    });
  }

  addShape(slide, "roundRect", { left: 755, top: 412, width: 457, height: 196 }, C.purple, {
    line: { style: "solid", fill: C.purple, width: 0 },
    borderRadius: 34,
    shadow: "shadow-xl",
  });
  addText(slide, "THE QUESTION UNDER\nEVERY OTHER QUESTION", {
    left: 790,
    top: 446,
    width: 390,
    height: 46,
  }, { fontSize: 11, bold: true, color: C.lilacDeep, align: "center" });
  addText(slide, "Does this return agency,\nor transfer it?", {
    left: 790,
    top: 510,
    width: 390,
    height: 78,
  }, { fontSize: 30, bold: true, color: C.white, typeface: DISPLAY, align: "center" });
  addText(slide, "nana", { left: 925, top: 83, width: 210, height: 64 }, {
    fontSize: 54,
    bold: true,
    color: C.purple,
    typeface: DISPLAY,
    align: "center",
  });
  addText(slide, "Independence through voice.\nProtection when it matters.", {
    left: 830,
    top: 180,
    width: 400,
    height: 72,
  }, { fontSize: 24, bold: true, color: C.ink, typeface: DISPLAY, align: "center" });
  addFooter(slide, "NANA · INDEPENDENCE THROUGH VOICE · PROTECTION WHEN IT MATTERS");
  addNotes(slide, [
    "docs/deck/nana-deck.pdf · slides 4, 7 and 15",
    "docs/business-model/nana-business-model.pdf · pages 2 and 9",
    "This decision filter is a synthesis of the team's stated product rules.",
  ]);
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(path.dirname(OUT), { recursive: true });
await fs.mkdir(RENDERED_DIR, { recursive: true });

for (const [index, slide] of presentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(
    path.join(RENDERED_DIR, `${stem}.png`),
    await presentation.export({ slide, format: "png", scale: 1 }),
  );
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(RENDERED_DIR, `${stem}.layout.json`), await layout.text());
}

await writeBlob(
  path.join(RENDERED_DIR, "montage.webp"),
  await presentation.export({ format: "webp", montage: true, scale: 1 }),
);

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(OUT);

const regeneratedPresentation = await PresentationFile.importPptx(await FileBlob.load(OUT));
await fs.mkdir(PPTX_RENDERED_DIR, { recursive: true });
for (const [index, slide] of regeneratedPresentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(
    path.join(PPTX_RENDERED_DIR, `${stem}.png`),
    await regeneratedPresentation.export({ slide, format: "png", scale: 1 }),
  );
}

const snapshot = await presentation.inspect({
  kind: "slide,textbox,shape,image,notes",
  maxChars: 16000,
});
await fs.writeFile(path.join(RENDERED_DIR, "inspect.ndjson"), snapshot.ndjson);

console.log(OUT);
