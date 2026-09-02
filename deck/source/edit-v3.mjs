import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE_DIR = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(SOURCE_DIR, "../..");
const WORK = path.join(ROOT, "tmp/deck-v3-aesthetic");
const STARTER = path.join(ROOT, "deck/archive/v2/Nana-Pitch-Deck-v2.pptx");
const OUT = path.join(ROOT, "deck/current/Nana-Pitch-Deck-v3.pptx");
const RENDER = path.join(WORK, "final-render");
const LAYOUT = path.join(WORK, "final-layout");

const C = {
  ink: "#17151B",
  muted: "#5A5860",
  purple: "#5639DD",
  purpleDark: "#3C239F",
  lilac: "#EDE8FF",
  lilacDeep: "#D8CBFF",
  white: "#FFFFFF",
  mint: "#DDF5E8",
  green: "#2F7D5A",
  blush: "#F8E2E7",
  red: "#B44959",
};

async function readBytes(filePath) {
  const buffer = await fs.readFile(filePath);
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

const assets = {
  cream: await readBytes(path.join(ROOT, "assets/deck-textures/nana-cream-bubble-top-right-deck.jpg")),
  lilac: await readBytes(path.join(ROOT, "assets/deck-textures/nana-lilac-bubble-bottom-left-deck.jpg")),
  purple: await readBytes(path.join(ROOT, "assets/deck-textures/nana-purple-bubble-top-right-deck.jpg")),
  bubble: await readBytes(path.join(ROOT, "assets/bubbles/nana-glass-bubble-deck.png")),
  transparent: new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X99Z5QAAAABJRU5ErkJggg==", "base64")),
};

const deck = await PresentationFile.importPptx(await FileBlob.load(STARTER));

function positionOf(item) {
  const p = item.position ?? item.frame;
  return { left: p.left, top: p.top, width: p.width, height: p.height };
}

function textOf(shape) {
  try {
    // Imported elements are lazy facades: touching geometry/position first resolves text reliably.
    void shape.geometry;
    void shape.position;
    return shape.text?.text ?? shape.text?.toString?.() ?? "";
  } catch {
    return "";
  }
}

function close(a, b, tolerance = 2) {
  return Math.abs(a - b) <= tolerance;
}

function shapeAt(shapes, left, top, width = null, height = null) {
  return shapes.find(({ shape }) => {
    const p = positionOf(shape);
    return close(p.left, left) && close(p.top, top)
      && (width === null || close(p.width, width))
      && (height === null || close(p.height, height));
  })?.shape;
}

function styleCard(shape, { fill = "white/70", line = "#5639DD/14", shadow = "0px 14px 36px #3C239F/10", radius = 28 } = {}) {
  if (!shape) return;
  shape.fill = fill;
  shape.line = { style: "solid", fill: line, width: 1 };
  shape.borderRadius = radius;
  shape.shadow = shadow;
}

function styleText(entry, style = {}) {
  if (!entry?.shape) return;
  const { shape } = entry;
  if (style.color) shape.text.color = style.color;
  if (style.fontSize) shape.text.fontSize = style.fontSize;
  if (style.bold !== undefined) shape.text.bold = style.bold;
  if (style.alignment) shape.text.alignment = style.alignment;
  if (style.verticalAlignment) shape.text.verticalAlignment = style.verticalAlignment;
  if (style.insets) shape.text.insets = style.insets;
}

function findText(entries, value, condition = () => true) {
  return entries.find((entry) => entry.text === value && condition(entry));
}

function suppressImage(image) {
  const oldFrame = image.frame;
  image.replace({
    blob: assets.transparent,
    contentType: "image/png",
    alt: "Decorative asset intentionally removed in Nana deck v3",
    fit: "contain",
  });
  image.frame = oldFrame;
}

function addBackground(slide, kind) {
  const image = slide.images.add({
    blob: assets[kind],
    contentType: "image/jpeg",
    alt: `Nana ${kind} atmospheric background with a faint grid and cropped translucent bubble`,
    prompt: `Minimal Nana Wallet 16:9 presentation background, ${kind} palette, ultra-faint geometric grid, one oversized translucent lavender glass bubble cropped by the canvas edge, ample negative space, no text.`,
    fit: "cover",
    position: { left: 0, top: 0, width: 1280, height: 720 },
  });
  slide.background.fill = { type: "image", imageReference: { id: image.referenceId } };
  image.delete();
}

function addExpansionBubble(slide, left, top, size) {
  const source = slide.images.add({
    blob: assets.bubble,
    contentType: "image/png",
    alt: "Nana translucent lavender glass bubble representing a product expansion stage",
    prompt: "A soft translucent lavender glass sphere with a milky upper-left highlight and subtle purple inset shadow, matching Nana's landing-page bubble texture.",
    fit: "contain",
    position: { left, top, width: size, height: size },
  });
  const bubble = slide.shapes.add({
    geometry: "rect",
    position: { left, top, width: size, height: size },
    fill: { type: "image", imageReference: { id: source.referenceId } },
    line: { style: "solid", fill: "none", width: 0 },
  });
  bubble.sendToBack();
  source.delete();
  return bubble;
}

function restyleHeader(slideNo, textEntries) {
  const dark = slideNo === 6;
  const wordmark = findText(textEntries, "nana", ({ shape }) => positionOf(shape).top < 80);
  styleText(wordmark, { color: dark ? C.white : C.ink });

  const page = textEntries.find(({ text, shape }) => /^\d{2}$/.test(text) && positionOf(shape).top < 80);
  if (page) {
    page.shape.position = { left: 1170, top: 685, width: 46, height: 18 };
    styleText(page, { color: dark ? C.lilacDeep : C.purple, fontSize: 10, alignment: "right", verticalAlignment: "middle" });
  }

  const kicker = textEntries.find(({ text, shape }) => {
    const p = positionOf(shape);
    return p.top >= 82 && p.top <= 100 && text === text.toUpperCase() && text.length > 2;
  });
  if (kicker) {
    const width = Math.max(116, Math.min(220, kicker.text.length * 7.2 + 30));
    kicker.shape.position = { left: 64, top: 84, width, height: 28 };
    kicker.shape.fill = dark ? "white/15" : "white/66";
    kicker.shape.line = { style: "solid", fill: dark ? "white/24" : "#5639DD/28", width: 1 };
    kicker.shape.borderRadius = 14;
    styleText(kicker, {
      color: dark ? C.white : C.purple,
      fontSize: 11,
      bold: true,
      alignment: "center",
      verticalAlignment: "middle",
      insets: { top: 0, right: 8, bottom: 0, left: 8 },
    });
  }

  const footer = findText(textEntries, "NANA · CONFIDENTIAL");
  styleText(footer, { color: dark ? C.lilacDeep : "#766E82", fontSize: 9 });
}

function suppressGridAndFooterRules(shapes) {
  for (const { shape } of shapes) {
    const p = positionOf(shape);
    const isGrid = (close(p.width, 0) && p.height > 500) || (close(p.height, 0) && p.width > 700);
    if (isGrid) shape.line = { style: "solid", fill: "none", width: 0 };
  }
}

function suppressRepeatedOrbs(slideNo, images) {
  for (const image of images) {
    const p = positionOf(image);
    const cornerTop = close(p.left, 1090) && close(p.top, 0) && close(p.width, 190);
    const cornerBottom = close(p.left, 0) && close(p.top, 580) && close(p.width, 140);
    const coverNani = slideNo === 1 && close(p.left, 728) && p.width > 400;
    const closeNani = slideNo === 16 && close(p.left, 733) && p.width > 395;
    const keepCharacter = coverNani || closeNani;
    if (slideNo === 1 && !keepCharacter) suppressImage(image);
    else if (slideNo === 16 && !keepCharacter) suppressImage(image);
    else if (slideNo === 6) suppressImage(image);
    else if (cornerTop || cornerBottom) suppressImage(image);
  }
}

function recolorAllText(textEntries, color) {
  for (const entry of textEntries) styleText(entry, { color });
}

const backgrounds = [
  "cream", "lilac", "cream", "lilac",
  "cream", "purple", "cream", "lilac",
  "lilac", "cream", "lilac", "cream",
  "lilac", "cream", "cream", "purple",
];

for (let index = 0; index < deck.slides.count; index += 1) {
  const slideNo = index + 1;
  const slide = deck.slides.getItem(index);
  const rawShapes = slide.shapes.items.slice();
  const shapes = rawShapes.map((shape) => ({ shape, text: textOf(shape) }));
  const textEntries = shapes.filter(({ text }) => Boolean(text));
  const images = slide.images.items.slice();

  suppressGridAndFooterRules(shapes);
  suppressRepeatedOrbs(slideNo, images);
  if (slideNo > 1 && slideNo < 16) restyleHeader(slideNo, textEntries);

  // The two purple slides work as visual breaths and carry the same language as the landing manifesto.
  if (slideNo === 6) {
    recolorAllText(textEntries, C.white);
    const title = textEntries.find(({ shape }) => close(positionOf(shape).top, 122));
    const subtitle = textEntries.find(({ shape }) => close(positionOf(shape).top, 195));
    styleText(title, { color: C.white });
    styleText(subtitle, { color: C.lilac });
    styleText(findText(textEntries, "DEFAULT"), { color: "#BDF3D5" });
    styleText(findText(textEntries, "EXCEPTION"), { color: "#FFD0DA" });
    styleText(findText(textEntries, "NANA · CONFIDENTIAL"), { color: C.lilacDeep });
    const first = shapeAt(shapes, 145, 292, 690, 280);
    const second = shapeAt(shapes, 632, 331, 380, 205);
    if (first) {
      first.fill = "white/14";
      first.line = { style: "solid", fill: "white/30", width: 1 };
      first.shadow = "0px 18px 44px #210F72/28";
    }
    if (second) {
      second.fill = "#F8E2E7/18";
      second.line = { style: "solid", fill: "#FFD0DA/42", width: 1 };
      second.shadow = "0px 18px 44px #210F72/24";
    }
  }

  if (slideNo === 2) {
    styleCard(shapeAt(shapes, 64, 305), { fill: "white/72" });
    styleCard(shapeAt(shapes, 64, 407), { fill: "white/72" });
    styleCard(shapeAt(shapes, 64, 509), { fill: "#EDE8FF/76" });
  }

  if (slideNo === 3) {
    for (const x of [64, 446, 828]) styleCard(shapeAt(shapes, x, 300), { fill: "white/66", radius: 32 });
    for (const metric of ["1 in 6", "1.4B", "1.3B"]) styleText(findText(textEntries, metric), { color: C.purple });
  }

  if (slideNo === 5) {
    styleCard(shapeAt(shapes, 74, 278), { fill: "white/68", radius: 34 });
    styleCard(shapeAt(shapes, 736, 278), { fill: "#EDE8FF/70", radius: 34 });
  }

  if (slideNo === 7) {
    styleCard(shapeAt(shapes, 64, 285), { fill: "#DDF5E8/78", line: "#2F7D5A/20", shadow: "0px 12px 34px #2F7D5A/08" });
    styleCard(shapeAt(shapes, 64, 457), { fill: "#F8E2E7/78", line: "#B44959/20", shadow: "0px 12px 34px #B44959/08" });
  }

  if (slideNo === 8) {
    for (const [i, x] of [68, 364, 660, 956].entries()) {
      styleCard(shapeAt(shapes, x, 315), { fill: i % 2 ? "white/68" : "#EDE8FF/70", radius: 30 });
    }
    for (const [x, y, w] of [[262, 574, 150], [445, 574, 170], [648, 574, 150], [831, 574, 194]]) {
      styleCard(shapeAt(shapes, x, y, w), { fill: "white/70", line: "#5639DD/22", shadow: "none", radius: 22 });
    }
    for (const label of ["BOUNDED", "TRANSPARENT", "REVOCABLE", "USER-CONTROLLED"]) styleText(findText(textEntries, label), { color: C.purple });
  }

  if (slideNo === 9) {
    for (const x of [64, 339, 614]) styleCard(shapeAt(shapes, x, 286), { fill: "white/68", radius: 28 });
    for (const metric of ["7", "2nd", "1"]) {
      const candidates = textEntries.filter(({ text, shape }) => text === metric && positionOf(shape).top > 280 && positionOf(shape).left < 900);
      for (const entry of candidates) styleText(entry, { color: C.purple });
    }
  }

  if (slideNo === 10) {
    styleCard(shapeAt(shapes, 64, 280), {
      fill: "linear(135deg, #4B2FC7 0%, #6D53F2 100%)",
      line: "white/18",
      shadow: "0px 20px 48px #3C239F/22",
      radius: 36,
    });
    styleCard(shapeAt(shapes, 674, 280), { fill: "white/70", radius: 36 });
  }

  if (slideNo === 11) {
    for (const [i, x] of [64, 446, 828].entries()) {
      styleCard(shapeAt(shapes, x, 286), { fill: i === 1 ? "#EDE8FF/72" : "white/68", radius: 30 });
    }
    styleCard(shapeAt(shapes, 225, 582), {
      fill: "linear(135deg, #4930C0 0%, #684CF6 100%)",
      line: "white/16",
      shadow: "0px 14px 34px #3C239F/18",
      radius: 26,
    });
  }

  if (slideNo === 12) {
    for (const [x, y, w, h] of [[86, 278, 1090, 340], [347, 319, 790, 258], [692, 357, 405, 184]]) {
      const ellipse = shapeAt(shapes, x, y, w, h);
      if (ellipse) {
        ellipse.fill = "transparent";
        ellipse.line = { style: "solid", fill: "none", width: 0 };
      }
    }
    addExpansionBubble(slide, 82, 260, 360);
    addExpansionBubble(slide, 358, 302, 318);
    addExpansionBubble(slide, 690, 335, 260);
    styleText(findText(textEntries, "03"), { color: C.purple });
    styleText(findText(textEntries, "Delegated\nfinance"), { color: C.ink });
  }

  if (slideNo === 13) {
    styleCard(shapeAt(shapes, 78, 295), { fill: "white/68", radius: 34 });
    styleCard(shapeAt(shapes, 464, 295), { fill: "#EDE8FF/70", radius: 34 });
    styleCard(shapeAt(shapes, 850, 295), {
      fill: "linear(135deg, #4930C0 0%, #684CF6 100%)",
      line: "white/16",
      shadow: "0px 18px 44px #3C239F/20",
      radius: 34,
    });
  }

  if (slideNo === 14) {
    styleCard(shapeAt(shapes, 64, 276), { fill: "#DDF5E8/76", line: "#2F7D5A/16", radius: 34 });
    styleCard(shapeAt(shapes, 666, 276), { fill: "#EDE8FF/74", line: "#5639DD/16", radius: 34 });
  }

  if (slideNo === 15) {
    for (const x of [66, 365, 664, 963]) styleCard(shapeAt(shapes, x, 276), { fill: "white/58", line: "#5639DD/12", radius: 30 });
  }

  if (slideNo === 16) {
    const panel = shapeAt(shapes, 64, 495, 535, 94);
    styleCard(panel, { fill: "white/88", line: "white/42", shadow: "0px 18px 44px #210F72/24", radius: 30 });
  }

  addBackground(slide, backgrounds[index]);
}

await fs.mkdir(path.dirname(OUT), { recursive: true });
await fs.mkdir(RENDER, { recursive: true });
await fs.mkdir(LAYOUT, { recursive: true });

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

for (let index = 0; index < deck.slides.count; index += 1) {
  const slide = deck.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  const png = await deck.export({ slide, format: "png", scale: 1 });
  await writeBlob(path.join(RENDER, `${stem}.png`), png);
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(LAYOUT, `${stem}.layout.json`), await layout.text());
}

const inspect = await deck.inspect({ kind: "slide,textbox,shape,image,notes,layout", maxChars: 200000 });
await fs.writeFile(path.join(WORK, "final-inspect.ndjson"), inspect.ndjson);

const pptx = await PresentationFile.exportPptx(deck);
await pptx.save(OUT);
console.log(OUT);
