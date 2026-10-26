// Theme, layouts and helpers for the v3 (simplified) Sakai proposal deck
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");

const THEME = {
  name: "GREEX Proposal",
  headFontFace: "Meiryo UI",
  bodyFontFace: "Meiryo UI",
  colors: {
    dk1: "262A3B", // body text
    lt1: "FFFFFF",
    dk2: "232B5C", // navy (titles, header blocks)
    lt2: "F4F4F8", // light gray panel
    accent1: "5149C1", // violet (message line, key figures)
    accent2: "8E89D6", // soft violet
    accent3: "2A8F7E", // teal (positive)
    accent4: "D9A23A", // amber (rare highlight)
    accent5: "C14F6C", // rose (risk)
    accent6: "7A7F93", // gray (captions)
    hlink: "5149C1",
    folHlink: "232B5C",
  },
};
const HEX = THEME.colors;
const LINE = "DCDDE6"; // hairline borders
const PANEL = "F4F4F8";
const TINT = "EEEDFA"; // pale violet

const W = 13.333;
const H = 7.5;
const MX = 0.6;
const CW = W - MX * 2;
const TOP = 1.7; // body top
const BOT = 6.9; // body bottom

function newDeck() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.author = "グリーエックス株式会社";
  pres.company = "グリーエックス株式会社";
  pres.title = "株式会社サカイ様 Google Workspace × Gemini 活用支援のご提案";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "COVER",
    background: { color: C.background1 },
    objects: [
      { rect: { x: 9.35, y: 0, w: W - 9.35, h: H, fill: { color: C.text2 }, line: { type: "none" } } },
      { placeholder: { options: { name: "client", type: "body", x: 0.9, y: 1.45, w: 8.0, h: 0.5, fontSize: 18, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 2.2, w: 8.2, h: 1.9, fontSize: 38, bold: true, color: C.text2, valign: "top", align: "left", margin: 0, lineSpacingMultiple: 1.1 }, text: "" } },
      { placeholder: { options: { name: "subtitle", type: "body", x: 0.9, y: 4.25, w: 8.0, h: 0.9, fontSize: 17, bold: true, color: C.accent1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "meta", type: "body", x: 0.9, y: 5.75, w: 8.0, h: 0.8, fontSize: 14, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.4, MX, 0.6, MX],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.32, w: CW, h: 0.6, fontSize: 25, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "lead", type: "body", x: MX, y: 0.98, w: CW, h: 0.5, fontSize: 15, bold: true, color: C.accent1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "sec", type: "body", x: MX + 2.3, y: 7.12, w: 6, h: 0.22, fontSize: 9, color: C.accent6, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "グリーエックス株式会社", options: { x: MX, y: 7.12, w: 2.2, h: 0.22, fontSize: 9, color: C.accent6, margin: 0, valign: "middle" } } },
    ],
    slideNumber: { x: W - MX - 0.6, y: 7.12, w: 0.6, h: 0.22, fontSize: 9, color: C.accent6, align: "right" },
  });
  return { pres, C };
}

const iconCache = new Map();
async function iconData(Icon, hex) {
  const key = (Icon.displayName || Icon.name) + hex;
  if (iconCache.has(key)) return iconCache.get(key);
  const svg = RDS.renderToStaticMarkup(React.createElement(Icon, { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  const data = "image/png;base64," + buf.toString("base64");
  iconCache.set(key, data);
  return data;
}

module.exports = { THEME, HEX, LINE, PANEL, TINT, W, H, MX, CW, TOP, BOT, newDeck, iconData };
