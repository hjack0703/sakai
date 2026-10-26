// Shared theme, layouts and drawing helpers for the Sakai proposal deck
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const { illus } = require("./illus");

const THEME = {
  name: "GREEX Proposal",
  headFontFace: "Meiryo UI",
  bodyFontFace: "Meiryo UI",
  colors: {
    dk1: "1F2233", // body text
    lt1: "FFFFFF",
    dk2: "2B2A5E", // deep indigo (titles, table headers)
    lt2: "F3F2FA", // pale lavender panel
    accent1: "5B4FC4", // violet (message line, primary accent)
    accent2: "3D7BC8", // blue
    accent3: "2A9785", // teal (positive / effect)
    accent4: "EFA23A", // amber highlight
    accent5: "D2557A", // pink (risk / issue)
    accent6: "7C8096", // muted gray
    hlink: "3D7BC8",
    folHlink: "5B4FC4",
  },
};
const HEX = THEME.colors;

const W = 13.333;
const H = 7.5;
const MX = 0.6; // side margin
const CW = W - MX * 2; // content width

// Body region. Slide code is written against the original 2.0–6.62 body band and remapped
// onto the taller band below, so the body runs down to the footer.
const OLD_TOP = 2.0, OLD_BOT = 6.62;
const NEW_TOP = 1.68, NEW_BOT = 6.9;
const K = (NEW_BOT - NEW_TOP) / (OLD_BOT - OLD_TOP);
const mapY = (y) => NEW_TOP + (y - OLD_TOP) * K;

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
      { placeholder: { options: { name: "client", type: "body", x: 0.8, y: 1.3, w: 7.4, h: 0.5, fontSize: 20, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.05, w: 7.6, h: 1.9, fontSize: 38, bold: true, color: C.text2, valign: "top", align: "left", margin: 0, lineSpacingMultiple: 1.1 }, text: "" } },
      { placeholder: { options: { name: "subtitle", type: "body", x: 0.8, y: 4.1, w: 7.4, h: 0.9, fontSize: 17, color: C.accent1, bold: true, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "meta", type: "body", x: 0.8, y: 5.6, w: 7.4, h: 0.9, fontSize: 14, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });

  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.4, MX, 0.6, MX],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.3, w: CW, h: 0.62, fontSize: 26, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "lead", type: "body", x: MX, y: 0.98, w: CW, h: 0.55, fontSize: 16, bold: true, color: C.accent1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "sec", type: "body", x: MX + 2.3, y: 7.17, w: 6, h: 0.22, fontSize: 9, color: C.accent6, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "グリーエックス株式会社", options: { x: MX, y: 7.17, w: 2.2, h: 0.22, fontSize: 9, color: C.accent6, margin: 0, valign: "middle" } } },
    ],
    slideNumber: { x: W - MX - 0.6, y: 7.17, w: 0.6, h: 0.22, fontSize: 9, color: C.accent6, align: "right" },
  });

  return { pres, C };
}

// Wrap a slide so body objects drawn in old coordinates land in the new body band.
function bodySlide(slide) {
  const keep = (o) => o && (o.keepSize || o.rotate);
  const remap = (o, isImage, type) => {
    if (!o || o.placeholder || o.noMap || o.y === undefined) return o;
    const round = isImage || type === "ellipse" || (o.shape && o.shape === "ellipse") || keep(o);
    if (round) {
      const cy = mapY(o.y + (o.h || 0) / 2);
      return { ...o, y: cy - (o.h || 0) / 2 };
    }
    const out = { ...o, y: mapY(o.y) };
    if (typeof o.h === "number") out.h = o.h * K;
    if (Array.isArray(o.rowH)) out.rowH = o.rowH.map((r) => r * K);
    else if (typeof o.rowH === "number") out.rowH = o.rowH * K;
    return out;
  };
  return new Proxy(slide, {
    get(target, prop) {
      if (prop === "addText") return (t, o) => target.addText(t, remap(o, false));
      if (prop === "addShape") return (type, o) => target.addShape(type, remap(o, false, type));
      if (prop === "addImage") return (o) => target.addImage(remap(o, true));
      if (prop === "addTable") return (rows, o) => target.addTable(rows, remap(o, false));
      if (prop === "addChart") return (type, data, o) => target.addChart(type, data, remap(o, false));
      const v = target[prop];
      return typeof v === "function" ? v.bind(target) : v;
    },
  });
}

// --- Icons & illustrations ---------------------------------------------------
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

// Font Awesome icon name -> Fluent Emoji Flat illustration
const ILLUS = {
  FaShieldHalved: "shield", FaWandMagicSparkles: "sparkles", FaPeopleGroup: "busts-in-silhouette",
  FaGoogle: "cloud", FaDropbox: "package", FaFileExcel: "page-facing-up", FaEnvelope: "envelope",
  FaLightbulb: "light-bulb", FaFolderOpen: "open-file-folder", FaChartLine: "chart-increasing",
  FaCircleCheck: "check-mark-button", FaTriangleExclamation: "warning", FaLink: "link",
  FaUserLock: "locked-with-key", FaBan: "prohibited", FaHourglassHalf: "hourglass-not-done",
  FaMagnifyingGlass: "magnifying-glass-tilted-left", FaBuilding: "card-index-dividers",
  FaUserNurse: "woman-health-worker", FaHandshake: "handshake", FaUserTie: "man-office-worker",
  FaMicrophone: "studio-microphone", FaBookOpen: "books", FaListCheck: "clipboard",
  FaPhotoFilm: "video-camera", FaPersonChalkboard: "woman-teacher", FaComments: "speech-balloon",
  FaClock: "stopwatch", FaUsers: "busts-in-silhouette", FaStar: "trophy", FaGoogleDrive: "card-index-dividers",
  FaYenSign: "yen-banknote", FaUserCheck: "woman-office-worker", FaHeadset: "man-technologist",
  FaSitemap: "office-building", FaClipboardList: "spiral-notepad", FaSeedling: "seedling",
};

// --- Drawing helpers -------------------------------------------------------
function header(slide, { kicker, title, lead }) {
  slide.addText(title, { placeholder: "title" });
  if (lead) slide.addText(lead, { placeholder: "lead" });
  if (kicker) slide.addText(kicker, { placeholder: "sec" });
}

function shadow() {
  return { type: "outer", color: "1F2233", opacity: 0.1, blur: 5, offset: 1.2, angle: 90 };
}

function card(slide, pres, { x, y, w, h, fill, line, radius = 0.1, shadowOn = true, name }) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: radius,
    fill: fill || { color: pres.SchemeColor.background1 },
    line: line || { color: "E2E1EE", width: 0.75 },
    shadow: shadowOn ? shadow() : undefined,
    objectName: name,
  });
}

// Draws the illustration mapped from an icon; falls back to an icon in a circle.
async function iconCircle(slide, pres, { Icon, x, y, d = 0.6, color, iconHex = "FFFFFF", name }) {
  const key = Icon.displayName || Icon.name;
  if (ILLUS[key]) {
    const s = d * 1.05, off = (s - d) / 2;
    slide.addImage({ data: await illus(ILLUS[key]), x: x - off, y: y - off, w: s, h: s, altText: name || ILLUS[key] });
    return;
  }
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: color || pres.SchemeColor.accent1 }, line: { type: "none" } });
  const pad = d * 0.24;
  slide.addImage({ data: await iconData(Icon, iconHex), x: x + pad, y: y + pad, w: d - pad * 2, h: d - pad * 2, altText: name || "icon" });
}

function text(slide, str, opts) {
  slide.addText(str, { isTextBox: true, margin: 0, valign: "top", fontSize: 14, color: opts.color, ...opts });
}

function source(slide, pres, str) {
  slide.addText(str, { isTextBox: true, noMap: true, x: MX, y: 6.93, w: CW, h: 0.22, fontSize: 9, color: pres.SchemeColor.accent6, margin: 0, valign: "middle", objectName: "source" });
}

module.exports = { THEME, HEX, W, H, MX, CW, newDeck, bodySlide, iconData, illus, header, card, iconCircle, text, source, shadow };
