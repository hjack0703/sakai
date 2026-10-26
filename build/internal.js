// 株式会社サカイ様 案件メモ（社内用） generator
// apply_theme.js writes THEME colors into the generated deck (path supplied via APPLY_THEME)
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");
const fa = require("react-icons/fa6");
const L = require("./lib");
const D = require("./data");
const { MX, CW, W, HEX, header, card, iconCircle, text, source, illus } = L;

const OUT = process.argv[2] || "internal.pptx";

(async () => {
  const { pres, C } = L.newDeck();

  // small helpers bound to this deck -------------------------------------
  const chip = (s, str, { x, y, w, color = C.accent1, h = 0.28, fontSize = 10 }) => {
    s.addText(str, { isTextBox: true, shape: pres.ShapeType.roundRect, rectRadius: 0.06, x, y, w, h, fill: { color }, color: C.background1, fontSize, bold: true, align: "center", valign: "middle", margin: 0 });
  };
  const badge = (s, n, { x, y, d = 0.5, color = C.accent1, fontSize = 16 }) => {
    s.addText(String(n), { isTextBox: true, shape: pres.ShapeType.ellipse, x, y, w: d, h: d, fill: { color }, color: C.background1, fontSize, bold: true, align: "center", valign: "middle", margin: 0 });
  };
  const tint = (color, t = 88) => ({ color, transparency: t });
  const panel = (s, { x, y, w, h, color = C.accent1, t = 90, radius = 0.1 }) => {
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: radius, fill: tint(color, t), line: { type: "none" } });
  };
  const bullets = (items, opts = {}) =>
    items.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: opts.gap ?? 5 } }));
  const arrowR = (s, { x, y, w = 0.35, h = 0.4, color = C.accent6 }) =>
    s.addShape(pres.ShapeType.rightArrow, { x, y, w, h, fill: { color }, line: { type: "none" } });
  const notes = (s, str) => s.addNotes(str);

  const section = (num, title) => {
    pres.addSection({ title: `${num} ${title}` });
    return `${num} ${title}`;
  };
  const content = (sec, h) => {
    const raw = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
    header(raw, h);
    return L.bodySlide(raw);
  };

  // 社内用（提出前に削除）
  {
    pres.addSection({ title: "社内用（提出前に削除）" });
    const s = content("社内用（提出前に削除）", { kicker: "社内用", title: "社内用｜案件管理情報", lead: "社内確認用の資料です。お客様には提出・投影しないでください" });
    const th = { bold: true, color: HEX.lt1, fill: { color: HEX.accent5 }, fontSize: 11, valign: "middle" };
    const td = { color: HEX.dk1, fill: { color: "FFFFFF" }, fontSize: 11, valign: "middle" };
    s.addTable(D.internal.map(([k, v]) => [{ text: k, options: th }, { text: v, options: td }]), { x: MX, y: 2.0, w: CW, colW: [2.3, CW - 2.3], rowH: D.internal.map(() => 0.45), border: { type: "solid", pt: 0.75, color: "DAD8E8" }, margin: [2, 8, 2, 8], fontFace: L.THEME.bodyFontFace });
    notes(s, "想定問答メモの出典: https://gree-x.com/news/20261023.html 。助成金: https://www.mhlw.go.jp/content/11800000/001731978.pdf 、https://it-shien.smrj.go.jp/pdf/it2026_kaitei_koubo_tsujyo.pdf");
  }
  {
    const s = content("社内用（提出前に削除）", { kicker: "社内用", title: "社内用｜挙手時のアピール文（原文）", lead: "案件への挙手時に送付したアピール文です。本日の説明内容との整合確認用にご活用ください" });
    card(s, pres, { x: MX, y: 2.0, w: CW, h: 4.6, shadowOn: false, line: { color: HEX.accent5, width: 1 } });
    text(s, D.appeal.map((t, i) => ({ text: t, options: { breakLine: i < D.appeal.length - 1, paraSpaceAfter: 8 } })), { x: MX + 0.35, y: 2.2, w: CW - 0.7, h: 4.2, fontSize: 14, color: C.text1 });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, L.THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
