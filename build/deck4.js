// 株式会社サカイ様 ご提案資料 generator (v4: v2 look, shortened v3 wording, no sources or notes)
// apply_theme.js writes THEME colors into the generated deck (path supplied via APPLY_THEME)
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");
const fa = require("react-icons/fa6");
const L = require("./lib");
const D = require("./data4");
const { MX, CW, W, HEX, header, card, iconCircle, text, source, illus } = L;

const OUT = process.argv[2] || "proposal.pptx";

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
  const notes = () => {};
  const source = () => {};

  const section = (num, title) => {
    pres.addSection({ title: `${num} ${title}` });
    return `${num} ${title}`;
  };
  const content = (sec, h) => {
    const raw = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
    header(raw, h);
    return L.bodySlide(raw);
  };

  // ======================================================================
  // 表紙
  pres.addSection({ title: "表紙" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "表紙" });
    s.addImage({ data: await illus("woman-health-worker"), x: 8.9, y: 1.7, w: 3.3, h: 3.3, altText: "介護職員のイラスト" });
    s.addImage({ data: await illus("laptop"), x: 8.35, y: 4.35, w: 1.9, h: 1.9, altText: "ノートパソコンのイラスト" });
    s.addImage({ data: await illus("sparkles"), x: 11.5, y: 1.35, w: 1.1, h: 1.1, altText: "きらめきのイラスト" });
    s.addImage({ data: await illus("locked-with-key"), x: 11.0, y: 4.75, w: 1.35, h: 1.35, altText: "鍵のイラスト" });
    s.addText("株式会社サカイ 御中", { placeholder: "client" });
    s.addText("「事務の時間」を、\nケアと未来の時間へ", { placeholder: "title" });
    s.addText("Google Workspace × Gemini 活用による\n業務効率化・Drive 移行のご提案", { placeholder: "subtitle" });
    s.addText("2026年10月27日\nグリーエックス株式会社　山口 宙輝", { placeholder: "meta" });
  }

  // はじめに: 本提案のポイント
  pres.addSection({ title: "はじめに" });
  {
    const s = content("はじめに", {
      kicker: "はじめに",
      title: "本提案のポイント",
      lead: "導入済みの Workspace／Gemini を活かし切り、安全なデータ管理と業務効率化を両立します",
    });
    const pts = [
      ["安全な Drive 移行を最優先に", "共有ルールと権限設計を先に整え、Dropbox からの切替を早期に完了します"],
      ["8名の業務から「型」をつくる", "効果を実感しやすい業務から、テンプレートと手順書を整備します"],
      ["50〜70名へ段階的に展開", "削減時間・利用率を可視化し、月次会議でご報告できる形に整えます"],
    ];
    pts.forEach(([h, b], i) => {
      const y = 2.1 + i * 1.5;
      badge(s, i + 1, { x: MX, y: y + 0.02, d: 0.55, color: C.accent1, fontSize: 18 });
      text(s, h, { x: MX + 0.8, y, w: 6.9, h: 0.45, fontSize: 19, bold: true, color: C.text2 });
      text(s, b, { x: MX + 0.8, y: y + 0.52, w: 6.9, h: 0.7, fontSize: 14, color: C.text1 });
    });
    // right: what we offer
    const px = 8.55, pw = W - MX - px;
    s.addShape(pres.ShapeType.roundRect, { x: px, y: 2.05, w: pw, h: 4.5, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" } });
    text(s, "6ヶ月後のゴール", { x: px + 0.3, y: 2.3, w: pw - 0.6, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });
    const goals = ["Dropbox 解約・Drive に一本化", "テンプレート・Gem を整備", "効果を毎月数値でご報告"];
    goals.forEach((g, i) => {
      const y = 3.0 + i * 1.1;
      badge(s, i + 1, { x: px + 0.3, y, d: 0.5, color: [C.accent1, C.accent2, C.accent3][i], fontSize: 15 });
      text(s, g, { x: px + 0.95, y: y - 0.05, w: pw - 1.1, h: 0.6, fontSize: 14, bold: true, color: C.text1, valign: "middle" });
    });
  }

  // アジェンダ
  {
    const s = content("はじめに", { kicker: "はじめに", title: "本日のアジェンダ", lead: "現状のご理解から、ご提案内容、他社事例、スケジュールまでを順にご説明します" });
    const items = [
      ["01", "現状のご理解", "業界環境と貴社の現状・課題"],
      ["02", "ご提案内容", "Drive 移行／業務別の AI 活用／社内展開"],
      ["03", "他社事例", "介護・福祉ほかの活用事例"],
      ["04", "スケジュール", "6ヶ月の進め方"],
      ["05", "本日ご相談したいこと", "確認事項と今後の進め方"],
    ];
    items.forEach(([n, t, d], i) => {
      const y = 2.05 + i * 0.9;
      text(s, n, { x: MX + 0.2, y, w: 1.0, h: 0.7, fontSize: 32, bold: true, color: C.accent1, valign: "middle" });
      text(s, t, { x: MX + 1.4, y: y + 0.02, w: 6.5, h: 0.4, fontSize: 20, bold: true, color: C.text2 });
      text(s, d, { x: MX + 1.4, y: y + 0.42, w: 6.5, h: 0.3, fontSize: 14, color: C.accent6 });
    });
    s.addImage({ data: await illus("spiral-notepad"), x: 9.4, y: 2.35, w: 2.9, h: 2.9, altText: "ノートのイラスト" });
  }

  // ======================================================================
  let sec = section("01", "現状のご理解", "業界の環境変化と、いただいた情報から貴社の現状・課題を整理しました");

  // ご相談内容の確認
  {
    const s = content(sec, {
      kicker: "01｜現状のご理解",
      title: "ご相談内容の確認",
      lead: "ご相談内容を次のように理解しています（認識に相違があれば、本日ご指摘ください）",
    });
    const th = { bold: true, color: HEX.dk2, fill: { color: "F3F2FA" }, fontSize: 13, valign: "middle" };
    const td = { color: HEX.dk1, fill: { color: "FFFFFF" }, fontSize: 13, valign: "middle" };
    s.addTable(D.rfp.map(([k, v]) => [{ text: k, options: th }, { text: v, options: td }]), { x: MX, y: 2.0, w: CW, colW: [2.2, CW - 2.2], rowH: D.rfp.map(() => 0.5), border: { type: "solid", pt: 0.75, color: "DAD8E8" }, margin: [3, 10, 3, 10], fontFace: L.THEME.bodyFontFace });
  }

  // 業界環境
  {
    const s = content(sec, {
      kicker: "01｜現状のご理解",
      title: "介護業界では「間接業務の削減」が経営テーマに",
      lead: "担い手不足が続くなか、「ケア以外の時間」をどれだけ減らせるかが競争力になっています",
    });
    const cw = (CW - 0.6) / 3;
    D.industry.stats.forEach((st, i) => {
      const x = MX + i * (cw + 0.3);
      card(s, pres, { x, y: 2.2, w: cw, h: 2.55 });
      text(s, st.label, { x: x + 0.3, y: 2.4, w: cw - 0.6, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
      text(s, [{ text: st.value, options: { fontSize: 44, bold: true, color: C.accent1 } }, { text: st.unit, options: { fontSize: 18, bold: true, color: C.accent1 } }], { x: x + 0.3, y: 2.8, w: cw - 0.6, h: 0.9, valign: "middle" });
      text(s, st.desc, { x: x + 0.3, y: 3.8, w: cw - 0.6, h: 0.6, fontSize: 13, color: C.accent6 });
    });
    panel(s, { x: MX, y: 5.0, w: CW, h: 1.5, color: C.accent1, t: 90 });
    await iconCircle(s, pres, { Icon: fa.FaLightbulb, x: MX + 0.3, y: 5.4, d: 0.7, color: C.accent4 });
    text(s, "貴社にとっての意味", { x: MX + 1.25, y: 5.2, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    text(s, D.industry.implication, { x: MX + 1.25, y: 5.58, w: CW - 1.55, h: 0.85, fontSize: 15, bold: true, color: C.text2 });
    source(s, pres, D.industry.source);
    notes(s, D.industry.notes);
  }

  // 貴社の事業から見た活用ポイント
  {
    const s = content(sec, {
      kicker: "01｜現状のご理解",
      title: "貴社の事業から見た、Workspace 活用のポイント",
      lead: D.company.lead,
    });
    const lw = 4.55;
    s.addShape(pres.ShapeType.roundRect, { x: MX, y: 2.1, w: lw, h: 4.5, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" } });
    text(s, "貴社の事業", { x: MX + 0.3, y: 2.3, w: lw - 0.6, h: 0.32, fontSize: 14, bold: true, color: C.accent1 });
    D.company.facts.forEach(([label, num, unit], i) => {
      const x = MX + 0.3 + i * 2.05;
      text(s, [{ text: num, options: { fontSize: 36, bold: true, color: C.text1 } }, { text: " " + unit, options: { fontSize: 12, color: C.text1 } }], { x, y: 2.68, w: 1.95, h: 0.62, valign: "middle" });
      text(s, label, { x, y: 3.3, w: 1.95, h: 0.3, fontSize: 11, color: C.text1 });
    });
    s.addShape(pres.ShapeType.line, { x: MX + 0.3, y: 3.9, w: lw - 0.6, h: 0, line: { color: "D5D2EE", width: 0.75 } });
    text(s, D.company.others, { x: MX + 0.3, y: 4.1, w: lw - 0.6, h: 0.45, fontSize: 14, bold: true, color: C.text2 });
    text(s, "エリア", { x: MX + 0.3, y: 4.8, w: lw - 0.6, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
    text(s, D.company.area.replace("尾張旭市・", "尾張旭市\n"), { x: MX + 0.3, y: 5.12, w: lw - 0.6, h: 0.75, fontSize: 13, color: C.text1 });
    const rx = MX + lw + 0.35, rw = CW - lw - 0.35, gw = (rw - 0.3) / 2, gh = 2.1;
    const icons = [fa.FaSitemap, fa.FaClipboardList, fa.FaComments, fa.FaSeedling];
    for (let i = 0; i < D.company.points.length; i++) {
      const pt = D.company.points[i];
      const x = rx + (i % 2) * (gw + 0.3), y = 2.1 + Math.floor(i / 2) * (gh + 0.3);
      card(s, pres, { x, y, w: gw, h: gh });
      await iconCircle(s, pres, { Icon: icons[i], x: x + 0.25, y: y + 0.25, d: 0.55, color: [C.accent1, C.accent2, C.accent3, C.accent4][i] });
      text(s, pt.t, { x: x + 0.95, y: y + 0.25, w: gw - 1.1, h: 0.55, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      text(s, pt.d, { x: x + 0.25, y: y + 0.95, w: gw - 0.45, h: 1.05, fontSize: 14, color: C.text1 });
    }
    source(s, pres, D.company.source);
    notes(s, D.company.notes);
  }

  // 貴社の現状
  {
    const s = content(sec, {
      kicker: "01｜現状のご理解",
      title: "貴社の現状：基盤は整い、「使いこなし」の段階へ",
      lead: "Workspace はすぐ使える一方、Dropbox・無料 Gmail・Office が併存し、データとルールが分散しています",
    });
    const tools = [
      { Icon: fa.FaGoogle, name: "Google Workspace", tag: "導入済", tagC: C.accent3, desc: "Drive で共有、カレンダーで実績記録、スプレッドシートで提出物・郵便物管理" },
      { Icon: fa.FaWandMagicSparkles, name: "Gemini", tag: "テスト導入中", tagC: C.accent2, desc: "担当部署8名で試行。資料作成で大幅な時間削減" },
      { Icon: fa.FaDropbox, name: "Dropbox", tag: "移行対象", tagC: C.accent5, desc: "現行のデータ保管先。早期に Drive へ切り替えたい" },
      { Icon: fa.FaFileExcel, name: "Office 2024", tag: "併用", tagC: C.accent6, desc: "買い切り版。Excel／Word との併用ルールが必要" },
      { Icon: fa.FaEnvelope, name: "Gmail（無料版）", tag: "要整理", tagC: C.accent5, desc: "一部業務で利用。組織の管理外のため集約を推奨" },
    ];
    const gw = 4.0, gh = 1.38, gx = MX, gy = 2.15;
    tools.forEach((t, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = gx + col * (gw + 0.25), y = gy + row * (gh + 0.2);
      card(s, pres, { x, y, w: i === 4 ? gw * 2 + 0.25 : gw, h: gh });
      return t;
    });
    for (let i = 0; i < tools.length; i++) {
      const t = tools[i];
      const col = i % 2, row = Math.floor(i / 2);
      const x = gx + col * (gw + 0.25), y = gy + row * (gh + 0.2);
      await iconCircle(s, pres, { Icon: t.Icon, x: x + 0.22, y: y + 0.22, d: 0.55, color: C.text2 });
      const cw2 = i === 4 ? gw * 2 + 0.25 : gw;
      text(s, t.name, { x: x + 0.92, y: y + 0.2, w: cw2 - 2.1, h: 0.32, fontSize: 14, bold: true, color: C.text2 });
      chip(s, t.tag, { x: x + cw2 - 1.08, y: y + 0.22, w: 0.92, color: t.tagC });
      text(s, t.desc, { x: x + 0.92, y: y + 0.6, w: cw2 - 1.1, h: 0.7, fontSize: 13, color: C.text1 });
    }
    // right callout
    const px = 9.2, pw = W - MX - px;
    s.addShape(pres.ShapeType.roundRect, { x: px, y: 2.15, w: pw, h: 4.36, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" } });
    text(s, "すでにある「成功体験」", { x: px + 0.3, y: 2.4, w: pw - 0.6, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
    text(s, "資料作成で AI を使い、\n数分で高品質な成果物ができた", { x: px + 0.3, y: 2.85, w: pw - 0.5, h: 1.3, fontSize: 15, bold: true, color: C.text2 });
    text(s, [
      { text: "テスト導入  ", options: { fontSize: 13, color: C.text1 } },
      { text: "8名", options: { fontSize: 24, bold: true, color: C.text1, breakLine: true } },
      { text: "展開予定  ", options: { fontSize: 13, color: C.text1 } },
      { text: "50〜70名", options: { fontSize: 24, bold: true, color: C.accent1 } },
    ], { x: px + 0.3, y: 4.3, w: pw - 0.6, h: 1.2 });
    text(s, "この体験を、管理部門・管理職の「当たり前」に", { x: px + 0.3, y: 5.75, w: pw - 0.6, h: 0.5, fontSize: 14, bold: true, color: C.text1 });
  }

  // 課題の構造化
  {
    const s = content(sec, {
      kicker: "01｜現状のご理解",
      title: "課題は3つのテーマに集約されます",
      lead: "背景要因は「ルール」「活用シーン」「展開体制」の未整備にあり、テーマごとに打ち手を定めます",
    });
    const cols = [
      { x: MX, w: 2.45, h: "テーマ" },
      { x: MX + 2.6, w: 5.6, h: "現在の課題" },
      { x: 0, w: 0, h: "" },
      { x: MX + 8.75, w: CW - 8.75, h: "打ち手の方向性" },
    ];
    cols.filter((c) => c.h).forEach((c) => text(s, c.h, { x: c.x, y: 2.0, w: c.w, h: 0.3, fontSize: 12, bold: true, color: C.accent6 }));
    const rows = [
      { Icon: fa.FaShieldHalved, color: C.accent1, theme: "A. データ管理・\nセキュリティ",
        issues: ["リンク共有で、一度開いた人が閲覧できてしまう", "個人情報を含むデータの移行が進まない"],
        cause: "共有範囲・権限の設計と、情報の区分ルールが未整備",
        dir: "管理者設定と共有ドライブで「仕組みで守る」状態にしてから移行" },
      { Icon: fa.FaWandMagicSparkles, color: C.accent2, theme: "B. AI 活用の\n具体化",
        issues: ["機能を使いこなせていない", "部門ごとに最適な使い方が見えない"],
        cause: "業務起点の活用シーン（ユースケース）が未定義",
        dir: "業務を棚卸しし、効果の大きい業務から「型」にする" },
      { Icon: fa.FaPeopleGroup, color: C.accent3, theme: "C. 定着・\n社内展開",
        issues: ["現場の抵抗と準備期間の負担", "社内で教えられる体制がない"],
        cause: "展開の手順・教材・推進役が整っていない",
        dir: "ご担当者様を「社内コーチ」に。教材を揃えて段階的に展開" },
    ];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const y = 2.38 + i * 1.42, h = 1.3;
      card(s, pres, { x: MX, y, w: CW, h, shadowOn: false, line: { color: "E2E1EE", width: 0.75 } });
      await iconCircle(s, pres, { Icon: r.Icon, x: MX + 0.18, y: y + 0.33, d: 0.62, color: r.color });
      text(s, r.theme, { x: MX + 0.95, y: y + 0.25, w: 1.55, h: 0.8, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      text(s, bullets(r.issues, { gap: 6 }), { x: cols[1].x, y: y + 0.17, w: cols[1].w, h: h - 0.3, fontSize: 14, color: C.text1, valign: "middle" });
      arrowR(s, { x: cols[3].x - 0.32, y: y + h / 2 - 0.17, w: 0.24, h: 0.34, color: r.color });
      s.addShape(pres.ShapeType.roundRect, { x: cols[3].x, y: y + 0.12, w: cols[3].w - 0.12, h: h - 0.24, rectRadius: 0.08, fill: tint(r.color, 88), line: { type: "none" } });
      text(s, r.dir, { x: cols[3].x + 0.2, y: y + 0.17, w: cols[3].w - 0.45, h: h - 0.34, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
    }
  }

  // 目指す姿
  {
    const s = content(sec, {
      kicker: "01｜現状のご理解",
      title: "6ヶ月後に目指す姿",
      lead: "ご相談時のゴールを、6ヶ月後の具体的な状態に置き換えました",
    });
    const rows = [
      [fa.FaFolderOpen, "Dropbox と Drive にデータが分散し、個人情報を移せない", "Dropbox を解約して Drive に一本化。個人情報は仕組みで保護"],
      [fa.FaWandMagicSparkles, "8名が個人の工夫で試行し、使い方が属人的", "部門別のテンプレートが揃い、資料・メール・議事録が日常的に AI 化"],
      [fa.FaChartLine, "効果が数値化されず、展開を判断しにくい", "50〜70名に展開し、削減時間を毎月可視化"],
    ];
    const lx = MX, lw = 4.9, rx = MX + 5.75, rw = CW - 5.75;
    text(s, "現在", { x: lx, y: 2.0, w: lw, h: 0.35, fontSize: 15, bold: true, color: C.accent6 });
    text(s, "6ヶ月後", { x: rx, y: 2.0, w: rw, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
    for (let i = 0; i < rows.length; i++) {
      const [Icon, before, after] = rows[i];
      const y = 2.45 + i * 1.38, h = 1.2;
      s.addShape(pres.ShapeType.roundRect, { x: lx, y, w: lw, h, rectRadius: 0.1, fill: { color: "F1F1F4" }, line: { type: "none" } });
      text(s, before, { x: lx + 0.3, y: y + 0.12, w: lw - 0.5, h: h - 0.24, fontSize: 14, color: C.text1, valign: "middle" });
      s.addShape(pres.ShapeType.chevron, { x: lx + lw + 0.22, y: y + 0.38, w: 0.4, h: 0.45, fill: { color: C.accent1 }, line: { type: "none" } });
      card(s, pres, { x: rx, y, w: rw, h, fill: tint(C.accent1, 90), line: { color: HEX.accent1, width: 1 }, shadowOn: false });
      await iconCircle(s, pres, { Icon, x: rx + 0.25, y: y + 0.3, d: 0.6, color: C.accent1 });
      text(s, after, { x: rx + 1.05, y: y + 0.12, w: rw - 1.25, h: h - 0.24, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    }
  }

  // ======================================================================
  sec = section("02", "ご提案内容", "3本の柱で、移行・活用・定着を一気通貫でご支援します");

  // 支援の全体像
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容",
      title: "ご支援の全体像：3本の柱",
      lead: "「安全なデータ基盤」→「業務別の活用の型」→「社内で広げる体制」の順に、一気通貫でご支援します",
    });
    // roof
    s.addShape(pres.ShapeType.triangle, { x: MX + 0.6, y: 1.95, w: CW - 1.2, h: 0.9, fill: { color: C.text2 }, line: { type: "none" } });
    text(s, "定型業務を減らし、ケアと新たな事業展開に時間を使える組織へ", { x: MX + 2.8, y: 2.38, w: CW - 5.6, h: 0.4, fontSize: 14, bold: true, color: C.background1, align: "center" });
    const pillars = [
      { Icon: fa.FaShieldHalved, color: C.accent1, n: "柱①", t: "Drive 移行・\n安全な共有", b: ["Dropbox から Drive への移行", "共有ドライブと権限の設計", "管理者設定と情報区分ルール"] },
      { Icon: fa.FaWandMagicSparkles, color: C.accent2, n: "柱②", t: "業務別の\n活用の「型」", b: ["8名の業務棚卸し", "テンプレート・Gem の作成", "提出物管理などの自動化"] },
      { Icon: fa.FaPeopleGroup, color: C.accent3, n: "柱③", t: "定着・\n社内展開", b: ["社内コーチの育成", "手順書・動画・FAQ の整備", "部門別の段階展開"] },
    ];
    const pw = 3.55, gap = 0.4, px0 = MX + (CW - (pw * 3 + gap * 2)) / 2;
    for (let i = 0; i < 3; i++) {
      const p = pillars[i], x = px0 + i * (pw + gap), y = 3.0, h = 2.6;
      card(s, pres, { x, y, w: pw, h });
      await iconCircle(s, pres, { Icon: p.Icon, x: x + 0.25, y: y + 0.25, d: 0.62, color: p.color });
      text(s, p.n, { x: x + 1.0, y: y + 0.2, w: 1.0, h: 0.28, fontSize: 12, bold: true, color: p.color });
      text(s, p.t, { x: x + 1.0, y: y + 0.45, w: pw - 1.15, h: 0.7, fontSize: 15, bold: true, color: C.text2 });
      text(s, bullets(p.b, { gap: 6 }), { x: x + 0.25, y: y + 1.3, w: pw - 0.45, h: 1.2, fontSize: 14, color: C.text1 });
    }
    const fy = 5.78;
    s.addShape(pres.ShapeType.roundRect, { x: px0, y: fy, w: pw * 3 + gap * 2, h: 0.36, rectRadius: 0.06, fill: tint(C.accent4, 70), line: { type: "none" } });
    text(s, "効果測定：削減時間・利用率を毎月可視化し、月次会議でご報告", { x: px0 + 0.2, y: fy, w: pw * 3 + gap * 2 - 0.4, h: 0.36, fontSize: 13, bold: true, color: C.text2, align: "center", valign: "middle" });
    s.addShape(pres.ShapeType.roundRect, { x: px0, y: fy + 0.44, w: pw * 3 + gap * 2, h: 0.36, rectRadius: 0.06, fill: tint(C.text2, 15), line: { type: "none" } });
    text(s, "ガバナンス：個人情報の取扱いルール・AI 利用ガイドライン・セキュリティ研修", { x: px0 + 0.2, y: fy + 0.44, w: pw * 3 + gap * 2 - 0.4, h: 0.36, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
  }

  // 柱① 移行ステップ
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱① Drive 移行・セキュアな共有",
      title: "Dropbox から Google Drive への移行ステップ",
      lead: "「先にルールと箱をつくり、次にデータを移す」順番で、部署単位で小さく確かめながら移行します",
    });
    const steps = D.migration.steps;
    const sw = (CW - 0.15 * 4) / 5;
    steps.forEach((st, i) => {
      const x = MX + i * (sw + 0.15);
      s.addShape(i === 0 ? pres.ShapeType.homePlate : pres.ShapeType.chevron, { x, y: 2.05, w: sw, h: 0.62, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      text(s, [{ text: `STEP ${i + 1}`, options: { fontSize: 10, breakLine: true } }, { text: st.t, options: { fontSize: 13 } }], { x: x + (i === 0 ? 0.2 : 0.45), y: 2.05, w: sw - (i === 0 ? 0.5 : 0.75), h: 0.62, bold: true, color: C.background1, valign: "middle" });
      card(s, pres, { x, y: 2.85, w: sw, h: 3.0, shadowOn: false });
      text(s, bullets(st.b, { gap: 8 }), { x: x + 0.15, y: 3.05, w: sw - 0.25, h: 1.9, fontSize: 13, color: C.text1 });
      text(s, "成果物", { x: x + 0.15, y: 5.05, w: sw - 0.3, h: 0.25, fontSize: 10, bold: true, color: C.accent1 });
      text(s, st.out, { x: x + 0.15, y: 5.3, w: sw - 0.3, h: 0.5, fontSize: 12, color: C.text2, bold: true });
    });
    panel(s, { x: MX, y: 6.0, w: CW, h: 0.62, color: C.accent3, t: 88 });
    await iconCircle(s, pres, { Icon: fa.FaCircleCheck, x: MX + 0.15, y: 6.08, d: 0.46, color: C.accent3 });
    text(s, D.migration.tool, { x: MX + 0.75, y: 6.0, w: CW - 0.95, h: 0.62, fontSize: 12, bold: true, color: C.text2, valign: "middle" });
    notes(s, D.migration.notes);
  }

  // 柱① リンク共有の不安
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱① Drive 移行・セキュアな共有",
      title: "「リンクを開いた人が見られる」不安は、仕組みで大きく減らせます",
      lead: D.sharing.lead,
    });
    // left: cause
    const lw = 4.45;
    card(s, pres, { x: MX, y: 2.1, w: lw, h: 4.45, fill: tint(C.accent5, 92), line: { type: "none" }, shadowOn: false });
    await iconCircle(s, pres, { Icon: fa.FaTriangleExclamation, x: MX + 0.25, y: 2.3, d: 0.55, color: C.accent5 });
    text(s, "よくある原因", { x: MX + 0.95, y: 2.38, w: lw - 1.2, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    D.sharing.causes.forEach((c, i) => {
      const y = 3.05 + i * 1.15;
      text(s, c.t, { x: MX + 0.3, y, w: lw - 0.6, h: 0.32, fontSize: 14, bold: true, color: C.accent5 });
      text(s, c.d, { x: MX + 0.3, y: y + 0.4, w: lw - 0.5, h: 0.6, fontSize: 13, color: C.text1 });
    });
    // right: measures
    const rx = MX + lw + 0.35, rw = CW - lw - 0.35;
    text(s, "対策：管理者設定 × 共有ドライブ × 運用ルール", { x: rx, y: 2.1, w: rw, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    const mIcons = [fa.FaLink, fa.FaUserLock, fa.FaBan, fa.FaHourglassHalf, fa.FaMagnifyingGlass];
    for (let i = 0; i < D.sharing.measures.length; i++) {
      const m = D.sharing.measures[i];
      const y = 2.55 + i * 0.8;
      card(s, pres, { x: rx, y, w: rw, h: 0.7, shadowOn: false });
      await iconCircle(s, pres, { Icon: mIcons[i], x: rx + 0.15, y: y + 0.11, d: 0.48, color: C.accent3 });
      text(s, m.t, { x: rx + 0.8, y: y + 0.05, w: rw - 2.55, h: 0.6, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      chip(s, m.ed, { x: rx + rw - 1.62, y: y + 0.21, w: 1.47, color: m.edC === "ok" ? C.accent3 : C.accent6, fontSize: 9 });
    }
    source(s, pres, D.sharing.source);
    notes(s, D.sharing.notes);
  }

  // 柱① 情報区分ルール
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱① Drive 移行・セキュアな共有",
      title: "情報区分と取扱いルール（たたき台）",
      lead: "移行前に「どこに置き、誰と共有し、AI に使ってよいか」を情報区分ごとに決めます",
    });
    const hdr = ["区分", "主なデータ", "保存場所", "共有範囲", "Gemini の利用"];
    const rows = D.classification;
    const colW = [1.9, 3.7, 2.45, 1.6, CW - 9.65];
    const th = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 12, valign: "middle" } });
    const lvColor = [HEX.accent5, HEX.accent4, HEX.accent2, HEX.accent3];
    const tbl = [hdr.map(th)].concat(rows.map((r, i) => r.map((c, j) => ({
      text: c,
      options: j === 0
        ? { bold: true, color: HEX.lt1, fill: { color: lvColor[i] }, fontSize: 14, valign: "middle" }
        : { color: HEX.dk1, fill: { color: i % 2 ? "FFFFFF" : "F7F6FC" }, fontSize: 13, valign: "middle" },
    }))));
    s.addTable(tbl, { x: MX, y: 2.1, w: CW, colW, rowH: [0.42, 0.95, 0.85, 0.85, 0.75], border: { type: "solid", pt: 0.75, color: "DAD8E8" }, margin: [4, 8, 4, 8], fontFace: L.THEME.bodyFontFace });
    text(s, D.classificationNote, { x: MX, y: 6.05, w: CW, h: 0.55, fontSize: 12, color: C.accent6 });
  }

  // 柱② 業務別活用マップ
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱② 業務別AI 活用の「型」づくり",
      title: "業務別 Gemini 活用マップ（例）",
      lead: "部門ごとに「誰の・どの業務に・どの機能を」使うかを具体化し、貴社版のマップに仕上げます",
    });
    const depts = D.useMap;
    const icons = [fa.FaBuilding, fa.FaUserNurse, fa.FaHandshake, fa.FaUserTie];
    const colors = [C.accent1, C.accent3, C.accent2, C.text2];
    const gw = (CW - 0.3) / 2, gh = 2.18;
    for (let i = 0; i < depts.length; i++) {
      const d = depts[i];
      const x = MX + (i % 2) * (gw + 0.3), y = 2.1 + Math.floor(i / 2) * (gh + 0.22);
      card(s, pres, { x, y, w: gw, h: gh });
      await iconCircle(s, pres, { Icon: icons[i], x: x + 0.2, y: y + 0.18, d: 0.5, color: colors[i] });
      text(s, d.name, { x: x + 0.85, y: y + 0.2, w: gw - 1.0, h: 0.45, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      d.items.forEach(([u, tool], k) => {
        const yy = y + 0.85 + k * 0.42;
        text(s, u, { x: x + 0.3, y: yy, w: gw - 2.45, h: 0.36, fontSize: 14, color: C.text1, valign: "middle" });
        chip(s, tool, { x: x + gw - 2.1, y: yy + 0.04, w: 1.9, h: 0.28, color: colors[i], fontSize: 10 });
      });
    }
  }

  // 柱② 優先度マトリクス
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱② 業務別AI 活用の「型」づくり",
      title: "効果を実感しやすい業務から着手します",
      lead: "「効果の大きさ」と「始めやすさ」で整理し、すぐ成果が出る業務から着手します",
    });
    const gx = MX + 0.5, gy = 2.15, gw = 6.6, gh = 4.25;
    // quadrants
    s.addShape(pres.ShapeType.rect, { x: gx, y: gy, w: gw / 2, h: gh / 2, fill: tint(C.accent2, 92), line: { color: "FFFFFF", width: 2 } });
    s.addShape(pres.ShapeType.rect, { x: gx + gw / 2, y: gy, w: gw / 2, h: gh / 2, fill: tint(C.accent3, 85), line: { color: "FFFFFF", width: 2 } });
    s.addShape(pres.ShapeType.rect, { x: gx, y: gy + gh / 2, w: gw / 2, h: gh / 2, fill: { color: "F3F3F6" }, line: { color: "FFFFFF", width: 2 } });
    s.addShape(pres.ShapeType.rect, { x: gx + gw / 2, y: gy + gh / 2, w: gw / 2, h: gh / 2, fill: tint(C.accent4, 88), line: { color: "FFFFFF", width: 2 } });
    text(s, "仕組み化（STEP2）", { x: gx + 0.12, y: gy + 0.08, w: 2.5, h: 0.28, fontSize: 11, bold: true, color: C.accent2 });
    text(s, "クイックウィン（STEP1）", { x: gx + gw / 2 + 0.12, y: gy + 0.08, w: 2.8, h: 0.28, fontSize: 11, bold: true, color: C.accent3 });
    text(s, "将来検討（STEP3）", { x: gx + 0.12, y: gy + gh / 2 + 0.08, w: 2.5, h: 0.28, fontSize: 11, bold: true, color: C.accent6 });
    text(s, "すき間時間で（随時）", { x: gx + gw / 2 + 0.12, y: gy + gh / 2 + 0.08, w: 2.8, h: 0.28, fontSize: 11, bold: true, color: "B07A1E" });
    // axes labels
    text(s, "効果の大きさ →", { x: gx - 0.3 - 1.0, y: gy + gh / 2 - 0.15, w: 2.0, h: 0.3, fontSize: 11, color: C.accent6, rotate: 270, align: "center" });
    text(s, "始めやすさ →", { x: gx, y: gy + gh + 0.05, w: gw, h: 0.28, fontSize: 11, color: C.accent6, align: "center" });
    D.priority.forEach((p) => {
      const cx = gx + p.x * gw, cy = gy + (1 - p.y) * gh;
      const d = 0.2;
      s.addShape(pres.ShapeType.ellipse, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: C.text2 }, line: { type: "none" } });
      text(s, p.t, { x: cx + 0.15, y: cy - 0.16, w: 1.9, h: 0.32, fontSize: 12, bold: true, color: C.text2, valign: "middle" });
    });
    // right: steps
    const rx = gx + gw + 0.45, rw = W - MX - rx;
    D.prioritySteps.forEach((st, i) => {
      const y = 2.15 + i * 1.45;
      card(s, pres, { x: rx, y, w: rw, h: 1.3, shadowOn: false });
      badge(s, i + 1, { x: rx + 0.2, y: y + 0.2, d: 0.45, color: [C.accent3, C.accent2, C.accent6][i], fontSize: 14 });
      text(s, st.t, { x: rx + 0.8, y: y + 0.18, w: rw - 1.0, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
      text(s, st.d, { x: rx + 0.8, y: y + 0.58, w: rw - 1.0, h: 0.6, fontSize: 13, color: C.text1 });
    });
  }

  // 柱② 活用シナリオ
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱② 業務別AI 活用の「型」づくり",
      title: "活用シナリオ例：まず着手をおすすめする3つの業務",
      lead: "いずれも追加ツールなしで実現でき、効果はテスト導入メンバーで実測します",
    });
    const sc = D.scenarios;
    const icons = [fa.FaMicrophone, fa.FaBookOpen, fa.FaListCheck];
    const colors = [C.accent1, C.accent2, C.accent3];
    const cw = (CW - 0.6) / 3;
    for (let i = 0; i < sc.length; i++) {
      const x = MX + i * (cw + 0.3), y = 2.1;
      card(s, pres, { x, y, w: cw, h: 4.5 });
      await iconCircle(s, pres, { Icon: icons[i], x: x + 0.22, y: y + 0.2, d: 0.55, color: colors[i] });
      text(s, sc[i].t, { x: x + 0.9, y: y + 0.18, w: cw - 0.95, h: 0.6, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      text(s, "いま（想定）", { x: x + 0.25, y: y + 0.95, w: 1.5, h: 0.25, fontSize: 10, bold: true, color: C.accent6 });
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.2, y: y + 1.2, w: cw - 0.4, h: 0.95, rectRadius: 0.08, fill: { color: "F1F1F4" }, line: { type: "none" } });
      text(s, sc[i].before, { x: x + 0.32, y: y + 1.25, w: cw - 0.64, h: 0.85, fontSize: 14, color: C.text1, valign: "middle" });
      s.addShape(pres.ShapeType.downArrow, { x: x + cw / 2 - 0.18, y: y + 2.22, w: 0.36, h: 0.3, fill: { color: colors[i] }, line: { type: "none" } });
      text(s, "これから", { x: x + 0.25, y: y + 2.5, w: 1.2, h: 0.25, fontSize: 10, bold: true, color: colors[i] });
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.2, y: y + 2.75, w: cw - 0.4, h: 1.15, rectRadius: 0.08, fill: tint(colors[i], 88), line: { type: "none" } });
      text(s, sc[i].after, { x: x + 0.32, y: y + 2.8, w: cw - 0.64, h: 1.05, fontSize: 14, color: C.text2, bold: true, valign: "middle" });
      text(s, "使う機能：" + sc[i].tools, { x: x + 0.25, y: y + 3.97, w: cw - 0.4, h: 0.45, fontSize: 11, color: C.text1 });
    }
    notes(s, D.scenarioNotes);
  }

  // 柱③ 定着・内製化
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱③ 定着・社内展開",
      title: "ご担当者様が「社内で教えられる」状態をつくります",
      lead: "ご担当者様を「社内コーチ」として育成し、教材と相談窓口を揃えて現場の負担を減らします",
    });
    const items = D.enablement;
    const icons = [fa.FaWandMagicSparkles, fa.FaPhotoFilm, fa.FaPersonChalkboard, fa.FaComments];
    const gw = 3.75, gh = 2.05;
    for (let i = 0; i < items.length; i++) {
      const x = MX + (i % 2) * (gw + 0.25), y = 2.1 + Math.floor(i / 2) * (gh + 0.25);
      card(s, pres, { x, y, w: gw, h: gh });
      await iconCircle(s, pres, { Icon: icons[i], x: x + 0.22, y: y + 0.22, d: 0.55, color: C.accent3 });
      text(s, items[i].t, { x: x + 0.9, y: y + 0.2, w: gw - 1.0, h: 0.6, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
      text(s, items[i].d, { x: x + 0.25, y: y + 0.95, w: gw - 0.45, h: 1.0, fontSize: 14, color: C.text1 });
    }
    const px = MX + gw * 2 + 0.55, pw = W - MX - px;
    s.addShape(pres.ShapeType.roundRect, { x: px, y: 2.1, w: pw, h: 4.35, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" } });
    text(s, "現場の負担を抑える3つの工夫", { x: px + 0.3, y: 2.3, w: pw - 0.6, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
    D.lowBurden.forEach(([t, d], i) => {
      const y = 2.9 + i * 1.12;
      badge(s, i + 1, { x: px + 0.3, y, d: 0.42, color: C.accent1, fontSize: 13 });
      text(s, t, { x: px + 0.85, y: y - 0.02, w: pw - 1.0, h: 0.35, fontSize: 14, bold: true, color: C.text1 });
      text(s, d, { x: px + 0.9, y: y + 0.4, w: pw - 1.15, h: 0.6, fontSize: 13, color: C.text1 });
    });
  }

  // 柱③ 段階展開
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容　柱③ 定着・社内展開",
      title: "段階展開のロードマップ：8名から50〜70名へ",
      lead: "担当部署 → 本部管理部門 → 営業・現場管理職と、効果を確認しながら広げます",
    });
    const ph = D.phases;
    const pw = (CW - 0.5) / 3;
    const heights = [1.3, 1.7, 2.1];
    ph.forEach((p, i) => {
      const x = MX + i * (pw + 0.25);
      const bh = heights[i], by = 4.05 - bh;
      s.addShape(pres.ShapeType.roundRect, { x, y: by, w: pw, h: bh, rectRadius: 0.08, fill: { color: [C.accent2, C.accent1, C.text2][i] }, line: { type: "none" } });
      text(s, p.when, { x: x + 0.25, y: by + 0.12, w: pw - 0.5, h: 0.3, fontSize: 12, bold: true, color: C.background1 });
      text(s, p.users, { x: x + 0.25, y: by + 0.42, w: pw - 0.5, h: 0.5, fontSize: 26, bold: true, color: C.background1, valign: "middle" });
      text(s, p.who, { x: x + 0.25, y: by + 0.92, w: pw - 0.5, h: 0.3, fontSize: 13, color: C.background1 });
      card(s, pres, { x, y: 4.2, w: pw, h: 2.35, shadowOn: false });
      text(s, p.goal, { x: x + 0.2, y: 4.32, w: pw - 0.4, h: 0.5, fontSize: 15, bold: true, color: C.text2 });
      text(s, bullets(p.acts, { gap: 8 }), { x: x + 0.2, y: 4.95, w: pw - 0.4, h: 1.55, fontSize: 14, color: C.text1 });
    });
  }

  // ご要件への対応
  {
    const s = content(sec, {
      kicker: "02｜ご提案内容",
      title: "ご要件と本提案の対応",
      lead: "いただいた必須要件・推奨要件のすべてに、本提案の具体的な取り組みでお応えします",
    });
    const th = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 12, valign: "middle" } });
    const rows = D.requirements.map((r, i) => [
      { text: r[0], options: { bold: true, color: HEX.lt1, fill: { color: r[0] === "必須" ? HEX.accent1 : HEX.accent3 }, align: "center", fontSize: 12, valign: "middle" } },
      { text: r[1], options: { color: HEX.dk1, fill: { color: i % 2 ? "FFFFFF" : "F7F6FC" }, fontSize: 13, valign: "middle" } },
      { text: r[2], options: { color: HEX.dk2, bold: true, fill: { color: i % 2 ? "FFFFFF" : "F7F6FC" }, fontSize: 13, valign: "middle" } },
    ]);
    s.addTable([[th("区分"), th("ご要件"), th("本提案での対応")]].concat(rows), { x: MX, y: 1.95, w: CW, colW: [0.9, 5.2, CW - 6.1], rowH: [0.4].concat(D.requirements.map(() => 0.55)), border: { type: "solid", pt: 0.75, color: "DAD8E8" }, margin: [3, 8, 3, 8], fontFace: L.THEME.bodyFontFace });
  }

  // ======================================================================
  sec = section("03", "他社事例", "介護・福祉の現場や他業種での Google Workspace／Gemini 活用事例をご紹介します");

  // 事例サマリー
  {
    const s = content(sec, {
      kicker: "03｜他社事例",
      title: D.caseSummary.title,
      lead: D.caseSummary.lead,
    });
    const hdr = ["組織（業種）", "取り組み", "成果"];
    const th = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 12, valign: "middle" } });
    const tbl = [hdr.map(th)].concat(D.caseSummary.rows.map((r, i) => r.map((c, j) => ({
      text: c,
      options: { color: j === 2 ? HEX.accent1 : HEX.dk1, bold: j !== 1, fill: { color: i % 2 ? "FFFFFF" : "F7F6FC" }, fontSize: 12, valign: "middle" },
    }))));
    s.addTable(tbl, { x: MX, y: 2.05, w: CW, colW: [3.4, 4.9, CW - 8.3], rowH: [0.38].concat(D.caseSummary.rows.map(() => 0.55)), border: { type: "solid", pt: 0.75, color: "DAD8E8" }, margin: [3, 8, 3, 8], fontFace: L.THEME.bodyFontFace });
    text(s, D.caseSummary.note, { x: MX, y: 6.4, w: CW, h: 0.25, fontSize: 11, color: C.accent6 });
    notes(s, D.caseSummary.notes);
  }

  // 事例詳細
  for (const cs of D.caseDetails) {
    const s = content(sec, { kicker: "03｜他社事例", title: cs.title, lead: cs.lead });
    const cw = (CW - 0.4) / 2;
    for (let i = 0; i < cs.cases.length; i++) {
      const c = cs.cases[i];
      const x = MX + i * (cw + 0.4), y = 2.1;
      card(s, pres, { x, y, w: cw, h: 4.45 });
      chip(s, c.tag, { x: x + 0.3, y: y + 0.25, w: 1.9, color: [C.accent1, C.accent2][i] });
      text(s, c.org, { x: x + 0.3, y: y + 0.62, w: cw - 0.6, h: 0.42, fontSize: 17, bold: true, color: C.text2 });
      text(s, c.profile, { x: x + 0.3, y: y + 1.05, w: cw - 0.6, h: 0.3, fontSize: 12, color: C.accent6 });
      // stat
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.3, y: y + 1.4, w: cw - 0.6, h: 0.9, rectRadius: 0.08, fill: tint([C.accent1, C.accent2][i], 90), line: { type: "none" } });
      text(s, [{ text: c.stat, options: { fontSize: 30, bold: true, color: [C.accent1, C.accent2][i] } }, { text: "  " + c.statUnit, options: { fontSize: 13, bold: true, color: C.text2 } }], { x: x + 0.5, y: y + 1.42, w: cw - 1.0, h: 0.86, valign: "middle" });
      text(s, bullets(c.points, { gap: 8 }), { x: x + 0.3, y: y + 2.5, w: cw - 0.6, h: 1.3, fontSize: 14, color: C.text1 });
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.3, y: y + 3.85, w: cw - 0.6, h: 0.45, rectRadius: 0.06, fill: tint(C.accent4, 82), line: { type: "none" } });
      text(s, "示唆：" + c.implication, { x: x + 0.42, y: y + 3.85, w: cw - 0.84, h: 0.45, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    }
    source(s, pres, cs.source);
    notes(s, cs.notes);
  }

  // ======================================================================
  sec = section("04", "スケジュール");

  // スケジュール
  {
    const s = content(sec, {
      kicker: "04｜スケジュール",
      title: "スケジュール（案）：6ヶ月で移行完了と対象部門への展開",
      lead: "Dropbox 切替を最優先に、6ヶ月で移行完了と対象部門への展開まで進めます",
    });
    const months = D.schedule.months;
    const lx = MX, lw = 3.3, gx = lx + lw, gw = CW - lw, mw = gw / months.length;
    months.forEach((m, i) => {
      s.addShape(pres.ShapeType.rect, { x: gx + i * mw, y: 2.05, w: mw, h: 0.38, fill: { color: i % 2 ? HEX.accent1 : HEX.dk2 }, line: { type: "none" } });
      text(s, m, { x: gx + i * mw, y: 2.05, w: mw, h: 0.38, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    });
    const rowH = 0.43;
    D.schedule.rows.forEach((r, i) => {
      const y = 2.55 + i * rowH;
      if (i % 2 === 0) s.addShape(pres.ShapeType.rect, { x: lx, y, w: CW, h: rowH, fill: { color: "F7F6FC" }, line: { type: "none" } });
      text(s, r.t, { x: lx + 0.1, y, w: lw - 0.2, h: rowH, fontSize: 13, color: C.text1, valign: "middle", bold: !!r.bold });
      const color = { 1: C.accent1, 2: C.accent2, 3: C.accent3, 0: C.accent4 }[r.p];
      s.addShape(pres.ShapeType.roundRect, { x: gx + r.s * mw + 0.05, y: y + 0.12, w: (r.e - r.s) * mw - 0.1, h: rowH - 0.24, rectRadius: 0.05, fill: { color }, line: { type: "none" } });
    });
    D.schedule.milestones.forEach((m) => {
      const x = gx + m.at * mw, my = 2.55 + m.row * rowH + 0.1;
      s.addShape(pres.ShapeType.diamond, { x: x - 0.13, y: my, w: 0.26, h: 0.26, fill: { color: C.accent5 }, line: { color: "FFFFFF", width: 1 } });
      text(s, m.t, { x: m.labelLeft ? x - 0.2 - 1.8 : x + 0.18, y: my - 0.02, w: 1.8, h: 0.3, fontSize: 11, bold: true, color: C.accent5, valign: "middle", align: m.labelLeft ? "right" : "left" });
    });
    const my = 2.1;
    text(s, "凡例", { x: lx + 0.1, y: my, w: 0.6, h: 0.28, fontSize: 11, bold: true, color: C.accent6, valign: "middle" });
    text(s, D.schedule.note, { x: MX, y: 6.42, w: CW, h: 0.22, fontSize: 11, color: C.accent6, align: "right" });
    [["柱①", C.accent1], ["柱②", C.accent2], ["柱③", C.accent3], ["共通", C.accent4]].forEach(([t, c], i) => {
      s.addShape(pres.ShapeType.roundRect, { x: lx + 0.7 + i * 0.65, y: my + 0.06, w: 0.18, h: 0.18, rectRadius: 0.03, fill: { color: c }, line: { type: "none" } });
      text(s, t, { x: lx + 0.92 + i * 0.65, y: my, w: 0.45, h: 0.28, fontSize: 10, color: C.text1, valign: "middle" });
    });
  }

  // ======================================================================
  sec = section("05", "本日ご相談したいこと", "最適な進め方をご提案するために、いくつか確認させてください");
  {
    const s = content(sec, {
      kicker: "05｜本日ご相談したいこと",
      title: "本日確認させていただきたいことと、今後の進め方",
      lead: "以下を伺えれば、1週間以内に正式なお見積りとスケジュールをご提示できます",
    });
    const lw = 7.1;
    D.questions.forEach((t, i) => {
      const y = 2.05 + i * 0.75;
      badge(s, i + 1, { x: MX, y: y + 0.04, d: 0.46, color: C.accent1, fontSize: 14 });
      text(s, t, { x: MX + 0.7, y, w: lw - 0.7, h: 0.55, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    });
    const rx = MX + lw + 0.4, rw = CW - lw - 0.4;
    s.addShape(pres.ShapeType.roundRect, { x: rx, y: 2.05, w: rw, h: 4.5, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" } });
    text(s, "今後の進め方", { x: rx + 0.3, y: 2.25, w: rw - 0.6, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
    D.nextSteps.forEach(([when, what], i) => {
      const y = 2.8 + i * 1.2;
      s.addShape(pres.ShapeType.ellipse, { x: rx + 0.35, y: y + 0.05, w: 0.22, h: 0.22, fill: { color: C.accent1 }, line: { type: "none" } });
      if (i < D.nextSteps.length - 1) s.addShape(pres.ShapeType.line, { x: rx + 0.46, y: y + 0.3, w: 0, h: 0.95, line: { color: HEX.accent4, width: 1.5 } });
      text(s, when, { x: rx + 0.75, y, w: rw - 1.0, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
      text(s, what, { x: rx + 0.75, y: y + 0.35, w: rw - 1.0, h: 0.6, fontSize: 15, bold: true, color: C.text2 });
    });
  }

  // ======================================================================
  pres.addSection({ title: "Appendix" });
  {
    const s = content("Appendix", { kicker: "Appendix", title: "参考：エディション別の主な機能", lead: D.editions.lead });
    const th = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 12, align: "center", valign: "middle" } });
    const tbl = [D.editions.header.map((h, i) => (i === 0 ? { text: h, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 12, valign: "middle" } } : th(h)))].concat(
      D.editions.rows.map((r, i) => r.map((c, j) => ({
        text: c,
        options: { color: c === "○" ? HEX.accent3 : c === "－" ? HEX.accent6 : HEX.dk1, bold: j === 0 || c === "○", align: j === 0 ? "left" : "center", fill: { color: i % 2 ? "FFFFFF" : "F7F6FC" }, fontSize: j === 0 ? 12 : 12, valign: "middle" },
      }))));
    s.addTable(tbl, { x: MX, y: 2.05, w: CW, colW: [4.533, 1.9, 1.9, 1.9, 1.9], rowH: 0.4, border: { type: "solid", pt: 0.75, color: "DAD8E8" }, margin: [2, 8, 2, 8], fontFace: L.THEME.bodyFontFace });
    source(s, pres, D.editions.source);
  }
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, L.THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
