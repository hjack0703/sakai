// 株式会社サカイ様 ご提案資料 generator (v3: fewer words, no illustrations, no on-slide sources)
// apply_theme.js writes THEME colors into the generated deck (path supplied via APPLY_THEME)
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");
const fa = require("react-icons/fa6");
const L = require("./lib3");
const { HEX, LINE, PANEL, TINT, W, MX, CW, iconData } = L;

const OUT = process.argv[2] || "proposal.pptx";

(async () => {
  const { pres, C } = L.newDeck();

  // --- helpers -------------------------------------------------------------
  const t = (s, str, o) => s.addText(str, { isTextBox: true, margin: 0, valign: "top", fontSize: 14, color: C.text1, ...o });
  const box = (s, o) => s.addShape(pres.ShapeType.rect, { line: { type: "none" }, ...o });
  const hline = (s, x, y, w) => s.addShape(pres.ShapeType.line, { x, y, w, h: 0, line: { color: LINE, width: 0.75 } });
  const bullets = (items, gap = 6) => items.map((v, i) => ({ text: v, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: gap } }));
  const icon = async (s, Icon, x, y, d, hex) => s.addImage({ data: await iconData(Icon, hex), x, y, w: d, h: d, altText: "icon" });
  const tbl = (s, rows, o) => s.addTable(rows, { border: { type: "solid", pt: 0.75, color: LINE }, margin: [4, 10, 4, 10], fontFace: L.THEME.bodyFontFace, ...o });
  const th = (v, extra = {}) => ({ text: v, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 12, valign: "middle", ...extra } });
  const td = (v, extra = {}) => ({ text: v, options: { color: HEX.dk1, fill: { color: "FFFFFF" }, fontSize: 13, valign: "middle", ...extra } });

  let sec = "はじめに";
  const section = (name) => { pres.addSection({ title: name }); sec = name; };
  const slide = (title, lead) => {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
    s.addText(title, { placeholder: "title" });
    if (lead) s.addText(lead, { placeholder: "lead" });
    s.addText(sec, { placeholder: "sec" });
    return s;
  };

  // ======================================================================
  pres.addSection({ title: "表紙" });
  {
    const s = pres.addSlide({ masterName: "COVER", sectionTitle: "表紙" });
    s.addText("株式会社サカイ 御中", { placeholder: "client" });
    s.addText("「事務の時間」を、\nケアと未来の時間へ", { placeholder: "title" });
    s.addText("Google Workspace × Gemini 活用による\n業務効率化・Drive 移行のご提案", { placeholder: "subtitle" });
    s.addText("2026年10月27日\nグリーエックス株式会社　山口 宙輝", { placeholder: "meta" });
  }

  // ======================================================================
  section("はじめに");
  {
    const s = slide("本提案のポイント", "導入済みの Workspace／Gemini を活かし切り、安全なデータ管理と業務効率化を両立します");
    const pts = [
      ["安全な Drive 移行を\n最優先に", "共有ルールと権限設計を先に整え、Dropbox からの切替を早期に完了します", "Dropbox 解約・Drive に一本化"],
      ["8名の業務から\n「型」をつくる", "効果を実感しやすい業務から、テンプレートと手順書を整備します", "テンプレート・Gem を整備"],
      ["50〜70名へ\n段階的に展開", "削減時間・利用率を可視化し、月次会議でご報告できる形に整えます", "効果を毎月数値でご報告"],
    ];
    const cw = (CW - 0.6) / 3;
    pts.forEach(([h, d, goal], i) => {
      const x = MX + i * (cw + 0.3), y = 1.8;
      box(s, { x, y, w: cw, h: 5.0, fill: { color: PANEL } });
      t(s, `0${i + 1}`, { x: x + 0.4, y: y + 0.4, w: 1.5, h: 0.6, fontSize: 32, bold: true, color: C.accent1 });
      t(s, h, { x: x + 0.4, y: y + 1.25, w: cw - 0.8, h: 1.0, fontSize: 19, bold: true, color: C.text2 });
      t(s, d, { x: x + 0.4, y: y + 2.5, w: cw - 0.8, h: 1.3, fontSize: 14 });
      hline(s, x + 0.4, y + 3.95, cw - 0.8);
      t(s, "ゴール", { x: x + 0.4, y: y + 4.1, w: 1.0, h: 0.3, fontSize: 11, bold: true, color: C.accent6 });
      t(s, goal, { x: x + 0.4, y: y + 4.4, w: cw - 0.8, h: 0.4, fontSize: 14, bold: true, color: C.text2 });
    });
  }
  {
    const s = slide("本日のアジェンダ", "現状のご理解から、ご提案内容、他社事例、スケジュールまでを順にご説明します");
    const items = [
      ["01", "現状のご理解", "業界環境と貴社の現状・課題"],
      ["02", "ご提案内容", "Drive 移行／業務別の AI 活用／社内展開"],
      ["03", "他社事例", "介護・福祉ほかの活用事例"],
      ["04", "スケジュール", "6ヶ月の進め方"],
      ["05", "本日ご相談したいこと", "確認事項と今後の進め方"],
    ];
    items.forEach(([n, ti, d], i) => {
      const y = 1.9 + i * 0.98;
      t(s, n, { x: MX + 0.1, y, w: 1.0, h: 0.75, fontSize: 30, bold: true, color: C.accent1, valign: "middle" });
      t(s, ti, { x: MX + 1.4, y, w: 4.6, h: 0.75, fontSize: 20, bold: true, color: C.text2, valign: "middle" });
      t(s, d, { x: MX + 6.3, y, w: CW - 6.3, h: 0.75, fontSize: 15, color: C.accent6, valign: "middle" });
      if (i < items.length - 1) hline(s, MX, y + 0.86, CW);
    });
  }

  // ======================================================================
  section("01 現状のご理解");
  {
    const s = slide("ご相談内容の確認", "ご相談内容を次のように理解しています（認識に相違があれば、本日ご指摘ください）");
    const rows = [
      ["ご相談の背景", "Workspace と Gemini で業務効率化を推進中。Dropbox 移行の過程で、セキュリティ面の課題と未活用の機能が判明"],
      ["現在の状況", "担当部署8名でテスト導入中。資料作成では大幅な時間削減を実感"],
      ["展開予定", "反応が良ければ、営業部門・現場管理職へ段階展開（最終50〜70名）"],
      ["注力したい領域", "バックオフィス・管理職の事務作業の削減。現場展開の前に担当者側で準備したい"],
      ["達成すべきゴール", "Workspace／Gemini の最大活用、部門別の AI 活用方法の確立、定型業務の削減、Dropbox → Drive 移行の完了"],
      ["ご依頼の範囲", "課題分析・戦略立案・実行支援（アドバイザリー型、導入済みのシステムで改善）"],
      ["期間・予算", "いずれもご提案希望。相場観を踏まえ、複数のパターンで検討"],
      ["スケジュール", "できるだけ早期に着手したい"],
      ["選定・決裁", "提案内容・価格・適合性で総合判断／ご担当者様 → 代表取締役 → 月次会議でご報告"],
    ];
    tbl(s, rows.map(([k, v]) => [td(k, { bold: true, color: HEX.dk2, fill: { color: PANEL } }), td(v)]), { x: MX, y: 1.75, w: CW, colW: [2.2, CW - 2.2], rowH: 0.56 });
  }
  {
    const s = slide("介護業界では「間接業務の削減」が経営テーマに", "担い手不足が続くなか、「ケア以外の時間」をどれだけ減らせるかが競争力になっています");
    const stats = [
      ["3.87", "倍", "介護関係職種の有効求人倍率", "全職業平均は1.05倍（2025年6月）"],
      ["約57", "万人", "2040年度に追加で必要な介護職員", "2022年度比"],
      ["81.9", "%", "ICT 導入で文書作成の時間が短縮", "ICT 導入支援事業（令和3年度）"],
    ];
    const cw = (CW - 0.6) / 3;
    stats.forEach(([v, u, label, sub], i) => {
      const x = MX + i * (cw + 0.3), y = 1.8;
      box(s, { x, y, w: cw, h: 3.0, fill: { color: PANEL } });
      t(s, [{ text: v, options: { fontSize: 52, bold: true, color: C.accent1 } }, { text: " " + u, options: { fontSize: 20, bold: true, color: C.accent1 } }], { x: x + 0.4, y: y + 0.35, w: cw - 0.8, h: 1.1, valign: "middle" });
      t(s, label, { x: x + 0.4, y: y + 1.65, w: cw - 0.8, h: 0.7, fontSize: 15, bold: true, color: C.text2 });
      t(s, sub, { x: x + 0.4, y: y + 2.4, w: cw - 0.8, h: 0.4, fontSize: 12, color: C.accent6 });
    });
    box(s, { x: MX, y: 5.05, w: CW, h: 1.85, fill: { color: C.text2 } });
    t(s, "管理職・バックオフィスの事務時間を減らし、ケアと人材育成に振り向ける。\nその手段として、導入済みの Google Workspace を活かし切ります。", { x: MX + 0.5, y: 5.05, w: CW - 1.0, h: 1.85, fontSize: 18, bold: true, color: C.background1, valign: "middle" });
  }
  {
    const s = slide("貴社の事業から見た、Workspace 活用のポイント", "多拠点を運営する貴社では、本部と拠点の情報共有と、管理者の事務軽減が効果の出やすい領域です");
    const lw = 4.3;
    box(s, { x: MX, y: 1.8, w: lw, h: 5.05, fill: { color: PANEL } });
    t(s, "貴社の事業", { x: MX + 0.4, y: 2.05, w: lw - 0.8, h: 0.35, fontSize: 13, bold: true, color: C.accent6 });
    [["11", "グループホーム"], ["7", "小規模多機能型居宅介護"]].forEach(([n, l], i) => {
      const y = 2.5 + i * 1.25;
      t(s, n, { x: MX + 0.4, y, w: 1.0, h: 0.9, fontSize: 48, bold: true, color: C.accent1, valign: "middle" });
      t(s, l, { x: MX + 1.4, y, w: lw - 1.6, h: 0.9, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    });
    hline(s, MX + 0.4, 5.0, lw - 0.8);
    t(s, "訪問看護・デイサービスも運営", { x: MX + 0.4, y: 5.2, w: lw - 0.8, h: 0.6, fontSize: 13, color: C.text1 });
    t(s, "名古屋市・春日井市・尾張旭市・江南市・刈谷市", { x: MX + 0.4, y: 5.85, w: lw - 0.8, h: 0.7, fontSize: 13, color: C.accent6 });
    const rx = MX + lw + 0.5, rw = CW - lw - 0.5;
    const pts = [
      ["本部と拠点をつなぐ", "施設会議の資料・議事録を共有し、欠席者もすぐに確認できる"],
      ["管理者の事務を軽くする", "加算・運営指導の書類準備や、データ集計を効率化する"],
      ["多様な職員を支える", "マニュアルの翻訳・やさしい日本語化、研修教材づくりに活用"],
      ["新拠点の立ち上げを速く", "開設準備のチェックリストや採用広報をテンプレート化する"],
    ];
    pts.forEach(([h, d], i) => {
      const y = 1.85 + i * 1.26;
      t(s, `0${i + 1}`, { x: rx, y, w: 0.8, h: 0.5, fontSize: 22, bold: true, color: C.accent1 });
      t(s, h, { x: rx + 0.9, y: y + 0.02, w: rw - 0.9, h: 0.45, fontSize: 17, bold: true, color: C.text2 });
      t(s, d, { x: rx + 0.9, y: y + 0.5, w: rw - 0.9, h: 0.45, fontSize: 14 });
      if (i < pts.length - 1) hline(s, rx, y + 1.1, rw);
    });
  }
  {
    const s = slide("貴社の現状：基盤は整い、「使いこなし」の段階へ", "Workspace はすぐ使える一方、Dropbox・無料 Gmail・Office が併存し、データとルールが分散しています");
    const lw = 8.4;
    const st = (v, color) => td(v, { bold: true, color, align: "center" });
    const rows = [
      [th("ツール"), th("状況", { align: "center" }), th("現在の使い方")],
      [td("Google Workspace", { bold: true }), st("導入済", HEX.accent3), td("Drive で共有、カレンダーで実績記録、スプレッドシートで提出物・郵便物管理")],
      [td("Gemini", { bold: true }), st("試行中", HEX.accent1), td("担当部署8名で試行。資料作成で大幅な時間削減")],
      [td("Dropbox", { bold: true }), st("移行対象", HEX.accent5), td("現行のデータ保管先。早期に Drive へ切り替えたい")],
      [td("Office 2024", { bold: true }), st("併用", HEX.accent6), td("買い切り版。Excel／Word との併用ルールが必要")],
      [td("Gmail（無料版）", { bold: true }), st("要整理", HEX.accent5), td("一部業務で利用。組織の管理外のため集約を推奨")],
    ];
    tbl(s, rows, { x: MX, y: 1.8, w: lw, colW: [2.3, 1.2, lw - 3.5], rowH: [0.45, 0.92, 0.92, 0.92, 0.92, 0.92] });
    const px = MX + lw + 0.35, pw = CW - lw - 0.35;
    box(s, { x: px, y: 1.8, w: pw, h: 5.05, fill: { color: C.text2 } });
    t(s, "すでにある成功体験", { x: px + 0.35, y: 2.1, w: pw - 0.7, h: 0.35, fontSize: 13, bold: true, color: C.accent2 });
    t(s, "資料作成で AI を使い、数分で高品質な成果物ができた", { x: px + 0.35, y: 2.6, w: pw - 0.7, h: 1.3, fontSize: 17, bold: true, color: C.background1 });
    [["テスト導入", "8名"], ["展開予定", "50〜70名"]].forEach(([k, v], i) => {
      t(s, [{ text: k + "　", options: { fontSize: 12, color: C.accent2 } }, { text: v, options: { fontSize: 24, bold: true, color: C.background1 } }], { x: px + 0.35, y: 4.0 + i * 0.6, w: pw - 0.7, h: 0.55, valign: "middle" });
    });
    t(s, "この体験を、管理部門・管理職の「当たり前」に", { x: px + 0.35, y: 5.5, w: pw - 0.7, h: 0.9, fontSize: 14, color: C.background1 });
  }
  {
    const s = slide("課題は3つのテーマに集約されます", "背景要因は「ルール」「活用シーン」「展開体制」の未整備にあり、テーマごとに打ち手を定めます");
    const rows = [
      [fa.FaShieldHalved, "A", "データ管理・\nセキュリティ", ["リンク共有で、一度開いた人が閲覧できてしまう", "個人情報を含むデータの移行が進まない"], "管理者設定と共有ドライブで「仕組みで守る」状態にしてから移行"],
      [fa.FaWandMagicSparkles, "B", "AI 活用の\n具体化", ["機能を使いこなせていない", "部門ごとに最適な使い方が見えない"], "業務を棚卸しし、効果の大きい業務から「型」にする"],
      [fa.FaPeopleGroup, "C", "定着・\n社内展開", ["現場の抵抗と準備期間の負担", "社内で教えられる体制がない"], "ご担当者様を「社内コーチ」に。教材を揃えて段階的に展開"],
    ];
    const c1 = 2.9, c2x = MX + c1 + 0.3, c2w = 4.9, c3x = c2x + c2w + 0.65, c3w = W - MX - c3x;
    t(s, "現在の課題", { x: c2x, y: 1.75, w: c2w, h: 0.3, fontSize: 12, bold: true, color: C.accent6 });
    t(s, "打ち手の方向性", { x: c3x, y: 1.75, w: c3w, h: 0.3, fontSize: 12, bold: true, color: C.accent6 });
    for (let i = 0; i < rows.length; i++) {
      const [Ic, k, name, issues, dir] = rows[i];
      const y = 2.15 + i * 1.6, h = 1.4;
      box(s, { x: MX, y, w: c1, h, fill: { color: C.text2 } });
      await icon(s, Ic, MX + 0.3, y + 0.45, 0.5, "FFFFFF");
      t(s, [{ text: k + "  ", options: { color: C.accent2 } }, { text: name, options: { color: C.background1 } }], { x: MX + 1.0, y, w: c1 - 1.1, h, fontSize: 16, bold: true, valign: "middle" });
      t(s, bullets(issues, 8), { x: c2x, y, w: c2w, h, fontSize: 14, valign: "middle" });
      s.addShape(pres.ShapeType.rightArrow, { x: c3x - 0.5, y: y + h / 2 - 0.17, w: 0.32, h: 0.34, fill: { color: C.accent2 }, line: { type: "none" } });
      box(s, { x: c3x, y, w: c3w, h, fill: { color: TINT } });
      t(s, dir, { x: c3x + 0.3, y, w: c3w - 0.6, h, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    }
  }
  {
    const s = slide("6ヶ月後に目指す姿", "ご相談時のゴールを、6ヶ月後の具体的な状態に置き換えました");
    const rows = [
      ["Dropbox と Drive にデータが分散し、個人情報を移せない", "Dropbox を解約して Drive に一本化。個人情報は仕組みで保護"],
      ["8名が個人の工夫で試行し、使い方が属人的", "部門別のテンプレートが揃い、資料・メール・議事録が日常的に AI 化"],
      ["効果が数値化されず、展開を判断しにくい", "50〜70名に展開し、削減時間を毎月可視化"],
    ];
    const lw = 5.0, rx = MX + lw + 0.9, rw = CW - lw - 0.9;
    t(s, "現在", { x: MX, y: 1.75, w: lw, h: 0.35, fontSize: 14, bold: true, color: C.accent6 });
    t(s, "6ヶ月後", { x: rx, y: 1.75, w: rw, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
    rows.forEach(([b, a], i) => {
      const y = 2.2 + i * 1.57, h = 1.37;
      box(s, { x: MX, y, w: lw, h, fill: { color: PANEL } });
      t(s, b, { x: MX + 0.35, y, w: lw - 0.7, h, fontSize: 15, valign: "middle" });
      s.addShape(pres.ShapeType.chevron, { x: MX + lw + 0.3, y: y + h / 2 - 0.25, w: 0.32, h: 0.5, fill: { color: C.accent1 }, line: { type: "none" } });
      box(s, { x: rx, y, w: rw, h, fill: { color: TINT } });
      t(s, a, { x: rx + 0.35, y, w: rw - 0.7, h, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    });
  }

  // ======================================================================
  section("02 ご提案内容");
  {
    const s = slide("ご支援の全体像：3本の柱", "「安全なデータ基盤」→「業務別の活用の型」→「社内で広げる体制」の順に、一気通貫でご支援します");
    box(s, { x: MX, y: 1.8, w: CW, h: 0.62, fill: { color: C.accent1 } });
    t(s, [{ text: "目的　", options: { color: C.accent2 } }, { text: "定型業務を減らし、ケアと新たな事業展開に時間を使える組織へ", options: { color: C.background1 } }], { x: MX + 0.4, y: 1.8, w: CW - 0.8, h: 0.62, fontSize: 16, bold: true, valign: "middle" });
    const pillars = [
      [fa.FaShieldHalved, "柱①", "Drive 移行・安全な共有", ["Dropbox から Drive への移行", "共有ドライブと権限の設計", "管理者設定と情報区分ルール"]],
      [fa.FaWandMagicSparkles, "柱②", "業務別の活用の「型」", ["8名の業務棚卸し", "テンプレート・Gem の作成", "提出物管理などの自動化"]],
      [fa.FaPeopleGroup, "柱③", "定着・社内展開", ["社内コーチの育成", "手順書・動画・FAQ の整備", "部門別の段階展開"]],
    ];
    const pw = (CW - 0.6) / 3;
    for (let i = 0; i < 3; i++) {
      const [Ic, n, ti, bs] = pillars[i];
      const x = MX + i * (pw + 0.3), y = 2.6;
      box(s, { x, y, w: pw, h: 0.95, fill: { color: C.text2 } });
      await icon(s, Ic, x + 0.3, y + 0.25, 0.45, "FFFFFF");
      t(s, n, { x: x + 0.95, y: y + 0.12, w: pw - 1.1, h: 0.3, fontSize: 11, bold: true, color: C.accent2 });
      t(s, ti, { x: x + 0.95, y: y + 0.4, w: pw - 1.05, h: 0.45, fontSize: 15, bold: true, color: C.background1 });
      box(s, { x, y: y + 0.95, w: pw, h: 2.3, fill: { color: PANEL } });
      t(s, bullets(bs, 10), { x: x + 0.35, y: y + 1.2, w: pw - 0.6, h: 1.9, fontSize: 15 });
    }
    [["効果測定", "削減時間・利用率を毎月可視化し、月次会議でご報告"], ["ガバナンス", "個人情報の取扱いルール・AI 利用ガイドライン・セキュリティ研修"]].forEach(([k, v], i) => {
      const y = 6.05 + i * 0.45;
      box(s, { x: MX, y, w: CW, h: 0.38, fill: { color: i ? "E6E6EE" : PANEL } });
      t(s, [{ text: k + "　", options: { color: C.accent1 } }, { text: v, options: { color: C.text2 } }], { x: MX + 0.4, y, w: CW - 0.8, h: 0.38, fontSize: 13, bold: true, valign: "middle" });
    });
  }
  {
    const s = slide("Dropbox から Google Drive への移行ステップ", "「先にルールと箱をつくり、次にデータを移す」順番で、部署単位で小さく確かめながら移行します");
    const steps = [
      ["現状調査", ["データ量・構成", "外部共有リンク", "個人情報の所在"], "現状調査シート"],
      ["設計", ["共有ドライブ構成", "メンバー・権限", "命名ルール"], "共有ドライブ設計書"],
      ["ルール策定", ["情報区分", "外部共有の承認", "管理者設定"], "運用ルール・設定一覧"],
      ["移行", ["共有リンクの整理", "パイロット移行", "部署単位で本移行"], "移行チェックリスト"],
      ["切替・定着", ["操作ガイド配布", "読み取り専用化", "Dropbox 解約"], "切替完了報告"],
    ];
    const sw = (CW - 0.15 * 4) / 5;
    steps.forEach(([ti, bs, out], i) => {
      const x = MX + i * (sw + 0.15);
      s.addShape(i === 0 ? pres.ShapeType.homePlate : pres.ShapeType.chevron, { x, y: 1.8, w: sw, h: 0.72, fill: { color: i % 2 ? C.accent1 : C.text2 }, line: { type: "none" } });
      t(s, [{ text: `STEP ${i + 1}`, options: { fontSize: 10, breakLine: true, color: C.accent2 } }, { text: ti, options: { fontSize: 15, color: C.background1 } }], { x: x + (i === 0 ? 0.25 : 0.45), y: 1.8, w: sw - (i === 0 ? 0.55 : 0.75), h: 0.72, bold: true, valign: "middle" });
      box(s, { x, y: 2.65, w: sw, h: 3.3, fill: { color: PANEL } });
      t(s, bullets(bs, 10), { x: x + 0.2, y: 2.85, w: sw - 0.35, h: 2.0, fontSize: 13 });
      hline(s, x + 0.2, 5.05, sw - 0.4);
      t(s, "成果物", { x: x + 0.2, y: 5.15, w: sw - 0.4, h: 0.28, fontSize: 10, bold: true, color: C.accent6 });
      t(s, out, { x: x + 0.2, y: 5.45, w: sw - 0.4, h: 0.4, fontSize: 13, bold: true, color: C.text2 });
    });
    box(s, { x: MX, y: 6.1, w: CW, h: 0.8, fill: { color: TINT } });
    t(s, "Google 公式の移行機能（2026年1月提供）で権限ごと移行できます。ただし「リンクを知っている全員」の共有も引き継がれるため、移行前の整理が欠かせません。", { x: MX + 0.4, y: 6.1, w: CW - 0.8, h: 0.8, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
  }
  {
    const s = slide("「リンクを開いた人が見られる」不安は、仕組みで大きく減らせます", "原因の多くは、リンクの公開範囲の設定と、共有時に付与されて残った閲覧権限にあります");
    const lw = 4.7;
    box(s, { x: MX, y: 1.8, w: lw, h: 5.05, fill: { color: PANEL } });
    t(s, "よくある原因", { x: MX + 0.4, y: 2.05, w: lw - 0.8, h: 0.35, fontSize: 15, bold: true, color: C.accent5 });
    const causes = [
      ["「リンクを知っている全員」のまま", "権限を外しても、URL があれば開ける"],
      ["送信時に付与した権限が残る", "メール等で送ると権限が付き、残り続ける"],
      ["フォルダの権限を引き継ぐ", "フォルダ内のファイルにも同じ権限が及ぶ"],
    ];
    causes.forEach(([h, d], i) => {
      const y = 2.65 + i * 1.35;
      t(s, String(i + 1), { x: MX + 0.4, y, w: 0.4, h: 0.4, fontSize: 18, bold: true, color: C.accent5 });
      t(s, h, { x: MX + 0.85, y: y + 0.02, w: lw - 1.2, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
      t(s, d, { x: MX + 0.85, y: y + 0.48, w: lw - 1.2, h: 0.6, fontSize: 13 });
    });
    const rx = MX + lw + 0.4, rw = CW - lw - 0.4;
    t(s, "対策：管理者設定 × 共有ドライブ × 運用ルール", { x: rx, y: 1.8, w: rw, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
    const ok = (v) => td(v, { align: "center", color: v.startsWith("全") ? HEX.accent3 : HEX.dk2, bold: true, fontSize: 12 });
    const rows = [
      [th("対策"), th("利用できるエディション", { align: "center" })],
      [td("組織外へのリンク共有を禁止・制限"), ok("全エディション")],
      [td("個人情報は共有ドライブでメンバー限定に"), ok("Business Standard 以上")],
      [td("閲覧者のダウンロード・印刷・コピーを禁止"), ok("全エディション")],
      [td("外部共有に有効期限を設定"), ok("全エディション")],
      [td("監査ログで共有状況を定期点検"), ok("全エディション")],
    ];
    tbl(s, rows, { x: rx, y: 2.3, w: rw, colW: [rw - 2.6, 2.6], rowH: [0.45, 0.85, 0.85, 0.85, 0.85, 0.85] });
  }
  {
    const s = slide("情報区分と取扱いルール（たたき台）", "移行前に「どこに置き、誰と共有し、AI に使ってよいか」を情報区分ごとに決めます");
    const cat = (v, color) => td(v, { bold: true, color: HEX.lt1, fill: { color }, fontSize: 14 });
    const rows = [
      [th("区分"), th("主なデータ"), th("保存場所"), th("共有範囲"), th("Gemini の利用")],
      [cat("機密（要配慮）", HEX.dk2), td("介護記録・ケアプラン、職員の健康情報・マイナンバー"), td("介護記録システム／限定共有ドライブ"), td("担当者のみ"), td("ルール整備後に限定利用")],
      [cat("社外秘", HEX.accent1), td("人事・給与、経営資料、契約書"), td("部門別の共有ドライブ"), td("部門メンバー"), td("社内ルールの範囲で利用")],
      [cat("社内", HEX.accent2), td("マニュアル・研修資料、議事録、様式"), td("全社共有ドライブ"), td("全職員"), td("積極活用")],
      [cat("公開", HEX.accent6), td("求人情報、施設案内、広報物"), td("公開用フォルダ"), td("制限なし"), td("積極活用")],
    ];
    tbl(s, rows, { x: MX, y: 1.8, w: CW, colW: [2.0, 3.7, 2.4, 1.55, CW - 9.65], rowH: [0.45, 0.95, 0.95, 0.95, 0.95] });
    t(s, "Workspace の Gemini は、顧客データを許可なくモデルの学習に使用しないと Google が公表しています。最終的なルールは、厚生労働省のガイダンス等を踏まえて確定します。", { x: MX, y: 6.15, w: CW, h: 0.7, fontSize: 13, color: C.accent6 });
  }
  {
    const s = slide("業務別 Gemini 活用マップ（例）", "部門ごとに「誰の・どの業務に・どの機能を」使うかを具体化し、貴社版のマップに仕上げます");
    const depts = [
      ["バックオフィス（人事・経理・総務）", [["求人票・スカウト文の作成", "Gmail・ドキュメント"], ["規程・マニュアルの質問応答", "Gemini Notebook"], ["提出物・郵便物管理の自動化", "フォーム・Studio"]]],
      ["事業所管理者・現場管理職", [["会議・委員会の議事録", "Meet"], ["研修資料・やさしい日本語化", "スライド"], ["稼働・加算データの集計", "スプレッドシート"]]],
      ["営業部門（入居相談・地域連携）", [["問い合わせ返信メールの下書き", "Gmail"], ["施設紹介・提案資料の作成", "スライド"], ["面談メモの要約", "ドキュメント"]]],
      ["経営・本部", [["月次会議資料の作成", "スライド"], ["制度改正情報の要点整理", "Gemini"], ["事業計画のたたき台", "ドキュメント"]]],
    ];
    const gw = (CW - 0.3) / 2, gh = 2.4;
    depts.forEach(([name, items], i) => {
      const x = MX + (i % 2) * (gw + 0.3), y = 1.8 + Math.floor(i / 2) * (gh + 0.25);
      box(s, { x, y, w: gw, h: 0.55, fill: { color: C.text2 } });
      t(s, name, { x: x + 0.3, y, w: gw - 0.6, h: 0.55, fontSize: 15, bold: true, color: C.background1, valign: "middle" });
      box(s, { x, y: y + 0.55, w: gw, h: gh - 0.55, fill: { color: PANEL } });
      items.forEach(([u, tool], k) => {
        const yy = y + 0.75 + k * 0.53;
        t(s, u, { x: x + 0.3, y: yy, w: gw - 2.6, h: 0.45, fontSize: 14, valign: "middle" });
        t(s, tool, { x: x + gw - 2.3, y: yy, w: 2.0, h: 0.45, fontSize: 12, bold: true, color: C.accent1, align: "right", valign: "middle" });
      });
    });
  }
  {
    const s = slide("効果を実感しやすい業務から着手します", "「効果の大きさ」と「始めやすさ」で整理し、すぐ成果が出る業務から着手します");
    const gx = MX + 0.45, gy = 1.85, gw = 6.6, gh = 4.75;
    const q = [[0, 0, PANEL, "仕組み化"], [1, 0, TINT, "クイックウィン"], [0, 1, "FAFAFC", "将来検討"], [1, 1, PANEL, "すき間時間で"]];
    q.forEach(([cx, cy, color, label]) => {
      s.addShape(pres.ShapeType.rect, { x: gx + cx * gw / 2, y: gy + cy * gh / 2, w: gw / 2, h: gh / 2, fill: { color }, line: { color: "FFFFFF", width: 2 } });
      t(s, label, { x: gx + cx * gw / 2 + 0.15, y: gy + cy * gh / 2 + 0.1, w: 2.5, h: 0.3, fontSize: 12, bold: true, color: label === "クイックウィン" ? C.accent1 : C.accent6 });
    });
    t(s, "効果の大きさ →", { x: gx - 0.3 - 1.0, y: gy + gh / 2 - 0.15, w: 2.0, h: 0.3, fontSize: 11, color: C.accent6, rotate: 270, align: "center" });
    t(s, "始めやすさ →", { x: gx, y: gy + gh + 0.03, w: gw, h: 0.25, fontSize: 11, color: C.accent6, align: "center" });
    const pts = [["議事録の自動作成", 0.62, 0.9], ["資料作成のたたき台", 0.56, 0.77], ["メール文面の作成", 0.66, 0.64], ["規程・マニュアルQ&A", 0.56, 0.55], ["提出物管理の自動化", 0.06, 0.84], ["月次レポート作成", 0.14, 0.62], ["記録システムとの連携", 0.08, 0.28], ["文章の校正・要約", 0.6, 0.28]];
    pts.forEach(([label, px, py]) => {
      const cx = gx + px * gw, cy = gy + (1 - py) * gh;
      s.addShape(pres.ShapeType.ellipse, { x: cx - 0.08, y: cy - 0.08, w: 0.16, h: 0.16, fill: { color: C.text2 }, line: { type: "none" } });
      t(s, label, { x: cx + 0.15, y: cy - 0.16, w: 2.0, h: 0.32, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    });
    const rx = gx + gw + 0.45, rw = W - MX - rx;
    const steps = [["1〜2ヶ月目", "クイックウィン", "議事録・メール・資料作成のテンプレートを整備"], ["3〜4ヶ月目", "仕組み化", "提出物管理・月次レポートを自動化"], ["5ヶ月目〜", "将来検討", "記録システムとの役割分担を整理"]];
    steps.forEach(([when, h, d], i) => {
      const y = 1.85 + i * 1.62;
      box(s, { x: rx, y, w: rw, h: 1.45, fill: { color: i === 0 ? TINT : PANEL } });
      t(s, when, { x: rx + 0.3, y: y + 0.2, w: rw - 0.6, h: 0.3, fontSize: 11, bold: true, color: C.accent6 });
      t(s, h, { x: rx + 0.3, y: y + 0.5, w: rw - 0.6, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
      t(s, d, { x: rx + 0.3, y: y + 0.92, w: rw - 0.6, h: 0.45, fontSize: 13 });
    });
  }
  {
    const s = slide("活用シナリオ例：まず着手をおすすめする3つの業務", "いずれも追加ツールなしで実現でき、効果はテスト導入メンバーで実測します");
    const sc = [
      ["会議・委員会の議事録", "会議後にメモから議事録を作成（1回30〜60分）", "Meet の自動メモで議事録と ToDo を作成。確認するだけに", "Meet・Gemini（Business Standard 以上）"],
      ["規程・マニュアルの問い合わせ", "本部への問い合わせに都度対応", "規程を Gemini Notebook に集約し、職員が自分で調べられる", "Gemini Notebook"],
      ["提出物・郵便物の管理", "手入力で記録し、未提出者に個別に催促", "フォームで提出・自動記録・自動リマインド", "フォーム・スプレッドシート・Workspace Studio"],
    ];
    const cw = (CW - 0.6) / 3;
    sc.forEach(([h, before, after, tools], i) => {
      const x = MX + i * (cw + 0.3);
      t(s, h, { x, y: 1.8, w: cw, h: 0.5, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
      t(s, "いま（想定）", { x, y: 2.45, w: cw, h: 0.28, fontSize: 11, bold: true, color: C.accent6 });
      box(s, { x, y: 2.75, w: cw, h: 1.2, fill: { color: PANEL } });
      t(s, before, { x: x + 0.3, y: 2.75, w: cw - 0.6, h: 1.2, fontSize: 14, valign: "middle" });
      s.addShape(pres.ShapeType.downArrow, { x: x + cw / 2 - 0.18, y: 4.03, w: 0.36, h: 0.3, fill: { color: C.accent2 }, line: { type: "none" } });
      t(s, "これから", { x, y: 4.35, w: cw, h: 0.28, fontSize: 11, bold: true, color: C.accent1 });
      box(s, { x, y: 4.65, w: cw, h: 1.45, fill: { color: TINT } });
      t(s, after, { x: x + 0.3, y: 4.65, w: cw - 0.6, h: 1.45, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
      t(s, "使う機能：" + tools, { x, y: 6.25, w: cw, h: 0.6, fontSize: 12, color: C.accent6 });
    });
  }
  {
    const s = slide("ご担当者様が「社内で教えられる」状態をつくります", "ご担当者様を「社内コーチ」として育成し、教材と相談窓口を揃えて現場の負担を減らします");
    const lw = 7.6;
    const items = [
      ["テンプレート・Gem 集", "部門別の指示文とカスタム AI を共有ドライブに整備"],
      ["手順書・短尺動画", "1テーマ3分の動画と1枚の手順書"],
      ["社内コーチ育成・リテラシー研修", "講師用資料を提供。セキュリティ・AI 利用ルールも含む"],
      ["相談窓口・FAQ", "Google Chat で質問を受け付け、FAQ に蓄積"],
    ];
    items.forEach(([h, d], i) => {
      const y = 1.85 + i * 1.26;
      t(s, `0${i + 1}`, { x: MX, y, w: 0.8, h: 0.5, fontSize: 22, bold: true, color: C.accent1 });
      t(s, h, { x: MX + 0.9, y: y + 0.02, w: lw - 0.9, h: 0.45, fontSize: 17, bold: true, color: C.text2 });
      t(s, d, { x: MX + 0.9, y: y + 0.5, w: lw - 0.9, h: 0.45, fontSize: 14 });
      if (i < items.length - 1) hline(s, MX, y + 1.1, lw);
    });
    const px = MX + lw + 0.4, pw = CW - lw - 0.4;
    box(s, { x: px, y: 1.8, w: pw, h: 5.05, fill: { color: C.text2 } });
    t(s, "現場の負担を抑える3つの工夫", { x: px + 0.35, y: 2.1, w: pw - 0.7, h: 0.4, fontSize: 14, bold: true, color: C.accent2 });
    [["「今の業務の延長」から始める", "いつもの資料・メールに AI を足すだけ"], ["1回30分のミニ勉強会", "業務の合間にオンラインで参加"], ["成功事例を毎月共有", "身近な活用例で自然に広げる"]].forEach(([h, d], i) => {
      const y = 2.75 + i * 1.3;
      t(s, h, { x: px + 0.35, y, w: pw - 0.7, h: 0.45, fontSize: 16, bold: true, color: C.background1 });
      t(s, d, { x: px + 0.35, y: y + 0.48, w: pw - 0.7, h: 0.45, fontSize: 13, color: C.accent2 });
    });
  }
  {
    const s = slide("段階展開のロードマップ：8名から50〜70名へ", "担当部署 → 本部管理部門 → 営業・現場管理職と、効果を確認しながら広げます");
    const ph = [
      ["Phase 0｜1〜2ヶ月目", "8名", "担当部署（テスト導入）", "効果の出る「型」をつくる", ["業務棚卸し・ユースケース選定", "テンプレート・Gem 作成", "Drive 移行設計・KPI 設定"]],
      ["Phase 1｜3〜4ヶ月目", "〜20名", "本部管理部門", "バックオフィスへ展開・移行完了", ["部門別テンプレートの拡充", "社内コーチの勉強会", "Dropbox からの本移行"]],
      ["Phase 2｜5〜6ヶ月目", "50〜70名", "営業・現場管理職", "対象部門への展開と効果のご報告", ["営業・管理職向けの展開", "FAQ・相談窓口の運用", "効果測定と次年度計画"]],
    ];
    const pw = (CW - 0.6) / 3, heights = [1.3, 1.7, 2.1], fills = [C.accent2, C.accent1, C.text2];
    ph.forEach(([when, users, who, goal, acts], i) => {
      const x = MX + i * (pw + 0.3), bh = heights[i], by = 3.95 - bh;
      box(s, { x, y: by, w: pw, h: bh, fill: { color: fills[i] } });
      t(s, when, { x: x + 0.3, y: by + 0.15, w: pw - 0.6, h: 0.3, fontSize: 11, bold: true, color: C.background1 });
      t(s, users, { x: x + 0.3, y: by + 0.45, w: pw - 0.6, h: 0.5, fontSize: 26, bold: true, color: C.background1, valign: "middle" });
      t(s, who, { x: x + 0.3, y: by + 0.92, w: pw - 0.6, h: 0.3, fontSize: 13, color: C.background1 });
      box(s, { x, y: 4.1, w: pw, h: 2.75, fill: { color: PANEL } });
      t(s, goal, { x: x + 0.3, y: 4.3, w: pw - 0.6, h: 0.5, fontSize: 15, bold: true, color: C.text2 });
      t(s, bullets(acts, 8), { x: x + 0.3, y: 4.95, w: pw - 0.6, h: 1.8, fontSize: 14 });
    });
  }
  {
    const s = slide("ご要件と本提案の対応", "いただいた必須要件・推奨要件のすべてに、本提案の具体的な取り組みでお応えします");
    const req = [
      ["必須", "Google Workspace・Gemini の専門知識", "管理者設定から Gemini 活用まで一貫して支援"],
      ["必須", "部門・担当者別の AI 活用の提案", "業務棚卸しから部門別の活用マップを作成"],
      ["必須", "資料・メール・議事録・共有・自動化の幅広い活用", "活用シナリオ、テンプレート・Gem 集、自動化"],
      ["必須", "セキュリティの基礎知識・リテラシー", "情報区分ルール・AI ガイドラインと職員研修"],
      ["推奨", "Workspace の活用支援への特化", "追加ツールなし。導入済みの範囲で設計・支援"],
      ["推奨", "少人数での習得から社内展開", "8名 → 本部 → 50〜70名の段階展開と社内コーチ育成"],
      ["推奨", "介護業界の理解と第三者視点の助言", "介護・福祉の事例と制度動向を踏まえて提案"],
      ["推奨", "費用対効果の明示", "削減時間・利用率を毎月計測し、金額換算してご報告"],
    ];
    const rows = [[th("区分", { align: "center" }), th("ご要件"), th("本提案での対応")]].concat(req.map(([k, a, b]) => [
      td(k, { bold: true, color: HEX.lt1, fill: { color: k === "必須" ? HEX.dk2 : HEX.accent2 }, align: "center" }), td(a), td(b, { bold: true, color: HEX.dk2 }),
    ]));
    tbl(s, rows, { x: MX, y: 1.8, w: CW, colW: [1.0, 5.2, CW - 6.2], rowH: [0.45].concat(req.map(() => 0.58)) });
  }

  // ======================================================================
  section("03 他社事例");
  {
    const s = slide("事例サマリー：介護・福祉の現場でも活用が進んでいます", "介護・福祉の法人から段階展開・Dropbox 移行まで、貴社に近い公開事例をご紹介します");
    const cs = [
      ["シーユーシー・ホスピス", "ホスピス型住宅・訪問看護", "Gem で音声・メモから介護記録の下書きを作成", "職員1人 月約3.5時間削減"],
      ["社会福祉法人ことの海会", "障がい・児童・高齢者支援", "LINE WORKS から Workspace へ一括切替", "管理職と現場でライセンスを使い分け"],
      ["ソラスト", "医療事務・介護・保育", "3万人以上の全職員で Workspace への移行を決定", "約3ヶ月で基盤を構築"],
      ["日本特殊陶業", "製造・名古屋", "15部門40名で1ヶ月トライアル", "週平均3.1時間削減／人"],
      ["船井総研ホールディングス", "経営コンサルティング", "50名 → 150名 → 全社と段階展開", "4ヶ月で全社導入・利用率97%"],
      ["星野リゾート", "宿泊", "企画・発信業務やアンケート分析に活用", "「魅力造成」の作業時間 約30%削減"],
      ["The Washington Post", "報道", "Dropbox から Workspace へ移行", "1,500名超・240TB超を期限内に移行"],
    ];
    const rows = [[th("組織"), th("業種"), th("取り組み"), th("成果")]].concat(cs.map(([a, b, c, d]) => [td(a, { bold: true, color: HEX.dk2 }), td(b, { color: HEX.accent6, fontSize: 12 }), td(c), td(d, { bold: true, color: HEX.accent1 })]));
    tbl(s, rows, { x: MX, y: 1.8, w: CW, colW: [2.8, 2.6, 3.7, CW - 9.1], rowH: [0.42].concat(cs.map(() => 0.6)) });
    t(s, "※ いずれも各社の公開事例です（当社の支援事例ではありません）", { x: MX, y: 6.55, w: CW, h: 0.3, fontSize: 11, color: C.accent6 });
  }
  const caseSlide = (title, lead, cases) => {
    const s = slide(title, lead);
    const cw = (CW - 0.4) / 2;
    cases.forEach((c, i) => {
      const x = MX + i * (cw + 0.4), y = 1.8;
      box(s, { x, y, w: cw, h: 5.05, fill: { color: PANEL } });
      t(s, c.tag, { x: x + 0.4, y: y + 0.3, w: cw - 0.8, h: 0.3, fontSize: 11, bold: true, color: C.accent1 });
      t(s, c.org, { x: x + 0.4, y: y + 0.62, w: cw - 0.8, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
      t(s, c.profile, { x: x + 0.4, y: y + 1.1, w: cw - 0.8, h: 0.3, fontSize: 12, color: C.accent6 });
      t(s, [{ text: c.stat, options: { fontSize: 40, bold: true, color: C.accent1 } }, { text: "  " + c.unit, options: { fontSize: 15, bold: true, color: C.text2 } }], { x: x + 0.4, y: y + 1.55, w: cw - 0.8, h: 0.95, valign: "middle" });
      hline(s, x + 0.4, y + 2.65, cw - 0.8);
      t(s, bullets(c.points, 10), { x: x + 0.4, y: y + 2.85, w: cw - 0.8, h: 1.2, fontSize: 14 });
      box(s, { x: x + 0.4, y: y + 4.2, w: cw - 0.8, h: 0.6, fill: { color: "FFFFFF" } });
      t(s, [{ text: "示唆　", options: { color: C.accent1 } }, { text: c.imp, options: { color: C.text2 } }], { x: x + 0.6, y: y + 4.2, w: cw - 1.2, h: 0.6, fontSize: 14, bold: true, valign: "middle" });
    });
  };
  caseSlide("介護・福祉の事例：記録業務と情報共有基盤の見直し", "同じ介護・福祉の現場で、記録業務の効率化と情報共有基盤の見直しが進んでいます", [
      { tag: "介護・ホスピス", org: "株式会社シーユーシー・ホスピス", profile: "ホスピス型住宅「ReHOPE」など全国59拠点", stat: "月約3.5時間", unit: "削減／職員1人", points: ["Gem で音声・メモから記録の下書きを作成", "記録1件 4.0分 → 3.3分（約20%短縮）"], imp: "ルール整備後の現場展開の候補に" },
      { tag: "社会福祉法人", org: "社会福祉法人ことの海会（長崎県）", profile: "障がい・児童発達・高齢者支援", stat: "一括導入", unit: "段階導入から切替", points: ["LINE WORKS から Google チャットへ切替", "管理職は Business Standard、現場は Frontline Starter"], imp: "ライセンスの使い分けは貴社の参考に" },
    ]);
  caseSlide("段階展開の事例：小さく試して、全社に広げる", "少人数で効果を測ってから広げる進め方は、貴社の「8名 → 50〜70名」と同じ型です", [
      { tag: "製造・愛知県", org: "日本特殊陶業株式会社", profile: "名古屋市に本社を置くメーカー", stat: "週3.1時間", unit: "削減／人（平均）", points: ["15部門40名で Gemini を1ヶ月試行", "60件のユースケースが集まる"], imp: "少人数でも、削減時間を測れば判断材料に" },
      { tag: "経営コンサルティング", org: "株式会社船井総研ホールディングス", profile: "中堅・中小企業向けコンサルティング", stat: "97%", unit: "の社員が Gemini を活用", points: ["50名 → 150名 → 全社と段階展開", "活用コンテストや勉強会で定着"], imp: "成功事例の共有が定着の鍵" },
    ]);
  caseSlide("データ移行とナレッジ活用の事例", "移行は「引っ越し」で終わらせず、蓄積した資料を AI で使える知識に変えます", [
      { tag: "医療用品メーカー", org: "ダイヤ工業株式会社（岡山県）", profile: "サポーター・コルセット等のメーカー", stat: "70%超", unit: "の社員が Gemini を日常利用", points: ["ファイルサーバーから Workspace へ移行", "会議資料と社内規定を NotebookLM に集約"], imp: "問い合わせ対応の効率化に応用できる" },
      { tag: "Dropbox からの移行", org: "The Washington Post（米国）", profile: "報道機関", stat: "240TB超", unit: "を期限内に移行", points: ["Dropbox の契約終了に合わせて移行", "複雑な権限・フォルダを事前に整理"], imp: "成否は移行前の整理で決まる" },
    ]);

  // ======================================================================
  section("04 スケジュール");
  {
    const s = slide("スケジュール（案）：6ヶ月で移行完了と対象部門への展開", "Dropbox 切替を最優先に、6ヶ月で移行完了と対象部門への展開まで進めます");
    const months = ["1ヶ月目", "2ヶ月目", "3ヶ月目", "4ヶ月目", "5ヶ月目", "6ヶ月目"];
    const lw = 3.4, gx = MX + lw, gw = CW - lw, mw = gw / 6;
    months.forEach((m, i) => {
      box(s, { x: gx + i * mw, y: 1.8, w: mw, h: 0.42, fill: { color: i % 2 ? HEX.accent1 : HEX.dk2 } });
      t(s, m, { x: gx + i * mw, y: 1.8, w: mw, h: 0.42, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    });
    const colors = { 1: HEX.dk2, 2: HEX.accent1, 3: HEX.accent2, 0: HEX.accent6 };
    const rows = [
      ["キックオフ・現状調査", 0, 0.6, 0], ["Drive 移行設計・ルール策定", 0.4, 1.6, 1], ["パイロット移行（1部署）", 1.4, 2.2, 1], ["本移行・Dropbox 切替", 2.2, 4.0, 1],
      ["業務棚卸し・テンプレート整備", 0.3, 2.5, 2], ["業務の自動化（提出物管理等）", 2.0, 4.0, 2], ["本部管理部門への展開", 2.0, 4.0, 3], ["営業・現場管理職への展開", 4.0, 6.0, 3], ["効果測定・月次報告", 0.6, 6.0, 0],
    ];
    const rh = 0.47;
    rows.forEach(([label, a, b, p], i) => {
      const y = 2.32 + i * rh;
      if (i % 2 === 0) box(s, { x: MX, y, w: CW, h: rh, fill: { color: PANEL } });
      t(s, label, { x: MX + 0.15, y, w: lw - 0.2, h: rh, fontSize: 13, valign: "middle" });
      box(s, { x: gx + a * mw + 0.05, y: y + 0.12, w: (b - a) * mw - 0.1, h: rh - 0.24, fill: { color: colors[p] } });
    });
    const my = 2.32 + 3 * rh + 0.1;
    s.addShape(pres.ShapeType.diamond, { x: gx + 4 * mw - 0.13, y: my, w: 0.26, h: 0.26, fill: { color: C.accent5 }, line: { color: "FFFFFF", width: 1 } });
    t(s, "Dropbox 解約", { x: gx + 4 * mw + 0.2, y: my - 0.02, w: 1.8, h: 0.3, fontSize: 12, bold: true, color: C.accent5, valign: "middle" });
    const ly = 2.32 + rows.length * rh + 0.15;
    [["柱① 移行", HEX.dk2], ["柱② 活用", HEX.accent1], ["柱③ 展開", HEX.accent2], ["共通", HEX.accent6]].forEach(([l, c], i) => {
      box(s, { x: MX + 0.15 + i * 1.6, y: ly + 0.07, w: 0.2, h: 0.2, fill: { color: c } });
      t(s, l, { x: MX + 0.42 + i * 1.6, y: ly, w: 1.2, h: 0.34, fontSize: 11, valign: "middle" });
    });
    t(s, "※ 開始時期はご決裁後に確定。Dropbox 解約は契約更新時期に合わせて調整", { x: MX + 6.4, y: ly, w: CW - 6.4, h: 0.34, fontSize: 11, color: C.accent6, align: "right", valign: "middle" });
  }

  // ======================================================================
  section("05 本日ご相談したいこと");
  {
    const s = slide("本日確認させていただきたいことと、今後の進め方", "以下を伺えれば、1週間以内に正式なお見積りとスケジュールをご提示できます");
    const lw = 7.7;
    const qs = ["ご契約エディション・ライセンス数・管理者", "Dropbox のプラン・データ量・契約更新時期", "個人情報を含むデータの範囲と「ロック」の操作", "テスト導入8名の業務と現在の使い方", "既存システム・お取引先との役割分担", "ご予算感・ご決裁の時期・定例の頻度"];
    qs.forEach((q, i) => {
      const y = 1.85 + i * 0.83;
      t(s, String(i + 1), { x: MX, y, w: 0.5, h: 0.6, fontSize: 20, bold: true, color: C.accent1, valign: "middle" });
      t(s, q, { x: MX + 0.6, y, w: lw - 0.6, h: 0.6, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
      if (i < qs.length - 1) hline(s, MX, y + 0.72, lw);
    });
    const px = MX + lw + 0.4, pw = CW - lw - 0.4;
    box(s, { x: px, y: 1.8, w: pw, h: 5.05, fill: { color: C.text2 } });
    t(s, "今後の進め方", { x: px + 0.35, y: 2.1, w: pw - 0.7, h: 0.4, fontSize: 14, bold: true, color: C.accent2 });
    [["本日", "課題・優先度のすり合わせ"], ["1週間以内", "お見積りのご提示"], ["ご決裁後", "キックオフ・現状調査"]].forEach(([w, d], i) => {
      const y = 2.75 + i * 1.3;
      t(s, w, { x: px + 0.35, y, w: pw - 0.7, h: 0.4, fontSize: 13, bold: true, color: C.accent2 });
      t(s, d, { x: px + 0.35, y: y + 0.42, w: pw - 0.7, h: 0.5, fontSize: 17, bold: true, color: C.background1 });
    });
  }

  // ======================================================================
  section("Appendix");
  {
    const s = slide("参考：エディション別の主な機能", "使える機能はエディションで異なるため、対象者に応じた組み合わせもご提案できます");
    const hdr = ["機能", "Business\nStarter", "Business\nStandard", "Business\nPlus", "Enterprise"];
    const data = [
      ["月額料金（年間プラン・1ユーザー・税別）", "800円", "1,600円", "2,500円", "要問合せ"],
      ["Gmail の Gemini（文章作成・要約）", "○", "○", "○", "○"],
      ["Gemini アプリ・Gem・Gemini Notebook", "○", "○", "○", "○"],
      ["ドキュメント／スプレッドシート／スライドの Gemini", "－", "○", "○", "○"],
      ["Meet の自動メモ（議事録作成）", "－", "○", "○", "○"],
      ["Workspace Studio（業務の自動化）", "○", "○", "○", "○"],
      ["組織外共有の制限・リンク共有の禁止", "○", "○", "○", "○"],
      ["共有ドライブの詳細設定（外部共有・DL 禁止）", "－", "○", "○", "○"],
      ["Vault（データの保持・検索）", "－", "－", "○", "○"],
      ["DLP（情報漏えい防止）・信頼ルール", "－", "－", "－", "○"],
    ];
    const rows = [hdr.map((h, i) => th(h, { align: i ? "center" : "left" }))].concat(data.map((r) => r.map((v, j) => td(v, j ? { align: "center", bold: v === "○", color: v === "○" ? HEX.accent1 : v === "－" ? HEX.accent6 : HEX.dk1 } : { bold: true, color: HEX.dk2 }))));
    tbl(s, rows, { x: MX, y: 1.8, w: CW, colW: [5.333, 1.7, 1.7, 1.7, 1.7], rowH: [0.55].concat(data.map(() => 0.44)) });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, L.THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
