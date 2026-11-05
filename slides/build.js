const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaUserAlt, FaBoxOpen } = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/2a88dc59-cd8e-4ef9-97fa-37ebf681ddc1_6b56ea32-91ed-4cc8-aae1-78a582c3da03/pptx/scripts/apply_theme.js");

const THEME = {
  name: "AirShower",
  headFontFace: "Meiryo",
  bodyFontFace: "Meiryo",
  colors: {
    dk1: "1F2A30", lt1: "FFFFFF", dk2: "0F4C5C", lt2: "E8F1F2",
    accent1: "1A7F8E", accent2: "E36414", accent3: "7A8C91",
    accent4: "F6D5BF", accent5: "5BA4AE", accent6: "C9D6D9",
    hlink: "1A7F8E", folHlink: "5F7A80",
  },
};

async function icon(Comp, hex) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "エアシャワーの分類";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "TITLE_ONLY",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.75,
          fontSize: 32, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.05, w: 12.1, h: 0.45,
          fontSize: 16, color: C.accent3, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent3, align: "right" },
  });

  const iconPerson = await icon(FaUserAlt, "FFFFFF");
  const iconBox = await icon(FaBoxOpen, "FFFFFF");

  // ---------- Slide 1: 分類マトリクス ----------
  pres.addSection({ title: "分類" });
  const s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "分類" });
  s.addText("エアシャワーは4つに分かれる", { placeholder: "title" });
  s.addText("「何を除塵するか（人／物）」×「何を守るか（中／外）」の2つの問いで分類", { placeholder: "body" });

  const mx = 0.6, my = 1.7;          // matrix origin
  const rowHeadW = 1.35, colHeadH = 0.8, gap = 0.1;
  const cellW = 3.2, cellH = 2.0;
  const c1x = mx + rowHeadW + gap, c2x = c1x + cellW + gap;
  const r1y = my + colHeadH + gap, r2y = r1y + cellH + gap;

  // corner
  s.addText([
    { text: "守るのは →", options: { breakLine: true } },
    { text: "除塵するのは ↓" },
  ], { isTextBox: true, objectName: "corner", x: mx, y: my, w: rowHeadW, h: colHeadH,
       fontSize: 10, color: C.accent3, align: "right", valign: "middle", margin: 4 });

  // column headers
  [["中を守る", "持ち込み防止", c1x], ["外を守る", "持ち出し防止", c2x]].forEach(([a, b, x], i) => {
    s.addText([
      { text: a, options: { fontSize: 18, bold: true, breakLine: true } },
      { text: "（" + b + "）", options: { fontSize: 13 } },
    ], { isTextBox: true, objectName: "colHead" + (i + 1), x, y: my, w: cellW, h: colHeadH,
         shape: pres.ShapeType.roundRect, rectRadius: 0.08, fill: { color: C.text2 },
         color: C.background1, align: "center", valign: "middle", margin: 2 });
  });

  // row headers
  [["人用", iconPerson, r1y], ["物用", iconBox, r2y]].forEach(([label, ic, y], i) => {
    s.addShape(pres.ShapeType.roundRect, { objectName: "rowHead" + (i + 1), x: mx, y, w: rowHeadW, h: cellH,
      rectRadius: 0.08, fill: { color: C.text2 }, line: { color: C.text2 } });
    s.addImage({ data: ic, objectName: "rowIcon" + (i + 1), x: mx + rowHeadW / 2 - 0.25, y: y + 0.55, w: 0.5, h: 0.5 });
    s.addText(label, { isTextBox: true, objectName: "rowLabel" + (i + 1), x: mx, y: y + 1.15, w: rowHeadW, h: 0.45,
      fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  });

  // cells
  const cells = [
    { n: "①", x: c1x, y: r1y, t: "クリーンルーム・食品工場の入口用",
      d: "人に付いた塵埃を室内へ持ち込ませない。最も一般的", tag: "愛豊工業：手がける", tagStyle: "outline" },
    { n: "②", x: c2x, y: r1y, t: "高濃度粉じん用",
      d: "焼却施設・アスベスト現場の有害粉じんを外へ持ち出させない", tag: "愛豊工業：主力", tagStyle: "solid", hl: true },
    { n: "③", x: c1x, y: r2y, t: "物品用",
      d: "コンベア型・台車用・エアシャワー付パスボックス", tag: "愛豊工業：公開情報で未確認", tagStyle: "muted" },
    { n: "④", x: c2x, y: r2y, t: "汚染区域からの搬出用",
      d: "汚染区域から出す機材・資材の除塵。製品例は少ない", tag: "愛豊工業：公開情報で未確認", tagStyle: "muted" },
  ];
  cells.forEach((c, i) => {
    s.addShape(pres.ShapeType.roundRect, { objectName: "cell" + (i + 1), x: c.x, y: c.y, w: cellW, h: cellH,
      rectRadius: 0.08, fill: { color: C.background2 },
      line: c.hl ? { color: C.accent2, width: 2.5 } : { color: C.background2, width: 0.75 } });
    s.addShape(pres.ShapeType.ellipse, { objectName: "num" + (i + 1), x: c.x + 0.2, y: c.y + 0.2, w: 0.5, h: 0.5,
      fill: { color: c.hl ? C.accent2 : C.accent1 }, line: { color: c.hl ? C.accent2 : C.accent1 } });
    s.addText(c.n.replace(/[①-④]/, m => String("①②③④".indexOf(m) + 1)), { isTextBox: true, objectName: "numText" + (i + 1),
      x: c.x + 0.2, y: c.y + 0.2, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: C.background1,
      align: "center", valign: "middle", margin: 0 });
    s.addText(c.t, { isTextBox: true, objectName: "cellTitle" + (i + 1), x: c.x + 0.8, y: c.y + 0.15, w: cellW - 0.95, h: 0.6,
      fontSize: 15, bold: true, color: C.text2, valign: "middle", margin: 0 });
    s.addText(c.d, { isTextBox: true, objectName: "cellDesc" + (i + 1), x: c.x + 0.2, y: c.y + 0.85, w: cellW - 0.4, h: 0.7,
      fontSize: 13, color: C.text1, valign: "top", margin: 0 });
    const tagFill = c.tagStyle === "solid" ? { color: C.accent2 } : { color: C.background1 };
    const tagLine = c.tagStyle === "solid" ? { color: C.accent2 } : c.tagStyle === "outline" ? { color: C.accent2, width: 1 } : { color: C.accent6, width: 1 };
    const tagColor = c.tagStyle === "solid" ? C.background1 : c.tagStyle === "outline" ? C.accent2 : C.accent3;
    s.addText(c.tag, { isTextBox: true, objectName: "tag" + (i + 1), x: c.x + 0.2, y: c.y + cellH - 0.5, w: cellW - 0.4, h: 0.32,
      shape: pres.ShapeType.roundRect, rectRadius: 0.16, fill: tagFill, line: tagLine,
      fontSize: 10, bold: true, color: tagColor, align: "center", valign: "middle", margin: 0 });
  });

  // spec panel
  const px = c2x + cellW + 0.4, pw = 12.73 - px, py = my, ph = r2y + cellH - my;
  s.addShape(pres.ShapeType.roundRect, { objectName: "specPanel", x: px, y: py, w: pw, h: ph,
    rectRadius: 0.08, fill: { color: C.background1 }, line: { color: C.accent6, width: 1 } });
  s.addText("仕様の違い", { isTextBox: true, objectName: "specHead", x: px + 0.25, y: py + 0.15, w: pw - 0.5, h: 0.45,
    fontSize: 18, bold: true, color: C.text2, margin: 0, valign: "middle" });
  s.addText("4区分のどれにも共通する選択肢として扱う", { isTextBox: true, objectName: "specSub", x: px + 0.25, y: py + 0.6, w: pw - 0.5, h: 0.35,
    fontSize: 11, color: C.accent3, margin: 0, valign: "middle" });
  const specs = [
    ["形", "箱型／トンネル（多連結）型／L型・T型／扉なし開放型／置き型・ポータブル／分割型／屋外型"],
    ["扉", "手動／自動／シートシャッター（防虫）"],
    ["材質", "鋼板焼付塗装／ステンレス（食品工場の湿式床）"],
    ["付加機能", "除菌（スモーク・二酸化塩素）／陰圧（抗菌型）／イオナイザー／靴底専用（エアブロー・乾燥機・吸着マット）"],
  ];
  const runs = [];
  specs.forEach(([k, v], i) => {
    runs.push({ text: k, options: { bold: true, color: C.accent1, fontSize: 14, breakLine: true } });
    runs.push({ text: v, options: { color: C.text1, fontSize: 13, breakLine: i < specs.length - 1, paraSpaceAfter: 10 } });
  });
  s.addText(runs, { isTextBox: true, objectName: "specBody", x: px + 0.25, y: py + 1.05, w: pw - 0.5, h: ph - 1.2,
    valign: "top", margin: 0, lineSpacingMultiple: 1.1 });

  s.addText("出典：愛豊工業 エアシャワー議論の引継ぎメモ（2026-11-05）。愛豊工業の位置づけは公開情報に基づく", {
    isTextBox: true, objectName: "source", x: 0.6, y: 7.0, w: 11.3, h: 0.3, fontSize: 10, color: C.accent3, margin: 0, valign: "middle" });

  s.addNotes("2つの問いで分ける。①何を除塵するか（人／物）、②何を守るか（中＝持ち込み防止／外＝持ち出し防止）。形・扉・材質・付加機能は4区分の中の仕様の違いとして扱う。愛豊工業の主力は②、①も手がける。③④は公開情報では確認できない。");

  const out = "エアシャワー分類.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
}
main().catch(e => { console.error(e); process.exit(1); });
