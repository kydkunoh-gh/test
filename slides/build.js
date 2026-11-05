// エアシャワー分類スライドの生成スクリプト
// 実行: NODE_PATH=<pptxgenjsのnode_modules> node build.js
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SCRIPTS
  ? path.join(process.env.PPTX_SCRIPTS, "apply_theme.js")
  : "./apply_theme.js");

// 色は黒・グレーと、見出し用の1色だけ
const THEME = {
  name: "AirShowerSimple",
  headFontFace: "Meiryo",
  bodyFontFace: "Meiryo",
  colors: {
    dk1: "262626", lt1: "FFFFFF", dk2: "1F4E5A", lt2: "F2F2F2",
    accent1: "1F4E5A", accent2: "7F7F7F", accent3: "BFBFBF",
    accent4: "D9D9D9", accent5: "595959", accent6: "A6A6A6",
    hlink: "1F4E5A", folHlink: "595959",
  },
};

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
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.7,
          fontSize: 30, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.1, w: 12.1, h: 0.4,
          fontSize: 16, color: C.accent5, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent2, align: "right" },
  });

  // ---------- Slide 1: 分類マトリクス ----------
  pres.addSection({ title: "分類" });
  const s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "分類" });
  s.addText("エアシャワーの4分類", { placeholder: "title" });
  s.addText("「何を除塵するか（人／物）」と「何を守るか（中／外）」の2つの問いで分ける", { placeholder: "body" });

  const border = { type: "solid", pt: 0.75, color: THEME.colors.accent3 };
  const head = (t, sub) => ({
    text: [
      { text: t, options: { bold: true, fontSize: 16, breakLine: true } },
      { text: sub, options: { fontSize: 12, color: C.accent5 } },
    ],
    options: { fill: { color: C.background2 }, align: "center", valign: "middle" },
  });
  const rowHead = t => ({ text: t, options: { bold: true, fontSize: 16, fill: { color: C.background2 }, align: "center", valign: "middle" } });
  const cell = (n, t, d) => ({
    text: [
      { text: n + " " + t, options: { bold: true, fontSize: 15, color: C.text2, breakLine: true, paraSpaceAfter: 6 } },
      { text: d, options: { fontSize: 13, color: C.text1 } },
    ],
    options: { valign: "top" },
  });

  s.addTable([
    [{ text: "", options: { fill: { color: C.background1 } } },
      head("中を守る", "持ち込み防止"), head("外を守る", "持ち出し防止")],
    [rowHead("人用"),
      cell("①", "クリーンルーム・食品工場の入口用", "人に付いた塵埃を室内へ持ち込ませない。最も一般的"),
      cell("②", "高濃度粉じん用", "焼却施設・アスベスト現場の有害粉じんを外へ持ち出させない")],
    [rowHead("物用"),
      cell("③", "物品用", "コンベア型・台車用・エアシャワー付パスボックス"),
      cell("④", "汚染区域からの搬出用", "汚染区域から出す機材・資材の除塵。製品例は少ない")],
  ], {
    objectName: "matrix", x: 0.6, y: 1.85, w: 7.7, colW: [1.1, 3.3, 3.3], rowH: [0.85, 1.75, 1.75],
    border, color: C.text1, margin: [0.12, 0.15, 0.12, 0.15],
  });

  // 仕様の違い（枠なし、余白で区切る）
  const px = 8.9, pw = 12.73 - px;
  s.addText("仕様の違い", { isTextBox: true, objectName: "specHead", x: px, y: 1.85, w: pw, h: 0.45,
    fontSize: 18, bold: true, color: C.text1, margin: 0, valign: "middle" });
  s.addText("4区分に共通する選択肢として扱う", { isTextBox: true, objectName: "specSub", x: px, y: 2.3, w: pw, h: 0.35,
    fontSize: 12, color: C.accent5, margin: 0, valign: "middle" });
  const specs = [
    ["形", "箱型／トンネル（多連結）型／L型・T型／扉なし開放型／置き型・ポータブル／分割型／屋外型"],
    ["扉", "手動／自動／シートシャッター（防虫）"],
    ["材質", "鋼板焼付塗装／ステンレス（食品工場の湿式床）"],
    ["付加機能", "除菌（スモーク・二酸化塩素）／陰圧（抗菌型）／イオナイザー／靴底専用（エアブロー・乾燥機・吸着マット）"],
  ];
  const runs = [];
  specs.forEach(([k, v], i) => {
    runs.push({ text: k, options: { bold: true, color: C.text2, fontSize: 14, breakLine: true } });
    runs.push({ text: v, options: { color: C.text1, fontSize: 13, breakLine: i < specs.length - 1, paraSpaceAfter: 12 } });
  });
  s.addText(runs, { isTextBox: true, objectName: "specBody", x: px, y: 2.8, w: pw, h: 3.9,
    valign: "top", margin: 0, lineSpacingMultiple: 1.1 });

  s.addText("出典：愛豊工業 エアシャワー議論の引継ぎメモ（2026-11-05）", {
    isTextBox: true, objectName: "source", x: 0.6, y: 7.0, w: 11.3, h: 0.3, fontSize: 10, color: C.accent2, margin: 0, valign: "middle" });

  s.addNotes("2つの問いで分ける。何を除塵するか（人／物）、何を守るか（中＝持ち込み防止／外＝持ち出し防止）。形・扉・材質・付加機能は4区分の中の仕様の違いとして扱う。");

  const out = "エアシャワー分類.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
}
main().catch(e => { console.error(e); process.exit(1); });
