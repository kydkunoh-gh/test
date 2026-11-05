// エアシャワー分類スライドの生成スクリプト
// 実行: PPTX_SCRIPTS=<pptxスキルのscriptsフォルダ> NODE_PATH=<pptxgenjs・sharpのnode_modules> node build.js
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip");
const sharp = require("sharp");
const path = require("path");
const fs = require("fs");
const { applyTheme } = require(path.join(process.env.PPTX_SCRIPTS || __dirname, "apply_theme.js"));

const FONT = "HGPｺﾞｼｯｸM";
const IMG = path.join(__dirname, "img");
const OUT = path.join(__dirname, "エアシャワー分類.pptx");

// 色は黒・グレーと、アクセントの1色だけ
const THEME = {
  name: "AirShowerSimple",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: "262626", lt1: "FFFFFF", dk2: "1F4E5A", lt2: "F3F4F4",
    accent1: "1F4E5A", accent2: "8C8C8C", accent3: "BFBFBF",
    accent4: "DADADA", accent5: "595959", accent6: "A6A6A6",
    hlink: "1F4E5A", folHlink: "595959",
  },
};

// ---------------- 具体例（内容は各URLのページで確認したもの） ----------------
const EXAMPLES = [
  {
    q: 0, title: "① 人用 × 持ち込み防止",
    lead: "クリーンルーム・食品工場の入口用。同じ区分の中でも、形・扉・材質で品種が分かれる",
    items: [
      { img: "p1a.png", tag: "標準機", maker: "日本エアーテック", name: "標準型エアーシャワー\nEAS-17型",
        desc: "天井ノズルで上から除塵。多連結・L字／T字・自動ドア・シートシャッター内蔵に対応",
        url: "https://mono.ipros.com/product/detail/2000740417" },
      { img: "p1b.png", tag: "形：L型・T型", maker: "日本エアーテック", name: "L型／T型エアーシャワー",
        desc: "出入口を直角（L型）や三方（T型）に配置。部屋の配置に合わせて向きを変える",
        url: "https://mono.ipros.com/product/detail/2000740482" },
      { img: "p1c.jpg", tag: "扉：シートシャッター", maker: "ニットー冷熱製作所", name: "シートシャッター型\nNAS-T161PSL2-AU",
        desc: "高速で自動開閉するオレンジ色のシートで虫の侵入を抑える。食品工場向け",
        url: "https://pr.mono.ipros.com/takatsu/product/detail/2001637962/" },
      { img: "p1d.jpg", tag: "材質：ステンレス", maker: "日立産機システム", name: "エアシャワー PCJ-S88JSM4",
        desc: "筐体をステンレス（SUS304）にして耐食性を高めた。湿式床の食品工場向け",
        url: "https://www.hitachi.com/ja-jp/press/articles/2019/08/0826/" },
    ],
  },
  {
    q: 1, title: "② 人用 × 持ち出し防止",
    lead: "焼却施設・アスベスト現場などの高濃度粉じん用。有害な粉じんを外へ持ち出させない",
    items: [
      { img: "p2a.png", tag: "清掃工場用", maker: "アクティオ", name: "清掃工場用エアシャワー\nJS-30",
        desc: "ダイオキシンなどの有害粉塵を工場の外へ持ち出さない。クリーンルームでは使用不可と明記",
        url: "https://www.eg.aktio.co.jp/wp-content/uploads/cata_JS-30.pdf" },
      { img: "p2b.jpg", tag: "形：分割型", maker: "アゼアス", name: "小型分割式エアシャワー\nVAS-0004C",
        desc: "3分割でき、狭い現場にも搬入・設置できる。アスベスト・石綿対策向け",
        url: "https://mono.ipros.com/product/detail/2001257139" },
      { img: "p2c.jpg", tag: "レンタル", maker: "流機エンジニアリング", name: "エアシャワー AS-1／AS-2\n（設置事例）",
        desc: "工事現場で作業服の粉じんを30m/sのジェットで払い落とし、二次飛散を抑えた",
        url: "https://pr.mono.ipros.com/ryuki/product/detail/2000777521/" },
      { img: "p2e.jpg", tag: "ダブルノズル", maker: "愛豊工業", name: "ダブルノズル\nエアシャワールーム",
        desc: "清浄空気と高圧パルスエアを組み合わせ、防護服・保護具に付いた粉じんを吹き飛ばす",
        url: "https://www.ipros.com/en/product/detail/2001395659/" },
    ],
  },
  {
    q: 2, title: "③ 物用 × 持ち込み防止",
    lead: "物品に付いた塵埃を、清浄な区域へ持ち込ませない。コンベア・台車・受け渡し口と組み合わせる",
    items: [
      { img: "p3a.png", tag: "コンベア連動", maker: "三宝電機", name: "物品用エアシャワー（画像右）",
        desc: "入口・出口のコンベアと連動し、搬入から除塵、搬出までを自動化する",
        url: "https://www.ipros.com/product/detail/2000692576" },
      { img: "p3b.jpg", tag: "コンベア・台車対応", maker: "ジェピック", name: "コンベアシャワー",
        desc: "コンベア・ローラー・台車に載せたまま、製品や資材に付いた異物をジェット風で除去",
        url: "https://mono.ipros.com/product/detail/2000411761" },
      { img: "p3c.jpg", tag: "パスボックス", maker: "ニットー冷熱製作所", name: "エアシャワー付パスボックス\nNPB-868J",
        desc: "受け渡し口に4個のジェットノズルを付け、受け渡す物品をさらに清浄化する（画像はシリーズ）",
        url: "https://mono.ipros.com/product/detail/2000148223" },
      { img: "p3d.png", tag: "形：多連結", maker: "日本エアーテック", name: "2連大型エアーシャワー",
        desc: "2台を連結した大空間型。搬入物と同時に入退室でき、台車の通過にも使われる",
        url: "https://mono.ipros.com/product/detail/2000740435" },
    ],
  },
  {
    q: 3, title: "④ 物用 × 持ち出し防止",
    lead: "汚染区域から出す物品の除塵。専用の製品は少なく、用途の一つとして挙げる例にとどまる",
    items: [
      { img: "p4a.jpg", tag: "用途例に「ワーク持ち出し防止」", maker: "アイディー技研", name: "エアシャワー",
        desc: "一般入室用のほか、コンベア組込みのワーク受け入れ用・台車搬入用にも対応。用途例に「ワーク持ち出し防止」とある",
        url: "https://mono.ipros.com/product/detail/287896004" },
      { img: "p4b.jpg", tag: "汚染物質の持ち出し防止を併記", maker: "セギジャパン", name: "オールステンレス製エアーシャワー\nST-AS02",
        desc: "貨物・台車の通用口の防塵・防虫と、室内の粉塵・汚染物質を外に持ち出さない用途を併記。物の持ち出し専用設計かは不明",
        url: "https://mono.ipros.com/product/detail/2000158603" },
    ],
    note: {
      head: "搬出専用の製品は\n見つからなかった",
      body: "焼却施設・アスベスト現場から機材・資材を搬出するための専用エアシャワーは見つからなかった。",
      body2: "現場では一般に、HEPA掃除機での清掃、湿らせた布での拭き取り、廃棄物の二重袋詰めで対応していると考えられる（推定）。",
      foot: "調査は網羅的ではない",
    },
  },
];

// 画像の出典（掲載者「ページの題名」掲載サイト）。画像は縦横比を保ち、切り抜き・加工はしない
const SRC = {
  "p1a.png": "日本エアーテック「標準型エアーシャワー『EAS-17型』」イプロスものづくり",
  "p1b.png": "日本エアーテック「L型/T型エアーシャワー」イプロスものづくり",
  "p1c.jpg": "高津伝動精機「エアシャワー シートシャッター型 NAS-T161PSL2-AU」イプロスものづくり",
  "p1d.jpg": "日立製作所「エアシャワー「PCJ-S88JSM4」を販売開始」ニュースリリース（2019年8月26日）",
  "p2a.png": "アクティオ「エアシャワー JS-30」製品カタログ（PDF）の掲載写真",
  "p2b.jpg": "アゼアス「エアシャワー 小型分割式 3分割可能」イプロスものづくり",
  "p2c.jpg": "流機エンジニアリング「【設置事例】作業者の作業服に付着した粉じんを払い落としたい」イプロスものづくり",
  "p2e.jpg": "愛豊工業「Double Nozzle Air Shower Room」IPROS",
  "p3a.png": "三宝電機「標準型エアシャワー／物品用エアシャワー」イプロス",
  "p3b.jpg": "ジェピック「コンベア対応型エアーシャワー『コンベアシャワー』」イプロスものづくり",
  "p3c.jpg": "ニットー冷熱製作所「クリーン機器 パスボックスシリーズ」イプロスものづくり",
  "p3d.png": "日本エアーテック「2連大型エアーシャワー」イプロスものづくり",
  "p4a.jpg": "アイディー技研「エアシャワー」イプロスものづくり",
  "p4b.jpg": "セギジャパン「オールステンレス製エアーシャワー手動扉 ST-AS02」イプロスものづくり",
};
const SOURCE_EX = "画像は各社が公表した資料からの引用（出典は各画像の下）。説明文は各資料をもとに作成。いずれも2026年11月5日閲覧";

// ---------------- 共通パーツ ----------------
function hairline(s, C, x, y, w, color, pt, name) {
  s.addShape("line", { objectName: name, x, y, w, h: 0, line: { color, width: pt } });
}
function vline(s, C, x, y, h, color, pt, name) {
  s.addShape("line", { objectName: name, x, y, w: 0, h, line: { color, width: pt } });
}

// 右上の位置表示（2×2の小さなマス。いまの区分だけ塗る）
function locator(s, C, q) {
  const size = 0.26, g = 0.05, x0 = 12.73 - (size * 2 + g), y0 = 0.5;
  for (let i = 0; i < 4; i++) {
    const x = x0 + (i % 2) * (size + g), y = y0 + Math.floor(i / 2) * (size + g);
    const on = i === q;
    s.addText(String(i + 1), { isTextBox: true, objectName: `loc${i + 1}`, x, y, w: size, h: size,
      shape: "rect", fill: { color: on ? C.accent1 : C.background2 }, line: { color: on ? C.accent1 : C.accent4, width: 0.5 },
      fontSize: 9, color: on ? C.background1 : C.accent2, align: "center", valign: "middle", margin: 0 });
  }
}

async function fitImage(s, file, box, name, alt, valign = "bottom") {
  const { width, height } = await sharp(file).metadata();
  const scale = Math.min(box.w / width, box.h / height);
  const w = width * scale, h = height * scale;
  const x = box.x + (box.w - w) / 2;
  const y = valign === "bottom" ? box.y + box.h - h : box.y + (box.h - h) / 2;
  s.addImage({ path: file, objectName: name, x, y, w, h, altText: alt });
}

// 製品カードの文字部分は1つの表にする（列＝製品、行＝タグ／会社名・製品名／説明／出典）。
// 文字を増減しても行の高さが自動で伸び縮みし、各製品の行がそろったまま編集できる
function cardTable(s, C, items, x0, w, gap, L) {
  const none = { type: "none" };
  const cell = (text, opts = {}) => ({ text, options: { valign: "top", margin: 0, border: [none, none, none, none], ...opts } });
  const gapCell = () => cell("");
  const row = f => items.flatMap((it, i) => (i ? [gapCell(), f(it)] : [f(it)]));
  const rows = [
    row(it => cell(it.tag, { fontSize: 12, bold: true, color: C.accent1 })),
    row(it => cell([
      { text: it.maker, options: { fontSize: 11, color: C.accent5, breakLine: true } },
      ...it.name.split("\n").map((t, k, a) => ({ text: t, options: { fontSize: 14, bold: true, color: C.text1, breakLine: k < a.length - 1 } })),
    ])),
    row(it => cell(it.desc, { fontSize: 13, color: C.text1, lineSpacingMultiple: 1.1 })),
    row(it => {
      if (!SRC[it.img]) throw new Error("出典がない画像: " + it.img);
      return cell([
        { text: "画像出典：" + SRC[it.img], options: { color: C.accent5, breakLine: true } },
        { text: it.url, options: { color: C.accent1, hyperlink: { url: it.url, tooltip: it.url } } },
      ], { fontSize: 9, lineSpacingMultiple: 1.15 });
    }),
  ];
  const colW = items.flatMap((_, i) => (i ? [gap, w] : [w]));
  s.addTable(rows, { objectName: "productTable", x: x0, y: L.tableY, w: colW.reduce((a, b) => a + b, 0), colW,
    rowH: L.rowH, color: C.text1, border: none, fill: { color: C.background1 } });
}

async function exampleSlide(pres, C, ex) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "具体例" });
  s.addText("具体例", { placeholder: "kicker" });
  s.addText(ex.title, { placeholder: "title" });
  s.addText(ex.lead, { placeholder: "body" });
  locator(s, C, ex.q);

  const L = { imgY: 1.9, imgH: 1.9, shelfY: 3.9, tableY: 4.0, rowH: [0.33, 0.7, 0.9, 0.9] };
  const x0 = 0.6, gap = 0.35;
  const areaW = ex.note ? 7.4 : 12.13;
  const n = ex.items.length, w = (areaW - gap * (n - 1)) / n;
  for (let i = 0; i < n; i++) {
    const it = ex.items[i], x = x0 + i * (w + gap);
    await fitImage(s, path.join(IMG, it.img), { x: x + 0.1, y: L.imgY, w: w - 0.2, h: L.imgH }, `img${i + 1}`, `${it.maker} ${it.name.replace(/\n/g, " ")}`);
  }
  hairline(s, C, x0, L.shelfY, areaW, C.accent3, 0.75, "shelf");
  cardTable(s, C, ex.items, x0, w, gap, L);

  if (ex.note) {
    const nx = 8.45, nw = 12.73 - nx, ny = 1.95, nh = 4.9;
    s.addShape("rect", { objectName: "notePanel", x: nx, y: ny, w: nw, h: nh, fill: { color: C.background2 }, line: { color: C.background2, width: 0 } });
    s.addText(ex.note.head, { isTextBox: true, objectName: "noteHead", x: nx + 0.35, y: ny + 0.35, w: nw - 0.7, h: 0.75,
      fontSize: 18, bold: true, color: C.text1, margin: 0, valign: "top" });
    s.addText([
      { text: ex.note.body, options: { breakLine: true, paraSpaceAfter: 10 } },
      { text: ex.note.body2 },
    ], { isTextBox: true, objectName: "noteBody", x: nx + 0.35, y: ny + 1.25, w: nw - 0.7, h: 2.9,
      fontSize: 14, color: C.text1, margin: 0, valign: "top", lineSpacingMultiple: 1.15 });
    s.addText(ex.note.foot, { isTextBox: true, objectName: "noteFoot", x: nx + 0.35, y: ny + nh - 0.6, w: nw - 0.7, h: 0.3,
      fontSize: 11, color: C.accent5, margin: 0, valign: "middle" });
  }
  s.addText(SOURCE_EX, { placeholder: "source" });
  s.addNotes(`${ex.title}。${ex.lead}。` + ex.items.map(it => `${it.maker}「${it.name.replace(/\n/g, " ")}」：${it.desc}（${it.url}）`).join(" "));
}

// ---------------- 1枚目：分類マトリクス ----------------
function matrixSlide(pres, C) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "分類" });
  s.addText("分類", { placeholder: "kicker" });
  s.addText("エアシャワーの4分類", { placeholder: "title" });
  s.addText("「何を除塵するか（人／物）」と「何を守るか（中／外）」の2つの問いで分ける", { placeholder: "body" });

  // 罫線（表の罫線は16進の色で指定する）
  const H = THEME.colors;
  const none = { type: "none" };
  const thin = { type: "solid", pt: 0.75, color: H.accent4 };
  const thick = { type: "solid", pt: 1.25, color: H.dk1 };

  // 2×2のマトリクスを1つの表にする（列：行見出し／番号／内容／番号／内容）
  const headY = 2.0, headH = 0.7, bodyH = 2.0;
  const head = (a, b, padL) => ({
    text: [
      { text: a, options: { fontSize: 18, bold: true, color: C.text1 } },
      { text: "　" + b, options: { fontSize: 13, color: C.accent5 } },
    ],
    options: { colspan: 2, valign: "bottom", margin: [0, 0, 0.1, padL], border: [none, none, thick, none] },
  });
  const rowHead = (t, topB) => ({ text: t, options: { fontSize: 18, bold: true, color: C.text1, valign: "middle", margin: 0,
    border: [topB, none, thin, none] } });
  const num = (n, divider, topB) => ({ text: n, options: { fontSize: 40, color: C.accent1, valign: "top",
    margin: [0.25, 0, 0, divider ? 0.15 : 0], border: [topB, none, thin, divider ? thin : none] } });
  const body = (t, d, topB) => ({
    text: [
      { text: t, options: { fontSize: 16, bold: true, color: C.text1, breakLine: true, paraSpaceAfter: 8 } },
      { text: d, options: { fontSize: 13, color: C.accent5, lineSpacingMultiple: 1.1 } },
    ],
    options: { valign: "top", margin: [0.32, 0.05, 0.1, 0.05], border: [topB, none, thin, none] },
  });
  s.addTable([
    [{ text: "", options: { border: [none, none, thick, none] } },
      head("中を守る", "持ち込み防止", 0), head("外を守る", "持ち出し防止", 0.15)],
    [rowHead("人用", thick),
      num("1", false, thick), body("クリーンルーム・食品工場の入口用", "人に付いた塵埃を室内へ持ち込ませない。最も一般的な形", thick),
      num("2", true, thick), body("高濃度粉じん用", "焼却施設・アスベスト現場の有害な粉じんを外へ持ち出させない", thick)],
    [rowHead("物用", thin),
      num("3", false, thin), body("物品用", "コンベア型・台車用・エアシャワー付パスボックス", thin),
      num("4", true, thin), body("汚染区域からの搬出用", "汚染区域から出す機材・資材の除塵。製品例は少ない", thin)],
  ], { objectName: "matrixTable", x: 0.6, y: headY, w: 7.8, colW: [1.2, 0.6, 2.6, 0.75, 2.65],
    rowH: [headH, bodyH, bodyH], color: C.text1, fill: { color: C.background1 } });

  // 右：仕様の違い（見出し＋1列の表）
  const px = 9.0, pw = 12.73 - px;
  s.addText([
    { text: "仕様の違い", options: { fontSize: 18, bold: true, color: C.text1, breakLine: true } },
    { text: "4区分に共通する選択肢", options: { fontSize: 12, color: C.accent5 } },
  ], { isTextBox: true, objectName: "specHead", x: px, y: headY - 0.1, w: pw, h: headH, margin: [0, 0, 7, 0], valign: "bottom" });
  const specs = [
    ["形", "箱型／トンネル（多連結）型／L型・T型／扉なし開放型／置き型・ポータブル／分割型／屋外型", 1.2],
    ["扉", "手動／自動／シートシャッター（防虫）", 0.75],
    ["材質", "鋼板焼付塗装／ステンレス（食品工場の湿式床）", 0.85],
    ["付加機能", "除菌（スモーク・二酸化塩素）／陰圧（抗菌型）／イオナイザー／靴底専用（エアブロー・乾燥機・吸着マット）", 1.2],
  ];
  s.addTable(specs.map(([k, v], i) => [{
    text: [
      { text: k, options: { fontSize: 12, bold: true, color: C.accent1, breakLine: true } },
      { text: v, options: { fontSize: 13, color: C.text1, lineSpacingMultiple: 1.1 } },
    ],
    options: { valign: "top", margin: [0.12, 0, 0.08, 0],
      border: [i === 0 ? thick : thin, none, i === specs.length - 1 ? none : thin, none] },
  }]), { objectName: "specTable", x: px, y: headY + headH, w: pw, colW: [pw], rowH: specs.map(r => r[2]),
    color: C.text1, fill: { color: C.background1 } });

  s.addText("出典：愛豊工業 エアシャワー議論の引継ぎメモ（2026-11-05）", { placeholder: "source" });
  s.addNotes("2つの問いで分ける。何を除塵するか（人／物）、何を守るか（中＝持ち込み防止／外＝持ち出し防止）。形・扉・材質・付加機能は4区分の中の仕様の違いとして扱う。");
}

// テーマの日本語フォント（ea と Jpan）も HGPｺﾞｼｯｸM にする
async function setJapaneseThemeFont(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/theme\/theme\d+\.xml$/.test(n))) {
    let xml = await zip.file(name).async("string");
    xml = xml.replace(/<a:(majorFont|minorFont)>([\s\S]*?)<\/a:\1>/g, (all, tag, inner) => {
      inner = inner
        .replace(/<a:latin typeface="[^"]*"/, `<a:latin typeface="${FONT}"`)
        .replace(/<a:ea typeface="[^"]*"/, `<a:ea typeface="${FONT}"`)
        .replace(/<a:font script="Jpan" typeface="[^"]*"/, `<a:font script="Jpan" typeface="${FONT}"`);
      return `<a:${tag}>${inner}</a:${tag}>`;
    });
    zip.file(name, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "エアシャワーの分類";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.45, w: 9, h: 0.3,
          fontSize: 12, bold: true, color: C.accent1, align: "left", valign: "middle", margin: 0, charSpacing: 3 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.75, w: 10.8, h: 0.65,
          fontSize: 30, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.4, w: 11.0, h: 0.4,
          fontSize: 15, color: C.accent5, align: "left", valign: "middle", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "source", type: "body", x: 0.6, y: 7.0, w: 11.0, h: 0.28,
          fontSize: 10, color: C.accent2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.28, fontSize: 10, color: C.accent2, align: "right" },
  });

  pres.addSection({ title: "分類" });
  matrixSlide(pres, C);
  pres.addSection({ title: "具体例" });
  for (const ex of EXAMPLES) await exampleSlide(pres, C, ex);

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  await setJapaneseThemeFont(OUT);
  console.log("wrote", OUT);
}
main().catch(e => { console.error(e); process.exit(1); });
