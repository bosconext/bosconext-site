/* =====================================================================
   2027年度 ヤマト福祉財団「障がい者福祉助成金」／ 企画書 生成スクリプト
   ---------------------------------------------------------------------
   ★ 文言を直したいときは、下の CONTENT 部分だけを書き換えてください。
   ★ 実行:  node build-kikakusho.js
   ===================================================================== */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType, VerticalAlign,
  Footer, Header, PageNumber, LineRuleType, HeadingLevel, PageBreak
} = require("docx");

/* ---------- 体裁の設定 ---------- */
const JP   = { ascii: "游ゴシック", eastAsia: "游ゴシック", hAnsi: "游ゴシック" };
const NAVY = "1F3864";
const RED  = "B32424";
const GRAY = "595959";
const LINE = "BFBFBF";

const PAGE_W = 11906, MARGIN = 1134;         // A4 縦 / 余白 20mm
const CW = PAGE_W - MARGIN * 2;              // 本文幅 9638 dxa

/* ---------- 部品 ---------- */
const t = (text, o = {}) => new TextRun(Object.assign({ text, font: JP, size: 20 }, o));

function p(text, o = {}) {
  const runs = Array.isArray(text) ? text : [t(text, o.run || {})];
  return new Paragraph({
    children: runs,
    alignment: o.align,
    indent: o.indent,
    spacing: Object.assign({ before: 0, after: 60, line: 280, lineRule: LineRuleType.AUTO }, o.spacing),
    border: o.border,
  });
}

/* 本文（1字下げ） */
const body = (text, o = {}) =>
  p([t(text, o.run || {})], Object.assign({ indent: { firstLine: 200 } }, o));

/* 章見出し  例）1. 事業の概要 */
const h1 = (num, text) => new Paragraph({
  children: [
    t(num + "　", { bold: true, size: 23, color: NAVY }),
    t(text, { bold: true, size: 23, color: NAVY }),
  ],
  spacing: { before: 240, after: 110, line: 268, lineRule: LineRuleType.AUTO },
  border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: NAVY, space: 4 } },
});

/* 小見出し 例）(1) ○○ */
const h2 = (text) => new Paragraph({
  children: [t("▍" + text, { bold: true, size: 21, color: NAVY })],
  spacing: { before: 150, after: 60, line: 268, lineRule: LineRuleType.AUTO },
});

/* 箇条書き（・） */
const li = (text, o = {}) => new Paragraph({
  children: [t("・" + text, o.run || {})],
  indent: { left: 200, hanging: 200 },
  spacing: { after: 30, line: 268, lineRule: LineRuleType.AUTO },
});

/* 入れ子の箇条書き（－） */
const li2 = (text) => new Paragraph({
  children: [t("－ " + text, { size: 19 })],
  indent: { left: 520, hanging: 240 },
  spacing: { after: 30, line: 270, lineRule: LineRuleType.AUTO },
});

/* 注記 */
const note = (text) => new Paragraph({
  children: [t(text, { size: 17, color: GRAY })],
  indent: { left: 200 },
  spacing: { before: 60, after: 80, line: 260, lineRule: LineRuleType.AUTO },
});

const spacer = (h = 120) => new Paragraph({ children: [t("")], spacing: { after: h } });

/* 表のセル */
function cell(content, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((c) =>
    typeof c === "string"
      ? new Paragraph({
          children: [t(c, { bold: o.bold, size: o.size || 19, color: o.color })],
          alignment: o.align,
          spacing: { before: 40, after: 40, line: 260, lineRule: LineRuleType.AUTO },
        })
      : c
  );
  return new TableCell({
    children: paras,
    width: { size: o.w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 110, right: 110 },
  });
}

/* 2列（項目／内容）の表 */
function kvTable(rows, wLabel) {
  const w1 = wLabel || 2000, w2 = CW - w1;
  return new Table({
    columnWidths: [w1, w2],
    width: { size: CW, type: WidthType.DXA },
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 6, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: LINE },
      left:   { style: BorderStyle.SINGLE, size: 6, color: LINE },
      right:  { style: BorderStyle.SINGLE, size: 6, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: LINE },
      insideVertical:   { style: BorderStyle.SINGLE, size: 6, color: LINE },
    },
    rows: rows.map(([k, v]) => new TableRow({
      cantSplit: true,
      children: [
        cell(k, { w: w1, bold: true, fill: "EEF1F7" }),
        cell(v, { w: w2 }),
      ],
    })),
  });
}

/* 見出し行つきの表 */
function gridTable(header, rows, widths) {
  const mk = (arr, isHead) => new TableRow({
    tableHeader: !!isHead,
    cantSplit: true,
    children: arr.map((c, i) =>
      cell(c, { w: widths[i], bold: isHead, fill: isHead ? NAVY : undefined, color: isHead ? "FFFFFF" : undefined, align: isHead ? AlignmentType.CENTER : undefined })
    ),
  });
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 6, color: LINE },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: LINE },
      left:   { style: BorderStyle.SINGLE, size: 6, color: LINE },
      right:  { style: BorderStyle.SINGLE, size: 6, color: LINE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: LINE },
      insideVertical:   { style: BorderStyle.SINGLE, size: 6, color: LINE },
    },
    rows: [mk(header, true)].concat(rows.map((r) => mk(r, false))),
  });
}

/* =====================================================================
   ここから本文（CONTENT）
   ===================================================================== */
const children = [];

/* --- 標題ブロック --- */
children.push(new Paragraph({
  children: [t("2027年度　公益財団法人ヤマト福祉財団　障がい者福祉助成金", { size: 18, color: GRAY })],
  alignment: AlignmentType.RIGHT,
  spacing: { after: 40 },
}));
children.push(new Paragraph({
  children: [t("企　画　書", { bold: true, size: 34, color: NAVY })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 60, after: 60 },
  border: { bottom: { style: BorderStyle.DOUBLE, size: 8, color: NAVY, space: 6 } },
}));
children.push(new Paragraph({
  children: [t("提出日：2026年　　月　　日　／　申請団体：BOSCO NEXT（ボスコ ネクスト）", { size: 18, color: GRAY })],
  alignment: AlignmentType.RIGHT,
  spacing: { before: 80, after: 220 },
}));

/* --- 事業名 --- */
children.push(new Paragraph({
  children: [t("事　業　名", { bold: true, size: 19, color: "FFFFFF" })],
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.CLEAR, color: "auto", fill: NAVY },
  spacing: { after: 0 },
}));
children.push(new Paragraph({
  children: [t("「どんな壁も、乗り越えられる。」", { bold: true, size: 26, color: RED })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 140, after: 40 },
}));
children.push(new Paragraph({
  children: [t("ソーシャルフットボールによる精神障がい当事者の", { bold: true, size: 24 })],
  alignment: AlignmentType.CENTER,
  spacing: { after: 20 },
}));
children.push(new Paragraph({
  children: [t("健康づくり・社会参加促進事業", { bold: true, size: 24 })],
  alignment: AlignmentType.CENTER,
  spacing: { after: 40 },
}));
children.push(new Paragraph({
  children: [t("― 地域に開かれた「はじめの一歩」の場づくりと、全国大会への挑戦 ―", { size: 19, color: GRAY })],
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: LINE, space: 8 } },
}));
children.push(spacer(180));

/* --- 1. 事業の概要 --- */
children.push(h1("1.", "事業の概要"));
children.push(kvTable([
  ["実施主体", "BOSCO NEXT（ボスコ ネクスト）／任意団体・当事者による自主運営"],
  ["事業の種類", "スポーツ活動・文化活動"],
  ["実施期間", "2027年4月1日 ～ 2028年2月29日（11か月）"],
  ["実施場所", "東京都大田区・目黒区・品川区および神奈川県内の体育館・フットサルコート"],
  ["対 象 者", "精神障がいのある方（未経験者を含む）、そのご家族、精神科デイケア・就労移行支援事業所・地域活動支援センター等の利用者および職員"],
  ["事業規模", "オープン体験会 年5回（延べ100名以上）／定例練習 年40回／強化練習会 年12回／ハンドブック400部の制作・配布"],
  ["総事業費", "1,360,000円"],
  ["助成申請額", "1,000,000円　（自己資金 360,000円）"],
], 1700));
children.push(note("※ 詳細は別紙1「スケジュール」および別紙2「費用積算表」をご参照ください。"));

/* --- 2. 申請団体の概要 --- */
children.push(h1("2.", "申請団体の概要"));
children.push(kvTable([
  ["団 体 名", "BOSCO NEXT（ボスコ ネクスト）"],
  ["種　　別", "任意団体（会則に基づく自主運営）"],
  ["設　　立", "2013年4月1日　※前身の「BOSCO」として2012年5月に活動開始"],
  ["代 表 者", "【代表者氏名を記入】"],
  ["所 在 地", "【〒　　－　　　　　　住所を記入】"],
  ["連 絡 先", "bosco_futsal@yahoo.co.jp ／ TEL【　　　　　　　】"],
  ["構 成 員", "選手15名・スタッフ1名（監督）　計16名 ※2026年9月現在"],
  ["活 動 日", "毎週土曜日（年間約40回）ほか、大会・合宿・遠征"],
  ["活動場所", "大田区・目黒区・品川区・神奈川県ほかの体育館、フットサルコート"],
  ["ウェブ", "https://weekendclub.github.io/bosconext/ ／ https://bosconext.jimdofree.com/"],
  ["S N S", "X：@BOSCONEXT ／ Instagram：@bosconext_futsal"],
], 1700));
children.push(spacer(60));
children.push(body("BOSCO NEXTは、精神障がいのある方を中心に、東京都内で活動するソーシャルフットボール（フットサル）のクラブチームです。障害の有無や種別、競技経験の枠にとらわれず、仲間とともに楽しみ、本気で勝利を目指すことを大切にしています。運営はメンバー自身が担う自主運営で、13年にわたり活動を継続してきました。"));

/* --- 3. これまでの活動実績 --- */
children.push(h1("3.", "これまでの活動実績"));
children.push(gridTable(
  ["年　月", "内　容"],
  [
    ["2012年5月", "精神科デイケアのフットサルチーム「BOSCO」として活動開始"],
    ["2013年4月", "デイケア卒業後も続けられる場をつくるため「BOSCO NEXT」として独立・発足"],
    ["2024年", "エスパルスハートフルカップ（静岡）出場"],
    ["2025年9月", "ソーシャルフットボール東京都大会2025　優勝（3試合無失点）"],
    ["2026年2月", "2026年度ソーシャルフットボール強化指定選手にメンバーが選出（JSFA）"],
    ["2026年4月", "長野合宿を実施（いいづなコネクトWEST ほか）"],
    ["2026年5月", "ソーシャルフットボール東京都大会2026　準優勝"],
    ["2026年7月", "JIFF 第6回ソーシャルフットボール全国大会 予選・関東大会　優勝　※得点王・最優秀選手賞もチームから受賞"],
    ["2026年11月", "JIFF 第6回ソーシャルフットボール全国大会（神奈川県）出場"],
  ],
  [1700, CW - 1700]
));
children.push(note("※ 応募要件「2025年4月から1年間以上の活動実績」を満たしています（2013年4月から継続活動）。"));

/* --- 4. 背景と課題 --- */
children.push(h1("4.", "事業の背景と課題"));

children.push(h2("(1) 精神障がいのある人の「からだの健康」という課題"));
children.push(body("精神疾患のある方は、症状や服薬の影響、生活リズムの乱れなどにより運動習慣を持ちにくく、肥満や糖尿病などの生活習慣病を併発しやすいことが、国内外で指摘されています。身体合併症は日々の生活の質を下げるだけでなく、平均余命の短さにもつながると報告されており、「安心して体を動かし続けられる場」は、こころの回復と同じくらい重要な、いのちに関わるテーマです。"));

children.push(h2("(2) 「デイケアを卒業したあと」の居場所が失われる"));
children.push(body("私たちのチームは、まさにこの課題から生まれました。医療機関のデイケアで行われるフットサルは、デイケアの卒業がそのままチームからの離脱を意味します。「卒業してもフットサルを続けたい」という当事者自身の声から、2013年に自分たちの手で立ち上げたのがBOSCO NEXTです。医療・福祉のプログラムの“外側”に、当事者が主体的に参加し続けられる場は、今も決して多くありません。"));

children.push(h2("(3) 「はじめの一歩」が踏み出しにくい"));
children.push(body("運動をしたい気持ちがあっても、「体力が続くか不安」「対人関係が怖い」「うまくできずに気まずくなりそう」といった不安から、一歩を踏み出せない当事者は少なくありません。既存チームの見学・体験も、その場限りの受け入れでは配慮が行き届かず、継続に結びつかないことがあります。必要なのは、失敗しても大丈夫だと体感できる、設計された「入口」です。"));

children.push(h2("(4) 担い手と運営基盤の弱さ"));
children.push(body("ソーシャルフットボールの多くのチームは、当事者と少数のボランティアによる自主運営です。会場使用料・保険料・遠征費といった固定的な支出は会費でまかなっており、活動を地域へ開いていくための余力を持ちにくいのが実情です。加えて近年は体育館・フットサルコートの使用料が上昇し、抽選倍率も高まっており、練習会場の確保そのものが年々難しくなっています。"));

/* --- 5. 目的 --- */
children.push(h1("5.", "事業の目的"));
children.push(body("本事業は、BOSCO NEXTが13年間かけて積み上げてきた「安心して本気になれる場」を、チームの内側から地域へ開くことを目的とします。具体的には、次の3点を実現します。"));
children.push(li("【健康】精神障がいのある方が、運動を安全に継続できる場を年間を通じて確保する。", { run: { bold: false } }));
children.push(li("【参加】デイケア・就労支援等を利用中／卒業後の当事者が、無理なく「はじめの一歩」を踏み出せる体験の入口をつくる。"));
children.push(li("【発信】当事者が自らチームを運営してきた実践知を言語化し、地域と全国に共有することで、新たな担い手を増やす。"));

/* --- 6. 事業の内容 --- */
children.push(h1("6.", "事業の内容"));

children.push(h2("柱1　地域に開かれたオープン体験会・交流会（年5回）"));
children.push(li("対象：精神障がいのある方（未経験可）、ご家族、連携機関の利用者・職員"));
children.push(li("規模：1回あたり20～30名程度（年間 延べ100名以上、うち初参加50名以上を目標）"));
children.push(li("プログラム："));
children.push(li2("ウォーミングアップとやさしいボール遊び（体力・経験を問わず参加できる内容）"));
children.push(li2("ミニゲーム（人数・時間・ルールを可変にし、「見学だけ」「途中参加」「途中退出」を明示的に歓迎）"));
children.push(li2("交流タイム（当事者スタッフが自らの“フットサルとの出会い”を語り、参加者どうしで対話）"));
children.push(li2("希望者への継続相談（近隣チームや活動情報の紹介を含む）"));
children.push(li("合理的配慮・安全面：静養できるスペースの常設、休憩の定例化、体調確認シートの活用、精神保健福祉士等の専門職の同席、写真撮影の可否を個別に確認"));

children.push(h2("柱2　通年の練習環境の確保と競技力の向上"));
children.push(li("毎週土曜日の定例練習（年40回）の会場を安定的に確保する"));
children.push(li("月1回の強化練習会（年12回）に外部指導者を招き、傷害予防を含む安全なトレーニングと戦術面の向上を図る"));
children.push(li("東京都大会・関東大会・全国大会への出場"));
children.push(li("全メンバーの傷害保険加入、救急用品の整備、会場のAED所在確認を含む安全管理体制の徹底"));
children.push(note("※ 勝利を目指して本気で闘う経験そのものが、当事者にとって大きな自己効力感の回復につながっています。2026年の関東大会優勝は、その何よりの証拠です。"));

children.push(h2("柱3　普及・啓発と実践知の共有"));
children.push(li("「はじめてのソーシャルフットボール ハンドブック」（A5判・24頁・400部）を制作する"));
children.push(li2("当事者・家族向け：参加のしかた、当日の流れ、よくある不安と答え"));
children.push(li2("支援者向け：安全な場のつくり方、合理的配慮のポイント、事故・体調不良時の対応"));
children.push(li2("運営者向け：当事者主体でチームを13年続けてきた工夫（会場確保、役割分担、担い手の交代）"));
children.push(li("医療機関・福祉事業所・自治体窓口など50か所以上へ配布する"));
children.push(li("ウェブサイト・SNS・noteで活動を継続的に発信し、年度末にオープン形式の活動報告会を開催する"));

/* --- 7. 実施体制 --- */
children.push(h1("7.", "実施体制"));
children.push(gridTable(
  ["役　割", "担　当", "主な業務"],
  [
    ["事業責任者", "代表", "全体統括、関係機関との調整、会計管理"],
    ["運営担当", "マネージャー", "会場確保、参加申込の受付、記録・広報"],
    ["指導担当", "監督・外部コーチ", "プログラム設計、当日の指導、傷害予防"],
    ["ピアサポート", "当事者スタッフ（選手）", "体験会の進行補助、参加者への声かけ、体験談の共有"],
    ["専門職", "精神保健福祉士 等", "当日の相談対応、安全管理・合理的配慮に関する助言"],
  ],
  [1600, 2400, CW - 4000]
));
children.push(spacer(60));
children.push(body("連携予定先：精神科デイケア、就労移行支援事業所、地域活動支援センター、日本ソーシャルフットボール協会（JSFA）、地域のフットサル施設。"));
children.push(note("※ 連携先については、内諾済みの団体名と、これから依頼する団体名を分けて記入してください。実名が入ると企画の実現性が大きく高まります。"));

/* --- 8. 効果と評価 ---（ページの切れ目を整えるための改ページ。不要なら次の1行を削除） */
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1("8.", "期待される効果と評価方法"));
children.push(gridTable(
  ["区　分", "指　標", "目　標"],
  [
    [["アウトプット", "（実施量）"], "オープン体験会の開催回数／参加者数", "年5回／延べ100名以上（うち初参加50名以上）"],
    ["", "定例練習・強化練習会の実施回数", "定例練習 年40回／強化練習会 年12回"],
    ["", "ハンドブックの配布", "400部・50か所以上"],
    [["アウトカム", "（変化）"], ["参加者アンケートの肯定回答率", "（「また参加したい」「体を動かすのが楽しかった」）"], "80％以上"],
    ["", "体験会をきっかけに運動・チーム活動を継続した人数", "10名以上"],
    ["", "参加前後の気分・主観的健康感の変化", "改善傾向を確認"],
    ["", "連携機関の職員から見た利用者の変化（聞き取り）", "肯定的な変化の報告を得る"],
  ],
  [1600, 4000, CW - 5600]
));
children.push(spacer(60));
children.push(body("評価方法：毎回の参加記録、参加前後の簡易アンケート、連携機関へのヒアリング、年度末の活動報告会での振り返りを行います。得られた結果はハンドブックの改訂と活動報告書に反映し、財団へご報告いたします。"));

/* --- 9. 継続性 --- */
children.push(h1("9.", "助成終了後の継続性"));
children.push(body("本事業で整備するトレーニング用具、ハンドブック、そして連携機関とのネットワークは、助成期間の終了後も継続して活用できる「チームの資産」となります。オープン体験会は、無理のない範囲の参加費と連携機関との共催によって継続します。また、運営ノウハウを文書として残すことで、担い手が交代してもチームが続いていく体制を整えます。"));
children.push(body("私たちは、デイケアの卒業とともに失われかけた場を、自分たちの手で13年続けてきました。この事業は、その経験を次の誰かのために開く取り組みです。"));

/* --- 10. 別紙 --- */
children.push(h1("10.", "添付書類"));
children.push(li("別紙1　スケジュール（2027年4月～2028年2月）"));
children.push(li("別紙2　費用積算表（総事業費 1,360,000円／助成申請額 1,000,000円）"));
children.push(li("会則（規約）および直近の総会議事録　※任意団体のため添付"));

children.push(spacer(200));
children.push(new Paragraph({
  children: [t("以　上", { size: 20 })],
  alignment: AlignmentType.RIGHT,
}));

/* =====================================================================
   文書の組み立て
   ===================================================================== */
const doc = new Document({
  creator: "BOSCO NEXT",
  title: "2027年度 ヤマト福祉財団 障がい者福祉助成金 企画書",
  description: "ソーシャルフットボールによる精神障がい当事者の健康づくり・社会参加促進事業",
  styles: {
    default: {
      document: {
        run: { font: JP, size: 20, color: "222222" },
        paragraph: { spacing: { line: 268, lineRule: LineRuleType.AUTO } },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN, header: 680, footer: 680 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [t("BOSCO NEXT｜2027年度 ヤマト福祉財団 障がい者福祉助成金 企画書", { size: 16, color: GRAY })],
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9", space: 2 } },
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            t("－ ", { size: 17, color: GRAY }),
            new TextRun({ children: [PageNumber.CURRENT], font: JP, size: 17, color: GRAY }),
            t(" －", { size: 17, color: GRAY }),
          ],
        })],
      }),
    },
    children,
  }],
});

const out = path.join(__dirname, "..", "01_企画書.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote " + out); });
