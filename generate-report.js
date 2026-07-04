#!/usr/bin/env node
/**
 * 定期巡回報告書 docx生成（雛形完全再現版）
 * 使い方: cat mail.txt | node generate-report.js
 *         node generate-report.js data.json
 */
const fs   = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
  PageNumber, TabStopType, TabStopPosition,
  Header, Footer,
} = require('docx');

// ─── データ読み込み ──────────────────────────────────────────
async function loadData() {
  let raw = '';
  if (process.argv[2] && fs.existsSync(process.argv[2])) {
    raw = fs.readFileSync(process.argv[2], 'utf8');
  } else if (!process.stdin.isTTY) {
    raw = await new Promise(res => {
      const c = []; process.stdin.on('data', d => c.push(d)); process.stdin.on('end', () => res(Buffer.concat(c).toString('utf8')));
    });
  } else { console.error('使い方: cat mail.txt | node generate-report.js'); process.exit(1); }

  // BEGIN/END抽出
  const bm = raw.match(/===BEGIN CSV===([\s\S]*?)===END CSV===/);
  const csv = bm ? bm[1].trim() : raw.trim();

  // #NEXTで複数物件に分割
  const blocks = csv.split(/^#NEXT$/m).map(b => b.trim()).filter(b => b);
  return blocks.map(parseBlock);
}

function parseLine(line) {
  const r = []; let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { if (inQ && line[i+1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
    else if (c === ',' && !inQ) { r.push(cur); cur = ''; }
    else cur += c;
  }
  r.push(cur); return r;
}

function parseBlock(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l && !l.startsWith('#PROPERTY,'));
  let header = {}, notes = '', items = [];
  lines.forEach(line => {
    if (line.startsWith('#JUNKAI,')) {
      const c = parseLine(line.slice(8));
      header = { propertyNo: c[0]||'', propertyName: c[1]||'', mgmtCompany: c[2]||'', managerName: c[3]||'', patrolDate: c[4]||'' };
    } else if (line.startsWith('#NOTES,')) {
      notes = parseLine(line.slice(7))[0] || '';
    } else if (!line.startsWith('大分類,') && !line.startsWith('#')) {
      const c = parseLine(line);
      if (c.length >= 3) items.push({ category: c[0], label: c[1], status: c[2], report: c[3]||'' });
    }
  });
  return { header, notes, items };
}

// ─── 色定数（雛形完全準拠） ─────────────────────────────────
const C = {
  titleBg:   'FFFFFF', // タイトルは白背景
  headerBg:  '1F4E78', // 表ヘッダー濃紺
  catBg:     'F2F2F2', // 大分類 薄灰
  labelBg:   'D9EAF7', // 物件情報ラベル 水色
  white:     'FFFFFF',
  black:     '000000',
  gray:      '555555',
  border:    '9E9E9E',
  ybg:       'FFF3A0', // 要指導
  rbg:       'FFE0E0', // 要改善
};

// ─── 罫線ヘルパー ────────────────────────────────────────────
const B  = (c, sz) => ({ style: BorderStyle.SINGLE, size: sz||4, color: c||C.border });
const BB = (c, sz) => { const b = B(c,sz); return { top:b, bottom:b, left:b, right:b }; };
const BN = { style: BorderStyle.NONE, size:0, color:'FFFFFF' };
const BNONE = { top:BN, bottom:BN, left:BN, right:BN };

// ─── セルヘルパー ────────────────────────────────────────────
function cell(runs, w, opts={}) {
  const { rs=1, cs=1, bg, vAlign=VerticalAlign.CENTER, align=AlignmentType.LEFT, borders } = opts;
  return new TableCell({
    rowSpan: rs, columnSpan: cs,
    verticalAlign: vAlign,
    borders: borders || BB(),
    shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    width: { size: w, type: WidthType.DXA },
    margins: { top: 40, bottom: 40, left: 80, right: 60 },
    children: [new Paragraph({
      alignment: align,
      children: Array.isArray(runs) ? runs : [runs],
    })],
  });
}

function txt(text, opts={}) {
  const { bold=false, color=C.black, size=18, font='Meiryo' } = opts;
  return new TextRun({ text: String(text||''), font, size, bold, color });
}

// ─── チェック文言マスタ（雛形準拠） ──────────────────────
const CHECK_MASTER = {
  'アプローチ':    '館銘板・排水溝・床等',
  'エントランス':  '天井・床・マット・レール溝等',
  'メールコーナー':'チラシ等が散乱または溢れていないか',
  '廊下、階段':   '床・手摺・玄関周り・排水溝・腰壁等',
  '各戸玄関周り': '玄関扉・インターフォン・PS扉等',
  '建物外周':     '植え込み周り（植栽・雑草）水周り等',
  'ゴミ置き場':   '整理整頓・清掃・臭気・害虫発生状況等',
  '給水設備':     '目視点検による不具合箇所の確認。水漏れ、異音等',
  'エレベーター':  '異常音・振動・点検中に清掃状況（マット・溝・他）',
  '自動火災報知機':'火災報知器の目視による不具合箇所の点検',
  '駐車場・駐輪場':'清掃状況や無断駐車の有無など（不具合・破損・故障がないか）',
  '宅配ボックス':  '清掃状況並びに不具合・破損・故障等',
  'タイマー':     '共用灯タイマーの状況（不具合・破損・故障等）',
  'エントランス照明':'点灯確認・故障の有無や清掃状況確認',
  '廊下照明':     '点灯確認・故障の有無や清掃状況確認',
  '階段照明':     '点灯確認・故障の有無や清掃状況確認',
  '外灯':         '点灯確認・故障の有無や清掃状況確認',
  'その他照明':    '点灯確認・故障の有無や清掃状況確認',
  'セキュリティ':  '監視カメラ具合確認・門扉の施錠状況',
  '粗大ごみ':     'シールが貼られていない物がないか',
  '掲示板':       '過去の掲示物が無い・掲示物の整理',
  '連絡ノートの記入状況':'連絡ノートの記入状況',
  '管理員室':     '', // 複数行あるため個別判定は不要
};

// ─── 1物件分のdocxセクションを生成 ─────────────────────────
function buildSection(data, isLast) {
  const { header: h, notes, items } = data;

  // A4, 余白 上下15mm / 左右13mm
  const PAGE_W  = 11906;
  const M_H     = 737;   // 13mm
  const M_V     = 851;   // 15mm
  const CONTENT = PAGE_W - M_H * 2; // 10432 DXA

  // 列幅（雛形の比率: A12.6/B13.6/C25.6/D10.6/E27.6 → 合計90.0）
  // DXA換算
  const TOTAL_CH = 12.6 + 13.6 + 25.6 + 10.6 + 27.6; // 90.0
  const W = {
    A: Math.round(CONTENT * 12.6 / TOTAL_CH),  // 大分類
    B: Math.round(CONTENT * 13.6 / TOTAL_CH),  // 項目
    C: Math.round(CONTENT * 25.6 / TOTAL_CH),  // 状況チェック
    D: Math.round(CONTENT * 10.6 / TOTAL_CH),  // 結果
    E: CONTENT - Math.round(CONTENT * 12.6 / TOTAL_CH)
               - Math.round(CONTENT * 13.6 / TOTAL_CH)
               - Math.round(CONTENT * 25.6 / TOTAL_CH)
               - Math.round(CONTENT * 10.6 / TOTAL_CH), // 報告
  };

  // ── 1. タイトル行（A1:E1 結合） ──
  const titleTable = new Table({
    width: { size: CONTENT, type: WidthType.DXA },
    columnWidths: [CONTENT],
    rows: [new TableRow({ height: { value: 480, rule: 'exact' }, children: [
      cell([txt('定期巡回報告書', { bold:true, size:26 })],
        CONTENT, { cs:1, align: AlignmentType.CENTER, borders: BB(C.border,4) })
    ]})]
  });

  // ── 2. 管理会社行（E2右寄せ、A-D空白） ──
  const mgmtTable = new Table({
    width: { size: CONTENT, type: WidthType.DXA },
    columnWidths: [W.A+W.B+W.C+W.D, W.E],
    rows: [new TableRow({ height: { value: 360, rule: 'exact' }, children: [
      cell([], W.A+W.B+W.C+W.D, { borders: BNONE }),
      cell([txt(h.mgmtCompany||'', { size:18 })],
        W.E, { align: AlignmentType.RIGHT, borders: BNONE }),
    ]})]
  });

  // ── 3. 物件情報テーブル（行3-4） ──
  const infoTable = new Table({
    width: { size: CONTENT, type: WidthType.DXA },
    columnWidths: [W.A, W.B+W.C, W.D, W.E],
    rows: [
      // 行3: 物件番号 | ●● | 巡回日 | 日付
      new TableRow({ height: { value: 380, rule: 'exact' }, children: [
        cell([txt('物件番号', { bold:true, size:17 })],
          W.A, { bg: C.labelBg, align: AlignmentType.CENTER }),
        cell([txt(h.propertyNo||'', { size:17 })], W.B+W.C, {}),
        cell([txt('巡回日', { bold:true, size:17 })],
          W.D, { bg: C.labelBg, align: AlignmentType.CENTER }),
        cell([txt(h.patrolDate||'', { size:17 })], W.E, {}),
      ]}),
      // 行4: 建物名 | ●● | 報告者氏名 | 氏名
      new TableRow({ height: { value: 380, rule: 'exact' }, children: [
        cell([txt('建物名', { bold:true, size:17 })],
          W.A, { bg: C.labelBg, align: AlignmentType.CENTER }),
        cell([txt(h.propertyName||'', { size:17 })], W.B+W.C, {}),
        cell([txt('報告者氏名', { bold:true, size:17 })],
          W.D, { bg: C.labelBg, align: AlignmentType.CENTER }),
        cell([txt(h.managerName||'', { size:17 })], W.E, {}),
      ]}),
    ]
  });

  // ── 4. 「下記の通り…」文 ──
  const introPara = new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [txt('下記の通りに定期巡回を行いましたので、御報告申し上げます。', { bold:true, size:19 })],
  });

  // ── 5. チェック表 ──
  // ヘッダー行
  const tHdr = new TableRow({
    tableHeader: true,
    height: { value: 440, rule: 'exact' },
    children: [
      cell([txt('大分類',    { bold:true, color:C.white, size:18 })], W.A, { bg: C.headerBg, align: AlignmentType.CENTER, borders: BB(C.headerBg) }),
      cell([txt('項目',      { bold:true, color:C.white, size:18 })], W.B, { bg: C.headerBg, align: AlignmentType.CENTER, borders: BB(C.headerBg) }),
      cell([txt('状況チェック',{ bold:true, color:C.white, size:18 })], W.C, { bg: C.headerBg, align: AlignmentType.CENTER, borders: BB(C.headerBg) }),
      cell([txt('結果',      { bold:true, color:C.white, size:18 })], W.D, { bg: C.headerBg, align: AlignmentType.CENTER, borders: BB(C.headerBg) }),
      cell([txt('報告',      { bold:true, color:C.white, size:18 })], W.E, { bg: C.headerBg, align: AlignmentType.CENTER, borders: BB(C.headerBg) }),
    ]
  });

  // カテゴリ別グループ化
  const catMap = {}, catOrder = [];
  items.forEach(it => {
    if (!catMap[it.category]) { catMap[it.category] = []; catOrder.push(it.category); }
    catMap[it.category].push(it);
  });

  const dataRows = [];
  catOrder.forEach(catName => {
    const catItems = catMap[catName];
    catItems.forEach((item, idx) => {
      const isWarn = item.status === '要指導';
      const isFix  = item.status === '要改善';
      const rText  = isWarn ? '要指導' : isFix ? '要改善' : '○ 良好';
      const rColor = isWarn ? '8B6914' : isFix ? 'CC0000' : C.black;
      const rBg    = isWarn ? C.ybg    : isFix ? C.rbg    : undefined;
      const rBold  = isWarn || isFix;

      // 大分類セル（最初の行のみ rowspan・改行対応）
      const catCell = idx === 0 ? new TableCell({
        rowSpan: catItems.length,
        verticalAlign: VerticalAlign.CENTER,
        borders: BB(),
        shading: { fill: C.catBg, type: ShadingType.CLEAR },
        width: { size: W.A, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 60, right: 40 },
        children: catName.split('\n').map(line => new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [txt(line, { bold:true, size:17 })],
        })),
      }) : null;

      // 結果セル
      const resCell = new TableCell({
        verticalAlign: VerticalAlign.TOP,
        borders: BB(),
        shading: rBg ? { fill: rBg, type: ShadingType.CLEAR } : undefined,
        width: { size: W.D, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 60, right: 40 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [txt(rText, { bold:rBold, color:rColor, size:17 })],
        })],
      });

      const cells = [];
      if (catCell) cells.push(catCell);
      cells.push(cell([txt(item.label,  { size:17 })], W.B, { vAlign: VerticalAlign.TOP }));
      cells.push(cell([txt(item.check || CHECK_MASTER[item.label] || '', { size:16, color:C.gray })], W.C, { vAlign: VerticalAlign.TOP }));
      cells.push(resCell);
      cells.push(cell([txt(item.report||'', { size:16, color:C.gray })], W.E, { vAlign: VerticalAlign.TOP }));

      dataRows.push(new TableRow({
        height: { value: 380, rule: 'atLeast' },
        children: cells,
      }));
    });
  });

  const checkTable = new Table({
    width: { size: CONTENT, type: WidthType.DXA },
    columnWidths: [W.A, W.B, W.C, W.D, W.E],
    rows: [tHdr, ...dataRows],
  });

  // ── 6. 処理実施内容 ──
  // A33:E33 結合（太字ラベル）
  const notesLabelTable = new Table({
    width: { size: CONTENT, type: WidthType.DXA },
    columnWidths: [CONTENT],
    rows: [new TableRow({ height: { value: 280, rule: 'exact' }, children: [
      cell([txt('処理実施内容及びその他報告内容', { bold:true, size:18 })],
        CONTENT, { borders: BB(), vAlign: VerticalAlign.CENTER }),
    ]})]
  });

  // A34:E34 結合（内容・高さ大）
  const notesContentTable = new Table({
    width: { size: CONTENT, type: WidthType.DXA },
    columnWidths: [CONTENT],
    rows: [new TableRow({ height: { value: 1600, rule: 'atLeast' }, children: [
      cell([txt(notes||'', { size:18 })],
        CONTENT, { borders: BB(), vAlign: VerticalAlign.TOP }),
    ]})]
  });

  const sp = (n) => new Paragraph({ children: [], spacing: { after: n||60 } });

  return [
    // タイトル
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [txt('定期巡回報告書', { bold:true, size:28 })],
    }),
    // 管理会社（右寄せ）
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { after: 40 },
      children: [txt(h.mgmtCompany||'', { size:17 })],
    }),
    // 物件情報テーブル
    infoTable,
    // 「下記の通り…」
    new Paragraph({
      spacing: { before: 80, after: 80 },
      children: [txt('下記の通りに定期巡回を行いましたので、御報告申し上げます。', { bold:true, size:19 })],
    }),
    // チェック表
    checkTable,
    // 処理実施内容ラベル
    notesLabelTable,
    // 処理実施内容本文
    notesContentTable,
  ];
}

// ─── メイン ─────────────────────────────────────────────────
(async () => {
  const dataList = await loadData();
  console.log(`📋 ${dataList.length}件の物件データを読み込みました`);

  // 複数物件をページブレークで連結
  const { PageBreak } = require('docx');
  const allChildren = [];
  dataList.forEach((data, idx) => {
    const section = buildSection(data, idx === dataList.length - 1);
    if (idx > 0) {
      // ページ区切り
      allChildren.push(new Paragraph({
        children: [new PageBreak()],
        spacing: { before: 0, after: 0 },
      }));
    }
    allChildren.push(...section);
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size:   { width: 11906, height: 16838 },
          margin: { top: 851, right: 737, bottom: 851, left: 737 },
        },
      },
      children: allChildren,
    }],
  });

  // 出力ファイル名
  const h0   = dataList[0].header;
  // 日付変換: 「2026年6月29日」→「20260629」
  const fmtDate = s => {
    const m = String(s||'').match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
    if(!m) return String(s||'').replace(/[\\/:*?"<>| ]/g,'');
    return m[1]+('0'+m[2]).slice(-2)+('0'+m[3]).slice(-2);
  };
  const safe = s => String(s||'').replace(/[\\/:*?"<>| ]/g,'');
  // ファイル名: {物件番号}{物件名}_{日付}_{担当者名}.docx
  const buildFname = h =>
    safe(h.propertyNo)+safe(h.propertyName)+'_'+fmtDate(h.patrolDate)+'_'+safe(h.managerName)+'.docx';

  const fname = dataList.length === 1
    ? buildFname(h0)
    : safe(h0.propertyNo)+safe(h0.propertyName)+'_ほか'+dataList.length+'件_'+fmtDate(h0.patrolDate)+'_'+safe(h0.managerName)+'.docx';

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(fname, buf);
  console.log(`✅ 生成完了: ${fname}`);
})().catch(e => { console.error('❌ エラー:', e.message); process.exit(1); });
