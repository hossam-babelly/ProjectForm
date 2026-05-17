/**
 * ─────────────────────────────────────────────────────────────
 *  نموذج طرح دراسة مشروع — Backend Server v2
 *  Express · docx · ExcelJS · Nodemailer
 * ─────────────────────────────────────────────────────────────
 */

const express    = require('express');
const cors       = require('cors');
const fs         = require('fs');
const path       = require('path');
const crypto     = require('crypto');
const nodemailer = require('nodemailer');
const ExcelJS    = require('exceljs');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell,
  AlignmentType, WidthType, ShadingType,
  BorderStyle, VerticalAlign, PageOrientation
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Directories ──────────────────────────────────────────────
const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

app.use(cors());
app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Helpers ──────────────────────────────────────────────────
function genId() {
  return Date.now().toString(36).toUpperCase() + '-' +
         crypto.randomBytes(3).toString('hex').toUpperCase();
}
function loadIndex() {
  const p = path.join(DATA_DIR, 'index.json');
  if (!fs.existsSync(p)) return [];
  try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch { return []; }
}
function saveIndex(index) {
  fs.writeFileSync(path.join(DATA_DIR, 'index.json'), JSON.stringify(index, null, 2), 'utf8');
}
const MONTHS = ['شهر 1','شهر 2','شهر 3','شهر 4','شهر 5','شهر 6',
                'شهر 7','شهر 8','شهر 9','شهر 10','شهر 11','شهر 12'];

// ════════════════════════════════════════════════════════════
//  EXCEL GENERATOR
// ════════════════════════════════════════════════════════════
async function generateExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'نموذج دراسة مشروع';
  wb.created = new Date();

  const C_HEADER = '1A3A5C';
  const C_ACCENT = 'C8A84B';
  const C_SUBTOT = 'EAF0F8';
  const C_TOTAL  = 'D1DBE8';

  function hStyle(bg) {
    return {
      font:      { bold:true, color:{argb:'FFFFFFFF'}, name:'Arial', size:11 },
      fill:      { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+bg} },
      alignment: { horizontal:'center', vertical:'middle', wrapText:true, readingOrder:2 },
      border:    { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} }
    };
  }
  function tStyle(bg) {
    return {
      font:      { bold:true, name:'Arial', size:11 },
      fill:      { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+bg} },
      alignment: { horizontal:'right', vertical:'middle', readingOrder:2 },
      border:    { top:{style:'medium'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} }
    };
  }
  function dStyle() {
    return {
      font:      { name:'Arial', size:10 },
      alignment: { horizontal:'center', vertical:'middle', readingOrder:2 },
      border:    { top:{style:'hair'}, bottom:{style:'hair'}, left:{style:'hair'}, right:{style:'hair'} }
    };
  }

  // ── Sheet 1: الملخص ─────────────────────────────────────
  const s0 = wb.addWorksheet('الملخص', { rightToLeft:true });
  s0.columns = [{ width:38 }, { width:28 }];
  s0.mergeCells('A1:B1');
  Object.assign(s0.getCell('A1'), {
    value: 'ملخص معلومات المشروع',
    font:  { bold:true, size:16, color:{argb:'FFFFFFFF'}, name:'Arial' },
    fill:  { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_HEADER} },
    alignment: { horizontal:'center', vertical:'middle', readingOrder:2 }
  });
  s0.getRow(1).height = 30;

  s0.mergeCells('A2:B2');
  Object.assign(s0.getCell('A2'), {
    value: 'قسم المشاريع التنموية',
    font:  { bold:true, size:12, color:{argb:'FF'+C_HEADER}, name:'Arial' },
    fill:  { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_ACCENT} },
    alignment: { horizontal:'center', readingOrder:2 }
  });

  [
    ['فكرة المشروع', data.projectIdea||''],
    ['إجمالي التكاليف التأسيسية', data.summary?.foundingTotal||'$0'],
    ['إجمالي الإيرادات المتوقعة (سنوياً)', data.summary?.revenueAnnual||'$0'],
    ['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0'],
    ['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual||'$0'],
    ['الاهتلاك (سنوياً)', data.summary?.depreciation||'$0'],
    ['الربح الصافي (سنوياً)', data.summary?.netProfit||'$0'],
    ['عدد الموظفين', data.summary?.employees||'0'],
  ].forEach(([label, val], i) => {
    const r = s0.getRow(i+3);
    r.height = 22;
    r.getCell(1).value = label;
    r.getCell(1).font  = { bold:true, name:'Arial', size:10 };
    r.getCell(1).alignment = { readingOrder:2 };
    r.getCell(2).value = String(val);
    r.getCell(2).alignment = { horizontal:'center', readingOrder:2 };
    r.getCell(2).font  = { name:'Arial', size:10 };
    [1,2].forEach(c => {
      r.getCell(c).border = { top:{style:'hair'}, bottom:{style:'hair'}, left:{style:'hair'}, right:{style:'hair'} };
    });
  });

  // ── Sheet 2: التكاليف التأسيسية ────────────────────────
  if (data.foundingRows?.length) {
    const s1 = wb.addWorksheet('التكاليف التأسيسية', { rightToLeft:true });
    s1.columns = [
      {width:6},{width:16},{width:24},{width:10},{width:8},{width:20},{width:20},{width:20}
    ];
    const hr = s1.getRow(1);
    hr.height = 24;
    ['#','الصنف','البيان','اهتلاك','العدد','ملاحظات','التكلفة للواحدة ($)','التكلفة الإجمالية ($)']
      .forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle(C_HEADER)); });

    let tot = 0;
    data.foundingRows.forEach((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot += v;
      const row = s1.getRow(i+2);
      row.height = 20;
      [r ? i+1 : '', r.cat, r.bayan, r.dep?'✓':'', r.qty, r.notes, parseFloat(r.price)||0, v]
        .forEach((val,c) => {
          row.getCell(c+1).value = val;
          Object.assign(row.getCell(c+1), dStyle());
          if(c===6||c===7) row.getCell(c+1).numFmt='"$"#,##0.00';
        });
    });
    const n = data.foundingRows.length;
    s1.mergeCells(`A${n+2}:G${n+2}`);
    s1.getCell(`A${n+2}`).value = 'الإجمالي';
    Object.assign(s1.getCell(`A${n+2}`), tStyle(C_TOTAL));
    s1.getCell(`H${n+2}`).value = tot;
    s1.getCell(`H${n+2}`).numFmt = '"$"#,##0.00';
    Object.assign(s1.getCell(`H${n+2}`), tStyle(C_TOTAL));
  }

  // ── Sheet 3: الإيرادات ────────────────────────────────
  const pids = Object.keys(data.products||{}).filter(p=>data.products[p]?.name);
  if (pids.length) {
    const s2 = wb.addWorksheet('الإيرادات المتوقعة', { rightToLeft:true });
    s2.columns = [{width:28},{width:10},...MONTHS.map(()=>({width:12}))];
    const hr = s2.getRow(1); hr.height=24;
    ['البيان','الواحدة',...MONTHS].forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle(C_HEADER)); });

    let rowIdx = 2;
    pids.forEach(pid => {
      const p = data.products[pid];
      const rev = data.revenueData?.[pid]||[];
      const qr = s2.getRow(rowIdx++);
      qr.getCell(1).value = p.name+' — الكمية';
      qr.getCell(1).font = {bold:true,name:'Arial',size:10};
      qr.getCell(2).value = p.unit||'';
      rev.forEach((m,i) => { qr.getCell(i+3).value=parseFloat(m.qty)||0; Object.assign(qr.getCell(i+3),dStyle()); });

      const ur = s2.getRow(rowIdx++);
      ur.getCell(1).value = p.name+' — سعر الوحدة';
      ur.getCell(2).value = '$';
      rev.forEach((m,i) => { ur.getCell(i+3).value=parseFloat(m.unitPrice)||0; ur.getCell(i+3).numFmt='"$"#,##0.00'; Object.assign(ur.getCell(i+3),dStyle()); });

      const tr = s2.getRow(rowIdx++);
      tr.getCell(1).value = p.name+' — إجمالي المبيعات';
      tr.getCell(1).font = {bold:true,name:'Arial',size:10};
      rev.forEach((m,i) => {
        const v = parseFloat(String(m.total||'0').replace(/[^0-9.]/g,''))||0;
        tr.getCell(i+3).value=v; tr.getCell(i+3).numFmt='"$"#,##0.00';
        Object.assign(tr.getCell(i+3),dStyle());
        tr.getCell(i+3).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C_SUBTOT}};
      });
    });
  }

  // ── Sheet 4: التكاليف الثابتة ─────────────────────────
  if (data.fixedRows?.length) {
    const s4 = wb.addWorksheet('التكاليف الثابتة', { rightToLeft:true });
    s4.columns = [{width:6},{width:16},{width:24},{width:8},{width:20},{width:22},{width:22}];
    const hr = s4.getRow(1); hr.height=24;
    ['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة ($)','التكلفة الشهرية الإجمالية ($)']
      .forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle(C_HEADER)); });

    let tot=0;
    data.fixedRows.forEach((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot+=v;
      const row = s4.getRow(i+2); row.height=20;
      [i+1, r.cat, r.bayan, r.qty, r.notes, parseFloat(r.price)||0, v]
        .forEach((val,c) => {
          row.getCell(c+1).value=val; Object.assign(row.getCell(c+1),dStyle());
          if(c===5||c===6) row.getCell(c+1).numFmt='"$"#,##0.00';
        });
    });
    const n = data.fixedRows.length;
    s4.mergeCells(`A${n+2}:F${n+2}`);
    s4.getCell(`A${n+2}`).value='الإجمالي الشهري'; Object.assign(s4.getCell(`A${n+2}`),tStyle(C_TOTAL));
    s4.getCell(`G${n+2}`).value=tot; s4.getCell(`G${n+2}`).numFmt='"$"#,##0.00'; Object.assign(s4.getCell(`G${n+2}`),tStyle(C_TOTAL));
  }

  // ── Sheet 5: الموارد البشرية ──────────────────────────
  if (data.hrRows?.length) {
    const s5 = wb.addWorksheet('الموارد البشرية', { rightToLeft:true });
    s5.columns = [{width:6},{width:20},{width:16},{width:8},{width:20},{width:22},{width:22}];
    const hr = s5.getRow(1); hr.height=24;
    ['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي ($)','الراتب الشهري الإجمالي ($)']
      .forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle(C_HEADER)); });

    let tot=0;
    data.hrRows.forEach((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot+=v;
      const row = s5.getRow(i+2); row.height=20;
      [i+1, r.position, r.type, r.qty, r.reports, parseFloat(r.salary)||0, v]
        .forEach((val,c) => {
          row.getCell(c+1).value=val; Object.assign(row.getCell(c+1),dStyle());
          if(c===5||c===6) row.getCell(c+1).numFmt='"$"#,##0.00';
        });
    });
    const n = data.hrRows.length;
    s5.mergeCells(`A${n+2}:F${n+2}`);
    s5.getCell(`A${n+2}`).value='إجمالي الرواتب'; Object.assign(s5.getCell(`A${n+2}`),tStyle(C_TOTAL));
    s5.getCell(`G${n+2}`).value=tot; s5.getCell(`G${n+2}`).numFmt='"$"#,##0.00'; Object.assign(s5.getCell(`G${n+2}`),tStyle(C_TOTAL));
  }

  return wb.xlsx.writeBuffer();
}

// ════════════════════════════════════════════════════════════
//  WORD GENERATOR  (v4 — pixel-perfect match)
// ════════════════════════════════════════════════════════════

// ── Color palette (from XML analysis of الملف المطلوب) ────────
const C = {
  DARK_BLUE:   '1F3864',  // section headers, total rows (dark blue)
  GREY_HDR:    'AEAAAA',  // section headers for Revenue/Ops/Products
  COL_HDR:     '262626',  // column header rows (near-black)
  COL_HDR2:    'D9E2F3',  // column header rows (light blue) - founding/HR/fixed/dep
  PROD1_DARK:  '8EA9DB',  // product 1 (odd) - name/unit cells dark
  PROD1_LIGHT: '8EAADB',  // product 1 (odd) - unit cell slight variant  
  PROD2_DARK:  'F4B084',  // product 2 (even) - name/unit cells
  PROD2_LIGHT: 'F4B083',  // product 2 (even) - unit cell slight variant
  MON_ODD:     'D9E1F2',  // month odd columns (light blue)
  MON_EVEN:    'B4C6E7',  // month even columns (slightly darker)
  MON_ODD2:    'FCE4D6',  // month odd for product 2
  MON_EVEN2:   'F8CBAD',  // month even for product 2
  SUBTOT:      '808080',  // subtotal rows (grey)
  SUMMARY_LBL: 'DEEAF6',  // summary label cells (very light blue)
  SUMMARY_VAL: 'FBE4D5',  // summary value cells (very light orange)
  ROW_ODD:     'FBE4D5',  // data rows odd (founding/HR/fixed/dep)
  ROW_EVEN:    'F2F2F2',  // data rows even
  WHITE:        'FFFFFF',
};

// Page: landscape A4, margins 0.5"
const PAGE = {
  size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
  margin: { top: 720, right: 720, bottom: 720, left: 720 }
};
// Content width = 16838 - 720 - 720 = 15398 DXA (landscape uses long edge)
const TW = 15398;

// ── Border helper ──────────────────────────────────────────────
const THIN  = { style: BorderStyle.SINGLE, size: 1, color: '808080' };
const NONE  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const borders = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };

// ── Cell factory ──────────────────────────────────────────────
function cell(text, opts = {}) {
  const {
    fill, bold = false, size = 28, color, align = AlignmentType.CENTER,
    w, vMergeRestart, vMerge, rowSpan, colSpan, vertical = VerticalAlign.CENTER
  } = opts;

  const shading = fill ? { type: ShadingType.CLEAR, fill, color: fill } : undefined;

  const tcPr = {};
  if (w) tcPr.width = { size: w, type: WidthType.DXA };
  if (colSpan) tcPr.columnSpan = colSpan;
  if (vMergeRestart) tcPr.rowSpan = rowSpan;   // docx-js uses rowSpan
  if (vMerge) tcPr.rowSpan = 0; // continuation
  tcPr.verticalAlign = vertical;

  return new TableCell({
    ...tcPr,
    shading,
    borders,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      bidirectional: true,
      alignment: align,
      spacing: { after: 0, line: 240, lineRule: 'auto' },
      children: [new TextRun({
        text: String(text ?? ''),
        bold,
        size,
        font: 'Arial',
        color: color || (fill === C.DARK_BLUE || fill === C.COL_HDR || fill === C.SUBTOT || fill === C.GREY_HDR ? C.WHITE : undefined),
      })],
    })],
  });
}

// ── Table wrapper ──────────────────────────────────────────────
function tbl(rows, colWidths) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: colWidths,
    bidirectional: true,
    rows,
  });
}

// ── Spacing paragraph ──────────────────────────────────────────
function spacer(before = 200) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [] });
}

function fmtN(v) {
  const n = parseFloat(String(v || '0').replace(/[^0-9.-]/g, ''));
  if (isNaN(n) || n === 0) return '0';
  return n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}
function fmtM(v) {
  const n = parseFloat(String(v || '0').replace(/[^0-9.-]/g, ''));
  if (isNaN(n)) return '0 $';
  return n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 }) + ' $';
}

// Product color palette cycling
const PROD_COLORS = [
  { dark: C.PROD1_DARK, mon_odd: C.MON_ODD, mon_even: C.MON_EVEN },
  { dark: C.PROD2_DARK, mon_odd: C.MON_ODD2, mon_even: C.MON_EVEN2 },
];
function prodColor(idx) { return PROD_COLORS[idx % 2]; }


// ════════════════════════════════════════════════════════════════
async function generateWord(data) {
  const pids = Object.keys(data.products || {}).filter(p => data.products[p]?.name);
  const ch = [];

  // ── Title ──────────────────────────────────────────────────
  ch.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: 'قسم المشاريع التنموية', bold: true, size: 34, font: 'Arial' })],
  }));
  ch.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [new TextRun({ text: 'نموذج طرح دراسة مشروع', bold: true, size: 28, font: 'Arial' })],
  }));

  // ════ 1. SUMMARY TABLE ═══════════════════════════════════════
  // col widths: label 7740, value 7648 (from XML: tbl1 has 2 cols)
  const SW = [7699, 7699];
  ch.push(tbl([
    // Section header spanning 2 cols
    new TableRow({ children: [cell('ملخص معلومات المشروع', { fill: C.DARK_BLUE, bold: true, size: 28, colSpan: 2, w: TW, align: AlignmentType.CENTER })] }),
    // Data rows
    ...[
      ['فكرة المشروع',                     data.projectIdea || ''],
      ['اسم مقدم المشروع',                 data.applicantName || ''],
      ['رقم الهاتف',                        data.applicantPhone || ''],
      ['إجمالي التكاليف التأسيسية',         data.summary?.foundingTotal || '$0'],
      ['إجمالي الإيرادات المتوقعة (سنوياً)',data.summary?.revenueAnnual || '$0'],
      ['إجمالي التكاليف التشغيلية (سنوياً)',data.summary?.opsAnnual || '$0'],
      ['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual || '$0'],
      ['الاهتلاك (سنوياً)',                data.summary?.depreciation || '$0'],
      ['الربح الصافي (سنوياً)',             data.summary?.netProfit || '$0'],
      ['عدد الموظفين في المشروع',           String(data.summary?.employees || '0')],
    ].map(([lbl, val]) => new TableRow({ children: [
      cell(lbl, { fill: C.SUMMARY_LBL, bold: true, size: 28, w: SW[0], align: AlignmentType.RIGHT }),
      cell(val, { fill: C.SUMMARY_VAL, bold: true, size: 28, w: SW[1], align: AlignmentType.CENTER }),
    ]})),
  ], SW));

  // ════ 2. FOUNDING TABLE ══════════════════════════════════════
  if (data.foundingRows?.length) {
    ch.push(spacer());
    // cols: #149 صنف269 بيان1368 اهتلاك380 عدد233 ملاحظات1521 تكلفة498 إجمالي582
    const FW = [149, 269, 1368, 380, 233, 1521, 1572, 1906];
    // normalize to TW
    const fSum = FW.reduce((a,b)=>a+b,0);
    const FWN = FW.map(w => Math.round(w * TW / fSum));
    FWN[FWN.length-1] = TW - FWN.slice(0,-1).reduce((a,b)=>a+b,0);

    const fRows = [
      new TableRow({ children: [cell('التكاليف التأسيسية', { fill: C.DARK_BLUE, bold: true, size: 28, colSpan: 8, w: TW, align: AlignmentType.CENTER })] }),
      new TableRow({ children: ['#','الصنف','البيان','الاهتلاك','العدد','ملاحظات','التكلفة للواحدة','التكلفة الإجمالية']
        .map((h,i) => cell(h, { fill: C.COL_HDR2, bold: true, size: 28, w: FWN[i], color: '000000' })) }),
      ...data.foundingRows.map((r, i) => new TableRow({ children: [
        cell(i+1,         { fill: C.ROW_ODD, bold: false, size: 28, w: FWN[0] }),
        cell(r.cat||'',   { fill: i%2===0?'':C.ROW_EVEN, size: 28, w: FWN[1], color: '000000' }),
        cell(r.bayan||'', { fill: i%2===0?'':C.ROW_EVEN, size: 28, w: FWN[2], color: '000000' }),
        cell(r.dep ? '✓' : '', { fill: i%2===0?'':C.ROW_EVEN, size: 28, w: FWN[3], color: '000000' }),
        cell(r.qty||'',   { fill: i%2===0?'':C.ROW_EVEN, size: 28, w: FWN[4], color: '000000' }),
        cell(r.notes||'', { fill: i%2===0?'':C.ROW_EVEN, size: 28, w: FWN[5], color: '000000' }),
        cell(fmtM(r.price),{ fill: i%2===0?'':C.ROW_EVEN, size: 28, w: FWN[6], color: '000000' }),
        cell(fmtM(r.total),{ fill: i%2===0?'':C.ROW_EVEN, size: 28, w: FWN[7], color: '000000' }),
      ]})),
      // Total row
      new TableRow({ children: [
        cell('الإجمالي', { fill: C.DARK_BLUE, bold: true, size: 28, colSpan: 7, w: FWN.slice(0,7).reduce((a,b)=>a+b,0), align: AlignmentType.RIGHT }),
        cell(fmtM(data.summary?.foundingTotal), { fill: C.DARK_BLUE, bold: true, size: 28, w: FWN[7] }),
      ]}),
    ];
    ch.push(tbl(fRows, FWN));
  }

  // ════ 3. PRODUCTS TABLE ══════════════════════════════════════
  if (pids.length) {
    ch.push(spacer());
    // cols from XML: 4822, 4841, 5705 → normalize
    const PW_raw = [4822, 4841, 5705];
    const pSum = PW_raw.reduce((a,b)=>a+b,0);
    const PW = PW_raw.map(w => Math.round(w * TW / pSum));
    PW[2] = TW - PW[0] - PW[1];

    const prodRows = [
      new TableRow({ children: [cell('جدول المنتجات', { fill: C.GREY_HDR, bold: true, size: 28, colSpan: 3, w: TW, align: AlignmentType.CENTER })] }),
      new TableRow({ children: [
        cell('البيان',    { fill: C.COL_HDR, bold: true, size: 28, w: PW[0] }),
        cell('الواحدة',  { fill: C.COL_HDR, bold: true, size: 28, w: PW[1] }),
        cell('المكونات', { fill: C.COL_HDR, bold: true, size: 28, w: PW[2] }),
      ]}),
    ];

    pids.forEach((pid, pidIdx) => {
      const p = data.products[pid];
      const comps = (p.components || []).filter(c => c);
      const pc = prodColor(pidIdx);
      comps.forEach((comp, ci) => {
        if (ci === 0) {
          prodRows.push(new TableRow({ children: [
            new TableCell({
              rowSpan: comps.length,
              shading: { type: ShadingType.CLEAR, fill: pc.dark, color: pc.dark },
              borders, margins: { top:60,bottom:60,left:80,right:80 },
              width: { size: PW[0], type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({ bidirectional:true, alignment:AlignmentType.CENTER, spacing:{after:0}, children:[new TextRun({text:p.name||'',bold:true,size:28,font:'Arial',color:C.WHITE})] })],
            }),
            new TableCell({
              rowSpan: comps.length,
              shading: { type: ShadingType.CLEAR, fill: pc.dark, color: pc.dark },
              borders, margins: { top:60,bottom:60,left:80,right:80 },
              width: { size: PW[1], type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({ bidirectional:true, alignment:AlignmentType.CENTER, spacing:{after:0}, children:[new TextRun({text:p.unit||'',bold:false,size:28,font:'Arial',color:C.WHITE})] })],
            }),
            cell(comp, { fill: pc.dark, size: 28, w: PW[2] }),
          ]}));
        } else {
          prodRows.push(new TableRow({ children: [
            cell(comp, { fill: pc.dark, size: 28, w: PW[2] }),
          ]}));
        }
      });
    });
    ch.push(tbl(prodRows, PW));
  }

  // ════ 4. REVENUE TABLE ═══════════════════════════════════════
  if (pids.length) {
    ch.push(spacer());
    // cols: بيان756 واحدة383 + 12 months ~309 each
    const RW_raw = [756, 383, 309,309,309,309,309,309,309,309,309,359,359,362];
    const rSum = RW_raw.reduce((a,b)=>a+b,0);
    const RW = RW_raw.map(w => Math.round(w * TW / rSum));
    RW[RW.length-1] = TW - RW.slice(0,-1).reduce((a,b)=>a+b,0);

    const revRows = [
      new TableRow({ children: [cell('الإيرادات المتوقعة', { fill: C.GREY_HDR, bold:true, size:28, colSpan: 14, w: TW, align: AlignmentType.CENTER })] }),
      new TableRow({ children: [
        cell('البيان',   { fill: C.COL_HDR, bold:true, size:28, w: RW[0] }),
        cell('الواحدة', { fill: C.COL_HDR, bold:true, size:28, w: RW[1] }),
        ...MONTHS.map((m,i) => cell(m, { fill: C.COL_HDR, bold:true, size:28, w: RW[i+2] })),
      ]}),
    ];

    pids.forEach((pid, pidIdx) => {
      const p = data.products[pid];
      const rev = data.revenueData?.[pid] || [];
      const pc = prodColor(pidIdx);
      // row 1: qty
      revRows.push(new TableRow({ children: [
        cell(p.name,  { fill: pc.dark,  bold:true, size:28, w: RW[0] }),
        cell(p.unit||'', { fill: pc.dark, size:28, w: RW[1] }),
        ...MONTHS.map((_,m) => cell(fmtN(rev[m]?.qty), { fill: m%2===0?pc.mon_odd:pc.mon_even, size:28, w:RW[m+2], color:'000000' })),
      ]}));
      // row 2: unit price
      revRows.push(new TableRow({ children: [
        cell('سعر مبيع الواحدة', { fill: pc.dark, size:28, w: RW[0] }),
        cell('$',               { fill: pc.dark, size:28, w: RW[1] }),
        ...MONTHS.map((_,m) => cell(fmtN(rev[m]?.unitPrice), { fill: m%2===0?pc.mon_odd:pc.mon_even, size:28, w:RW[m+2], color:'000000' })),
      ]}));
      // row 3: total
      revRows.push(new TableRow({ children: [
        cell('سعر المبيع الإجمالي', { fill: pc.dark, size:28, w: RW[0] }),
        cell('$',                   { fill: pc.dark, size:28, w: RW[1] }),
        ...MONTHS.map((_,m) => cell(fmtN(rev[m]?.total||0), { fill: m%2===0?pc.mon_odd:pc.mon_even, size:28, w:RW[m+2], color:'000000' })),
      ]}));
    });
    // grand total row
    const monthTotals = MONTHS.map((_,m) => {
      let t = 0;
      pids.forEach(pid => { t += parseFloat(String(data.revenueData?.[pid]?.[m]?.total||'0').replace(/[^0-9.-]/g,''))||0; });
      return fmtM(t);
    });
    revRows.push(new TableRow({ children: [
      cell('الإجمالي', { fill: C.COL_HDR, bold:true, size:28, w: RW[0] }),
      cell('',          { fill: C.COL_HDR, bold:true, size:28, w: RW[1] }),
      ...MONTHS.map((_,m) => cell(monthTotals[m], { fill: C.COL_HDR, bold:true, size:28, w:RW[m+2] })),
    ]}));
    ch.push(tbl(revRows, RW));
  }

  // ════ 5. OPS TABLE ═══════════════════════════════════════════
  if (pids.length && data.opsData) {
    ch.push(spacer());
    // cols: بيان433 واحدة285 مكوّن395 + 12 months
    const OW_raw = [433, 285, 395, 319,324,324,324,324,342,342,342,342,357,357,350];
    const oSum = OW_raw.reduce((a,b)=>a+b,0);
    const OW = OW_raw.map(w => Math.round(w * TW / oSum));
    OW[OW.length-1] = TW - OW.slice(0,-1).reduce((a,b)=>a+b,0);

    const opsRows = [
      new TableRow({ children: [cell('التكاليف التشغيلية', { fill: C.GREY_HDR, bold:true, size:28, colSpan:15, w:TW, align:AlignmentType.CENTER })] }),
      new TableRow({ children: [
        cell('البيان',    { fill: C.COL_HDR, bold:true, size:28, w: OW[0] }),
        cell('الواحدة',  { fill: C.COL_HDR, bold:true, size:28, w: OW[1] }),
        cell('التفاصيل', { fill: C.COL_HDR, bold:true, size:28, w: OW[2] }),
        ...MONTHS.map((m,i) => cell(m, { fill: C.COL_HDR, bold:true, size:28, w: OW[i+3] })),
      ]}),
    ];

    pids.forEach((pid, pidIdx) => {
      const p = data.products[pid];
      const comps = (p.components || []).filter(c => c);
      const opsD = data.opsData?.[pid] || {};
      const pc = prodColor(pidIdx);

      comps.forEach((comp, ci) => {
        const vals = MONTHS.map((_,m) => fmtN(opsD[`${ci}_${m}`]));
        if (ci === 0) {
          opsRows.push(new TableRow({ children: [
            new TableCell({
              rowSpan: comps.length,
              shading: { type: ShadingType.CLEAR, fill: pc.dark, color: pc.dark },
              borders, margins:{top:60,bottom:60,left:80,right:80},
              width:{size:OW[0],type:WidthType.DXA},
              verticalAlign: VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},children:[new TextRun({text:p.name||'',bold:true,size:28,font:'Arial',color:C.WHITE})]})],
            }),
            new TableCell({
              rowSpan: comps.length,
              shading: { type: ShadingType.CLEAR, fill: pc.dark, color: pc.dark },
              borders, margins:{top:60,bottom:60,left:80,right:80},
              width:{size:OW[1],type:WidthType.DXA},
              verticalAlign: VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},children:[new TextRun({text:p.unit||'',size:28,font:'Arial',color:C.WHITE})]})],
            }),
            cell(comp, { fill: pc.dark, size:28, w: OW[2] }),
            ...vals.map((v,m) => cell(v, { fill: m%2===0?pc.mon_odd:pc.mon_even, size:28, w:OW[m+3], color:'000000' })),
          ]}));
        } else {
          opsRows.push(new TableRow({ children: [
            cell(comp, { fill: pc.dark, size:28, w: OW[2] }),
            ...vals.map((v,m) => cell(v, { fill: m%2===0?pc.mon_odd:pc.mon_even, size:28, w:OW[m+3], color:'000000' })),
          ]}));
        }
      });
      // subtotal row
      const subVals = MONTHS.map((_,m) => fmtM(opsD[`sub_${m}`]||0));
      opsRows.push(new TableRow({ children: [
        cell(`إجمالي ${p.name}`, { fill: C.SUBTOT, bold:true, size:28, colSpan:3, w:OW.slice(0,3).reduce((a,b)=>a+b,0), align:AlignmentType.RIGHT }),
        ...subVals.map((v,m) => cell(v, { fill: C.SUBTOT, bold:true, size:28, w:OW[m+3] })),
      ]}));
    });
    // grand total
    const opsTotals = MONTHS.map((_,m) => {
      let t=0;
      pids.forEach(pid => { t += parseFloat(String(data.opsData?.[pid]?.[`sub_${m}`]||'0').replace(/[^0-9.-]/g,''))||0; });
      return fmtM(t);
    });
    opsRows.push(new TableRow({ children: [
      cell('الإجمالي', { fill: C.COL_HDR, bold:true, size:28, colSpan:3, w:OW.slice(0,3).reduce((a,b)=>a+b,0), align:AlignmentType.RIGHT }),
      ...opsTotals.map((v,m) => cell(v, { fill: C.COL_HDR, bold:true, size:28, w:OW[m+3] })),
    ]}));
    ch.push(tbl(opsRows, OW));
  }

  // ════ 6. HR TABLE ════════════════════════════════════════════
  if (data.hrRows?.length) {
    ch.push(spacer());
    const HW_raw = [149, 907, 876, 321, 1288, 2325, 2731];
    const hSum = HW_raw.reduce((a,b)=>a+b,0);
    const HW = HW_raw.map(w => Math.round(w * TW / hSum));
    HW[HW.length-1] = TW - HW.slice(0,-1).reduce((a,b)=>a+b,0);

    let hrTotal = 0;
    data.hrRows.forEach(r => { hrTotal += parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0; });

    const hrRows = [
      new TableRow({ children: [cell('الموارد البشرية', { fill: C.DARK_BLUE, bold:true, size:28, colSpan:7, w:TW, align:AlignmentType.CENTER })] }),
      new TableRow({ children: ['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي','الراتب الشهري الإجمالي']
        .map((h,i) => cell(h, { fill: C.COL_HDR2, bold:true, size:28, w:HW[i], color:'000000' })) }),
      ...data.hrRows.map((r,i) => new TableRow({ children: [
        cell(i+1,             { fill: C.ROW_ODD,              size:28, w:HW[0], color:'000000' }),
        cell(r.position||'',  { fill: i%2===0?'':C.ROW_EVEN,  size:28, w:HW[1], color:'000000' }),
        cell(r.type||'',      { fill: i%2===0?'':C.ROW_EVEN,  size:28, w:HW[2], color:'000000' }),
        cell(r.qty||'',       { fill: i%2===0?'':C.ROW_EVEN,  size:28, w:HW[3], color:'000000' }),
        cell(r.reports||'—',  { fill: i%2===0?'':C.ROW_EVEN,  size:28, w:HW[4], color:'000000' }),
        cell(fmtM(r.salary),  { fill: i%2===0?'':C.ROW_EVEN,  size:28, w:HW[5], color:'000000' }),
        cell(fmtM(r.total),   { fill: i%2===0?'':C.ROW_EVEN,  size:28, w:HW[6], color:'000000' }),
      ]})),
      new TableRow({ children: [
        cell('الإجمالي', { fill: C.DARK_BLUE, bold:true, size:28, colSpan:6, w:HW.slice(0,6).reduce((a,b)=>a+b,0), align:AlignmentType.RIGHT }),
        cell(fmtM(hrTotal), { fill: C.DARK_BLUE, bold:true, size:28, w:HW[6] }),
      ]}),
    ];
    ch.push(tbl(hrRows, HW));
  }

  // ════ 7. FIXED COSTS TABLE ═══════════════════════════════════
  if (data.fixedRows?.length) {
    ch.push(spacer());
    const FXW_raw = [149, 500, 1000, 300, 1400, 1400, 1400];
    const fxSum = FXW_raw.reduce((a,b)=>a+b,0);
    const FXW = FXW_raw.map(w => Math.round(w * TW / fxSum));
    FXW[FXW.length-1] = TW - FXW.slice(0,-1).reduce((a,b)=>a+b,0);

    let fxTotal = 0;
    data.fixedRows.forEach(r => { fxTotal += parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0; });

    const fxRows = [
      new TableRow({ children: [cell('التكاليف الثابتة', { fill: C.DARK_BLUE, bold:true, size:28, colSpan:7, w:TW, align:AlignmentType.CENTER })] }),
      new TableRow({ children: ['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة','التكلفة الشهرية الإجمالية']
        .map((h,i) => cell(h, { fill: C.COL_HDR2, bold:true, size:28, w:FXW[i], color:'000000' })) }),
      ...data.fixedRows.map((r,i) => new TableRow({ children: [
        cell(i+1,           { fill: C.ROW_ODD,             size:28, w:FXW[0], color:'000000' }),
        cell(r.cat||'',     { fill: i%2===0?'':C.ROW_EVEN, size:28, w:FXW[1], color:'000000' }),
        cell(r.bayan||'',   { fill: i%2===0?'':C.ROW_EVEN, size:28, w:FXW[2], color:'000000' }),
        cell(r.qty||'',     { fill: i%2===0?'':C.ROW_EVEN, size:28, w:FXW[3], color:'000000' }),
        cell(r.notes||'',   { fill: i%2===0?'':C.ROW_EVEN, size:28, w:FXW[4], color:'000000' }),
        cell(fmtM(r.price), { fill: i%2===0?'':C.ROW_EVEN, size:28, w:FXW[5], color:'000000' }),
        cell(fmtM(r.total), { fill: i%2===0?'':C.ROW_EVEN, size:28, w:FXW[6], color:'000000' }),
      ]})),
      new TableRow({ children: [
        cell('الإجمالي', { fill: C.DARK_BLUE, bold:true, size:28, colSpan:6, w:FXW.slice(0,6).reduce((a,b)=>a+b,0), align:AlignmentType.RIGHT }),
        cell(fmtM(fxTotal), { fill: C.DARK_BLUE, bold:true, size:28, w:FXW[6] }),
      ]}),
    ];
    ch.push(tbl(fxRows, FXW));
  }

  // ════ 8. DEPRECIATION TABLE ══════════════════════════════════
  if (data.depRows?.length) {
    ch.push(spacer());
    const DW_raw = [149, 280, 1350, 450, 225, 1500, 1500, 1500];
    const dSum = DW_raw.reduce((a,b)=>a+b,0);
    const DW = DW_raw.map(w => Math.round(w * TW / dSum));
    DW[DW.length-1] = TW - DW.slice(0,-1).reduce((a,b)=>a+b,0);

    let depTotal = 0;
    data.depRows.forEach(r => { depTotal += parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0; });

    const dRows = [
      new TableRow({ children: [cell('الاهتلاك', { fill: C.DARK_BLUE, bold:true, size:28, colSpan:8, w:TW, align:AlignmentType.CENTER })] }),
      new TableRow({ children: ['#','الصنف','البيان','نسبة الاهتلاك','العدد','ملاحظات','قيمة الاهتلاك للواحدة','قيمة الاهتلاك الإجمالية']
        .map((h,i) => cell(h, { fill: C.COL_HDR2, bold:true, size:28, w:DW[i], color:'000000' })) }),
      ...data.depRows.map((r,i) => new TableRow({ children: [
        cell(i+1,                    { fill: C.ROW_ODD,             size:28, w:DW[0], color:'000000' }),
        cell(r.cat||'',              { fill: i%2===0?'':C.ROW_EVEN, size:28, w:DW[1], color:'000000' }),
        cell(r.bayan||'',            { fill: i%2===0?'':C.ROW_EVEN, size:28, w:DW[2], color:'000000' }),
        cell((r.pct||'0') + ' %',   { fill: i%2===0?'':C.ROW_EVEN, size:28, w:DW[3], color:'000000' }),
        cell(r.qty||'',              { fill: i%2===0?'':C.ROW_EVEN, size:28, w:DW[4], color:'000000' }),
        cell(r.notes||'',            { fill: i%2===0?'':C.ROW_EVEN, size:28, w:DW[5], color:'000000' }),
        cell(fmtM(r.perUnit),        { fill: i%2===0?'':C.ROW_EVEN, size:28, w:DW[6], color:'000000' }),
        cell(fmtM(r.total),          { fill: i%2===0?'':C.ROW_EVEN, size:28, w:DW[7], color:'000000' }),
      ]})),
      new TableRow({ children: [
        cell('إجمالي قيمة الاهتلاك', { fill: C.DARK_BLUE, bold:true, size:28, colSpan:7, w:DW.slice(0,7).reduce((a,b)=>a+b,0), align:AlignmentType.RIGHT }),
        cell(fmtM(depTotal), { fill: C.DARK_BLUE, bold:true, size:28, w:DW[7] }),
      ]}),
    ];
    ch.push(tbl(dRows, DW));
  }

  // Footer
  ch.push(spacer(200));
  ch.push(new Paragraph({
    bidirectional: true,
    spacing: { after: 0 },
    children: [new TextRun({ text: `تاريخ الإرسال: ${new Date(data.submittedAt||Date.now()).toLocaleString('ar-SA')}`, size: 20, font: 'Arial', color: '888888' })],
  }));

  const doc = new Document({
    sections: [{ properties: { page: PAGE }, children: ch }],
  });
  return Packer.toBuffer(doc);
}

// ════════════════════════════════════════════════════════════
//  EMAIL SENDER
// ════════════════════════════════════════════════════════════
async function sendEmail(data, id, excelBuf, wordBuf) {
  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
    console.log('⚠️  متغيرات الإيميل غير مضبوطة — تخطي الإرسال');
    return;
  }
  const transporter = nodemailer.createTransport({
    service:'gmail',
    auth:{ user:process.env.EMAIL_USER, pass:process.env.EMAIL_PASS },
  });
  const to = process.env.RECEIVER_EMAIL || process.env.EMAIL_USER;
  await transporter.sendMail({
    from:    `"نظام دراسة المشاريع" <${process.env.EMAIL_USER}>`,
    to,
    subject: `📋 مشروع جديد — ${(data.projectIdea||'').substring(0,50)}`,
    html: `
<div dir="rtl" style="font-family:Arial;max-width:620px;margin:auto;border:1px solid #ddd;border-radius:12px;overflow:hidden">
  <div style="background:#1a3a5c;color:white;padding:20px 24px">
    <h2 style="margin:0;font-size:18px">📋 نموذج طرح دراسة مشروع جديد</h2>
    <p style="margin:6px 0 0;opacity:.75;font-size:13px">رقم الطلب: ${id}</p>
  </div>
  <div style="padding:20px 24px">
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#eaf0f8"><td style="padding:10px;font-weight:bold">فكرة المشروع</td><td style="padding:10px">${data.projectIdea||''}</td></tr>
      <tr><td style="padding:10px;font-weight:bold">اسم مقدم المشروع</td><td style="padding:10px">${data.applicantName||''}</td></tr>
      <tr style="background:#eaf0f8"><td style="padding:10px;font-weight:bold">رقم الهاتف</td><td style="padding:10px">${data.applicantPhone||''}</td></tr>
      <tr><td style="padding:10px;font-weight:bold">التكاليف التأسيسية</td><td style="padding:10px">${data.summary?.foundingTotal||'$0'}</td></tr>
      <tr style="background:#eaf0f8"><td style="padding:10px;font-weight:bold">الإيرادات السنوية</td><td style="padding:10px">${data.summary?.revenueAnnual||'$0'}</td></tr>
      <tr><td style="padding:10px;font-weight:bold">التكاليف التشغيلية السنوية</td><td style="padding:10px">${data.summary?.opsAnnual||'$0'}</td></tr>
      <tr style="background:#eaf0f8"><td style="padding:10px;font-weight:bold">التكاليف الثابتة السنوية</td><td style="padding:10px">${data.summary?.fixedAnnual||'$0'}</td></tr>
      <tr><td style="padding:10px;font-weight:bold">الاهتلاك السنوي</td><td style="padding:10px">${data.summary?.depreciation||'$0'}</td></tr>
      <tr style="background:#1a3a5c;color:white"><td style="padding:10px;font-weight:bold">الربح الصافي السنوي</td><td style="padding:10px;font-weight:bold">${data.summary?.netProfit||'$0'}</td></tr>
      <tr><td style="padding:10px;font-weight:bold">عدد الموظفين</td><td style="padding:10px">${data.summary?.employees||'0'}</td></tr>
    </table>
  </div>
  <div style="background:#f5f7fa;padding:12px 24px;font-size:11px;color:#888;border-top:1px solid #eee">
    تاريخ الإرسال: ${new Date().toLocaleString('ar-SA')} | مرفق: ملف Excel + ملف Word
  </div>
</div>`,
    attachments:[
      { filename:`مشروع_${id}.xlsx`, content:excelBuf, contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      { filename:`مشروع_${id}.docx`, content:wordBuf,  contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
    ],
  });
  console.log(`📧 إيميل أُرسل: ${id} → ${to}`);
}

// ════════════════════════════════════════════════════════════
//  ROUTES
// ════════════════════════════════════════════════════════════
app.post('/api/submit', async (req, res) => {
  try {
    const body = req.body;
    if (!body?.projectIdea) return res.status(400).json({ success:false, message:'فكرة المشروع مطلوبة' });

    const id = genId();
    const submission = { id, submittedAt: body.submittedAt||new Date().toISOString(), ...body };
    fs.writeFileSync(path.join(DATA_DIR,`${id}.json`), JSON.stringify(submission,null,2), 'utf8');
    const index = loadIndex();
    index.unshift({ id, projectIdea:body.projectIdea.substring(0,80), submittedAt:submission.submittedAt, summary:body.summary||{} });
    saveIndex(index);

    const [excelBuf, wordBuf] = await Promise.all([generateExcel(body), generateWord(body)]);
    fs.writeFileSync(path.join(DATA_DIR,`${id}.xlsx`), excelBuf);
    fs.writeFileSync(path.join(DATA_DIR,`${id}.docx`), wordBuf);

    await sendEmail(body, id, excelBuf, wordBuf).catch(e => console.error('Email error:', e.message));

    console.log(`✅ طلب جديد: ${id}`);
    return res.json({ success:true, id });
  } catch(err) {
    console.error('❌', err);
    return res.status(500).json({ success:false, message:err.message });
  }
});

app.get('/api/submissions', (req, res) => res.json({ success:true, submissions:loadIndex() }));

app.get('/api/submissions/:id', (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp = path.join(DATA_DIR,`${sid}.json`);
  if (!fs.existsSync(fp)) return res.status(404).json({ success:false });
  res.json({ success:true, submission:JSON.parse(fs.readFileSync(fp,'utf8')) });
});

app.get('/api/download/:id/excel', (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp = path.join(DATA_DIR,`${sid}.xlsx`);
  if (!fs.existsSync(fp)) return res.status(404).send('غير موجود');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition',`attachment; filename="project_${sid}.xlsx"`);
  res.send(fs.readFileSync(fp));
});

app.get('/api/download/:id/word', (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp = path.join(DATA_DIR,`${sid}.docx`);
  if (!fs.existsSync(fp)) return res.status(404).send('غير موجود');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition',`attachment; filename="project_${sid}.docx"`);
  res.send(fs.readFileSync(fp));
});

app.delete('/api/submissions/:id', (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  ['json','xlsx','docx'].forEach(ext => {
    const fp = path.join(DATA_DIR,`${sid}.${ext}`);
    if(fs.existsSync(fp)) fs.unlinkSync(fp);
  });
  saveIndex(loadIndex().filter(s=>s.id!==sid));
  res.json({ success:true });
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname,'public','index.html')));

app.listen(PORT, () => console.log(`🚀 يعمل على المنفذ ${PORT}`));
module.exports = app;
