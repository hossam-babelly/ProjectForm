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
//  WORD GENERATOR  (v5 — pixel-perfect + page breaks + org chart)
// ════════════════════════════════════════════════════════════

// ── Color palette (exact from XML of الملف المطلوب) ──────────
const C = {
  DARK_BLUE:   '1F3864',
  GREY_HDR:    'AEAAAA',
  COL_HDR_BLK: '262626',   // near-black col headers
  COL_HDR_BLU: 'D9E2F3',   // light-blue col headers
  PROD1_DARK:  '8EA9DB',
  PROD2_DARK:  'F4B084',
  PROD3_DARK:  '8EAADB',
  MON_ODD:     'D9E1F2',
  MON_EVEN:    'B4C6E7',
  MON_ODD2:    'FCE4D6',
  MON_EVEN2:   'F8CBAD',
  SUBTOT:      '808080',
  SUM_LBL:     'DEEAF6',
  SUM_VAL:     'FBE4D5',
  ROW_ODD:     'FBE4D5',
  ROW_EVEN:    'F2F2F2',
  WHITE:       'FFFFFF',
  ORG_ADM:     '4472C4',   // إداري — blue
  ORG_EXE:     '70AD47',   // تنفيذي — green
  ORG_SVC:     'ED7D31',   // مزود خدمة — orange
  ORG_OTH:     '5B9BD5',   // other
};

// Page: A4 landscape, 0.5" margins
const PAGE_LANDSCAPE = {
  size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
  margin: { top: 720, right: 720, bottom: 720, left: 720 }
};
const PAGE_PORTRAIT = {
  size: { width: 11906, height: 16838 },
  margin: { top: 720, right: 720, bottom: 720, left: 720 }
};

// Content width landscape (long edge - margins) = 16838 - 1440 = 15398
const TW_L = 15398;
// Content width portrait = 11906 - 1440 = 10466
const TW_P = 10466;

// Fonts
const FONT = 'Sakkal Majalla';

// Borders
const BDR = (color='auto', sz=12) => ({ style: BorderStyle.SINGLE, size: sz, color });
const BORDERS_AUTO = { top: BDR(), bottom: BDR(), left: BDR(), right: BDR() };
const BORDERS_BLK  = { top: BDR('000000',8), bottom: BDR('000000',8), left: BDR('000000',8), right: BDR('000000',8) };

// ── Cell factory ─────────────────────────────────────────────
function C_(text, opts={}) {
  const {
    fill, bold=false, sz=28, color,
    align=AlignmentType.CENTER,
    w, colSpan, rowSpan,
    vAlign=VerticalAlign.CENTER,
    borders=BORDERS_AUTO
  } = opts;

  const shading = fill ? { type: ShadingType.CLEAR, fill, color: fill } : undefined;
  const textColor = color || (
    fill === C.DARK_BLUE || fill === C.COL_HDR_BLK || fill === C.SUBTOT || fill === C.GREY_HDR
      ? C.WHITE : undefined
  );

  const cellProps = { shading, borders, margins:{top:60,bottom:60,left:80,right:80}, verticalAlign: vAlign };
  if (w) cellProps.width = { size: w, type: WidthType.DXA };
  if (colSpan) cellProps.columnSpan = colSpan;
  if (rowSpan > 1) cellProps.rowSpan = rowSpan;

  return new TableCell({
    ...cellProps,
    children: [new Paragraph({
      bidirectional: true,
      alignment: align,
      spacing: { after: 0, line: 240, lineRule: 'auto' },
      children: [new TextRun({ text: String(text??''), bold, size: sz, font: FONT, color: textColor })],
    })],
  });
}

// ── Table factory ─────────────────────────────────────────────
function T_(rows, colWidths, tw) {
  return new Table({
    width: { size: tw || TW_L, type: WidthType.DXA },
    columnWidths: colWidths,
    bidirectional: true,
    rows,
  });
}

// ── Section header row (dark blue, full width) ───────────────
function secHdr(title, cols, fill=C.DARK_BLUE) {
  return new TableRow({ tableHeader: true, children: [
    C_(title, { fill, bold:true, sz:32, colSpan:cols, w:TW_L, align:AlignmentType.CENTER,
      color: C.WHITE })
  ]});
}

// ── Column header row ─────────────────────────────────────────
function colHdr(labels, widths, fill=C.COL_HDR_BLK, textColor=C.WHITE) {
  return new TableRow({ tableHeader: true, children:
    labels.map((h,i) => C_(h, { fill, bold:true, sz:28, w:widths[i], color:textColor }))
  });
}

// ── Total row (colspan + value) ───────────────────────────────
function totRow(label, value, cols, lastW, fill=C.DARK_BLUE) {
  return new TableRow({ children: [
    C_(label, { fill, bold:true, sz:28, colSpan:cols-1, w:TW_L-lastW, align:AlignmentType.RIGHT }),
    C_(value, { fill, bold:true, sz:28, w:lastW }),
  ]});
}

// ── Spacer paragraph ──────────────────────────────────────────
function SP() { return new Paragraph({ spacing:{before:0,after:0}, children:[] }); }

// ── Page break paragraph ──────────────────────────────────────
function PB() {
  return new Paragraph({
    children: [new TextRun({ break: 1 })],
    spacing: { before:0, after:0 },
  });
}

// ── Number formatters ─────────────────────────────────────────
function fN(v) {
  const n = parseFloat(String(v||'0').replace(/[^0-9.-]/g,''));
  if (!n || isNaN(n)) return '0';
  return n.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:2});
}
function fM(v) {
  const n = parseFloat(String(v||'0').replace(/[^0-9.-]/g,''));
  if (isNaN(n)) return '0 $';
  return n.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:2})+' $';
}

// Product color palette (cycles for 3+ products)
const PROD_PAL = [
  { dark: C.PROD1_DARK, mo: C.MON_ODD,  me: C.MON_EVEN  },
  { dark: C.PROD2_DARK, mo: C.MON_ODD2, me: C.MON_EVEN2 },
  { dark: C.PROD3_DARK, mo: C.MON_ODD,  me: C.MON_EVEN  },
];
function PC(i) { return PROD_PAL[i % PROD_PAL.length]; }

// ── Org chart color by type ───────────────────────────────────
function orgColor(type) {
  const m = {'إداري':C.ORG_ADM,'تنفيذي':C.ORG_EXE,'مزود خدمة':C.ORG_SVC};
  return m[type] || C.ORG_OTH;
}

// ── Normalize column widths to sum exactly to TW ─────────────
function norm(arr, tw) {
  const sum = arr.reduce((a,b)=>a+b,0);
  const r = arr.map(w=>Math.round(w*tw/sum));
  r[r.length-1] = tw - r.slice(0,-1).reduce((a,b)=>a+b,0);
  return r;
}

// ════════════════════════════════════════════════════════════════
async function generateWord(data) {
  const pids = Object.keys(data.products||{}).filter(p=>data.products[p]?.name);

  // Each table gets its own section (= its own page)
  const sections = [];

  // ── Title section ──────────────────────────────────────────
  sections.push({
    properties: { page: PAGE_LANDSCAPE },
    children: [
      new Paragraph({ bidirectional:true, alignment:AlignmentType.CENTER, spacing:{before:3000,after:120},
        children:[new TextRun({text:'قسم المشاريع التنموية',bold:true,size:40,font:FONT,color:C.DARK_BLUE})] }),
      new Paragraph({ bidirectional:true, alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
        children:[new TextRun({text:'نموذج طرح دراسة مشروع',bold:true,size:32,font:FONT})] }),
    ],
  });

  // ── Helper: wrap table in a section ─────────────────────────
  function tableSection(tblObj, landscape=true) {
    return {
      properties: { page: landscape ? PAGE_LANDSCAPE : PAGE_PORTRAIT },
      children: [SP(), tblObj, SP()],
    };
  }

  // ════ 1. SUMMARY ══════════════════════════════════════════
  {
    const W = norm([7699,7699], TW_L);
    const rows = [
      new TableRow({ children: [C_('ملخص معلومات المشروع', { fill:C.DARK_BLUE, bold:true, sz:32, colSpan:2, w:TW_L, align:AlignmentType.CENTER, color:C.WHITE })] }),
      ...[
        ['فكرة المشروع',                      data.projectIdea||''],
        ['اسم مقدم المشروع',                  data.applicantName||''],
        ['رقم الهاتف',                         data.applicantPhone||''],
        ['إجمالي التكاليف التأسيسية',          data.summary?.foundingTotal||'$0'],
        ['إجمالي الإيرادات المتوقعة (سنوياً)', data.summary?.revenueAnnual||'$0'],
        ['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0'],
        ['إجمالي التكاليف الثابتة (سنوياً)',   data.summary?.fixedAnnual||'$0'],
        ['الاهتلاك (سنوياً)',                  data.summary?.depreciation||'$0'],
        ['الربح الصافي (سنوياً)',               data.summary?.netProfit||'$0'],
        ['عدد الموظفين في المشروع',             String(data.summary?.employees||'0')],
      ].map(([lbl,val]) => new TableRow({ children:[
        C_(lbl, { fill:C.SUM_LBL, bold:true, sz:28, w:W[0], align:AlignmentType.RIGHT }),
        C_(val, { fill:C.SUM_VAL, bold:true, sz:28, w:W[1] }),
      ]})),
    ];
    sections.push(tableSection(T_(rows, W)));
  }

  // ════ 2. FOUNDING ═════════════════════════════════════════
  if (data.foundingRows?.length) {
    const W = norm([149,269,1368,380,233,1521,1572,1906], TW_L);
    const hdrs = ['#','الصنف','البيان','الاهتلاك','العدد','ملاحظات','التكلفة للواحدة','التكلفة الإجمالية'];
    let tot = 0;
    data.foundingRows.forEach(r=>{ tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0; });
    const rows = [
      secHdr('التكاليف التأسيسية', 8),
      colHdr(hdrs, W, C.COL_HDR_BLU, '000000'),
      ...data.foundingRows.map((r,i)=>new TableRow({children:[
        C_(i+1,          {fill:C.ROW_ODD,                   sz:28,w:W[0]}),
        C_(r.cat||'',    {fill:i%2===0?C.WHITE:C.ROW_EVEN,  sz:28,w:W[1],color:'000000'}),
        C_(r.bayan||'',  {fill:i%2===0?C.WHITE:C.ROW_EVEN,  sz:28,w:W[2],color:'000000'}),
        C_(r.dep?'✓':'', {fill:i%2===0?C.WHITE:C.ROW_EVEN,  sz:28,w:W[3],color:r.dep?'70AD47':'000000'}),
        C_(r.qty||'',    {fill:i%2===0?C.WHITE:C.ROW_EVEN,  sz:28,w:W[4],color:'000000'}),
        C_(r.notes||'',  {fill:i%2===0?C.WHITE:C.ROW_EVEN,  sz:28,w:W[5],color:'000000'}),
        C_(fM(r.price),  {fill:i%2===0?C.WHITE:C.ROW_EVEN,  sz:28,w:W[6],color:'000000'}),
        C_(fM(r.total),  {fill:i%2===0?C.WHITE:C.ROW_EVEN,  sz:28,w:W[7],color:'000000'}),
      ]})),
      totRow('الإجمالي', fM(tot), 8, W[7]),
    ];
    sections.push(tableSection(T_(rows, W)));
  }

  // ════ 3. PRODUCTS ═════════════════════════════════════════
  if (pids.length) {
    const W = norm([4822,4841,5705], TW_L);
    const prodRows = [
      secHdr('جدول المنتجات', 3, C.GREY_HDR),
      colHdr(['البيان','الواحدة','المكونات'], W, C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi) => {
      const p = data.products[pid];
      const comps = (p.components||[]).filter(c=>c);
      const pc = PC(pi);
      comps.forEach((comp,ci) => {
        if (ci===0) {
          prodRows.push(new TableRow({ children:[
            new TableCell({ rowSpan:comps.length, shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS_AUTO, margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[0],type:WidthType.DXA}, verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.name||'',bold:true,size:28,font:FONT,color:C.WHITE})]})] }),
            new TableCell({ rowSpan:comps.length, shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS_AUTO, margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[1],type:WidthType.DXA}, verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.unit||'',size:28,font:FONT,color:C.WHITE})]})] }),
            C_(comp, {fill:pc.dark, sz:28, w:W[2]}),
          ]}));
        } else {
          prodRows.push(new TableRow({children:[C_(comp,{fill:pc.dark,sz:28,w:W[2]})]}));
        }
      });
    });
    sections.push(tableSection(T_(prodRows, W)));
  }

  // ════ 4. REVENUE ══════════════════════════════════════════
  if (pids.length) {
    const W = norm([756,383,309,309,309,309,309,309,309,309,309,359,359,362], TW_L);
    const revRows = [
      secHdr('الإيرادات المتوقعة', 14, C.GREY_HDR),
      colHdr(['البيان','الواحدة',...MONTHS], W, C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi) => {
      const p   = data.products[pid];
      const rev = data.revenueData?.[pid]||[];
      const pc  = PC(pi);
      revRows.push(new TableRow({children:[
        C_(p.name,   {fill:pc.dark,bold:true,sz:28,w:W[0]}),
        C_(p.unit||'',{fill:pc.dark,sz:28,w:W[1]}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.qty),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2],color:'000000'})),
      ]}));
      revRows.push(new TableRow({children:[
        C_('سعر مبيع الواحدة',{fill:pc.dark,sz:28,w:W[0]}),
        C_('$',{fill:pc.dark,sz:28,w:W[1]}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.unitPrice),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2],color:'000000'})),
      ]}));
      revRows.push(new TableRow({children:[
        C_('سعر المبيع الإجمالي',{fill:pc.dark,sz:28,w:W[0]}),
        C_('$',{fill:pc.dark,sz:28,w:W[1]}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.total||0),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2],color:'000000'})),
      ]}));
    });
    const mTots = MONTHS.map((_,m)=>{let t=0;pids.forEach(pid=>{t+=parseFloat(String(data.revenueData?.[pid]?.[m]?.total||'0').replace(/[^0-9.-]/g,''))||0;});return fM(t);});
    revRows.push(new TableRow({children:[
      C_('الإجمالي',{fill:C.COL_HDR_BLK,bold:true,sz:28,w:W[0]}),
      C_('',{fill:C.COL_HDR_BLK,sz:28,w:W[1]}),
      ...mTots.map((v,m)=>C_(v,{fill:C.COL_HDR_BLK,bold:true,sz:28,w:W[m+2]})),
    ]}));
    sections.push(tableSection(T_(revRows, W)));
  }

  // ════ 5. OPS ══════════════════════════════════════════════
  if (pids.length && data.opsData) {
    const W = norm([433,285,395,319,324,324,324,324,342,342,342,342,357,357,350], TW_L);
    const opsRows = [
      secHdr('التكاليف التشغيلية', 15, C.GREY_HDR),
      colHdr(['البيان','الواحدة','التفاصيل',...MONTHS], W, C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi) => {
      const p = data.products[pid];
      const comps = (p.components||[]).filter(c=>c);
      const opsD = data.opsData?.[pid]||{};
      const pc = PC(pi);
      comps.forEach((comp,ci) => {
        const vals = MONTHS.map((_,m)=>fN(opsD[`${ci}_${m}`]));
        if (ci===0) {
          opsRows.push(new TableRow({children:[
            new TableCell({ rowSpan:comps.length, shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS_AUTO, margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[0],type:WidthType.DXA}, verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.name||'',bold:true,size:28,font:FONT,color:C.WHITE})]})] }),
            new TableCell({ rowSpan:comps.length, shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS_AUTO, margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[1],type:WidthType.DXA}, verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.unit||'',size:28,font:FONT,color:C.WHITE})]})] }),
            C_(comp,{fill:pc.dark,sz:28,w:W[2]}),
            ...vals.map((v,m)=>C_(v,{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+3],color:'000000'})),
          ]}));
        } else {
          opsRows.push(new TableRow({children:[
            C_(comp,{fill:pc.dark,sz:28,w:W[2]}),
            ...vals.map((v,m)=>C_(v,{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+3],color:'000000'})),
          ]}));
        }
      });
      const subVals = MONTHS.map((_,m)=>fM(opsD[`sub_${m}`]||0));
      opsRows.push(new TableRow({children:[
        C_(`إجمالي ${p.name}`,{fill:C.SUBTOT,bold:true,sz:28,colSpan:3,w:W.slice(0,3).reduce((a,b)=>a+b,0),align:AlignmentType.RIGHT}),
        ...subVals.map((v,m)=>C_(v,{fill:C.SUBTOT,bold:true,sz:28,w:W[m+3]})),
      ]}));
    });
    const opsTots = MONTHS.map((_,m)=>{let t=0;pids.forEach(pid=>{t+=parseFloat(String(data.opsData?.[pid]?.[`sub_${m}`]||'0').replace(/[^0-9.-]/g,''))||0;});return fM(t);});
    opsRows.push(new TableRow({children:[
      C_('الإجمالي',{fill:C.COL_HDR_BLK,bold:true,sz:28,colSpan:3,w:W.slice(0,3).reduce((a,b)=>a+b,0),align:AlignmentType.RIGHT}),
      ...opsTots.map((v,m)=>C_(v,{fill:C.COL_HDR_BLK,bold:true,sz:28,w:W[m+3]})),
    ]}));
    sections.push(tableSection(T_(opsRows, W)));
  }

  // ════ 6. HR ═══════════════════════════════════════════════
  if (data.hrRows?.length) {
    const W = norm([149,907,876,321,1288,2325,2731], TW_L);
    const hdrs = ['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي','الراتب الشهري الإجمالي'];
    let tot=0; data.hrRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    const rows = [
      secHdr('الموارد البشرية', 7),
      colHdr(hdrs, W, C.COL_HDR_BLU, '000000'),
      ...data.hrRows.map((r,i)=>new TableRow({children:[
        C_(i+1,             {fill:C.ROW_ODD,                  sz:28,w:W[0]}),
        C_(r.position||'',  {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[1],color:'000000'}),
        C_(r.type||'',      {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[2],color:'000000'}),
        C_(r.qty||'',       {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[3],color:'000000'}),
        C_(r.reports||'—',  {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[4],color:'000000'}),
        C_(fM(r.salary),    {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[5],color:'000000'}),
        C_(fM(r.total),     {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[6],color:'000000'}),
      ]})),
      totRow('الإجمالي', fM(tot), 7, W[6]),
    ];
    sections.push(tableSection(T_(rows, W)));
  }

  // ════ 6b. ORG CHART (portrait page) ══════════════════════
  if (data.hrRows?.length) {
    const orgChildren = buildOrgChartSection(data.hrRows);
    sections.push({ properties:{ page: PAGE_PORTRAIT }, children: orgChildren });
  }

  // ════ 7. FIXED COSTS ══════════════════════════════════════
  if (data.fixedRows?.length) {
    const W = norm([149,500,1000,300,1400,1400,1400], TW_L);
    const hdrs = ['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة','التكلفة الشهرية الإجمالية'];
    let tot=0; data.fixedRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    const rows = [
      secHdr('التكاليف الثابتة', 7),
      colHdr(hdrs, W, C.COL_HDR_BLU, '000000'),
      ...data.fixedRows.map((r,i)=>new TableRow({children:[
        C_(i+1,          {fill:C.ROW_ODD,                  sz:28,w:W[0]}),
        C_(r.cat||'',    {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[1],color:'000000'}),
        C_(r.bayan||'',  {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[2],color:'000000'}),
        C_(r.qty||'',    {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[3],color:'000000'}),
        C_(r.notes||'',  {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[4],color:'000000'}),
        C_(fM(r.price),  {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[5],color:'000000'}),
        C_(fM(r.total),  {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[6],color:'000000'}),
      ]})),
      totRow('الإجمالي', fM(tot), 7, W[6]),
    ];
    sections.push(tableSection(T_(rows, W)));
  }

  // ════ 8. DEPRECIATION ════════════════════════════════════
  if (data.depRows?.length) {
    const W = norm([149,280,1350,450,225,1500,1500,1500], TW_L);
    const hdrs = ['#','الصنف','البيان','نسبة الاهتلاك','العدد','ملاحظات','قيمة الاهتلاك للواحدة','قيمة الاهتلاك الإجمالية'];
    let tot=0; data.depRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    const rows = [
      secHdr('الاهتلاك', 8),
      colHdr(hdrs, W, C.COL_HDR_BLU, '000000'),
      ...data.depRows.map((r,i)=>new TableRow({children:[
        C_(i+1,                  {fill:C.ROW_ODD,                  sz:28,w:W[0]}),
        C_(r.cat||'',            {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[1],color:'000000'}),
        C_(r.bayan||'',          {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[2],color:'000000'}),
        C_((r.pct||'0')+' %',    {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[3],color:'000000'}),
        C_(r.qty||'',            {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[4],color:'000000'}),
        C_(r.notes||'',          {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[5],color:'000000'}),
        C_(fM(r.perUnit),        {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[6],color:'000000'}),
        C_(fM(r.total),          {fill:i%2===0?C.WHITE:C.ROW_EVEN, sz:28,w:W[7],color:'000000'}),
      ]})),
      totRow('إجمالي قيمة الاهتلاك', fM(tot), 8, W[7]),
    ];
    sections.push(tableSection(T_(rows, W)));
  }

  const doc = new Document({ sections });
  return Packer.toBuffer(doc);
}

// ── Org Chart Section Builder ─────────────────────────────────
function buildOrgChartSection(hrRows) {
  // Parse employees
  const employees = hrRows.map(r => ({
    pos:     r.position||'',
    type:    r.type||'',
    qty:     parseInt(r.qty)||1,
    reports: r.reports||'',
  })).filter(e=>e.pos);

  // Build tree map
  const tree = {};
  employees.forEach(e => {
    const key = e.reports||'__root__';
    if (!tree[key]) tree[key] = [];
    tree[key].push(e);
  });

  const allPos = employees.map(e=>e.pos);
  const roots = employees.filter(e=>!e.reports||!allPos.includes(e.reports));

  const children = [];

  // Title
  children.push(new Paragraph({
    bidirectional:true, alignment:AlignmentType.CENTER, spacing:{before:400,after:300},
    children:[new TextRun({text:'الهيكل التنظيمي',bold:true,size:36,font:FONT,color:C.DARK_BLUE})],
  }));

  // Render tree using tables: each level is a row of cells
  // We'll render level by level
  function getLevels(nodeList, depth=0) {
    const levels = [];
    levels.push({ depth, nodes: nodeList });
    nodeList.forEach(n => {
      const ch = tree[n.pos]||[];
      if (ch.length) {
        const sub = getLevels(ch, depth+1);
        sub.forEach((l,i) => {
          const existing = levels.find(x=>x.depth===l.depth);
          if (existing) existing.nodes.push(...l.nodes);
          else levels.push(l);
        });
      }
    });
    return levels;
  }

  const levels = getLevels(roots);
  // Sort levels by depth
  levels.sort((a,b)=>a.depth-b.depth);

  levels.forEach(level => {
    const nodes = level.nodes;
    // Each node = one box
    const cellW = Math.floor(TW_P / Math.max(nodes.length, 1));
    const lastW = TW_P - cellW * (nodes.length-1);

    const row = new TableRow({ children: nodes.map((e,i) => {
      const bg = orgColor(e.type);
      const label = e.qty>1 ? `${e.pos} (${e.qty})` : e.pos;
      return new TableCell({
        width:{size: i===nodes.length-1?lastW:cellW, type:WidthType.DXA},
        shading:{type:ShadingType.CLEAR, fill:bg, color:bg},
        borders: BORDERS_AUTO,
        margins:{top:120,bottom:120,left:120,right:120},
        children:[
          new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
            children:[new TextRun({text:`👤 ${label}`,bold:true,size:24,font:FONT,color:C.WHITE})]}),
          new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
            children:[new TextRun({text:e.type,size:20,font:FONT,color:C.WHITE})]}),
        ],
      });
    })});

    const tbl = new Table({
      width:{size:TW_P,type:WidthType.DXA},
      columnWidths: nodes.map((_,i)=>i===nodes.length-1?lastW:cellW),
      bidirectional:true,
      rows:[row],
    });
    children.push(tbl);

    // Add connector line if not last level
    if (level.depth < levels[levels.length-1].depth) {
      children.push(new Paragraph({
        bidirectional:true, alignment:AlignmentType.CENTER,
        spacing:{before:0,after:0},
        children:[new TextRun({text:'│', size:28, font:FONT, color:C.DARK_BLUE})],
      }));
    }
  });

  return children;
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
