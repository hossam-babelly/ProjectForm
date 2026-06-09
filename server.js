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
const JSZip      = require('jszip');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell,
  AlignmentType, WidthType, ShadingType,
  BorderStyle, VerticalAlign, PageOrientation,
  Header, ImageRun, PageBreak,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom,
  HorizontalPositionAlign
} = require('docx');

// ── Assets (letterhead image + SmartArt diagram templates) ───
const ASSETS_DIR = path.join(__dirname, 'assets');
function readAsset(rel) {
  try { return fs.readFileSync(path.join(ASSETS_DIR, rel)); } catch { return null; }
}
const LETTERHEAD_PNG = readAsset('letterhead.png');

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
      {width:6},{width:16},{width:24},{width:10},{width:20},{width:20},{width:8},{width:20}
    ];
    const hr = s1.getRow(1);
    hr.height = 24;
    ['#','الصنف','البيان','اهتلاك','ملاحظات','التكلفة للواحدة ($)','العدد','التكلفة الإجمالية ($)']
      .forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle(C_HEADER)); });

    let tot = 0;
    data.foundingRows.forEach((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot += v;
      const row = s1.getRow(i+2);
      row.height = 20;
      [r ? i+1 : '', r.cat, r.bayan, r.dep?'✓':'', r.notes, parseFloat(r.price)||0, r.qty, v]
        .forEach((val,c) => {
          row.getCell(c+1).value = val;
          Object.assign(row.getCell(c+1), dStyle());
          if(c===5||c===7) row.getCell(c+1).numFmt='"$"#,##0.00';
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
    s4.columns = [{width:6},{width:16},{width:24},{width:20},{width:22},{width:8},{width:22}];
    const hr = s4.getRow(1); hr.height=24;
    ['#','الصنف','البيان','ملاحظات','التكلفة الشهرية للواحدة ($)','العدد','التكلفة الشهرية الإجمالية ($)']
      .forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle(C_HEADER)); });

    const isSalaryRow=(r)=>(r.cat==='رواتب') && (String(r.bayan||'').includes('الموظفين') || String(r.notes||'').includes('تلقائي'));
    let tot=0;
    data.fixedRows.forEach((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot += v;   // all rows monthly now
      const row = s4.getRow(i+2); row.height=20;
      [i+1, r.cat, r.bayan, r.notes, parseFloat(r.price)||0, r.qty, v]
        .forEach((val,c) => {
          row.getCell(c+1).value=val; Object.assign(row.getCell(c+1),dStyle());
          if(c===4||c===6) row.getCell(c+1).numFmt='"$"#,##0.00';
        });
    });
    const n = data.fixedRows.length;
    s4.mergeCells(`A${n+2}:F${n+2}`);
    s4.getCell(`A${n+2}`).value='الإجمالي'; Object.assign(s4.getCell(`A${n+2}`),tStyle(C_TOTAL));
    s4.getCell(`G${n+2}`).value=tot; s4.getCell(`G${n+2}`).numFmt='"$"#,##0.00'; Object.assign(s4.getCell(`G${n+2}`),tStyle(C_TOTAL));
  }

  // ── Sheet 5: الموارد البشرية ──────────────────────────
  if (data.hrRows?.length) {
    const s5 = wb.addWorksheet('الموارد البشرية', { rightToLeft:true });
    s5.columns = [{width:6},{width:20},{width:16},{width:20},{width:20},{width:14},{width:8},{width:22}];
    const hr = s5.getRow(1); hr.height=24;
    ['#','المنصب','النوع','تابع لـ','الراتب الشهري الفردي ($)','عدد أشهر الدوام في السنة','العدد','الراتب الشهري الإجمالي ($)']
      .forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle(C_HEADER)); });

    let tot=0;
    data.hrRows.forEach((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      const q=parseFloat(r.qty)||0;
      const s=parseFloat(String(r.salary||'0').replace(/[^0-9.-]/g,''))||0;
      const mo=parseFloat(r.months)||12;
      tot += q*s*mo;
      const row = s5.getRow(i+2); row.height=20;
      [i+1, r.position, r.type, r.reports, parseFloat(r.salary)||0, mo, r.qty, v]
        .forEach((val,c) => {
          row.getCell(c+1).value=val; Object.assign(row.getCell(c+1),dStyle());
          if(c===4||c===7) row.getCell(c+1).numFmt='"$"#,##0.00';
        });
    });
    const n = data.hrRows.length;
    s5.mergeCells(`A${n+2}:G${n+2}`);
    s5.getCell(`A${n+2}`).value='إجمالي الرواتب (سنوياً)'; Object.assign(s5.getCell(`A${n+2}`),tStyle(C_TOTAL));
    s5.getCell(`H${n+2}`).value=tot; s5.getCell(`H${n+2}`).numFmt='"$"#,##0.00'; Object.assign(s5.getCell(`H${n+2}`),tStyle(C_TOTAL));
  }

  return wb.xlsx.writeBuffer();
}

// ════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════
//  WORD GENERATOR  (v7 — matches "الملف المطلوب.docx" exactly)
//  Sakkal Majalla · full-width RTL tables · letterhead header
//  · real Organisation-Chart SmartArt coloured by employee type
// ════════════════════════════════════════════════════════════

const C = {
  DARK_BLUE:'1F3864', GREY_HDR:'AEAAAA', COL_HDR_BLK:'262626',
  COL_HDR_BLU:'D9E2F3', PROD1:'8EA9DB', PROD2:'F4B084', PROD3:'8EAADB',
  MON_ODD:'D9E1F2', MON_EVEN:'B4C6E7', MON_ODD2:'FCE4D6', MON_EVEN2:'F8CBAD',
  SUBTOT:'808080', SUM_LBL:'DEEAF6', SUM_VAL:'FBE4D5',
  ROW_ODD:'FBE4D5', ROW_EVEN:'F2F2F2', WHITE:'FFFFFF',
  NUM_TINT:'FBE4D5',                 // peach tint of the "#" column cells
  TITLE_BLUE:'2F5496', TITLE_ORANGE:'ED7D31',
  // org-chart node colours by employee type (theme accents)
  ORG_ADM:'accent1', ORG_EXE:'accent6', ORG_SVC:'accent2',
};

const PAGE = {
  size:{ width:11906, height:16838, orientation:PageOrientation.LANDSCAPE },
  margin:{ top:720, right:720, bottom:720, left:720, header:720, footer:1545 },
};
const TW   = 15388;            // table content width (DXA)
const FONT = 'Sakkal Majalla';

// exact per-table column widths (DXA) extracted from الملف المطلوب
const COLW = {
  summary : [7740,7648],
  founding: [454,857,4206,1165,4677,1529,713,1787],
  products: [4822,4841,5705],
  revenue : [2313,1167,940,940,940,940,940,982,982,982,941,1094,1094,1113],
  ops     : [1332,877,1213,964,986,986,986,987,1036,987,987,1036,990,990,1011],
  hr      : [459,2791,1800,2600,2100,2300,988,2240],
  fixed   : [609,1530,3268,4318,2342,902,2419],
  dep     : [411,857,4164,1397,4638,1486,687,1748],
};

const BDR = (c='auto',s=4) => ({style:BorderStyle.SINGLE,size:s,color:c});
const BORDERS = {top:BDR(),bottom:BDR(),left:BDR(),right:BDR()};

// ── Cell ──────────────────────────────────────────────────────
function C_(text, o={}) {
  const {fill,bold=true,sz=28,color,align=AlignmentType.CENTER,w,colSpan,rowSpan,
         vAlign=VerticalAlign.CENTER,font=FONT} = o;
  const tc = {};
  if (fill) tc.shading = {type:ShadingType.CLEAR, fill, color:'auto'};
  if (w)    tc.width   = {size:w, type:WidthType.DXA};
  if (colSpan) tc.columnSpan = colSpan;
  if (rowSpan>1) tc.rowSpan = rowSpan;
  tc.borders = BORDERS;
  tc.margins = {top:40,bottom:40,left:80,right:80};
  tc.verticalAlign = vAlign;
  const textColor = color || (
    (fill===C.DARK_BLUE||fill===C.COL_HDR_BLK||fill===C.SUBTOT) ? C.WHITE : '000000'
  );
  return new TableCell({...tc, children:[new Paragraph({
    bidirectional:true, alignment:align, spacing:{after:0,line:240,lineRule:'auto'},
    children:[new TextRun({text:String(text??''),bold,size:sz,font,color:textColor})],
  })]});
}

// ── Table ─────────────────────────────────────────────────────
function T_(rows, W) {
  return new Table({width:{size:TW,type:WidthType.DXA},columnWidths:W,
                    bidirectional:true,visuallyRightToLeft:true,rows});
}
// section title row
function secHdr(title, cols, fill=C.DARK_BLUE) {
  return new TableRow({tableHeader:true, children:[
    C_(title,{fill,bold:true,sz:32,colSpan:cols,w:TW,align:AlignmentType.CENTER,
      color:fill===C.GREY_HDR?'000000':C.WHITE})]});
}
// column header row
function colHdr(labels, W, fill=C.COL_HDR_BLU) {
  return new TableRow({tableHeader:true, children:
    labels.map((h,i)=>C_(h,{fill,bold:true,sz:28,w:W[i],
      color:fill===C.COL_HDR_BLK?C.WHITE:'000000'}))});
}
// total row (label spans all but last column)
function totRow(label, value, cols, lastW, fill=C.DARK_BLUE) {
  return new TableRow({children:[
    C_(label,{fill,bold:true,sz:28,colSpan:cols-1,w:TW-lastW,align:AlignmentType.CENTER}),
    C_(value,{fill,bold:true,sz:28,w:lastW})]});
}
// vertical spacer paragraph
function SP(before=200,sz=28) {
  return new Paragraph({spacing:{before,after:0},
    children:[new TextRun({text:'',size:sz,font:FONT})]});
}
function PB(){ return new Paragraph({children:[new PageBreak()]}); }

function fN(v){const n=parseFloat(String(v||'0').replace(/[^0-9.-]/g,''));
  if(!n||isNaN(n))return '0';
  return n.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:2});}
function fM(v){const n=parseFloat(String(v||'0').replace(/[^0-9.-]/g,''));
  if(isNaN(n))return '0 $';
  return n.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:2})+' $';}

const PAL=[{dark:C.PROD1,mo:C.MON_ODD ,me:C.MON_EVEN },
           {dark:C.PROD2,mo:C.MON_ODD2,me:C.MON_EVEN2},
           {dark:C.PROD3,mo:C.MON_ODD ,me:C.MON_EVEN }];
function PC(i){return PAL[i%PAL.length];}
function norm(arr,tw){const s=arr.reduce((a,b)=>a+b,0);
  const r=arr.map(w=>Math.round(w*tw/s));
  r[r.length-1]=tw-r.slice(0,-1).reduce((a,b)=>a+b,0);return r;}

// "#" cell (peach) + alternating data-row fill
function dFill(i){ return i%2===0 ? undefined : C.ROW_EVEN; }

// ── letterhead header (anchored full-page background image) ──
function buildHeader() {
  if (!LETTERHEAD_PNG) return undefined;
  return new Header({children:[ new Paragraph({children:[ new ImageRun({
    data: LETTERHEAD_PNG,
    transformation:{ width:1122, height:791 },   // EMU 10690495×7535545 ÷ 9525
    floating:{
      horizontalPosition:{ relative:HorizontalPositionRelativeFrom.PAGE, align:HorizontalPositionAlign.LEFT },
      verticalPosition:{ relative:VerticalPositionRelativeFrom.PARAGRAPH, offset:-449580 },
      behindDocument:true, allowOverlap:true,
    },
  }) ] }) ]});
}

// ════════════════════════════════════════════════════════════════
async function generateWord(data) {
  const pids = Object.keys(data.products||{}).filter(p=>data.products[p]?.name);
  const ch = [];

  // ── Title (the logo sits in the page header) ─────────────
  ch.push(
    SP(0,26),
    new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:1100,after:0},
      children:[new TextRun({text:'قسم المشاريع التنموية',bold:true,size:34,font:FONT,color:C.TITLE_BLUE})]}),
    new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:60,after:260},
      children:[new TextRun({text:'نموذج طرح دراسة مشروع',bold:true,size:28,font:FONT,color:C.TITLE_ORANGE})]}),
  );

  // each table lives in its OWN section, vertically centred on its page
  const sections = [];
  const mkSec = (children, valign=VerticalAlign.CENTER) =>
    ({ properties:{ page:PAGE, verticalAlign:valign }, headers:{default:buildHeader()}, children });
  const page = (...tbls)=>{ sections.push(mkSec(tbls, VerticalAlign.CENTER)); };

  // ══ 1. SUMMARY ══════════════════════════════════════════
  {
    const W = norm(COLW.summary,TW);
    ch.push(T_([
      new TableRow({children:[C_('ملخص معلومات المشروع',{fill:C.DARK_BLUE,bold:true,sz:32,colSpan:2,w:TW,align:AlignmentType.CENTER,color:C.WHITE})]}),
      ...[
        ['فكرة المشروع',data.projectIdea||''],
        ['اسم مقدم المشروع',data.applicantName||''],
        ['رقم الهاتف',data.applicantPhone||''],
        ['إجمالي التكاليف التأسيسية',data.summary?.foundingTotal||'$0'],
        ['إجمالي الإيرادات المتوقعة (سنوياً)',data.summary?.revenueAnnual||'$0'],
        ['إجمالي التكاليف التشغيلية (سنوياً)',data.summary?.opsAnnual||'$0'],
        ['إجمالي التكاليف الثابتة (سنوياً)',data.summary?.fixedAnnual||'$0'],
        ['الاهتلاك (سنوياً)',data.summary?.depreciation||'$0'],
        ['الربح الصافي (سنوياً)',data.summary?.netProfit||'$0'],
        ['عدد الموظفين في المشروع',String(data.summary?.employees||'0')],
      ].map(([l,v])=>new TableRow({children:[
        C_(l,{fill:C.SUM_LBL,bold:true,sz:28,w:W[0],align:AlignmentType.CENTER}),
        C_(v,{fill:C.SUM_VAL,bold:true,sz:28,w:W[1],align:AlignmentType.CENTER}),
      ]})),
    ],W));
  }
  // section 1 = title + summary (top-aligned)
  sections.push(mkSec(ch, VerticalAlign.TOP));

  // ══ 2. FOUNDING ═════════════════════════════════════════
  if (data.foundingRows?.length) {
    const W = norm(COLW.founding,TW);
    let tot=0; data.foundingRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    page(T_([
      secHdr('التكاليف التأسيسية',8),
      colHdr(['#','الصنف','البيان','الاهتلاك','ملاحظات','التكلفة للواحدة','العدد','التكلفة الإجمالية'],W,C.COL_HDR_BLU),
      ...data.foundingRows.map((r,i)=>new TableRow({children:[
        C_(i+1,         {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.cat||'',   {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.bayan||'', {fill:dFill(i),sz:28,w:W[2]}),
        C_(r.dep?'a':'r',{fill:dFill(i),sz:28,w:W[3],font:'Marlett',color:'000000'}),
        C_(r.notes||'', {fill:dFill(i),sz:28,w:W[4]}),
        C_(fM(r.price), {fill:dFill(i),sz:28,w:W[5]}),
        C_(r.qty||'',   {fill:dFill(i),sz:28,w:W[6]}),
        C_(fM(r.total), {fill:dFill(i),sz:28,w:W[7]}),
      ]})),
      totRow('الإجمالي',fM(tot),8,W[7]),
    ],W));
  }

  // ══ 3. PRODUCTS ═════════════════════════════════════════
  if (pids.length) {
    const W = norm(COLW.products,TW);
    const prodRows = [
      secHdr('جدول المنتجات',3,C.GREY_HDR),
      colHdr(['البيان','الواحدة','المكونات'],W,C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi)=>{
      const p=data.products[pid]; const comps=(p.components||[]).filter(c=>c); const pc=PC(pi);
      const n=Math.max(comps.length,1);
      comps.forEach((comp,ci)=>{
        if(ci===0){
          prodRows.push(new TableRow({children:[
            C_(p.name||'',{fill:pc.dark,sz:28,w:W[0],color:'000000',rowSpan:n}),
            C_(p.unit||'',{fill:pc.dark,sz:28,w:W[1],color:'000000',rowSpan:n}),
            C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000'}),
          ]}));
        } else {
          prodRows.push(new TableRow({children:[C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000'})]}));
        }
      });
    });
    page(T_(prodRows,W));
  }

  // ══ 4. REVENUE ══════════════════════════════════════════
  if (pids.length) {
    const W = norm(COLW.revenue,TW);
    const revRows = [
      secHdr('الإيرادات المتوقعة',14,C.GREY_HDR),
      colHdr(['البيان','الواحدة',...MONTHS],W,C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi)=>{
      const p=data.products[pid]; const rev=data.revenueData?.[pid]||[]; const pc=PC(pi);
      revRows.push(new TableRow({children:[
        C_(p.name,{fill:pc.dark,bold:true,sz:28,w:W[0],color:'000000'}),
        C_(p.unit||'',{fill:pc.dark,bold:true,sz:28,w:W[1],color:'000000'}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.qty),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2]})),
      ]}));
      revRows.push(new TableRow({children:[
        C_('سعر مبيع الواحدة',{fill:pc.dark,sz:28,w:W[0],color:'000000'}),
        C_('$',{fill:pc.dark,sz:28,w:W[1],color:'000000'}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.unitPrice),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2]})),
      ]}));
      revRows.push(new TableRow({children:[
        C_('سعر المبيع الإجمالي',{fill:pc.dark,sz:28,w:W[0],color:'000000'}),
        C_('$',{fill:pc.dark,sz:28,w:W[1],color:'000000'}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.total||0),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2]})),
      ]}));
    });
    const mTots = MONTHS.map((_,m)=>{let t=0;pids.forEach(pid=>{t+=parseFloat(String(data.revenueData?.[pid]?.[m]?.total||'0').replace(/[^0-9.-]/g,''))||0;});return fM(t);});
    revRows.push(new TableRow({children:[
      C_('الإجمالي',{fill:C.COL_HDR_BLK,bold:true,sz:28,colSpan:2,w:W[0]+W[1],align:AlignmentType.CENTER}),
      ...mTots.map((v,m)=>C_(v,{fill:C.COL_HDR_BLK,bold:true,sz:28,w:W[m+2]})),
    ]}));
    page(T_(revRows,W));
  }

  // ══ 5. OPS ══════════════════════════════════════════════
  if (pids.length && data.opsData) {
    const W = norm(COLW.ops,TW);
    const opsRows = [
      secHdr('التكاليف التشغيلية',15,C.GREY_HDR),
      colHdr(['البيان','الواحدة','التفاصيل',...MONTHS],W,C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi)=>{
      const p=data.products[pid]; const comps=(p.components||[]).filter(c=>c);
      const opsD=data.opsData?.[pid]||{}; const pc=PC(pi); const n=Math.max(comps.length,1);
      comps.forEach((comp,ci)=>{
        const vals=MONTHS.map((_,m)=>fM(opsD[`${ci}_${m}`]));
        if(ci===0){
          opsRows.push(new TableRow({children:[
            C_(p.name||'',{fill:pc.dark,sz:28,w:W[0],color:'000000',rowSpan:n}),
            C_(p.unit||'',{fill:pc.dark,sz:28,w:W[1],color:'000000',rowSpan:n}),
            C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000'}),
            ...vals.map((v,m)=>C_(v,{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+3],color:'000000'})),
          ]}));
        } else {
          opsRows.push(new TableRow({children:[
            C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000'}),
            ...vals.map((v,m)=>C_(v,{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+3],color:'000000'})),
          ]}));
        }
      });
      const subVals=MONTHS.map((_,m)=>fM(opsD[`sub_${m}`]||0));
      opsRows.push(new TableRow({children:[
        C_(`إجمالي ${p.name}`,{fill:C.SUBTOT,bold:true,sz:28,colSpan:3,w:W[0]+W[1]+W[2],align:AlignmentType.CENTER}),
        ...subVals.map((v,m)=>C_(v,{fill:C.SUBTOT,bold:true,sz:28,w:W[m+3]})),
      ]}));
    });
    const opsTots=MONTHS.map((_,m)=>{let t=0;pids.forEach(pid=>{t+=parseFloat(String(data.opsData?.[pid]?.[`sub_${m}`]||'0').replace(/[^0-9.-]/g,''))||0;});return fM(t);});
    opsRows.push(new TableRow({children:[
      C_('الإجمالي',{fill:C.COL_HDR_BLK,bold:true,sz:28,colSpan:3,w:W[0]+W[1]+W[2],align:AlignmentType.CENTER}),
      ...opsTots.map((v,m)=>C_(v,{fill:C.COL_HDR_BLK,bold:true,sz:28,w:W[m+3]})),
    ]}));
    page(T_(opsRows,W));
  }

  // ══ 6. HR ════════════════════════════════════════════════
  if (data.hrRows?.length) {
    const W=norm(COLW.hr,TW);
    let tot=0; data.hrRows.forEach(r=>{
      const q=parseFloat(r.qty)||0;
      const s=parseFloat(String(r.salary||'0').replace(/[^0-9.-]/g,''))||0;
      const mo=parseFloat(r.months)||12;
      tot += q*s*mo;
    });
    page(T_([
      secHdr('الموارد البشرية',8),
      colHdr(['#','المنصب','النوع','تابع لـ','الراتب الشهري الفردي','عدد أشهر الدوام في السنة','العدد','الراتب الشهري الإجمالي'],W,C.COL_HDR_BLU),
      ...data.hrRows.map((r,i)=>new TableRow({children:[
        C_(i+1,            {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.position||'', {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.type||'',     {fill:dFill(i),sz:28,w:W[2]}),
        C_(r.reports||'----------', {fill:dFill(i),sz:28,w:W[3]}),
        C_(fM(r.salary),   {fill:dFill(i),sz:28,w:W[4]}),
        C_(r.months||'12', {fill:dFill(i),sz:28,w:W[5]}),
        C_(r.qty||'',      {fill:dFill(i),sz:28,w:W[6]}),
        C_(fM(r.total),    {fill:dFill(i),sz:28,w:W[7]}),
      ]})),
      totRow('الإجمالي (سنوياً)',fM(tot),8,W[7]),
    ],W));

    // org-chart page (own section; SmartArt-like shapes injected post-process)
    sections.push(mkSec([
      new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:200,after:200},
        children:[new TextRun({text:'الهيكل التنظيمي',bold:true,size:32,font:FONT,color:C.DARK_BLUE})]}),
      new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:0,after:0},
        children:[new TextRun({text:'[[ORGCHART]]',size:2,font:FONT,color:'FFFFFF'})]}),
    ], VerticalAlign.TOP));
  }

  // ══ 7. FIXED ════════════════════════════════════════════
  if (data.fixedRows?.length) {
    const W=norm(COLW.fixed,TW);
    const isSalaryRow=(r)=>(r.cat==='رواتب') && (String(r.bayan||'').includes('الموظفين') || String(r.notes||'').includes('تلقائي'));
    let tot=0; data.fixedRows.forEach(r=>{
      const t=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;
      tot += t;   // all rows are monthly now (salary & marketing rows have qty=1)
    });
    page(T_([
      secHdr('التكاليف الثابتة',7),
      colHdr(['#','الصنف','البيان','ملاحظات','التكلفة الشهرية للواحدة','العدد','التكلفة الشهرية الإجمالية'],W,C.COL_HDR_BLU),
      ...data.fixedRows.map((r,i)=>new TableRow({children:[
        C_(i+1,         {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.cat||'',   {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.bayan||'', {fill:dFill(i),sz:28,w:W[2]}),
        C_(r.notes||'', {fill:dFill(i),sz:28,w:W[3]}),
        C_(fM(r.price), {fill:dFill(i),sz:28,w:W[4]}),
        C_(r.qty||'',   {fill:dFill(i),sz:28,w:W[5]}),
        C_(fM(r.total), {fill:dFill(i),sz:28,w:W[6]}),
      ]})),
      totRow('الإجمالي',fM(tot),7,W[6]),
    ],W));
  }

  // ══ 8. DEP ══════════════════════════════════════════════
  if (data.depRows?.length) {
    const W=norm(COLW.dep,TW);
    let tot=0; data.depRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    page(T_([
      secHdr('الاهتلاك',8),
      colHdr(['#','الصنف','البيان','نسبة الاهتلاك','ملاحظات','قيمة الاهتلاك للواحدة','العدد','قيمة الاهتلاك الإجمالية'],W,C.COL_HDR_BLU),
      ...data.depRows.map((r,i)=>new TableRow({children:[
        C_(i+1,               {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.cat||'',         {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.bayan||'',       {fill:dFill(i),sz:28,w:W[2]}),
        C_((r.pct||'0')+' %', {fill:dFill(i),sz:28,w:W[3]}),
        C_(r.notes||'',       {fill:dFill(i),sz:28,w:W[4]}),
        C_(fM(r.perUnit),     {fill:dFill(i),sz:28,w:W[5]}),
        C_(r.qty||'',         {fill:dFill(i),sz:28,w:W[6]}),
        C_(fM(r.total),       {fill:dFill(i),sz:28,w:W[7]}),
      ]})),
      totRow('إجمالي قيمة الاهتلاك',fM(tot),8,W[7]),
    ],W));
  }

  const buf = await Packer.toBuffer(new Document({sections}));
  const withChart = await injectOrgChart(buf, data.hrRows);
  return finalizeDocx(withChart);
}

// docx may emit word/fontTable.xml without a relationship → add it so the
// package validates cleanly (otherwise MS Word may flag "unreadable content").
async function finalizeDocx(buffer){
  try{
    const zip=await JSZip.loadAsync(buffer);
    const relsPath='word/_rels/document.xml.rels';
    const relsFile=zip.file(relsPath), ftFile=zip.file('word/fontTable.xml');
    if(relsFile && ftFile){
      let rels=await relsFile.async('string');
      if(!rels.includes('fontTable.xml')){
        const ids=[...rels.matchAll(/Id="rId(\d+)"/g)].map(m=>parseInt(m[1]));
        const nid=(ids.length?Math.max(...ids):0)+1;
        rels=rels.replace('</Relationships>',
          `<Relationship Id="rId${nid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>`);
        zip.file(relsPath, rels);
        return zip.generateAsync({type:'nodebuffer', compression:'DEFLATE'});
      }
    }
  }catch(_){ /* return original on any issue */ }
  return buffer;
}

// ════════════════════════════════════════════════════════════
//  SmartArt — real Organisation Chart (orgChart1 / accent1_2)
//  Reuses the genuine layout/colours/quick-style template parts
//  and generates data1.xml with per-node colour by employee type.
// ════════════════════════════════════════════════════════════
function gid(){ return '{'+crypto.randomUUID().toUpperCase()+'}'; }

// light fill (≈ accent lighter 40%) + darker border, by employee type
function orgFill(type){
  if(type==='إداري')     return {fill:'8EAADB', line:'4472C4'}; // blue
  if(type==='مزود خدمة') return {fill:'F4B083', line:'ED7D31'}; // orange
  return {fill:'A9D18E', line:'70AD47'};                        // green (تنفيذي + default)
}

// expand employees by quantity → individual nodes + position→id map
function buildPersons(hrRows){
  const persons=[];
  (hrRows||[]).forEach(r=>{
    const qty=Math.max(parseInt(r.qty)||1,1);
    for(let k=0;k<qty;k++)
      persons.push({id:gid(),pos:(r.position||'').trim(),type:(r.type||'').trim(),reports:(r.reports||'').trim()});
  });
  const posId={}; persons.forEach(p=>{ if(p.pos && !(p.pos in posId)) posId[p.pos]=p.id; });
  return {persons,posId};
}

// build hierarchy tree (roots + children) from persons
function buildOrgTree(persons,posId){
  const byId={}; persons.forEach(p=>byId[p.id]={...p,children:[]});
  const roots=[];
  persons.forEach(p=>{
    const pid=(p.reports && posId[p.reports] && posId[p.reports]!==p.id) ? posId[p.reports] : null;
    if(pid && byId[pid]) byId[pid].children.push(byId[p.id]); else roots.push(byId[p.id]);
  });
  return roots;
}

// ── org-chart layout (positions in EMU; parents centred over children) ──
function orgLayout(hrRows){
  const {persons,posId}=buildPersons(hrRows);
  const roots=buildOrgTree(persons,posId);
  const boxW=1750000, boxH=820000, hGap=240000, vGap=620000;
  let cursor=0;
  (function place(nodes,depth){
    nodes.forEach(n=>{
      n.depth=depth;
      if(!n.children.length){ n.x=cursor*(boxW+hGap); cursor++; }
      else { place(n.children,depth+1); n.x=(n.children[0].x+n.children[n.children.length-1].x)/2; }
      n.y=depth*(boxH+vGap);
    });
  })(roots,0);
  const all=[]; (function col(nodes){nodes.forEach(n=>{all.push(n); n.children.length&&col(n.children);});})(roots);
  if(!all.length) return null;
  const W=Math.max(...all.map(n=>n.x))+boxW;
  const H=Math.max(...all.map(n=>n.y))+boxH;
  return {all,boxW,boxH,vGap,W,H};
}

// ── SmartArt-like org chart: individual anchored shapes (text renders everywhere),
//    boxes + elbow connectors, vertically centred on the page (recomputed per size).
function buildOrgChartParagraphXml(hrRows){
  const L=orgLayout(hrRows);
  if(!L) return '';
  const EMU_IN=914400, TWIP=635;
  const pageW=16838*TWIP, pageH=11906*TWIP;
  const safeTop=1.75*EMU_IN, safeBot=1.2*EMU_IN;
  const availH=(pageH-safeBot)-safeTop;
  const targetW=pageW-2*520000;
  const scale=Math.min(1, targetW/L.W, availH/L.H);
  const sx=v=>Math.round(v*scale);
  const chartW=Math.round(L.W*scale), chartH=Math.round(L.H*scale);
  const xOff=Math.round((pageW-chartW)/2);
  const yOff=Math.round(safeTop+(availH-chartH)/2);
  const px=x=>xOff+sx(x), py=y=>yOff+sx(y);
  const nameSz=Math.max(16,Math.min(26,Math.round(26*scale)));
  const typeSz=Math.max(12,Math.min(20,Math.round(20*scale)));
  const F=`<w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>`;
  const lineClr='4472C4';
  let z=8000, runs='';

  function anchorWrap(x,y,w,h,inner,name){
    const zi=z++;
    return `<w:r><w:drawing>`+
      `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${zi}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`+
      `<wp:simplePos x="0" y="0"/>`+
      `<wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>`+
      `<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>`+
      `<wp:extent cx="${w}" cy="${h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>`+
      `<wp:docPr id="${zi}" name="${name}${zi}"/><wp:cNvGraphicFramePr/>`+
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`+
      `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">`+
      `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">`+
      `<wps:cNvPr id="${zi}" name="${name}${zi}"/>${inner}`+
      `</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  }
  function lineAnchor(x1,y1,x2,y2){
    const x=Math.min(x1,x2), y=Math.min(y1,y2), w=Math.max(Math.abs(x2-x1),1), h=Math.max(Math.abs(y2-y1),1);
    return anchorWrap(x,y,w,h,
      `<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`+
      `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`+
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="${lineClr}"/></a:solidFill></a:ln></wps:spPr>`+
      `<wps:bodyPr/>`, 'ln');
  }
  function boxAnchor(n){
    const c=orgFill(n.type), w=sx(L.boxW), h=sx(L.boxH);
    const txt=`<wps:txbx><w:txbxContent>`+
      `<w:p><w:pPr><w:bidi/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>`+
        `<w:r><w:rPr>${F}<w:b/><w:bCs/><w:sz w:val="${nameSz}"/><w:szCs w:val="${nameSz}"/><w:color w:val="000000"/></w:rPr><w:t xml:space="preserve">${escXml(n.pos)}</w:t></w:r></w:p>`+
      `<w:p><w:pPr><w:bidi/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>`+
        `<w:r><w:rPr>${F}<w:sz w:val="${typeSz}"/><w:szCs w:val="${typeSz}"/><w:color w:val="000000"/></w:rPr><w:t xml:space="preserve">(${escXml(n.type)})</w:t></w:r></w:p>`+
      `</w:txbxContent></wps:txbx>`;
    return anchorWrap(px(n.x),py(n.y),w,h,
      `<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`+
      `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>`+
      `<a:solidFill><a:srgbClr val="${c.fill}"/></a:solidFill>`+
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="${c.line}"/></a:solidFill></a:ln></wps:spPr>`+
      txt+
      `<wps:bodyPr rot="0" wrap="square" lIns="18000" tIns="9000" rIns="18000" bIns="9000" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`, 'box');
  }

  // connectors first (lower z), then boxes (higher z, drawn on top)
  L.all.forEach(n=>{
    if(!n.children.length) return;
    const pcx=px(n.x+L.boxW/2), pbot=py(n.y+L.boxH), busY=py(n.y+L.boxH+L.vGap/2);
    runs+=lineAnchor(pcx,pbot,pcx,busY);
    const cxs=n.children.map(c=>px(c.x+L.boxW/2));
    runs+=lineAnchor(Math.min(pcx,...cxs),busY,Math.max(pcx,...cxs),busY);
    n.children.forEach(c=>{const ccx=px(c.x+L.boxW/2); runs+=lineAnchor(ccx,busY,ccx,py(c.y));});
  });
  L.all.forEach(n=>{ runs+=boxAnchor(n); });

  return `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${runs}</w:p>`;
}

// replace the [[ORGCHART]] marker paragraph with the anchored-shapes paragraph
async function injectOrgChart(buffer, hrRows){
  const {persons}=buildPersons(hrRows);
  if(!persons.length) return buffer;
  const para=buildOrgChartParagraphXml(hrRows);
  if(!para) return buffer;
  const zip=await JSZip.loadAsync(buffer);
  let xml=await zip.file('word/document.xml').async('string');
  xml=xml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>).)*?\[\[ORGCHART\]\](?:(?!<\/w:p>).)*?<\/w:p>/s, para);
  zip.file('word/document.xml', xml);
  return zip.generateAsync({type:'nodebuffer', compression:'DEFLATE'});
}

function escXml(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}


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
// build a friendly download filename: "<applicant> - <date> - <serial>.docx"
function fileBase(sub){
  const name = String(sub.applicantName||'').trim() || 'مقدم المشروع';
  const date = String(sub.submittedAt||new Date().toISOString()).slice(0,10); // YYYY-MM-DD
  const serial = String(sub.serial||0).padStart(4,'0');
  return `${name} - ${date} - ${serial}`.replace(/[\\\/:*?"<>|\r\n]/g,'_').trim();
}

app.post('/api/submit', async (req, res) => {
  try {
    const body = req.body;
    if (!body?.projectIdea) return res.status(400).json({ success:false, message:'فكرة المشروع مطلوبة' });

    const id = genId();
    const index = loadIndex();
    const serial = index.reduce((m,s)=>Math.max(m, parseInt(s.serial)||0), 0) + 1;
    const submission = { id, serial, submittedAt: body.submittedAt||new Date().toISOString(), ...body };
    fs.writeFileSync(path.join(DATA_DIR,`${id}.json`), JSON.stringify(submission,null,2), 'utf8');
    index.unshift({ id, serial, projectIdea:body.projectIdea.substring(0,80), applicantName:body.applicantName||'', submittedAt:submission.submittedAt, summary:body.summary||{} });
    saveIndex(index);

    // generate + save files (needed for the immediate download)
    const [excelBuf, wordBuf] = await Promise.all([generateExcel(submission), generateWord(submission)]);
    fs.writeFileSync(path.join(DATA_DIR,`${id}.xlsx`), excelBuf);
    fs.writeFileSync(path.join(DATA_DIR,`${id}.docx`), wordBuf);

    console.log(`✅ طلب جديد: ${id}`);
    // respond NOW so the browser can download immediately …
    res.json({ success:true, id, fileName: fileBase(submission)+'.docx' });
    // … and send the email in the background (don't block the response)
    sendEmail(submission, id, excelBuf, wordBuf).catch(e => console.error('Email error:', e.message));
  } catch(err) {
    console.error('❌', err);
    if (!res.headersSent) return res.status(500).json({ success:false, message:err.message });
  }
});

app.get('/api/submissions', (req, res) => res.json({ success:true, submissions:loadIndex() }));

app.get('/api/submissions/:id', (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp = path.join(DATA_DIR,`${sid}.json`);
  if (!fs.existsSync(fp)) return res.status(404).json({ success:false });
  res.json({ success:true, submission:JSON.parse(fs.readFileSync(fp,'utf8')) });
});

app.get('/api/download/:id/excel', async (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp  = path.join(DATA_DIR, `${sid}.json`);
  if (!fs.existsSync(fp)) return res.status(404).send('الطلب غير موجود');
  try {
    const data = JSON.parse(fs.readFileSync(fp,'utf8'));
    const cached = path.join(DATA_DIR, `${sid}.xlsx`);
    const buf = fs.existsSync(cached) ? fs.readFileSync(cached) : await generateExcel(data);
    const fn = fileBase(data)+'.xlsx';
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="project_${sid}.xlsx"; filename*=UTF-8''${encodeURIComponent(fn)}`);
    res.send(buf);
  } catch(e) {
    console.error('Excel gen error:', e.message);
    res.status(500).send('خطأ في توليد الملف: ' + e.message);
  }
});

app.get('/api/download/:id/word', async (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp  = path.join(DATA_DIR, `${sid}.json`);
  if (!fs.existsSync(fp)) return res.status(404).send('الطلب غير موجود');
  try {
    const data = JSON.parse(fs.readFileSync(fp,'utf8'));
    const cached = path.join(DATA_DIR, `${sid}.docx`);
    const buf = fs.existsSync(cached) ? fs.readFileSync(cached) : await generateWord(data);
    const fn = fileBase(data)+'.docx';
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="project_${sid}.docx"; filename*=UTF-8''${encodeURIComponent(fn)}`);
    res.send(buf);
  } catch(e) {
    console.error('Word gen error:', e.message);
    res.status(500).send('خطأ في توليد الملف: ' + e.message);
  }
});

// PATCH /api/submissions/:id/update — update basic fields and regenerate files
app.patch('/api/submissions/:id/update', async (req, res) => {
  try {
    const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
    const fp  = path.join(DATA_DIR, `${sid}.json`);
    if (!fs.existsSync(fp)) return res.status(404).json({ success:false, message:'الطلب غير موجود' });

    const submission = JSON.parse(fs.readFileSync(fp,'utf8'));
    const { projectIdea, applicantName, applicantPhone } = req.body;

    // Update fields
    if (projectIdea)   submission.projectIdea   = projectIdea;
    if (applicantName !== undefined)  submission.applicantName  = applicantName;
    if (applicantPhone !== undefined) submission.applicantPhone = applicantPhone;

    // Save updated JSON
    fs.writeFileSync(fp, JSON.stringify(submission, null, 2), 'utf8');

    // Update index entry
    const index = loadIndex();
    const idx = index.findIndex(s=>s.id===sid);
    if (idx>=0) index[idx].projectIdea = submission.projectIdea.substring(0,80);
    saveIndex(index);

    // Regenerate Excel and Word files
    const [excelBuf, wordBuf] = await Promise.all([generateExcel(submission), generateWord(submission)]);
    fs.writeFileSync(path.join(DATA_DIR,`${sid}.xlsx`), excelBuf);
    fs.writeFileSync(path.join(DATA_DIR,`${sid}.docx`), wordBuf);

    return res.json({ success:true });
  } catch(err) {
    console.error('PATCH error:', err);
    return res.status(500).json({ success:false, message: err.message });
  }
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

// PUT /api/submissions/:id — update a submission and regenerate files
app.put('/api/submissions/:id', async (req, res) => {
  try {
    const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
    const fp = path.join(DATA_DIR, `${sid}.json`);
    if (!fs.existsSync(fp)) return res.status(404).json({ success:false, message:'غير موجود' });
    
    const body = req.body;
    if (!body?.projectIdea) return res.status(400).json({ success:false, message:'فكرة المشروع مطلوبة' });
    
    const existing = JSON.parse(fs.readFileSync(fp,'utf8'));
    const serial = existing.serial || (loadIndex().reduce((m,s)=>Math.max(m, parseInt(s.serial)||0), 0) + 1);
    const updated = { ...existing, ...body, id: sid, serial, submittedAt: existing.submittedAt||body.submittedAt, updatedAt: new Date().toISOString() };
    
    fs.writeFileSync(fp, JSON.stringify(updated, null, 2), 'utf8');
    
    // Update index
    const index = loadIndex();
    const idx = index.findIndex(s=>s.id===sid);
    if (idx>=0) {
      index[idx] = { ...index[idx], serial, projectIdea: body.projectIdea.substring(0,80), applicantName: body.applicantName||'', summary: body.summary||{}, updatedAt: updated.updatedAt };
      saveIndex(index);
    }
    
    // Regenerate files
    const [excelBuf, wordBuf] = await Promise.all([generateExcel(updated), generateWord(updated)]);
    fs.writeFileSync(path.join(DATA_DIR,`${sid}.xlsx`), excelBuf);
    fs.writeFileSync(path.join(DATA_DIR,`${sid}.docx`), wordBuf);
    
    console.log(`✅ تم تحديث الطلب: ${sid}`);
    return res.json({ success:true, id: sid, fileName: fileBase(updated)+'.docx' });
  } catch(err) {
    console.error('❌ خطأ في التحديث:', err);
    return res.status(500).json({ success:false, message: err.message });
  }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname,'public','index.html')));

app.listen(PORT, () => console.log(`🚀 يعمل على المنفذ ${PORT}`));
module.exports = app;
