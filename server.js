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
  founding: [454,857,4206,1165,713,4677,1529,1787],
  products: [4822,4841,5705],
  revenue : [2313,1167,940,940,940,940,940,982,982,982,941,1094,1094,1113],
  ops     : [1332,877,1213,964,986,986,986,987,1036,987,987,1036,990,990,1011],
  hr      : [459,2791,2696,988,3964,2250,2240],
  fixed   : [609,1530,3268,902,4318,2342,2419],
  dep     : [411,857,4164,1397,687,4638,1486,1748],
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

  // add a new page then a top spacer, before each subsequent table
  const page = (...tbls)=>{ ch.push(PB(), SP(1300,28), ...tbls); };

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
        C_(l,{fill:C.SUM_LBL,bold:true,sz:28,w:W[0],align:AlignmentType.RIGHT}),
        C_(v,{fill:C.SUM_VAL,bold:true,sz:28,w:W[1]}),
      ]})),
    ],W));
  }

  // ══ 2. FOUNDING ═════════════════════════════════════════
  if (data.foundingRows?.length) {
    const W = norm(COLW.founding,TW);
    let tot=0; data.foundingRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    page(T_([
      secHdr('التكاليف التأسيسية',8),
      colHdr(['#','الصنف','البيان','الاهتلاك','العدد','ملاحظات','التكلفة للواحدة','التكلفة الإجمالية'],W,C.COL_HDR_BLU),
      ...data.foundingRows.map((r,i)=>new TableRow({children:[
        C_(i+1,         {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.cat||'',   {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.bayan||'', {fill:dFill(i),sz:28,w:W[2]}),
        C_(r.dep?'a':'r',{fill:dFill(i),sz:28,w:W[3],font:'Marlett',color:'000000'}),
        C_(r.qty||'',   {fill:dFill(i),sz:28,w:W[4]}),
        C_(r.notes||'', {fill:dFill(i),sz:28,w:W[5]}),
        C_(fM(r.price), {fill:dFill(i),sz:28,w:W[6]}),
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
    let tot=0; data.hrRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    page(T_([
      secHdr('الموارد البشرية',7),
      colHdr(['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي','الراتب الشهري الإجمالي'],W,C.COL_HDR_BLU),
      ...data.hrRows.map((r,i)=>new TableRow({children:[
        C_(i+1,            {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.position||'', {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.type||'',     {fill:dFill(i),sz:28,w:W[2]}),
        C_(r.qty||'',      {fill:dFill(i),sz:28,w:W[3]}),
        C_(r.reports||'----------', {fill:dFill(i),sz:28,w:W[4]}),
        C_(fM(r.salary),   {fill:dFill(i),sz:28,w:W[5]}),
        C_(fM(r.total),    {fill:dFill(i),sz:28,w:W[6]}),
      ]})),
      totRow('الإجمالي',fM(tot),7,W[6]),
    ],W));

    // org-chart page (native table — colours by type, children under parents)
    ch.push(
      PB(), SP(1100,28),
      new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:0,after:240},
        children:[new TextRun({text:'الهيكل التنظيمي',bold:true,size:32,font:FONT,color:C.DARK_BLUE})]}),
      buildOrgChartTable(data.hrRows),
    );
  }

  // ══ 7. FIXED ════════════════════════════════════════════
  if (data.fixedRows?.length) {
    const W=norm(COLW.fixed,TW);
    let tot=0; data.fixedRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    page(T_([
      secHdr('التكاليف الثابتة',7),
      colHdr(['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة','التكلفة الشهرية الإجمالية'],W,C.COL_HDR_BLU),
      ...data.fixedRows.map((r,i)=>new TableRow({children:[
        C_(i+1,         {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.cat||'',   {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.bayan||'', {fill:dFill(i),sz:28,w:W[2]}),
        C_(r.qty||'',   {fill:dFill(i),sz:28,w:W[3]}),
        C_(r.notes||'', {fill:dFill(i),sz:28,w:W[4]}),
        C_(fM(r.price), {fill:dFill(i),sz:28,w:W[5]}),
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
      colHdr(['#','الصنف','البيان','نسبة الاهتلاك','العدد','ملاحظات','قيمة الاهتلاك للواحدة','قيمة الاهتلاك الإجمالية'],W,C.COL_HDR_BLU),
      ...data.depRows.map((r,i)=>new TableRow({children:[
        C_(i+1,               {fill:C.NUM_TINT,sz:28,w:W[0]}),
        C_(r.cat||'',         {fill:dFill(i),sz:28,w:W[1]}),
        C_(r.bayan||'',       {fill:dFill(i),sz:28,w:W[2]}),
        C_((r.pct||'0')+' %', {fill:dFill(i),sz:28,w:W[3]}),
        C_(r.qty||'',         {fill:dFill(i),sz:28,w:W[4]}),
        C_(r.notes||'',       {fill:dFill(i),sz:28,w:W[5]}),
        C_(fM(r.perUnit),     {fill:dFill(i),sz:28,w:W[6]}),
        C_(fM(r.total),       {fill:dFill(i),sz:28,w:W[7]}),
      ]})),
      totRow('إجمالي قيمة الاهتلاك',fM(tot),8,W[7]),
    ],W));
  }

  const sections = [{ properties:{page:PAGE}, headers:{default:buildHeader()}, children:ch }];
  const buf = await Packer.toBuffer(new Document({sections}));
  return finalizeDocx(buf);
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

// ── org chart as a native Word TABLE (renders text + colours everywhere) ──
// Children sit directly under their parent (column-span by leaf count).
function buildOrgChartTable(hrRows){
  const {persons,posId}=buildPersons(hrRows);
  const roots=buildOrgTree(persons,posId);

  // assign leaf columns + subtree span + depth
  let col=0, maxDepth=0;
  (function walk(nodes,d){
    nodes.forEach(n=>{
      n.depth=d; if(d>maxDepth) maxDepth=d;
      if(!n.children.length){ n.c0=col; n.c1=col; col++; }
      else { walk(n.children,d+1); n.c0=n.children[0].c0; n.c1=n.children[n.children.length-1].c1; }
    });
  })(roots,0);
  const ncols=Math.max(col,1);

  // group nodes by depth
  const levels=[];
  (function collect(nodes){ nodes.forEach(n=>{ (levels[n.depth]=levels[n.depth]||[]).push(n); n.children.length&&collect(n.children); }); })(roots);

  const colW=Math.floor(TW/ncols);
  const widths=Array.from({length:ncols},(_,i)=> i===ncols-1 ? TW-colW*(ncols-1) : colW);
  const NB={top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};

  function boxCell(n){
    const col=orgFill(n.type);
    const span=n.c1-n.c0+1;
    return new TableCell({
      columnSpan:span,
      shading:{type:ShadingType.CLEAR,fill:col.fill,color:'auto'},
      borders:{top:BDR(col.line,12),bottom:BDR(col.line,12),left:BDR(col.line,12),right:BDR(col.line,12)},
      margins:{top:140,bottom:140,left:80,right:80},
      verticalAlign:VerticalAlign.CENTER,
      children:[
        new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:30,line:240,lineRule:'auto'},
          children:[new TextRun({text:n.pos||'',bold:true,size:24,font:FONT,color:'000000'})]}),
        new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0,line:240,lineRule:'auto'},
          children:[new TextRun({text:'('+(n.type||'')+')',size:18,font:FONT,color:'000000'})]}),
      ],
    });
  }
  function spacerCell(span){
    return new TableCell({columnSpan:Math.max(span,1),borders:NB,
      children:[new Paragraph({spacing:{after:0},children:[]})]});
  }
  function gapRow(){
    return new TableRow({children:[new TableCell({columnSpan:ncols,borders:NB,
      children:[new Paragraph({spacing:{before:0,after:0},
        children:[new TextRun({text:'',size:10,font:FONT})]})]})]});
  }

  const rows=[];
  levels.forEach((lvl,di)=>{
    if(di>0) rows.push(gapRow());
    const nodes=lvl.slice().sort((a,b)=>a.c0-b.c0);
    const cells=[]; let cursor=0;
    nodes.forEach(n=>{
      if(n.c0>cursor) cells.push(spacerCell(n.c0-cursor));
      cells.push(boxCell(n));
      cursor=n.c1+1;
    });
    if(cursor<ncols) cells.push(spacerCell(ncols-cursor));
    rows.push(new TableRow({children:cells}));
  });

  return new Table({width:{size:TW,type:WidthType.DXA},columnWidths:widths,
                    bidirectional:true,visuallyRightToLeft:true,
                    borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
                             left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
                             insideHorizontal:{style:BorderStyle.NONE},insideVertical:{style:BorderStyle.NONE}},
                    rows});
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

app.get('/api/download/:id/excel', async (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp  = path.join(DATA_DIR, `${sid}.json`);
  if (!fs.existsSync(fp)) return res.status(404).send('الطلب غير موجود');
  try {
    const data = JSON.parse(fs.readFileSync(fp,'utf8'));
    const buf  = await generateExcel(data);
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="project_${sid}.xlsx"`);
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
    const buf  = await generateWord(data);
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="project_${sid}.docx"`);
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
    const updated = { ...existing, ...body, id: sid, updatedAt: new Date().toISOString() };
    
    fs.writeFileSync(fp, JSON.stringify(updated, null, 2), 'utf8');
    
    // Update index
    const index = loadIndex();
    const idx = index.findIndex(s=>s.id===sid);
    if (idx>=0) {
      index[idx] = { ...index[idx], projectIdea: body.projectIdea.substring(0,80), summary: body.summary||{}, updatedAt: updated.updatedAt };
      saveIndex(index);
    }
    
    // Regenerate files
    const [excelBuf, wordBuf] = await Promise.all([generateExcel(body), generateWord(body)]);
    fs.writeFileSync(path.join(DATA_DIR,`${sid}.xlsx`), excelBuf);
    fs.writeFileSync(path.join(DATA_DIR,`${sid}.docx`), wordBuf);
    
    console.log(`✅ تم تحديث الطلب: ${sid}`);
    return res.json({ success:true, id: sid });
  } catch(err) {
    console.error('❌ خطأ في التحديث:', err);
    return res.status(500).json({ success:false, message: err.message });
  }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname,'public','index.html')));

app.listen(PORT, () => console.log(`🚀 يعمل على المنفذ ${PORT}`));
module.exports = app;
