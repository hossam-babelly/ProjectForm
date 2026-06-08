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
//  WORD GENERATOR  (v6 — all fixes)
// ════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════
//  WORD GENERATOR  (v6 — all fixes applied)
// ════════════════════════════════════════════════════════════

const C = {
  DARK_BLUE:'1F3864', GREY_HDR:'AEAAAA', COL_HDR_BLK:'262626',
  COL_HDR_BLU:'D9E2F3', PROD1:'8EA9DB', PROD2:'F4B084', PROD3:'8EAADB',
  MON_ODD:'D9E1F2', MON_EVEN:'B4C6E7', MON_ODD2:'FCE4D6', MON_EVEN2:'F8CBAD',
  SUBTOT:'808080', SUM_LBL:'DEEAF6', SUM_VAL:'FBE4D5',
  ROW_ODD:'FBE4D5', ROW_EVEN:'F2F2F2', WHITE:'FFFFFF',
  ORG_ADM:'4472C4', ORG_EXE:'70AD47', ORG_SVC:'ED7D31', ORG_OTH:'5B9BD5',
};

const PAGE_L = { size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE}, margin:{top:720,right:720,bottom:720,left:720} };
const PAGE_P = { size:{width:11906,height:16838}, margin:{top:720,right:720,bottom:720,left:720} };
const TW = 15398;  // landscape content width DXA

const FONT = 'Sakkal Majalla';

const BDR = (c='auto',s=12) => ({style:BorderStyle.SINGLE,size:s,color:c});
const BORDERS = {top:BDR(),bottom:BDR(),left:BDR(),right:BDR()};

// ── Cell ──────────────────────────────────────────────────────
function C_(text, o={}) {
  const {fill,bold=true,sz=28,color,align=AlignmentType.CENTER,w,colSpan,rowSpan,vAlign=VerticalAlign.CENTER} = o;
  const tc = {};
  if (fill) tc.shading = {type:ShadingType.CLEAR,fill,color:fill};
  if (w)    tc.width   = {size:w,type:WidthType.DXA};
  if (colSpan) tc.columnSpan = colSpan;
  if (rowSpan>1) tc.rowSpan = rowSpan;
  tc.borders = BORDERS;
  tc.margins = {top:60,bottom:60,left:80,right:80};
  tc.verticalAlign = vAlign;

  const textColor = color || (
    fill===C.DARK_BLUE||fill===C.COL_HDR_BLK||fill===C.SUBTOT||fill===C.GREY_HDR
      ? C.WHITE : '000000'
  );

  return new TableCell({...tc, children:[new Paragraph({
    bidirectional:true, alignment:align, spacing:{after:0,line:240,lineRule:'auto'},
    children:[new TextRun({text:String(text??''),bold,size:sz,font:FONT,color:textColor})],
  })]});
}

// ── Table ─────────────────────────────────────────────────────
function T_(rows, W) {
  return new Table({width:{size:TW,type:WidthType.DXA},columnWidths:W,bidirectional:true,visuallyRightToLeft:true,rows});
}

// ── Section header row ────────────────────────────────────────
function secHdr(title, cols, fill=C.DARK_BLUE) {
  return new TableRow({tableHeader:true, children:[
    C_(title,{fill,bold:true,sz:32,colSpan:cols,w:TW,align:AlignmentType.CENTER,
      color:fill===C.GREY_HDR?'000000':C.WHITE})
  ]});
}

// ── Column header row ─────────────────────────────────────────
function colHdr(labels, W, fill=C.COL_HDR_BLK) {
  return new TableRow({tableHeader:true, children:
    labels.map((h,i)=>C_(h,{fill,bold:true,sz:28,w:W[i],
      color:fill===C.COL_HDR_BLU?'000000':C.WHITE}))
  });
}

// ── Total row ─────────────────────────────────────────────────
function totRow(label, value, cols, lastW, fill=C.DARK_BLUE) {
  return new TableRow({children:[
    C_(label,{fill,bold:true,sz:28,colSpan:cols-1,w:TW-lastW,align:AlignmentType.CENTER}),
    C_(value,{fill,bold:true,sz:28,w:lastW}),
  ]});
}

function SP(before=200) { return new Paragraph({spacing:{before,after:0},children:[]}); }

function fN(v) {
  const n = parseFloat(String(v||'0').replace(/[^0-9.-]/g,''));
  if(!n||isNaN(n)) return '0';
  return n.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:2});
}
function fM(v) {
  const n = parseFloat(String(v||'0').replace(/[^0-9.-]/g,''));
  if(isNaN(n)) return '0 $';
  return n.toLocaleString('en-US',{minimumFractionDigits:0,maximumFractionDigits:2})+' $';
}

const PAL = [
  {dark:C.PROD1,mo:C.MON_ODD, me:C.MON_EVEN},
  {dark:C.PROD2,mo:C.MON_ODD2,me:C.MON_EVEN2},
  {dark:C.PROD3,mo:C.MON_ODD, me:C.MON_EVEN},
];
function PC(i){return PAL[i%PAL.length];}

function norm(arr,tw) {
  const s=arr.reduce((a,b)=>a+b,0);
  const r=arr.map(w=>Math.round(w*tw/s));
  r[r.length-1]=tw-r.slice(0,-1).reduce((a,b)=>a+b,0);
  return r;
}


// ════════════════════════════════════════════════════════════════
async function generateWord(data) {
  const pids = Object.keys(data.products||{}).filter(p=>data.products[p]?.name);
  const sections = [];

  function tableSection(tblObj) {
    return { properties:{page:PAGE_L}, children:[SP(2000),tblObj,SP()] };
  }

  // ── Title ────────────────────────────────────────────────
  sections.push({ properties:{page:PAGE_L}, children:[
    new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:3000,after:120},
      children:[new TextRun({text:'قسم المشاريع التنموية',bold:true,size:40,font:FONT,color:C.DARK_BLUE})]}),
    new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{before:0,after:0},
      children:[new TextRun({text:'نموذج طرح دراسة مشروع',bold:true,size:32,font:FONT})]}),
  ]});

  // ══ 1. SUMMARY ══════════════════════════════════════════
  {
    const W = norm([7699,7699],TW);
    sections.push(tableSection(T_([
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
    ],W)));
  }

  // ══ 2. FOUNDING ═════════════════════════════════════════
  if (data.foundingRows?.length) {
    const W = norm([149,269,1368,380,233,1521,1572,1906],TW);
    let tot=0; data.foundingRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    sections.push(tableSection(T_([
      secHdr('التكاليف التأسيسية',8),
      colHdr(['#','الصنف','البيان','الاهتلاك','العدد','ملاحظات','التكلفة للواحدة','التكلفة الإجمالية'],W,C.COL_HDR_BLU),
      ...data.foundingRows.map((r,i)=>new TableRow({children:[
        C_(i+1,         {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[0]}),
        C_(r.cat||'',   {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[1]}),
        C_(r.bayan||'', {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[2]}),
        C_(r.dep?'✓':'',{fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[3],color:r.dep?'70AD47':'000000'}),
        C_(r.qty||'',   {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[4]}),
        C_(r.notes||'', {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[5]}),
        C_(fM(r.price), {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[6]}),
        C_(fM(r.total), {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[7]}),
      ]})),
      // Fix 3: "الإجمالي" centered in last row
      totRow('الإجمالي',fM(tot),8,W[7]),
    ],W)));
  }

  // ══ 3. PRODUCTS ═════════════════════════════════════════
  if (pids.length) {
    const W = norm([4822,4841,5705],TW);
    const prodRows = [
      secHdr('جدول المنتجات',3,C.GREY_HDR),
      // Fix 4: col header row white on black; data rows all black text
      colHdr(['البيان','الواحدة','المكونات'],W,C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi)=>{
      const p=data.products[pid]; const comps=(p.components||[]).filter(c=>c); const pc=PC(pi);
      comps.forEach((comp,ci)=>{
        if(ci===0){
          prodRows.push(new TableRow({children:[
            new TableCell({rowSpan:comps.length,shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS,margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[0],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.name||'',bold:true,size:28,font:FONT,color:'000000'})]})]}),
            new TableCell({rowSpan:comps.length,shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS,margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[1],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.unit||'',bold:true,size:28,font:FONT,color:'000000'})]})]}),
            C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000'}),
          ]}));
        } else {
          prodRows.push(new TableRow({children:[C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000'})]}));
        }
      });
    });
    sections.push(tableSection(T_(prodRows,W)));
  }

  // ══ 4. REVENUE ══════════════════════════════════════════
  if (pids.length) {
    const W = norm([756,383,309,309,309,309,309,309,309,309,309,359,359,362],TW);
    const revRows = [
      // Fix 5: "الإيرادات المتوقعة" black text; last row first 2 cells merged
      secHdr('الإيرادات المتوقعة',14,C.GREY_HDR),
      colHdr(['البيان','الواحدة',...MONTHS],W,C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi)=>{
      const p=data.products[pid]; const rev=data.revenueData?.[pid]||[]; const pc=PC(pi);
      revRows.push(new TableRow({children:[
        C_(p.name,{fill:pc.dark,bold:true,sz:28,w:W[0]}),
        C_(p.unit||'',{fill:pc.dark,bold:true,sz:28,w:W[1]}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.qty),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2]})),
      ]}));
      revRows.push(new TableRow({children:[
        C_('سعر مبيع الواحدة',{fill:pc.dark,sz:28,w:W[0]}),
        C_('$',{fill:pc.dark,sz:28,w:W[1]}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.unitPrice),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2]})),
      ]}));
      revRows.push(new TableRow({children:[
        C_('سعر المبيع الإجمالي',{fill:pc.dark,sz:28,w:W[0]}),
        C_('$',{fill:pc.dark,sz:28,w:W[1]}),
        ...MONTHS.map((_,m)=>C_(fN(rev[m]?.total||0),{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+2]})),
      ]}));
    });
    // Fix 5: merge first two cells of total row
    const mTots = MONTHS.map((_,m)=>{let t=0;pids.forEach(pid=>{t+=parseFloat(String(data.revenueData?.[pid]?.[m]?.total||'0').replace(/[^0-9.-]/g,''))||0;});return fM(t);});
    revRows.push(new TableRow({children:[
      C_('الإجمالي',{fill:C.COL_HDR_BLK,bold:true,sz:28,colSpan:2,w:W[0]+W[1],align:AlignmentType.CENTER}),
      ...mTots.map((v,m)=>C_(v,{fill:C.COL_HDR_BLK,bold:true,sz:28,w:W[m+2]})),
    ]}));
    sections.push(tableSection(T_(revRows,W)));
  }

  // ══ 5. OPS ══════════════════════════════════════════════
  if (pids.length && data.opsData) {
    const W = norm([433,285,395,319,324,324,324,324,342,342,342,342,357,357,350],TW);
    // Fix 6: all black except row2 (months) and last row (totals)
    const opsRows = [
      secHdr('التكاليف التشغيلية',15,C.GREY_HDR),
      colHdr(['البيان','الواحدة','التفاصيل',...MONTHS],W,C.COL_HDR_BLK),
    ];
    pids.forEach((pid,pi)=>{
      const p=data.products[pid]; const comps=(p.components||[]).filter(c=>c);
      const opsD=data.opsData?.[pid]||{}; const pc=PC(pi);
      comps.forEach((comp,ci)=>{
        const vals=MONTHS.map((_,m)=>fN(opsD[`${ci}_${m}`]));
        if(ci===0){
          opsRows.push(new TableRow({children:[
            new TableCell({rowSpan:comps.length,shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS,margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[0],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.name||'',bold:true,size:28,font:FONT,color:'000000'})]})]}),
            new TableCell({rowSpan:comps.length,shading:{type:ShadingType.CLEAR,fill:pc.dark,color:pc.dark},
              borders:BORDERS,margins:{top:60,bottom:60,left:80,right:80},
              width:{size:W[1],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:0},
                children:[new TextRun({text:p.unit||'',bold:true,size:28,font:FONT,color:'000000'})]})]}),
            C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000',align:AlignmentType.CENTER}),
            ...vals.map((v,m)=>C_(v,{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+3],color:'000000'})),
          ]}));
        } else {
          opsRows.push(new TableRow({children:[
            C_(comp,{fill:pc.dark,sz:28,w:W[2],color:'000000',align:AlignmentType.CENTER}),
            ...vals.map((v,m)=>C_(v,{fill:m%2===0?pc.mo:pc.me,sz:28,w:W[m+3],color:'000000'})),
          ]}));
        }
      });
      const subVals=MONTHS.map((_,m)=>fM(opsD[`sub_${m}`]||0));
      opsRows.push(new TableRow({children:[
        C_(`إجمالي ${p.name}`,{fill:C.SUBTOT,bold:true,sz:28,colSpan:3,w:W.slice(0,3).reduce((a,b)=>a+b,0),align:AlignmentType.CENTER}),
        ...subVals.map((v,m)=>C_(v,{fill:C.SUBTOT,bold:true,sz:28,w:W[m+3]})),
      ]}));
    });
    const opsTots=MONTHS.map((_,m)=>{let t=0;pids.forEach(pid=>{t+=parseFloat(String(data.opsData?.[pid]?.[`sub_${m}`]||'0').replace(/[^0-9.-]/g,''))||0;});return fM(t);});
    opsRows.push(new TableRow({children:[
      C_('الإجمالي',{fill:C.COL_HDR_BLK,bold:true,sz:28,colSpan:3,w:W.slice(0,3).reduce((a,b)=>a+b,0),align:AlignmentType.CENTER}),
      ...opsTots.map((v,m)=>C_(v,{fill:C.COL_HDR_BLK,bold:true,sz:28,w:W[m+3]})),
    ]}));
    sections.push(tableSection(T_(opsRows,W)));
  }

  // ══ 6. HR ════════════════════════════════════════════════
  if (data.hrRows?.length) {
    const W=norm([149,907,876,321,1288,2325,2731],TW);
    let tot=0; data.hrRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    sections.push(tableSection(T_([
      secHdr('الموارد البشرية',7),
      colHdr(['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي','الراتب الشهري الإجمالي'],W,C.COL_HDR_BLU),
      ...data.hrRows.map((r,i)=>new TableRow({children:[
        C_(i+1,            {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[0]}),
        C_(r.position||'', {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[1]}),
        C_(r.type||'',     {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[2]}),
        C_(r.qty||'',      {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[3]}),
        C_(r.reports||'—', {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[4]}),
        C_(fM(r.salary),   {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[5]}),
        C_(fM(r.total),    {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[6]}),
      ]})),
      // Fix 7: "الإجمالي" centered
      totRow('الإجمالي',fM(tot),7,W[6]),
    ],W)));
  }

  // ══ 6b. ORG CHART PAGE — title paragraph, SmartArt injected post-process
  if (data.hrRows?.length) {
    sections.push({
      properties:{page:PAGE_P},
      children: [
        new Paragraph({
          bidirectional:true, alignment:AlignmentType.CENTER,
          spacing:{before:800,after:400},
          children:[new TextRun({text:'الهيكل التنظيمي',bold:true,size:36,font:FONT,color:C.DARK_BLUE})],
        }),
        new Paragraph({children:[],spacing:{before:0,after:0}}),
      ],
    });
  }

  // ══ 7. FIXED ════════════════════════════════════════════
  if (data.fixedRows?.length) {
    const W=norm([149,500,1000,300,1400,1400,1400],TW);
    let tot=0; data.fixedRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    sections.push(tableSection(T_([
      secHdr('التكاليف الثابتة',7),
      colHdr(['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة','التكلفة الشهرية الإجمالية'],W,C.COL_HDR_BLU),
      ...data.fixedRows.map((r,i)=>new TableRow({children:[
        C_(i+1,         {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[0]}),
        C_(r.cat||'',   {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[1]}),
        C_(r.bayan||'', {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[2]}),
        C_(r.qty||'',   {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[3]}),
        C_(r.notes||'', {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[4]}),
        C_(fM(r.price), {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[5]}),
        C_(fM(r.total), {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[6]}),
      ]})),
      // Fix 9: "الإجمالي" centered — الإجمالي الشهري
      totRow('الإجمالي',fM(tot),7,W[6]),
    ],W)));
  }

  // ══ 8. DEP ══════════════════════════════════════════════
  if (data.depRows?.length) {
    const W=norm([149,280,1350,450,225,1500,1500,1500],TW);
    let tot=0; data.depRows.forEach(r=>{tot+=parseFloat(String(r.total||'0').replace(/[^0-9.-]/g,''))||0;});
    sections.push(tableSection(T_([
      secHdr('الاهتلاك',8),
      colHdr(['#','الصنف','البيان','نسبة الاهتلاك','العدد','ملاحظات','قيمة الاهتلاك للواحدة','قيمة الاهتلاك الإجمالية'],W,C.COL_HDR_BLU),
      ...data.depRows.map((r,i)=>new TableRow({children:[
        C_(i+1,               {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[0]}),
        C_(r.cat||'',         {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[1]}),
        C_(r.bayan||'',       {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[2]}),
        C_((r.pct||'0')+' %', {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[3]}),
        C_(r.qty||'',         {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[4]}),
        C_(r.notes||'',       {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[5]}),
        C_(fM(r.perUnit),     {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[6]}),
        C_(fM(r.total),       {fill:i%2===0?C.WHITE:C.ROW_EVEN,sz:28,w:W[7]}),
      ]})),
      totRow('إجمالي قيمة الاهتلاك',fM(tot),8,W[7]),
    ],W)));
  }

  // Build the base docx buffer
  const baseBuffer = await Packer.toBuffer(new Document({sections}));

  // Inject SmartArt into the docx zip if we have HR data
  if (data.hrRows?.length) {
    return injectSmartArt(baseBuffer, data.hrRows);
  }
  return baseBuffer;
}

// ── SmartArt injection via JSZip ──────────────────────────────
async function injectSmartArt(docxBuffer, hrRows) {
  const zip = await JSZip.loadAsync(docxBuffer);

  // Build data1.xml with org hierarchy
  const dataXml = buildSmartArtDataXml(hrRows);

  // Layout: Hierarchy / Organisation Chart (standard MS layout URI)
  const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/orgchart1"
  minVer="12.0">
<dgm:title lang="" val=""/>
<dgm:desc lang="" val=""/>
<dgm:catLst><dgm:cat type="hierarchy" pri="10100"/></dgm:catLst>
<dgm:layoutNode name="root"><dgm:varLst><dgm:var name="dir" val="norm"/><dgm:var name="animLvl" val="lvl"/><dgm:var name="animOne" val="one"/><dgm:var name="hierBranch" val="std"/></dgm:varLst></dgm:layoutNode>
</dgm:layoutDef>`;

  const colorsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/colorful3"
  minVer="12.0">
<dgm:catLst><dgm:cat type="mainScheme" pri="10100"/></dgm:catLst>
<dgm:styleLbl name="node0"><dgm:fillClrLst><a:schemeClr val="accent1"/></dgm:fillClrLst><dgm:linClrLst><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></dgm:linClrLst><dgm:effectClrLst><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></dgm:effectClrLst><dgm:txLinClrLst><a:schemeClr val="lt1"/></dgm:txLinClrLst><dgm:txFillClrLst><a:schemeClr val="lt1"/></dgm:txFillClrLst><dgm:txEffectClrLst><a:schemeClr val="lt1"/></dgm:txEffectClrLst></dgm:styleLbl>
</dgm:colorsDef>`;

  const styleXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/qs1"
  minVer="12.0">
<dgm:catLst><dgm:cat type="mainScheme" pri="10100"/></dgm:catLst>
<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>
<dgm:styleLbl name="node0"><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="1"/><a:fillRef idx="2"/><a:effectRef idx="0"/><a:fontRef idx="minor"/></dgm:style></dgm:styleLbl>
</dgm:styleDef>`;

  // Drawing placeholder (dimensions match A4 portrait content width)
  const cx = 8229600; // ~9.1 inches in EMU
  const cy = 3500000; // ~3.9 inches
  const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
  xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/></dsp:spTree>
</dsp:drawing>`;

  // Add diagram files to zip
  const diagramFolder = 'word/diagrams/';
  zip.file(`${diagramFolder}data1.xml`, dataXml);
  zip.file(`${diagramFolder}layout1.xml`, layoutXml);
  zip.file(`${diagramFolder}colors1.xml`, colorsXml);
  zip.file(`${diagramFolder}quickStyle1.xml`, styleXml);
  zip.file(`${diagramFolder}drawing1.xml`, drawingXml);

  // Add relationship file for diagrams
  const diagramRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>
</Relationships>`;

  // Update document.xml.rels to add diagram relationships
  const docRelsPath = 'word/_rels/document.xml.rels';
  let relsXml = await zip.file(docRelsPath).async('string');

  // Find the next available rId number
  const existingIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m=>parseInt(m[1]));
  const maxId = existingIds.length ? Math.max(...existingIds) : 0;
  const dm = `rId${maxId+1}`, lo = `rId${maxId+2}`, qs = `rId${maxId+3}`, cs = `rId${maxId+4}`, dw = `rId${maxId+5}`;

  relsXml = relsXml.replace('</Relationships>',
    `<Relationship Id="${dm}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
<Relationship Id="${lo}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="diagrams/layout1.xml"/>
<Relationship Id="${qs}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/>
<Relationship Id="${cs}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="diagrams/colors1.xml"/>
<Relationship Id="${dw}" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>
</Relationships>`
  );
  zip.file(docRelsPath, relsXml);

  // Find the org chart title paragraph in document.xml and insert SmartArt drawing after it
  let docXml = await zip.file('word/document.xml').async('string');

  const smartArtParagraph = `<w:p>
<w:pPr><w:bidi w:val="1"/><w:jc w:val="center"/><w:rPr><w:rtl/></w:rPr></w:pPr>
<w:r><w:drawing>
<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
<wp:extent cx="${cx}" cy="${cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="9001" name="الهيكل التنظيمي"/>
<wp:cNvGraphicFramePr/>
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:dm="${dm}" r:lo="${lo}" r:qs="${qs}" r:cs="${cs}"/>
</a:graphicData>
</a:graphic>
</wp:inline>
</w:r></w:drawing></w:p>`;

  // Find the "الهيكل التنظيمي" text in the doc and insert SmartArt after its paragraph
  docXml = docXml.replace(
    /(<w:p[^>]*>(?:<[^>]+>)*<w:t[^>]*>\s*الهيكل التنظيمي\s*<\/w:t>.*?<\/w:p>)/s,
    `$1\n${smartArtParagraph}`
  );

  zip.file('word/document.xml', docXml);

  return zip.generateAsync({type:'nodebuffer', compression:'DEFLATE'});
}

function buildSmartArtDataXml(hrRows) {
  // Expand employees by qty
  const nodes = [];
  hrRows.forEach(r => {
    const qty = parseInt(r.qty)||1;
    for(let i=0;i<qty;i++) {
      const idx = nodes.length;
      const pad = idx.toString(16).toUpperCase().padStart(8,'0');
      nodes.push({
        id: `{${pad}-AAAA-BBBB-CCCC-${pad.padStart(12,'0')}}`,
        pos: r.position||'',
        type: r.type||'',
        reports: r.reports||'',
        isAsst: r.type === 'مزود خدمة',
      });
    }
  });

  const DOC_ID = '{DOCID000-0000-0000-0000-FFFFFFFFFFFF}';
  const posToId = {};
  nodes.forEach(n => { if(!posToId[n.pos]) posToId[n.pos] = n.id; });

  let ptXml = `<dgm:pt modelId="${DOC_ID}" type="doc"><dgm:prSet/></dgm:pt>\n`;
  nodes.forEach((n,i) => {
    const typeAttr = n.isAsst ? ' type="asst"' : '';
    const par = `{PAR${i.toString(16).toUpperCase().padStart(8,'0')}}`;
    const sib = `{SIB${i.toString(16).toUpperCase().padStart(8,'0')}}`;
    ptXml += `<dgm:pt modelId="${n.id}"${typeAttr}>`;
    ptXml += `<dgm:prSet lang="ar-SA" phldrT=""/>`;
    ptXml += `<dgm:spPr/>`;
    ptXml += `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">`;
    ptXml += `<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`;
    ptXml += `<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`;
    ptXml += `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${escXml(n.pos)}</a:t></a:r></a:p>`;
    ptXml += `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${escXml(n.type)}</a:t></a:r></a:p>`;
    ptXml += `</p:txBody></dgm:pt>`;
    ptXml += `<dgm:pt modelId="${par}" type="parTrans"><dgm:prSet lang="ar-SA"/></dgm:pt>`;
    ptXml += `<dgm:pt modelId="${sib}" type="sibTrans"><dgm:prSet lang="ar-SA"/></dgm:pt>`;
  });

  let cxnXml = '';
  let ci = 1000;
  nodes.forEach((n,i) => {
    const cid = `{CXN${ci.toString(16).toUpperCase().padStart(8,'0')}}`;
    ci++;
    if (!n.reports) {
      cxnXml += `<dgm:cxn modelId="${cid}" srcId="${DOC_ID}" destId="${n.id}" srcOrd="${i}" destOrd="0"/>`;
    } else {
      const pid = posToId[n.reports];
      if(pid) cxnXml += `<dgm:cxn modelId="${cid}" srcId="${pid}" destId="${n.id}" srcOrd="${i}" destOrd="0"/>`;
    }
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<dgm:ptLst>${ptXml}</dgm:ptLst>
<dgm:cxnLst>${cxnXml}</dgm:cxnLst>
<dgm:bg/>
<dgm:whole/>
</dgm:dataModel>`;
}

function escXml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
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
    const index = loadIndex();
    const seq = index.length + 1;  // تسلسل النموذج
    const submission = { id, seq, submittedAt: body.submittedAt||new Date().toISOString(), ...body };

    // احفظ JSON والفهرس فوراً
    fs.writeFileSync(path.join(DATA_DIR,`${id}.json`), JSON.stringify(submission,null,2), 'utf8');
    index.unshift({ id, seq, projectIdea:body.projectIdea.substring(0,80), applicantName:body.applicantName||'', submittedAt:submission.submittedAt, summary:body.summary||{} });
    saveIndex(index);

    console.log(`✅ طلب جديد: ${id} (تسلسل ${seq})`);

    // ردّ فوري على المتصفح — لا ننتظر توليد الملفات
    res.json({ success:true, id, seq });

    // توليد الملفات والبريد في الخلفية (بعد الرد)
    setImmediate(async () => {
      try {
        const [excelBuf, wordBuf] = await Promise.all([generateExcel(submission), generateWord(submission)]);
        fs.writeFileSync(path.join(DATA_DIR,`${id}.xlsx`), excelBuf);
        fs.writeFileSync(path.join(DATA_DIR,`${id}.docx`), wordBuf);
        await sendEmail(submission, id, excelBuf, wordBuf).catch(e => console.error('Email error:', e.message));
        console.log(`📄 تم توليد ملفات الطلب: ${id}`);
      } catch(e) {
        console.error('❌ خطأ في توليد الملفات (خلفية):', e.message);
      }
    });
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

// بناء اسم ملف بصيغة: اسم المقدم - التاريخ - التسلسل
function buildFileName(data, ext) {
  const name = (data.applicantName || 'مشروع').trim().replace(/[\/\\:*?"<>|]/g, '').replace(/\s+/g, '_');
  const dateObj = new Date(data.submittedAt || Date.now());
  const date = `${dateObj.getFullYear()}-${String(dateObj.getMonth()+1).padStart(2,'0')}-${String(dateObj.getDate()).padStart(2,'0')}`;
  const seq = data.seq ? String(data.seq).padStart(4,'0') : (data.id || '');
  return `${name} - ${date} - ${seq}.${ext}`;
}

// تعيين رأس Content-Disposition يدعم العربية (UTF-8)
function setDownloadName(res, filename) {
  const encoded = encodeURIComponent(filename);
  res.setHeader('Content-Disposition', `attachment; filename="${encoded}"; filename*=UTF-8''${encoded}`);
}

app.get('/api/download/:id/excel', async (req, res) => {
  const sid = req.params.id.replace(/[^A-Z0-9\-]/g,'');
  const fp  = path.join(DATA_DIR, `${sid}.json`);
  if (!fs.existsSync(fp)) return res.status(404).send('الطلب غير موجود');
  try {
    const data = JSON.parse(fs.readFileSync(fp,'utf8'));
    const buf  = await generateExcel(data);
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    setDownloadName(res, buildFileName(data, 'xlsx'));
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
    setDownloadName(res, buildFileName(data, 'docx'));
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
