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
  wb.creator = 'نموذج دراسة مشروع'; wb.created = new Date();

  const F='Sakkal Majalla';
  const NAVY='FF1F3864', ORANGE='FFED7D31', BLUE='FF2F5496', LIGHT='FFF1F5F9',
        SUMLBL='FFDEEAF6', SUMVAL='FFFBE4D5', GREENc='FF538135', GRAYc='FF94A3B8',
        WHITE='FFFFFFFF', COLHDR='FFD9E2F3', STRIPE='FFEFF3FA';
  const INPUT='FF0000FF', FORMULA='FF000000', LINK='FF008000';
  const PRODCLR=['FFD9E2F3','FFFCE4D6','FFE2EFDA','FFFFF2CC','FFDBE5F1'];
  const MONEY='"$"#,##0.00;[Red]("$"#,##0.00);"-"', MONEY0='"$"#,##0;[Red]("$"#,##0);"-"', PCT='0.0%';
  const pm = s => parseFloat(String(s==null?'0':s).replace(/[^0-9.\-]/g,''))||0;
  const thin=()=>({top:{style:'thin',color:{argb:'FFBFBFBF'}},bottom:{style:'thin',color:{argb:'FFBFBFBF'}},left:{style:'thin',color:{argb:'FFBFBFBF'}},right:{style:'thin',color:{argb:'FFBFBFBF'}}});
  const fill=c=>({type:'pattern',pattern:'solid',fgColor:{argb:c}});
  const colL=n=>{let s='';while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=(n-m-1)/26;}return s;};
  const Q=name=>`'${name}'`;

  function band(ws,lastCol,text,sub){
    ws.mergeCells(1,1,1,lastCol);
    const t=ws.getCell(1,1); t.value=text; t.font={bold:true,size:18,color:{argb:WHITE},name:F};
    t.fill=fill(NAVY); t.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; ws.getRow(1).height=34;
    if(sub!==undefined){ ws.mergeCells(2,1,2,lastCol); const s=ws.getCell(2,1); s.value=sub;
      s.font={bold:true,size:12,color:{argb:NAVY},name:F}; s.fill=fill('FFF4D9B0');
      s.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; ws.getRow(2).height=20; }
  }
  function colHeaders(ws,rowIdx,headers,widths){
    if(widths) ws.columns=widths.map(w=>({width:w}));
    const hr=ws.getRow(rowIdx); hr.height=30;
    headers.forEach((h,i)=>{ const c=hr.getCell(i+1); c.value=h;
      c.font={bold:true,color:{argb:WHITE},name:F,size:12}; c.fill=fill(NAVY);
      c.alignment={horizontal:'center',vertical:'middle',wrapText:true,readingOrder:2}; c.border=thin(); });
  }
  function dcell(ws,r,c,val,o){ o=o||{}; const cell=ws.getCell(r,c); cell.value=val;
    const isNum=(typeof val==='number')||(val&&typeof val==='object'&&'formula' in val);
    cell.font={name:F,size:12,bold:o.bold||isNum,color:{argb:o.color||FORMULA}};
    cell.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true};
    cell.border=thin(); if(o.numFmt)cell.numFmt=o.numFmt; if(o.bg)cell.fill=fill(o.bg); return cell; }
  function totalRow(ws,rowIdx,fromC,toC,label,valCol,value,numFmt){ numFmt=numFmt||MONEY;
    ws.mergeCells(rowIdx,fromC,rowIdx,toC);
    const l=ws.getCell(rowIdx,fromC); l.value=label; l.font={bold:true,name:F,size:13,color:{argb:NAVY}};
    l.fill=fill(SUMLBL); l.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; l.border=thin();
    const v=ws.getCell(rowIdx,valCol); v.value=value; v.numFmt=numFmt; v.font={bold:true,name:F,size:13,color:{argb:NAVY}};
    v.fill=fill(SUMVAL); v.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; v.border=thin();
    ws.getRow(rowIdx).height=24; }

  const SN={app:'معلومات المقدّم',sum:'الخلاصة المالية',found:'التكاليف التأسيسية',prod:'المنتجات',
    mkt:'التسويق والمبيع',rev:'الإيرادات',ops:'التكاليف التشغيلية',hr:'الموارد البشرية',fixed:'التكاليف الثابتة',dep:'الاهتلاك'};
  const wsApp=wb.addWorksheet(SN.app,{views:[{rightToLeft:true}]});
  const wsSum=wb.addWorksheet(SN.sum,{views:[{rightToLeft:true}]});
  const wsFound=wb.addWorksheet(SN.found,{views:[{rightToLeft:true}]});
  const wsProd=wb.addWorksheet(SN.prod,{views:[{rightToLeft:true}]});
  const wsMkt=wb.addWorksheet(SN.mkt,{views:[{rightToLeft:true}]});
  const wsRev=wb.addWorksheet(SN.rev,{views:[{rightToLeft:true}]});
  const wsOps=wb.addWorksheet(SN.ops,{views:[{rightToLeft:true}]});
  const wsHR=wb.addWorksheet(SN.hr,{views:[{rightToLeft:true}]});
  const wsFixed=wb.addWorksheet(SN.fixed,{views:[{rightToLeft:true}]});
  const wsDep=wb.addWorksheet(SN.dep,{views:[{rightToLeft:true}]});

  const refs={}; const revQtyRow={};

  // FOUNDING
  { const ws=wsFound; band(ws,8,'التكاليف التأسيسية');
    colHeaders(ws,3,['#','الصنف','البيان','اهتلاك','ملاحظات','التكلفة للواحدة ($)','العدد','التكلفة الإجمالية ($)'],[6,16,24,9,22,18,9,20]);
    const rows=data.foundingRows||[];
    rows.forEach((r,i)=>{ const R=4+i, sbg=i%2?STRIPE:undefined;
      dcell(ws,R,1,i+1,{bg:sbg}); dcell(ws,R,2,r.cat||'',{bg:sbg}); dcell(ws,R,3,r.bayan||'',{bg:sbg});
      dcell(ws,R,4,r.dep?'✓':'',{bg:sbg}); dcell(ws,R,5,r.notes||'',{bg:sbg});
      dcell(ws,R,6,pm(r.price),{numFmt:MONEY,color:INPUT,bg:sbg}); dcell(ws,R,7,pm(r.qty),{color:INPUT,bg:sbg});
      dcell(ws,R,8,{formula:`F${R}*G${R}`},{numFmt:MONEY,bg:sbg}); });
    const tr=4+rows.length; totalRow(ws,tr,1,7,'الإجمالي',8,rows.length?{formula:`SUM(H4:H${tr-1})`}:0);
    refs.founding=`${Q(SN.found)}!$H$${tr}`; }

  // HR
  { const ws=wsHR; band(ws,8,'الموارد البشرية');
    colHeaders(ws,3,['#','المنصب','النوع','تابع لـ','الراتب الشهري الفردي ($)','أشهر الدوام/السنة','العدد','الراتب الشهري الإجمالي ($)'],[6,20,16,18,18,15,9,20]);
    const rows=data.hrRows||[];
    rows.forEach((r,i)=>{ const R=4+i, sbg=i%2?STRIPE:undefined;
      dcell(ws,R,1,i+1,{bg:sbg}); dcell(ws,R,2,r.position||'',{bg:sbg}); dcell(ws,R,3,r.type||'',{bg:sbg}); dcell(ws,R,4,r.reports||'',{bg:sbg});
      dcell(ws,R,5,pm(r.salary),{numFmt:MONEY,color:INPUT,bg:sbg}); dcell(ws,R,6,pm(r.months)||12,{color:INPUT,bg:sbg});
      dcell(ws,R,7,pm(r.qty),{color:INPUT,bg:sbg}); dcell(ws,R,8,{formula:`E${R}*G${R}`},{numFmt:MONEY,bg:sbg}); });
    const tr=4+rows.length;
    totalRow(ws,tr,1,7,'إجمالي الرواتب (سنوياً)',8, rows.length?{formula:`SUMPRODUCT(E4:E${tr-1},G4:G${tr-1},F4:F${tr-1})`}:0);
    refs.hrAnnual=`${Q(SN.hr)}!$H$${tr}`;
    const ocR=tr+2; ws.mergeCells(ocR,1,ocR,8); const oc=ws.getCell(ocR,1);
    oc.value='الهيكل التنظيمي'; oc.font={bold:true,color:{argb:WHITE},name:F,size:14}; oc.fill=fill(NAVY);
    oc.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; ws.getRow(ocR).height=30; }

  // MARKETING
  { const ws=wsMkt; band(ws,4,'التسويق والمبيع'); ws.columns=[{width:24},{width:28},{width:28},{width:28}];
    let R=3;
    const lv=(label,val,o)=>{ o=o||{}; ws.mergeCells(R,2,R,4);
      const l=ws.getCell(R,1); l.value=label; l.font={bold:true,name:F,size:12,color:{argb:NAVY}}; l.fill=fill(SUMLBL); l.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; l.border=thin();
      const v=ws.getCell(R,2); v.value=val; const isNum=(typeof val==='number')||(val&&typeof val==='object'&&'formula' in val);
      v.font={name:F,size:12,color:{argb:o.color||FORMULA},bold:o.bold||isNum}; v.fill=fill(o.bg||SUMVAL);
      v.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; v.border=thin(); if(o.numFmt)v.numFmt=o.numFmt;
      ws.getRow(R).height=o.h||24; R++; };
    lv('خطة التسويق', data.marketingPlan||'—',{h:40});
    const mktRow=R; lv('كلفة التسويق الشهرية ($)', pm(data.marketingCost), {numFmt:MONEY,color:INPUT});
    lv('كلفة التسويق السنوية ($)', {formula:`B${mktRow}*12`}, {numFmt:MONEY});
    lv('قنوات المبيع', data.salesChannels||'—',{h:36});
    refs.mktMonthly=`${Q(SN.mkt)}!$B$${mktRow}`; R++;
    const comps=(data.competitors||[]).filter(c=>c&&(c.name||c.strengths||c.weaknesses));
    if(comps.length){ ws.mergeCells(R,1,R,4); const h=ws.getCell(R,1); h.value='جدول المنافسين'; h.font={bold:true,color:{argb:WHITE},name:F,size:13}; h.fill=fill(NAVY); h.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; ws.getRow(R).height=26; R++;
      ['#','اسم المنافس','نقاط القوة','نقاط الضعف'].forEach((t,i)=>{ const c=ws.getCell(R,i+1); c.value=t; c.font={bold:true,color:{argb:WHITE},name:F,size:12}; c.fill=fill(BLUE); c.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; c.border=thin(); }); R++;
      comps.forEach((cp,i)=>{ const sbg=i%2?STRIPE:undefined; dcell(ws,R,1,i+1,{bg:sbg}); dcell(ws,R,2,cp.name||'',{bg:sbg}); dcell(ws,R,3,cp.strengths||'',{bg:sbg}); dcell(ws,R,4,cp.weaknesses||'',{bg:sbg}); R++; }); } }

  // PRODUCTS (Word-style)
  const pids=Object.keys(data.products||{}).filter(p=>data.products[p]?.name);
  { const ws=wsProd; band(ws,3,'المنتجات'); colHeaders(ws,3,['البيان','الواحدة','المكوّن'],[28,14,52]);
    let R=4;
    pids.forEach((pid,pi)=>{ const p=data.products[pid]; const comps=(p.components||[]).filter(c=>c);
      const clr=PRODCLR[pi%PRODCLR.length]; const n=Math.max(comps.length,1); const start=R;
      (comps.length?comps:['—']).forEach((comp)=>{ dcell(ws,R,3,comp,{bg:clr,color:NAVY}); R++; });
      ws.mergeCells(start,1,start+n-1,1); ws.mergeCells(start,2,start+n-1,2);
      const nameC=ws.getCell(start,1); nameC.value=p.name||''; nameC.font={bold:true,name:F,size:12,color:{argb:NAVY}}; nameC.fill=fill(clr); nameC.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; nameC.border=thin();
      const unitC=ws.getCell(start,2); unitC.value=p.unit||''; unitC.font={bold:true,name:F,size:12,color:{argb:NAVY}}; unitC.fill=fill(clr); unitC.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; unitC.border=thin(); }); }

  // REVENUE
  { const ws=wsRev; band(ws,15,'الإيرادات المتوقعة');
    colHeaders(ws,3,['البيان','الواحدة',...MONTHS,'الإجمالي السنوي ($)'],[26,9,...MONTHS.map(()=>11),16]);
    let R=4; const prodAnnual=[];
    pids.forEach((pid,pi)=>{ const p=data.products[pid]; const rev=data.revenueData?.[pid]||[];
      const qR=R,pR=R+1,tR=R+2; revQtyRow[pid]=qR; const clr=PRODCLR[pi%PRODCLR.length];
      dcell(ws,qR,1,(p.name||'')+' — الكمية',{bg:clr}); dcell(ws,qR,2,p.unit||'',{bg:clr});
      dcell(ws,pR,1,(p.name||'')+' — سعر الوحدة',{bg:LIGHT}); dcell(ws,pR,2,'$',{bg:LIGHT});
      dcell(ws,tR,1,(p.name||'')+' — الإجمالي',{bg:SUMLBL,bold:true}); dcell(ws,tR,2,'',{bg:SUMLBL});
      for(let m=0;m<12;m++){ const col=3+m,Lc=colL(col);
        dcell(ws,qR,col,(rev[m]&&pm(rev[m].qty))||0,{color:INPUT,bg:clr});
        dcell(ws,pR,col,(rev[m]&&pm(rev[m].unitPrice))||0,{numFmt:MONEY,color:INPUT,bg:LIGHT});
        dcell(ws,tR,col,{formula:`${Lc}${qR}*${Lc}${pR}`},{numFmt:MONEY,bg:SUMLBL}); }
      dcell(ws,qR,15,'',{bg:clr}); dcell(ws,pR,15,'',{bg:LIGHT});
      dcell(ws,tR,15,{formula:`SUM(C${tR}:N${tR})`},{numFmt:MONEY,bg:SUMVAL,bold:true});
      prodAnnual.push(`O${tR}`); R+=3; });
    const tr=R; totalRow(ws,tr,1,14,'إجمالي الإيرادات السنوية',15, prodAnnual.length?{formula:prodAnnual.join('+')}:0);
    refs.revAnnual=`${Q(SN.rev)}!$O$${tr}`; }

  // OPERATING COSTS (unit cost × revenue qty)
  { const ws=wsOps; band(ws,15,'التكاليف التشغيلية');
    colHeaders(ws,3,['المنتج / المكوّن','',...MONTHS,'الإجمالي السنوي ($)'],[26,9,...MONTHS.map(()=>11),16]);
    let R=4; const prodAnnual=[];
    pids.forEach((pid,pi)=>{ const p=data.products[pid]; const comps=(p.components||[]).filter(c=>c);
      const od=data.opsData?.[pid]||{}; const clr=PRODCLR[pi%PRODCLR.length];
      dcell(ws,R,1,p.name||'',{bg:clr,bold:true}); dcell(ws,R,2,'',{bg:clr}); for(let m=0;m<13;m++) dcell(ws,R,3+m,'',{bg:clr}); R++;
      const cStart=R;
      comps.forEach((comp,ci)=>{ const sbg=ci%2?STRIPE:undefined;
        dcell(ws,R,1,comp+' (كلفة الوحدة)',{bg:sbg}); dcell(ws,R,2,'$',{bg:sbg});
        for(let m=0;m<12;m++){ const col=3+m; dcell(ws,R,col,pm(od[`${ci}_${m}`])||0,{numFmt:MONEY,color:INPUT,bg:sbg}); }
        dcell(ws,R,15,'',{bg:sbg}); R++; });
      const cEnd=R-1;
      if(comps.length){ dcell(ws,R,1,'إجمالي شهري (الكلفة × الكمية)',{bg:SUMLBL,bold:true}); dcell(ws,R,2,'',{bg:SUMLBL});
        for(let m=0;m<12;m++){ const col=3+m,Lc=colL(col);
          dcell(ws,R,col,{formula:`SUM(${Lc}${cStart}:${Lc}${cEnd})*${Q(SN.rev)}!${Lc}${revQtyRow[pid]}`},{numFmt:MONEY,bg:SUMLBL}); }
        dcell(ws,R,15,{formula:`SUM(C${R}:N${R})`},{numFmt:MONEY,bg:SUMVAL,bold:true}); prodAnnual.push(`O${R}`); R++; } });
    const tr=R; totalRow(ws,tr,1,14,'إجمالي التكاليف التشغيلية السنوية',15, prodAnnual.length?{formula:prodAnnual.join('+')}:0);
    refs.opsAnnual=`${Q(SN.ops)}!$O$${tr}`; }

  // FIXED
  { const ws=wsFixed; band(ws,7,'التكاليف الثابتة');
    colHeaders(ws,3,['#','الصنف','البيان','ملاحظات','التكلفة الشهرية للواحدة ($)','العدد','التكلفة الشهرية الإجمالية ($)'],[6,16,24,22,20,9,22]);
    const isSalary=r=>(r.cat==='رواتب')&&(String(r.bayan||'').includes('الموظفين')||String(r.notes||'').includes('تلقائي'));
    const isMkt=r=>(r.cat==='تسويق')||(String(r.bayan||'').includes('التسويق')&&String(r.notes||'').includes('تلقائي'));
    const rows=data.fixedRows||[];
    rows.forEach((r,i)=>{ const R=4+i, auto=isSalary(r)||isMkt(r), sbg=auto?LIGHT:(i%2?STRIPE:undefined);
      dcell(ws,R,1,i+1,{bg:sbg}); dcell(ws,R,2,r.cat||'',{bg:sbg}); dcell(ws,R,3,r.bayan||'',{bg:sbg}); dcell(ws,R,4,r.notes||'',{bg:sbg});
      if(isSalary(r)){ dcell(ws,R,5,{formula:`${refs.hrAnnual}/12`},{numFmt:MONEY,color:LINK,bg:sbg}); dcell(ws,R,6,1,{bg:sbg}); }
      else if(isMkt(r)){ dcell(ws,R,5,{formula:`${refs.mktMonthly}`},{numFmt:MONEY,color:LINK,bg:sbg}); dcell(ws,R,6,1,{bg:sbg}); }
      else { dcell(ws,R,5,pm(r.price),{numFmt:MONEY,color:INPUT,bg:sbg}); dcell(ws,R,6,pm(r.qty),{color:INPUT,bg:sbg}); }
      dcell(ws,R,7,{formula:`E${R}*F${R}`},{numFmt:MONEY,bg:sbg}); });
    const tr=4+rows.length; totalRow(ws,tr,1,6,'الإجمالي الشهري',7,rows.length?{formula:`SUM(G4:G${tr-1})`}:0);
    const ar=tr+1; totalRow(ws,ar,1,6,'الإجمالي السنوي',7,{formula:`G${tr}*12`});
    refs.fixedMonthly=`${Q(SN.fixed)}!$G$${tr}`; refs.fixedAnnual=`${Q(SN.fixed)}!$G$${ar}`; }

  // DEPRECIATION
  { const ws=wsDep; band(ws,8,'الاهتلاك');
    colHeaders(ws,3,['#','الصنف','البيان','نسبة الاهتلاك (%)','ملاحظات','قيمة الاهتلاك للواحدة ($)','العدد','قيمة الاهتلاك الإجمالية ($)'],[6,16,24,16,20,20,9,22]);
    const rows=data.depRows||[];
    rows.forEach((r,i)=>{ const R=4+i, sbg=i%2?STRIPE:undefined;
      dcell(ws,R,1,i+1,{bg:sbg}); dcell(ws,R,2,r.cat||'',{bg:sbg}); dcell(ws,R,3,r.bayan||'',{bg:sbg});
      dcell(ws,R,4,pm(r.pct)/100,{numFmt:PCT,color:INPUT,bg:sbg}); dcell(ws,R,5,r.notes||'',{bg:sbg});
      const base=pm(r.perUnit)&&pm(r.pct)?pm(r.perUnit)/(pm(r.pct)/100):pm(r.price);
      dcell(ws,R,6,{formula:`(${base||0})*D${R}`},{numFmt:MONEY,bg:sbg}); dcell(ws,R,7,pm(r.qty),{color:INPUT,bg:sbg});
      dcell(ws,R,8,{formula:`F${R}*G${R}`},{numFmt:MONEY,bg:sbg}); });
    const tr=4+rows.length; totalRow(ws,tr,1,7,'إجمالي قيمة الاهتلاك',8,rows.length?{formula:`SUM(H4:H${tr-1})`}:0);
    refs.depTotal=`${Q(SN.dep)}!$H$${tr}`; }

  // FINANCIAL SUMMARY
  { const ws=wsSum; ws.columns=[{width:34},{width:24},{width:30},{width:3},{width:10},{width:10},{width:8},{width:8}];
    band(ws,3,'الخلاصة المالية','ملخّص ديناميكي — يتحدّث تلقائياً عند تعديل أي مُدخل');
    let R=3;
    const kv=(label,formula,o)=>{ o=o||{}; const numFmt=o.numFmt||MONEY0, color=o.color||LINK, big=o.big;
      const l=ws.getCell(R,1); l.value=label; l.font={bold:true,name:F,size:big?13:12,color:{argb:NAVY}}; l.fill=fill(SUMLBL); l.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; l.border=thin();
      ws.mergeCells(R,2,R,3); const v=ws.getCell(R,2); v.value=formula; if(numFmt)v.numFmt=numFmt;
      v.font={bold:true,name:F,size:big?18:13,color:{argb:big?ORANGE:color}}; v.fill=fill(big?NAVY:SUMVAL);
      v.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; v.border=thin(); ws.getRow(R).height=big?32:24;
      const rr=R; R++; return rr; };
    kv('الإيرادات السنوية المتوقعة',{formula:`=${refs.revAnnual}`});
    kv('إجمالي التكاليف التشغيلية السنوية',{formula:`=${refs.opsAnnual}`});
    kv('التكاليف الثابتة السنوية',{formula:`=${refs.fixedAnnual}`});
    kv('إجمالي التكاليف التأسيسية',{formula:`=${refs.founding}`});
    kv('الاهتلاك السنوي',{formula:`=${refs.depTotal}`});
    const rev=refs.revAnnual, ops=refs.opsAnnual, fix=refs.fixedAnnual, fnd=refs.founding, dep=refs.depTotal;
    const netRow=kv('صافي الربح السنوي',{formula:`${rev}-${ops}-${fix}-${dep}`},{numFmt:MONEY0,big:true});
    const netRef=`B${netRow}`;
    const payRow=R;
    const l=ws.getCell(payRow,1); l.value='فترة استرداد رأس المال'; l.font={bold:true,name:F,size:12,color:{argb:NAVY}}; l.fill=fill(SUMLBL); l.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; l.border=thin();
    ws.getCell(payRow,5).value={formula:`IF(${netRef}>0,${fnd}/${netRef},-1)`};
    ws.getCell(payRow,6).value={formula:`IF(E${payRow}>0,CEILING(E${payRow}*12,1),0)`};
    ws.getCell(payRow,7).value={formula:`INT(F${payRow}/12)`};
    ws.getCell(payRow,8).value={formula:`MOD(F${payRow},12)`};
    const y=`G${payRow}`, mo=`H${payRow}`;
    const yw=`IF(${y}=0,"",${y}&IF(${y}=1," سنة",IF(${y}<=10," سنوات"," سنة")))`;
    const mw=`IF(${mo}=0,"",${mo}&IF(${mo}=1," شهر",IF(${mo}<=10," أشهر"," شهر")))`;
    const phrase=`IF(E${payRow}<=0,"—",IF(AND(${y}=0,${mo}=0),"أقل من شهر",TRIM(${yw}&IF(AND(${y}>0,${mo}>0)," و ","")&${mw})))`;
    ws.mergeCells(payRow,2,payRow,3); const pv=ws.getCell(payRow,2); pv.value={formula:phrase};
    pv.font={bold:true,name:F,size:13,color:{argb:FORMULA}}; pv.fill=fill(SUMVAL); pv.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; pv.border=thin(); ws.getRow(payRow).height=24; R++;
    [5,6,7,8].forEach(c=>ws.getColumn(c).hidden=true);
    const marRow=kv('هامش الربح الصافي',{formula:`IF(${rev}>0,${netRef}/${rev},0)`},{numFmt:PCT,color:FORMULA});
    const roiRow=kv('العائد على الاستثمار',{formula:`IF(${fnd}>0,${netRef}/${fnd},0)`},{numFmt:PCT,color:FORMULA});
    [marRow,roiRow].forEach(rr=>{ try{ ws.addConditionalFormatting({ref:`B${rr}`,rules:[{type:'iconSet',iconSet:'5Quarters',reverse:false,showValue:true,cfvo:[{type:'num',value:0},{type:'num',value:0.2},{type:'num',value:0.4},{type:'num',value:0.6},{type:'num',value:0.8}]}]}); }catch(e){} });
    R++;
    ws.mergeCells(R,1,R,3); const bh=ws.getCell(R,1); bh.value='مقارنة الإيرادات والتكاليف والربح'; bh.font={bold:true,color:{argb:WHITE},name:F,size:13}; bh.fill=fill(NAVY); bh.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; ws.getRow(R).height=26; R++;
    const barStart=R;
    const barRows=[['الإيرادات',`${rev}`,BLUE],['التكاليف (تشغيلية+ثابتة+اهتلاك)',`${ops}+${fix}+${dep}`,GRAYc],['صافي الربح',`${netRef}`,GREENc]];
    barRows.forEach(([lab,fm,clr])=>{ const l2=ws.getCell(R,1); l2.value=lab; l2.font={bold:true,name:F,size:12,color:{argb:NAVY}}; l2.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; l2.border=thin();
      ws.mergeCells(R,2,R,3); const v=ws.getCell(R,2); v.value={formula:fm}; v.numFmt=MONEY0; v.font={name:F,size:12,bold:true,color:{argb:NAVY}}; v.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; v.border=thin(); ws.getRow(R).height=26;
      try{ ws.addConditionalFormatting({ref:`B${R}`,rules:[{type:'dataBar',gradient:false,border:false,showValue:true,color:{argb:clr},cfvo:[{type:'num',value:0},{type:'formula',value:`$B$${barStart}`}]}]}); }catch(e){}
      R++; });
    R++; ws.mergeCells(R,1,R,3); const note=ws.getCell(R,1);
    note.value='الخلايا الزرقاء مُدخلات قابلة للتعديل · الخضراء روابط بين الأوراق · السوداء معادلات محسوبة';
    note.font={italic:true,name:F,size:10,color:{argb:'FF808080'}}; note.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; }

  // APPLICANT
  { const ws=wsApp; ws.columns=[{width:30},{width:48}]; band(ws,2,'معلومات مقدّم المشروع');
    let R=3;
    const lv=(label,val)=>{ const l=ws.getCell(R,1); l.value=label; l.font={bold:true,name:F,size:12,color:{argb:NAVY}}; l.fill=fill(SUMLBL); l.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; l.border=thin();
      const v=ws.getCell(R,2); v.value=val||''; v.font={name:F,size:12,color:{argb:FORMULA}}; v.fill=fill(SUMVAL); v.alignment={horizontal:'center',vertical:'middle',readingOrder:2,wrapText:true}; v.border=thin(); ws.getRow(R).height=24; R++; };
    lv('اسم مقدم المشروع',data.applicantName); lv('رقم الهاتف',data.applicantPhone); lv('تاريخ الميلاد',data.applicantBirth);
    lv('مكان الإقامة',data.applicantResidence); lv('العمل الحالي',data.applicantJob); lv('عدد سنوات الخبرة',data.applicantExpYears);
    lv('الخبرة المرتبطة بالمشروع',data.applicantExperience);
    R++; ws.mergeCells(R,1,R,2); const h=ws.getCell(R,1); h.value='تفاصيل المشروع'; h.font={bold:true,color:{argb:WHITE},name:F,size:13}; h.fill=fill(NAVY); h.alignment={horizontal:'center',vertical:'middle',readingOrder:2}; ws.getRow(R).height=26; R++;
    lv('عنوان المشروع',data.projectTitle); lv('فكرة المشروع',data.projectIdea); lv('الشرائح المستهدفة',data.targetSegments);
    lv('نوع المكان المقترح',data.placeType); lv('وضع المكان',data.placeStatus); lv('عنوان / موقع المشروع',data.projectAddress); }

  [wsFound,wsHR,wsRev,wsOps,wsFixed,wsDep].forEach(ws=>{ ws.views=[{rightToLeft:true,state:'frozen',ySplit:3}]; });
  const xbuf = await wb.xlsx.writeBuffer();
  return injectExcelOrgChart(xbuf, data.hrRows);
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
// label/value table in the same visual identity as the summary table
function lblValTable(title, rows){
  const W = norm(COLW.summary, TW);
  return T_([
    new TableRow({children:[C_(title,{fill:C.DARK_BLUE,bold:true,sz:32,colSpan:2,w:TW,align:AlignmentType.CENTER,color:C.WHITE})]}),
    ...rows.map(([l,v])=>new TableRow({children:[
      C_(l,{fill:C.SUM_LBL,bold:true,sz:28,w:W[0],align:AlignmentType.CENTER}),
      C_(String(v==null?'':v),{fill:C.SUM_VAL,bold:true,sz:28,w:W[1],align:AlignmentType.CENTER}),
    ]})),
  ], W);
}

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

  // ══ PAGE 2: الخلاصة المالية (visual dashboard, injected) ═══
  sections.push(mkSec([
    new Paragraph({bidirectional:true, alignment:AlignmentType.CENTER, spacing:{before:120,after:0},
      children:[new TextRun({text:'الخلاصة المالية', bold:true, size:36, font:FONT, color:C.DARK_BLUE})]}),
    new Paragraph({bidirectional:true, alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
      children:[new TextRun({text:'[[FINANCE]]', size:2, font:FONT, color:'FFFFFF'})]}),
  ], VerticalAlign.TOP));

  // ══ PAGE 3: معلومات مقدّم المشروع ═══════════════════════
  page(lblValTable('معلومات مقدّم المشروع', [
    ['اسم مقدم المشروع', data.applicantName||''],
    ['رقم الهاتف', data.applicantPhone||''],
    ['تاريخ الميلاد', data.applicantBirth||''],
    ['مكان الإقامة', data.applicantResidence||''],
    ['العمل الحالي', data.applicantJob||''],
    ['عدد سنوات الخبرة', data.applicantExpYears||''],
    ['الخبرة المرتبطة بالمشروع', data.applicantExperience||''],
  ]));

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

  // ══ PAGE 6: التسويق والمبيع ═════════════════════════════
  {
    const pm = s => parseFloat(String(s||'0').replace(/[^0-9.-]/g,''))||0;
    const mkM = pm(data.marketingCost), mkA = mkM*12;
    const mkTable = lblValTable('التسويق والمبيع', [
      ['خطة التسويق', data.marketingPlan||'—'],
      ['كلفة التسويق الشهرية', fM(mkM)],
      ['كلفة التسويق السنوية', fM(mkA)],
      ['قنوات المبيع', data.salesChannels||'—'],
    ]);
    const comps = (data.competitors||[]).filter(c=>c && (c.name||c.strengths||c.weaknesses));
    const kids = [mkTable];
    if (comps.length) {
      const Wc = norm([900,4200,5144,5144], TW);
      kids.push(SP(220,20), T_([
        secHdr('جدول المنافسين',4),
        colHdr(['#','اسم المنافس','نقاط القوة','نقاط الضعف'],Wc,C.COL_HDR_BLU),
        ...comps.map((c,i)=>new TableRow({children:[
          C_(i+1,            {fill:C.NUM_TINT,sz:28,w:Wc[0]}),
          C_(c.name||'',     {fill:dFill(i),sz:28,w:Wc[1]}),
          C_(c.strengths||'',{fill:dFill(i),sz:28,w:Wc[2]}),
          C_(c.weaknesses||'',{fill:dFill(i),sz:28,w:Wc[3]}),
        ]})),
      ],Wc));
    }
    sections.push(mkSec(kids, VerticalAlign.CENTER));
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
  const withFinance = await injectFinance(withChart, data.summary);
  return finalizeDocx(withFinance);
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
function paybackPhrase(years){
  if(!isFinite(years) || years<=0) return '—';
  const total=Math.ceil(years*12); const y=Math.floor(total/12), m=total%12;
  const yw = y===0?'':(y+(y===1?' سنة':(y<=10?' سنوات':' سنة')));
  const mw = m===0?'':(m+(m===1?' شهر':(m<=10?' أشهر':' شهر')));
  if(!yw && !mw) return 'أقل من شهر';
  return (yw&&mw)?(yw+' و'+mw):(yw||mw);
}
// ── Excel org chart: real drawing shapes injected into the HR sheet ──
// Each shape is anchored to the actual column under its position (screen-from-right),
// avoiding the RTL offset-clamping that piles shapes at the right edge in Excel.
function buildExcelOrgDrawingXml(hrRows, top0){
  const L=orgLayout(hrRows); if(!L) return null;
  const targetW=7000000;
  const scale=Math.min(1, targetW/L.W);
  const sx=v=>Math.round(v*scale);
  const bw=sx(L.boxW), bh=sx(L.boxH);
  const colChars=[6,20,16,18,18,15,9,20];
  const wEmu=i=>{ const ch=i<colChars.length?colChars[i]:8.43; return Math.round((ch*7+5)*9525); };
  function sfr(emu){ let c=0, acc=0; while(c<400){ const w=wEmu(c); if(acc+w>emu) return {col:c, off:Math.max(0,Math.round(emu-acc))}; acc+=w; c++; } return {col:c, off:0}; }
  const RH=190500;   // default row height (15pt) in EMU
  function srow(y){ let r=top0, acc=0; while(acc+RH<=y && r<top0+800){ acc+=RH; r++; } return {row:r, off:Math.max(0,Math.round(y-acc))}; }
  const lineClr='4472C4'; let id=1;
  function place(xR,yT,w,h,inner){
    const f=sfr(xR), t=sfr(xR+w), fr=srow(yT), tr=srow(yT+h);
    return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${f.col}</xdr:col><xdr:colOff>${f.off}</xdr:colOff><xdr:row>${fr.row}</xdr:row><xdr:rowOff>${fr.off}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${t.col}</xdr:col><xdr:colOff>${t.off}</xdr:colOff><xdr:row>${tr.row}</xdr:row><xdr:rowOff>${tr.off}</xdr:rowOff></xdr:to>${inner}<xdr:clientData/></xdr:twoCellAnchor>`;
  }
  const rectLine=(xR,yT,w,h)=>{ const i=++id; w=Math.max(w,9525); h=Math.max(h,9525);
    return place(xR,yT,w,h,`<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${i}" name="ln${i}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${lineClr}"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ar-SA"/></a:p></xdr:txBody></xdr:sp>`); };
  const box=(n)=>{ const i=++id; const c=orgFill(n.type);
    const nameSz=Math.max(900,Math.min(1300,Math.round(1300*scale))), typeSz=Math.max(800,Math.min(1050,Math.round(1050*scale)));
    const body=`<xdr:txBody><a:bodyPr rot="0" wrap="square" lIns="18000" tIns="9000" rIns="18000" bIns="9000" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>`+
      `<a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr lang="ar-SA" sz="${nameSz}" b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${FONT}"/><a:cs typeface="${FONT}"/></a:rPr><a:t>${escXml(n.pos)}</a:t></a:r></a:p>`+
      `<a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr lang="ar-SA" sz="${typeSz}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${FONT}"/><a:cs typeface="${FONT}"/></a:rPr><a:t>(${escXml(n.type)})</a:t></a:r></a:p>`+
      `</xdr:txBody>`;
    return place(sx(n.x), sx(n.y), bw, bh,`<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${i}" name="box${i}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${bw}" cy="${bh}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${c.fill}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="${c.line}"/></a:solidFill></a:ln></xdr:spPr>${body}</xdr:sp>`); };
  const parts=[];
  L.all.forEach(n=>{ if(!n.children.length) return;
    const pc=sx(n.x+L.boxW/2), pbot=sx(n.y+L.boxH), busY=sx(n.y+L.boxH+L.vGap/2);
    parts.push(rectLine(pc-6350, pbot, 12700, busY-pbot));
    const cxs=n.children.map(ch=>sx(ch.x+L.boxW/2)); const xa=Math.min(pc,...cxs), xb=Math.max(pc,...cxs);
    parts.push(rectLine(xa, busY-6350, xb-xa, 12700));
    n.children.forEach(ch=>{ const cc=sx(ch.x+L.boxW/2); parts.push(rectLine(cc-6350, busY, 12700, sx(ch.y)-busY)); });
  });
  L.all.forEach(n=>parts.push(box(n)));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${parts.join('')}</xdr:wsDr>`;
}
async function injectExcelOrgChart(buffer, hrRows){
  try{
    const {persons}=buildPersons(hrRows); if(!persons.length) return buffer;
    const nH=(hrRows||[]).length, totalR=4+nH, titleR=totalR+2, top0=titleR;
    const drawing=buildExcelOrgDrawingXml(hrRows, top0); if(!drawing) return buffer;
    const zip=await JSZip.loadAsync(buffer);
    const wbXml=await zip.file('xl/workbook.xml').async('string');
    const relsXml=await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const m=wbXml.match(/<sheet[^>]*name="الموارد البشرية"[^>]*r:id="([^"]+)"/);
    if(!m) return buffer;
    const rm=relsXml.match(new RegExp('<Relationship[^>]*Id="'+m[1]+'"[^>]*Target="([^"]+)"'));
    if(!rm) return buffer;
    const baseName=rm[1].split('/').pop();
    const sheetFile='xl/worksheets/'+baseName;
    const drawCount=Object.keys(zip.files).filter(f=>/^xl\/drawings\/drawing\d+\.xml$/.test(f)).length;
    const drawName='drawing'+(drawCount+1)+'.xml';
    zip.file('xl/drawings/'+drawName, drawing);
    const relPath='xl/worksheets/_rels/'+baseName+'.rels';
    let sheetRels = zip.file(relPath) ? await zip.file(relPath).async('string')
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const drawRelId='rIdOrg'+(drawCount+1);
    sheetRels=sheetRels.replace('</Relationships>', `<Relationship Id="${drawRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawName}"/></Relationships>`);
    zip.file(relPath, sheetRels);
    let sxml=await zip.file(sheetFile).async('string');
    if(!/<drawing /.test(sxml)) sxml=sxml.replace('</worksheet>', `<drawing r:id="${drawRelId}"/></worksheet>`);
    zip.file(sheetFile, sxml);
    let ct=await zip.file('[Content_Types].xml').async('string');
    if(!ct.includes('/xl/drawings/'+drawName)){
      ct=ct.replace('</Types>', `<Override PartName="/xl/drawings/${drawName}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
      zip.file('[Content_Types].xml', ct);
    }
    return zip.generateAsync({type:'nodebuffer', compression:'DEFLATE'});
  }catch(e){ return buffer; }
}

// ── financial dashboard (anchored shapes injected onto page 2) ──────
const FW = { card:'F1F5F9', track:'E2E8F0', muted:'64748B', netLbl:'CDD9EE',
  blue:'2F5496', orange:'ED7D31', navy:'1F3864', gray:'94A3B8', green:'63991F' };
let _fwz = 7000;
function fwAnchor(x,y,w,h,inner,name){
  const z=_fwz++;
  return `<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${z}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${Math.round(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${Math.round(y)}</wp:posOffset></wp:positionV><wp:extent cx="${Math.round(w)}" cy="${Math.round(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${z}" name="${name}${z}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvPr id="${z}" name="${name}${z}"/>${inner}</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}
function fwTxt(lines){
  const ps=lines.map(l=>`<w:p><w:pPr><w:bidi/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>${l.bold?'<w:b/><w:bCs/>':''}<w:sz w:val="${l.sz}"/><w:szCs w:val="${l.sz}"/><w:color w:val="${l.color}"/></w:rPr><w:t xml:space="preserve">${escXml(l.text)}</w:t></w:r></w:p>`).join('');
  return `<wps:txbx><w:txbxContent>${ps}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" wrap="square" lIns="36000" tIns="18000" rIns="36000" bIns="18000" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>`;
}
function fwRect(x,y,w,h,fill,lines,opts={}){
  const hasLine = opts.line && opts.line!=='none';
  const ln = hasLine ? `<a:ln w="${opts.lineW||9525}"><a:solidFill><a:srgbClr val="${opts.line}"/></a:solidFill></a:ln>` : `<a:ln><a:noFill/></a:ln>`;
  const geom=opts.geom||'roundRect';
  const av=geom==='roundRect'?`<a:avLst><a:gd name="adj" fmla="val ${opts.radius||12000}"/></a:avLst>`:'<a:avLst/>';
  const fillXml = (fill==='none')?'<a:noFill/>':`<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`;
  return fwAnchor(x,y,w,h,`<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="${geom}">${av}</a:prstGeom>${fillXml}${ln}</wps:spPr>${lines?fwTxt(lines):'<wps:bodyPr/>'}`,'fw');
}
function fwRing(cx,cy,r,pct,color,label,valText){
  const x=cx-r,y=cy-r,d=2*r;
  const start=16200000, sweep=Math.max(0,Math.min(99.9,pct))/100*360, a2=Math.round((start+sweep*60000)%21600000);
  const ell=(rr,fill)=>fwRect(cx-rr,cy-rr,2*rr,2*rr,fill,null,{geom:'ellipse'});
  let out=ell(r,FW.track);                                   // track disc
  if(sweep>0) out+=fwAnchor(x,y,d,d,`<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(d)}" cy="${Math.round(d)}"/></a:xfrm><a:prstGeom prst="pie"><a:avLst><a:gd name="adj1" fmla="val ${start}"/><a:gd name="adj2" fmla="val ${a2}"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/>`,'arc');
  out+=ell(r*0.60,'FFFFFF');                                 // hole
  out+=fwAnchor(x,cy-330000,d,660000,`<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(d)}" cy="660000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>${fwTxt([{text:valText,sz:30,bold:true,color:FW.navy},{text:label,sz:17,bold:false,color:FW.muted}])}`,'rt');
  return out;
}
function fwBar(x,y,w,h,pct,color){
  let out=fwRect(x,y,w,h,'EEF2F7',null,{geom:'roundRect',radius:50000});
  const fw=Math.max(0,Math.min(100,pct))/100*w;
  if(fw>3000) out+=fwRect(x,y,fw,h,color,null,{geom:'roundRect',radius:50000});
  return out;
}
function buildFinanceDashXml(summary){
  summary=summary||{};
  const pm=s=>parseFloat(String(s||'0').replace(/[^0-9.-]/g,''))||0;
  const rv=pm(summary.revenueAnnual), opsA=pm(summary.opsAnnual), fxA=pm(summary.fixedAnnual),
        ft=pm(summary.foundingTotal), net=pm(summary.netProfit), dep=pm(summary.depreciation);
  const f1=x=>{const r=Math.round(x*10)/10;return Number.isInteger(r)?String(r):r.toFixed(1);};
  const payback=(net>0&&ft>0)?paybackPhrase(ft/net):'—';
  const margin=rv>0?net/rv*100:0, roi=ft>0?net/ft*100:0;
  const costs=opsA+fxA+dep, costsPct=rv>0?costs/rv*100:0, profitPct=rv>0?net/rv*100:0;
  const PW=16838*635, MX=600000, CW=PW-2*MX, gap=170000;
  _fwz=7000; let s='';
  // 1) KPI cards
  const cw=(CW-3*gap)/4, ch=940000, top=1230000;
  const cards=[['الإيرادات السنوية',summary.revenueAnnual||'$0'],['التكاليف التشغيلية',summary.opsAnnual||'$0'],
               ['التكاليف الثابتة',summary.fixedAnnual||'$0'],['التكاليف التأسيسية',summary.foundingTotal||'$0']];
  cards.forEach((c,i)=>{ s+=fwRect(MX+i*(cw+gap),top,cw,ch,FW.card,[{text:c[0],sz:18,bold:false,color:FW.muted},{text:c[1],sz:32,bold:true,color:FW.navy}],{radius:14000}); });
  // 2) net profit wide card
  const ny=top+ch+gap, nh=690000;
  s+=fwRect(MX,ny,CW,nh,FW.navy,[{text:'صافي الربح السنوي',sz:19,bold:false,color:FW.netLbl},{text:summary.netProfit||'$0',sz:40,bold:true,color:FW.orange}],{radius:14000});
  // 3) payback card + two donut gauges
  const by=ny+nh+gap, bw=(CW-2*gap)/3, bh=1380000;
  s+=fwRect(MX,by,bw,bh,FW.card,[{text:'فترة استرداد رأس المال',sz:18,bold:false,color:FW.muted},{text:payback,sz:26,bold:true,color:FW.navy}],{radius:14000});
  s+=fwRect(MX+bw+gap,by,bw,bh,'FFFFFF',null,{line:FW.track,radius:14000});
  s+=fwRing(MX+bw+gap+bw/2, by+bh/2, 560000, margin, FW.blue, 'هامش الربح', f1(margin)+'%');
  s+=fwRect(MX+2*(bw+gap),by,bw,bh,'FFFFFF',null,{line:FW.track,radius:14000});
  s+=fwRing(MX+2*(bw+gap)+bw/2, by+bh/2, 560000, roi, FW.orange, 'العائد على الاستثمار', f1(roi)+'%');
  // 4) comparison bars (revenue / costs / profit) scaled to revenue
  const cy0=by+bh+gap, rowH=270000, rowGap=120000, labW=CW*0.26, barX=MX+CW*0.30, barW=CW*0.66;
  const rows=[['الإيرادات',100,FW.blue],['التكاليف',costsPct,FW.gray],['صافي الربح',profitPct,FW.green]];
  rows.forEach((r,i)=>{ const yy=cy0+i*(rowH+rowGap);
    s+=fwRect(MX,yy-30000,labW,rowH+60000,'none',[{text:r[0],sz:19,bold:false,color:FW.muted}]);
    s+=fwBar(barX,yy,barW,rowH,r[1],r[2]); });
  return `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${s}</w:p>`;
}
async function injectFinance(buffer, summary){
  try{
    const para=buildFinanceDashXml(summary);
    const zip=await JSZip.loadAsync(buffer);
    let xml=await zip.file('word/document.xml').async('string');
    if(!xml.includes('[[FINANCE]]')) return buffer;
    xml=xml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>).)*?\[\[FINANCE\]\](?:(?!<\/w:p>).)*?<\/w:p>/s, para);
    zip.file('word/document.xml', xml);
    return zip.generateAsync({type:'nodebuffer', compression:'DEFLATE'});
  }catch(_){ return buffer; }
}


// ════════════════════════════════════════════════════════════
//  EMAIL SENDER
// ════════════════════════════════════════════════════════════
function gmailTransport(port){
  return nodemailer.createTransport({
    host:'smtp.gmail.com', port, secure: port===465, requireTLS: port===587,
    auth:{ user:process.env.EMAIL_USER, pass:process.env.EMAIL_PASS },
    connectionTimeout:15000, greetingTimeout:10000, socketTimeout:20000,
  });
}
// try port 587 (STARTTLS) first, then 465 (SSL) — whichever the host allows
async function sendViaGmail(mail){
  let lastErr;
  for(const port of [587,465]){
    try { await gmailTransport(port).sendMail(mail); console.log('📧 أُرسل عبر المنفذ', port); return port; }
    catch(e){ lastErr=e; console.error('تعذّر الإرسال عبر المنفذ', port, '—', e.code||e.message); }
  }
  throw lastErr;
}
async function sendEmail(data, id, excelBuf, wordBuf) {
  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
    console.log('⚠️  متغيرات الإيميل غير مضبوطة — تخطي الإرسال');
    return;
  }
  const to = process.env.RECEIVER_EMAIL || process.env.EMAIL_USER;
  const cc = (process.env.CC_EMAIL || '').trim();
  const projName = (data.projectTitle || data.projectIdea || 'مشروع').toString().substring(0,60);
  const applicant = (data.applicantName || '—').toString().substring(0,40);
  const base = fileBase(data);
  await sendViaGmail({
    from:    `"نظام دراسة المشاريع" <${process.env.EMAIL_USER}>`,
    to,
    ...(cc ? { cc } : {}),
    subject: `📋 نموذج جديد — ${projName} — ${applicant}`,
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
      { filename:`${base}.xlsx`, content:excelBuf, contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      { filename:`${base}.docx`, content:wordBuf,  contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
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
    // respond NOW so the browser downloads instantly …
    res.json({ success:true, id, fileName: fileBase(submission)+'.docx' });
    // … then send the email in the background with clear logging
    sendEmail(submission, id, excelBuf, wordBuf)
      .then(() => console.log(`📧 تم إرسال الإيميل بنجاح: ${id}`))
      .catch(e => console.error(`❌ فشل إرسال الإيميل (${id}):`, e && (e.message||e), 'code=', e && e.code));
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

// ── تشخيص الإيميل: افتح /api/email-test لمعرفة سبب المشكلة بالضبط ──
app.get('/api/email-test', async (req, res) => {
  const present = {
    EMAIL_USER: process.env.EMAIL_USER ? '✓ مضبوط' : '✗ غير مضبوط',
    EMAIL_PASS: process.env.EMAIL_PASS ? `✓ مضبوط (${process.env.EMAIL_PASS.length} حرف)` : '✗ غير مضبوط',
    RECEIVER_EMAIL: process.env.RECEIVER_EMAIL || '(سيُستخدم EMAIL_USER)',
    CC_EMAIL: process.env.CC_EMAIL || '(لا يوجد)',
  };
  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
    return res.json({ ok:false, step:'config', message:'EMAIL_USER أو EMAIL_PASS غير مضبوطة في Render → Environment', present });
  }
  // try to connect/authenticate on both ports and report each
  const ports = {};
  let working = null;
  for (const port of [587, 465]) {
    try { await gmailTransport(port).verify(); ports['port_'+port] = '✓ يعمل'; if(!working) working = port; }
    catch(e){ ports['port_'+port] = `✗ ${e.code||e.message}`; }
  }
  if (!working) {
    return res.json({ ok:false, step:'verify',
      message:'تعذّر الاتصال بـ Gmail على المنفذين 587 و465. إذا كان الخطأ ETIMEDOUT فالاستضافة تحجب منافذ SMTP — الحل التحويل لخدمة إيميل عبر HTTP (مثل Brevo/Resend). إذا كان EAUTH فالمشكلة بكلمة مرور التطبيق.',
      ports, present });
  }
  try {
    const to = process.env.RECEIVER_EMAIL || process.env.EMAIL_USER;
    const cc = (process.env.CC_EMAIL||'').trim();
    const info = await gmailTransport(working).sendMail({
      from:`"اختبار النظام" <${process.env.EMAIL_USER}>`, to, ...(cc?{cc}:{}),
      subject:'✅ اختبار إيميل — نظام دراسة المشاريع',
      text:'إذا وصلتك هذه الرسالة، فإعدادات الإيميل تعمل بنجاح وكل النماذج القادمة ستصلك.' });
    return res.json({ ok:true, step:'sent', message:`تم إرسال إيميل اختبار بنجاح عبر المنفذ ${working} — تفقّد صندوق الوارد (والـ CC).`, working_port:working, to, cc:cc||'(لا يوجد)', messageId:info.messageId, ports, present });
  } catch(e){ return res.json({ ok:false, step:'send', message:'نجح الاتصال لكن فشل إرسال الرسالة', error:e.message, code:e.code, ports, present }); }
});

// مسار التقاط الكل (يجب أن يبقى آخر مسار) — يرجّع صفحة التقديم لأي رابط غير معروف
app.get('*', (req, res) => res.sendFile(path.join(__dirname,'public','index.html')));

app.listen(PORT, () => console.log(`🚀 يعمل على المنفذ ${PORT}`));
module.exports = app;
