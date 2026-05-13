/**
 * ─────────────────────────────────────────────────────────────
 * نموذج طرح دراسة مشروع — Al-Majd Foundation Edition
 * ─────────────────────────────────────────────────────────────
 */

const express    = require('express');
const cors       = require('cors');
const fs         = require('fs');
const path       = require('path');
const crypto     = require('crypto');
const ExcelJS    = require('exceljs');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell,
  AlignmentType, WidthType, PageOrientation,
  HeadingLevel, BorderStyle,
  Header, Footer, PageNumber
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

app.use(cors());
app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, 'public')));

function genId() { return Date.now().toString(36).toUpperCase() + '-' + crypto.randomBytes(3).toString('hex').toUpperCase(); }
function loadIndex() {
  const p = path.join(DATA_DIR, 'index.json');
  if (!fs.existsSync(p)) return [];
  try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch { return []; }
}
function saveIndex(index) { fs.writeFileSync(path.join(DATA_DIR, 'index.json'), JSON.stringify(index, null, 2), 'utf8'); }

const MONTHS = ['شهر 1','شهر 2','شهر 3','شهر 4','شهر 5','شهر 6', 'شهر 7','شهر 8','شهر 9','شهر 10','شهر 11','شهر 12'];
function fmtMoney(n) { return '$' + Number(n||0).toLocaleString('en-US', {minimumFractionDigits:0, maximumFractionDigits:2}); }

// ════════════════════════════════════════════════════════════
//  EXCEL GENERATOR
// ════════════════════════════════════════════════════════════
async function generateExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'مؤسسة المجد التنموية';
  wb.created = new Date();

  const C_HEADER = '1A3A5C', C_ACCENT = 'C8A84B', C_TOTAL = 'D1DBE8', C_ZEBRA = 'F4F7FA';

  function applyZebra(sheet, rowStart, rowEnd, colCount) {
    for(let i=rowStart; i<=rowEnd; i++) {
      if ((i - rowStart) % 2 === 1) {
        const r = sheet.getRow(i);
        for(let c=1; c<=colCount; c++) { if(!r.getCell(c).fill) r.getCell(c).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_ZEBRA} }; }
      }
    }
  }

  function hStyle() {
    return { font: { bold:true, color:{argb:'FFFFFFFF'}, name:'Arial', size:12 }, fill: { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_HEADER} }, alignment: { horizontal:'center', vertical:'middle', wrapText:true, readingOrder:2 }, border: { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} } };
  }

  function dStyle() {
    return { font: { name:'Arial', size:11 }, alignment: { horizontal:'center', vertical:'middle', readingOrder:2 }, border: { top:{style:'hair', color:{argb:'FFD1DBE8'}}, bottom:{style:'hair', color:{argb:'FFD1DBE8'}}, left:{style:'hair', color:{argb:'FFD1DBE8'}}, right:{style:'hair', color:{argb:'FFD1DBE8'}} } };
  }

  function buildDataSheet(name, columns, headers, rowsData, totalRowData) {
    const sheet = wb.addWorksheet(name, { views: [{ rightToLeft:true, state:'frozen', ySplit:1 }] });
    sheet.columns = columns.map(w => ({width:w}));
    const hr = sheet.getRow(1); hr.height = 30;
    headers.forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle()); });
    
    rowsData.forEach((rowVals, i) => {
      const row = sheet.getRow(i+2); row.height = 22;
      rowVals.forEach((val, c) => {
        row.getCell(c+1).value = val; Object.assign(row.getCell(c+1), dStyle());
        if(typeof val === 'number' && (headers[c].includes('($)') || headers[c].includes('الإجمالي'))) row.getCell(c+1).numFmt = '_("$"* #,##0.00_)';
      });
    });
    applyZebra(sheet, 2, rowsData.length+1, headers.length);

    if (totalRowData) {
      const tr = sheet.getRow(rowsData.length + 2); tr.height = 25;
      totalRowData.forEach((val, c) => {
        if(val !== null) {
          tr.getCell(c+1).value = val; tr.getCell(c+1).font = { bold:true, name:'Arial', size:12, color:{argb:'FF1A3A5C'} };
          tr.getCell(c+1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_TOTAL} }; tr.getCell(c+1).alignment = { horizontal:'center', vertical:'middle', readingOrder:2 };
          if(typeof val === 'number') tr.getCell(c+1).numFmt = '_("$"* #,##0.00_)';
        }
      });
      sheet.mergeCells(`A${rowsData.length + 2}:${String.fromCharCode(64 + headers.length - 1)}${rowsData.length + 2}`);
      tr.getCell(1).alignment = { horizontal:'right', vertical:'middle' };
    }
  }

  // الملخص
  const s0 = wb.addWorksheet('الملخص والتفاصيل', { views: [{ rightToLeft:true }] });
  s0.columns = [{ width:45 }, { width:35 }];
  s0.mergeCells('A1:B1');
  s0.getCell('A1').value = 'مؤسسة المجد التنموية - ملخص معلومات المشروع';
  s0.getCell('A1').font = { bold:true, size:18, color:{argb:'FFFFFFFF'}, name:'Arial' }; s0.getCell('A1').fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_HEADER} }; s0.getCell('A1').alignment = { horizontal:'center', vertical:'middle', readingOrder:2 }; s0.getRow(1).height = 40;
  
  s0.mergeCells('A2:B2'); s0.getCell('A2').value = 'رقم الطلب: ' + data.id + ' | تاريخ الإرسال: ' + new Date(data.submittedAt||Date.now()).toLocaleString('ar-SA');
  s0.getCell('A2').font = { bold:true, size:11, color:{argb:'FF'+C_HEADER}, name:'Arial' }; s0.getCell('A2').fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_ACCENT} }; s0.getCell('A2').alignment = { horizontal:'center', readingOrder:2 }; s0.getRow(2).height = 25;

  const summaryData = [
    ['اسم مقدم المشروع', data.submitterName||''], ['رقم الهاتف', data.submitterPhone||''], ['فكرة المشروع', data.projectIdea||''], ['إجمالي التكاليف التأسيسية', data.summary?.foundingTotal||'$0'], ['إجمالي الإيرادات المتوقعة (سنوياً)', data.summary?.revenueAnnual||'$0'], ['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0'], ['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual||'$0'], ['الاهتلاك (سنوياً)', data.summary?.depreciation||'$0'], ['الربح الصافي (سنوياً)', data.summary?.netProfit||'$0'], ['عدد الموظفين', data.summary?.employees||'0'],
  ];

  summaryData.forEach(([label, val], i) => {
    const r = s0.getRow(i+4); r.height = 28;
    r.getCell(1).value = label; r.getCell(1).font = { bold:true, name:'Arial', size:12, color:{argb:'FF1A3A5C'} }; r.getCell(1).alignment = { horizontal:'right', vertical:'middle', readingOrder:2 }; r.getCell(1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF9FAFC'} };
    r.getCell(2).value = String(val); r.getCell(2).alignment = { horizontal:'center', vertical:'middle', readingOrder:2 }; r.getCell(2).font = { bold: i===8, name:'Arial', size:12, color: i===8 ? {argb:'FF16A34A'} : {argb:'FF000000'} };
    [1,2].forEach(c => r.getCell(c).border = { top:{style:'thin', color:{argb:'FFD1DBE8'}}, bottom:{style:'thin', color:{argb:'FFD1DBE8'}}, left:{style:'thin', color:{argb:'FFD1DBE8'}}, right:{style:'thin', color:{argb:'FFD1DBE8'}} });
  });

  // التأسيسية
  if (data.foundingRows?.length) {
    let tot = 0;
    const rData = data.foundingRows.map((r,i) => { const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0; tot+=v; return [i+1, r.cat, r.bayan, r.dep?'✓':'', parseFloat(r.qty)||0, r.notes, parseFloat(r.price)||0, v]; });
    buildDataSheet('التكاليف التأسيسية', [8,18,28,10,10,22,22,22], ['#','الصنف','البيان','اهتلاك','العدد','ملاحظات','التكلفة للواحدة ($)','التكلفة الإجمالية ($)'], rData, ['الإجمالي المالي للتأسيس', null, null, null, null, null, null, tot]);
  }

  // المنتجات والإيرادات والتشغيلية والاهتلاك
  const prodsList = Object.values(data.products||{});
  if(prodsList.length > 0) {
    buildDataSheet('المنتجات', [8,25,15,40], ['#','المنتج','الواحدة','المكونات'], prodsList.map((p,i)=>[i+1, p.name, p.unit, p.components.filter(c=>c).join('، ')]));
    
    const revRows = [];
    Object.entries(data.products||{}).forEach(([pid, p]) => {
        const rev = data.revenueData?.[pid]||{};
        const r1=[p.name,'الكمية']; let t1=0; for(let i=0;i<12;i++){ const q=Number(rev[i]?.qty||0); r1.push(q); t1+=q; } r1.push(t1); revRows.push(r1);
        const r2=[p.name,'إجمالي السعر ($)']; let t2=0; for(let i=0;i<12;i++){ const t=Number(rev[i]?.qty||0)*Number(rev[i]?.price||0); r2.push(t); t2+=t; } r2.push(t2); revRows.push(r2);
    });
    buildDataSheet('الإيرادات', [25,15, 12,12,12,12,12,12,12,12,12,12,12,12, 18], ['المنتج','البيان',...MONTHS,'الإجمالي ($)'], revRows);

    const opsRows = [];
    Object.entries(data.products||{}).forEach(([pid, p]) => {
        p.components.filter(c=>c).forEach((c, ci) => {
            const r = [p.name, c]; let cTot=0;
            for(let i=0;i<12;i++){ const v=Number(data.opsData?.[pid]?.[ci]?.[i]||0)*Number(data.revenueData?.[pid]?.[i]?.qty||0); r.push(v); cTot+=v; }
            r.push(cTot); opsRows.push(r);
        });
    });
    buildDataSheet('التكاليف التشغيلية', [20,20, 12,12,12,12,12,12,12,12,12,12,12,12, 18], ['المنتج','المكون',...MONTHS,'الإجمالي ($)'], opsRows);
  }

  // الموارد البشرية
  if (data.hrRows?.length) {
    let tot = 0; const rData = data.hrRows.map((r,i) => { const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0; tot+=v; return [i+1, r.position, r.type, parseFloat(r.qty)||0, r.reports, parseFloat(r.salary)||0, v]; });
    buildDataSheet('الموارد البشرية', [8,22,18,10,22,25,25], ['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي ($)','الراتب الشهري الإجمالي ($)'], rData, ['إجمالي الرواتب الشهري', null, null, null, null, null, tot]);
  }

  // الثابتة
  if (data.fixedRows?.length) {
    let tot = 0; const rData = data.fixedRows.map((r,i) => { const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0; tot+=v; return [i+1, r.cat, r.bayan, parseFloat(r.qty)||0, r.notes, parseFloat(r.price)||0, v]; });
    buildDataSheet('التكاليف الثابتة', [8,18,28,10,22,25,25], ['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة ($)','التكلفة الشهرية الإجمالية ($)'], rData, ['الإجمالي الشهري', null, null, null, null, null, tot]);
  }

  // الاهتلاك
  if (data.depData) {
    const depRows = []; let dIdx=1; let dTot=0;
    data.foundingRows?.forEach((fr, i) => {
        if(fr.dep) {
            const pct = Number(data.depData?.[i+1]||0); const depU = Number(fr.price||0)*(pct/100); const depT = depU*Number(fr.qty||0);
            dTot+=depT; depRows.push([dIdx++, fr.cat, fr.bayan, pct, Number(fr.qty||0), depU, depT]);
        }
    });
    if(depRows.length>0) buildDataSheet('الاهتلاك', [8,20,30,12,10,20,20], ['#','الصنف','البيان','النسبة (%)','العدد','إهلاك الواحدة ($)','الإجمالي ($)'], depRows, ['الإجمالي',null,null,null,null,null,dTot]);
  }

  return wb.xlsx.writeBuffer();
}

// ════════════════════════════════════════════════════════════
//  WORD GENERATOR (AL-MAJD EDITION - LANDSCAPE - RTL)
// ════════════════════════════════════════════════════════════
function sectionTitle(title) {
  return new Paragraph({
    bidirectional: true, heading: HeadingLevel.HEADING_1,
    spacing: { before:400, after:200 }, border: { bottom: { color: "C8A84B", space: 1, value: "single", size: 12 } },
    children: [new TextRun({ text:title, bold:true, size:30, font:'Arial', color:'1A3A5C' })],
  });
}

function makeProTable(headers, rowsData) {
  return new Table({
    visuallyRightToLeft: true,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C" }, right: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "D1DBE8" }, insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "D1DBE8" },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map(h => new TableCell({
          shading: { fill: "1A3A5C" }, margins: { top: 120, bottom: 120, left: 100, right: 100 },
          children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(h), color: "FFFFFF", bold: true, font: "Arial", size: 22 })] })]
        }))
      }),
      ...rowsData.map((row, rIdx) => new TableRow({
        children: row.map(cell => new TableCell({
          shading: { fill: rIdx % 2 === 0 ? "F4F7FA" : "FFFFFF" }, margins: { top: 100, bottom: 100, left: 100, right: 100 },
          children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(cell ?? ''), font: "Arial", size: 20 })] })]
        }))
      }))
    ]
  });
}

async function generateWord(data) {
  const ch = [];

  ch.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { before: 1500, after: 400 }, children: [new TextRun({ text: "دراسة مشروع", bold: true, size: 64, color: "1A3A5C", font:"Arial" })] }));
  ch.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 1200 }, children: [new TextRun({ text: data.projectIdea || 'غير محدد', bold: true, size: 36, color: "C8A84B", font:"Arial" })] }));
  ch.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "تاريخ الإرسال: " + new Date(data.submittedAt||Date.now()).toLocaleString('ar-SA'), size: 24, color: "555555", font:"Arial" })] }));
  ch.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "رقم الطلب: " + (data.id||"---"), size: 20, color: "888888", font:"Arial" })] }));
  ch.push(new Paragraph({ pageBreakBefore: true }));

  ch.push(sectionTitle('ملخص معلومات المشروع'));
  ch.push(makeProTable(['البند','القيمة'], [
    ['اسم مقدم المشروع', data.submitterName||''], ['رقم الهاتف', data.submitterPhone||''],
    ['فكرة المشروع', data.projectIdea||''], ['إجمالي التكاليف التأسيسية', data.summary?.foundingTotal||'$0'],
    ['إجمالي الإيرادات السنوية', data.summary?.revenueAnnual||'$0'], ['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0'],
    ['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual||'$0'], ['الاهتلاك السنوي', data.summary?.depreciation||'$0'],
    ['الربح الصافي السنوي', data.summary?.netProfit||'$0'], ['عدد الموظفين', data.summary?.employees||'0'],
  ]));

  if (data.foundingRows?.length) {
    ch.push(sectionTitle('التكاليف التأسيسية'));
    ch.push(makeProTable(['#','الصنف','البيان','العدد','التكلفة للواحدة','الإجمالي'], data.foundingRows.map((r,i)=>[i+1, r.cat, r.bayan, r.qty, fmtMoney(r.price), r.total])));
  }

  const prodsList = Object.values(data.products||{});
  if(prodsList.length > 0) {
    ch.push(sectionTitle('المنتجات والمكونات'));
    ch.push(makeProTable(['اسم المنتج', 'الواحدة', 'المكونات'], prodsList.map(p=>[p.name, p.unit, p.components.filter(c=>c).join('، ')])));

    ch.push(sectionTitle('الإيرادات المتوقعة (شهرياً)'));
    const revRows = [];
    Object.entries(data.products||{}).forEach(([pid, p]) => {
        const rev = data.revenueData?.[pid]||{};
        const r1=[p.name+' (كمية)', p.unit]; let t1=0; for(let i=0;i<12;i++){ const q=Number(rev[i]?.qty||0); r1.push(q); t1+=q; } r1.push(t1); revRows.push(r1);
        const r2=[p.name+' (السعر)', '$']; let t2=0; for(let i=0;i<12;i++){ const t=Number(rev[i]?.qty||0)*Number(rev[i]?.price||0); r2.push(fmtMoney(t)); t2+=t; } r2.push(fmtMoney(t2)); revRows.push(r2);
    });
    ch.push(makeProTable(['البيان', 'الواحدة', ...MONTHS, 'الإجمالي'], revRows));

    ch.push(sectionTitle('التكاليف التشغيلية (شهرياً)'));
    const opsRows = [];
    Object.entries(data.products||{}).forEach(([pid, p]) => {
        p.components.filter(c=>c).forEach((c, ci) => {
            const r = [p.name, c]; let cTot=0;
            for(let i=0;i<12;i++){ const v=Number(data.opsData?.[pid]?.[ci]?.[i]||0)*Number(data.revenueData?.[pid]?.[i]?.qty||0); r.push(fmtMoney(v)); cTot+=v; }
            r.push(fmtMoney(cTot)); opsRows.push(r);
        });
    });
    ch.push(makeProTable(['المنتج', 'المكون', ...MONTHS, 'الإجمالي'], opsRows));
  }

  if (data.hrRows?.length) {
    ch.push(sectionTitle('الموارد البشرية'));
    ch.push(makeProTable(['#','المنصب','النوع','العدد','الراتب','الإجمالي'], data.hrRows.map((r,i)=>[i+1, r.position, r.type, r.qty, fmtMoney(r.salary), r.total])));

    ch.push(sectionTitle('الهيكل التنظيمي للموارد البشرية'));
    const nodes = data.hrRows.map(r => ({ ...r, children: [] }));
    const tree = []; const lookup = {};
    nodes.forEach(n => lookup[n.position] = n);
    nodes.forEach(n => { if(n.reports && lookup[n.reports]) lookup[n.reports].children.push(n); else tree.push(n); });

    function addNodes(nodesList, level) {
        nodesList.forEach(node => {
            ch.push(new Paragraph({
                indent: { start: level * 720 }, bidirectional: true, spacing: { before: 100, after: 100 },
                children: [
                    new TextRun({ text: "▪ ", color: "C8A84B", bold: true, size: 28 }),
                    new TextRun({ text: `${node.position} (العدد: ${node.qty}) `, bold: true, color: "1A3A5C", size: 24, font: "Arial" }),
                    new TextRun({ text: `[${node.type}]`, color: "888888", size: 20, font: "Arial" })
                ]
            }));
            addNodes(node.children, level + 1);
        });
    }
    addNodes(tree, 0);
  }

  if (data.fixedRows?.length) {
    ch.push(sectionTitle('التكاليف الثابتة (الشهرية)'));
    ch.push(makeProTable(['#','الصنف','البيان','العدد','الواحدة','الإجمالي'], data.fixedRows.map((r,i)=>[i+1, r.cat, r.bayan, r.qty, fmtMoney(r.price), r.total])));
  }

  if (data.depData) {
    const dRows = []; let dIdx=1;
    data.foundingRows?.forEach((fr, i) => {
        if(fr.dep) {
            const pct = Number(data.depData?.[i+1]||0); const depU = Number(fr.price||0)*(pct/100); const depT = depU*Number(fr.qty||0);
            dRows.push([dIdx++, fr.cat, fr.bayan, pct+'%', Number(fr.qty||0), fmtMoney(depU), fmtMoney(depT)]);
        }
    });
    if(dRows.length>0) {
      ch.push(sectionTitle('جدول الاهتلاك'));
      ch.push(makeProTable(['#', 'الصنف', 'البيان', 'النسبة (%)', 'العدد', 'إهلاك الواحدة', 'الإجمالي'], dRows));
    }
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial" }, paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT } } } },
    sections: [{
      // تخطيط عرضي (Landscape) للوورد ليتسع لجداول 12 شهر بشكل رائع
      properties: { page: { size: { orientation: PageOrientation.LANDSCAPE, width: 16838, height: 11906 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              bidirectional: true, alignment: AlignmentType.CENTER, border: { bottom: { color: "C8A84B", space: 1, value: "single", size: 12 } }, spacing: { after: 200 },
              children: [new TextRun({ text: "مؤسسة المجد التنموية - قسم المشاريع التنموية", color: "1A3A5C", bold: true, size: 24, font:"Arial" })]
            })
          ]
        })
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              bidirectional: true, alignment: AlignmentType.CENTER, border: { top: { color: "C8A84B", space: 1, value: "single", size: 12 } }, spacing: { before: 200 },
              children: [
                new TextRun({ text: "صفحة ", color: "888888", size: 18, font:"Arial" }), new TextRun({ children: [PageNumber.CURRENT], color: "888888", size: 18, font:"Arial" }),
                new TextRun({ text: " من ", color: "888888", size: 18, font:"Arial" }), new TextRun({ children: [PageNumber.TOTAL_PAGES], color: "888888", size: 18, font:"Arial" })
              ]
            })
          ]
        })
      },
      children: ch
    }]
  });
  return Packer.toBuffer(doc);
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
