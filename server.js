/**
 * ─────────────────────────────────────────────────────────────
 * نموذج طرح دراسة مشروع — Backend Server v3 (Premium Edition)
 * Express · docx · ExcelJS · Nodemailer
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
  AlignmentType, WidthType,
  HeadingLevel, ShadingType, BorderStyle,
  Header, Footer, PageNumber
} = require('docx');

const app  = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

app.use(cors());
app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, 'public')));

function genId() {
  return Date.now().toString(36).toUpperCase() + '-' + crypto.randomBytes(3).toString('hex').toUpperCase();
}
function loadIndex() {
  const p = path.join(DATA_DIR, 'index.json');
  if (!fs.existsSync(p)) return [];
  try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch { return []; }
}
function saveIndex(index) {
  fs.writeFileSync(path.join(DATA_DIR, 'index.json'), JSON.stringify(index, null, 2), 'utf8');
}
const MONTHS = ['شهر 1','شهر 2','شهر 3','شهر 4','شهر 5','شهر 6', 'شهر 7','شهر 8','شهر 9','شهر 10','شهر 11','شهر 12'];

// ════════════════════════════════════════════════════════════
//  EXCEL GENERATOR (PREMIUM)
// ════════════════════════════════════════════════════════════
async function generateExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'نظام المشاريع التنموية';
  wb.created = new Date();

  const C_HEADER = '1A3A5C';
  const C_ACCENT = 'C8A84B';
  const C_SUBTOT = 'EAF0F8';
  const C_TOTAL  = 'D1DBE8';
  const C_ZEBRA  = 'F4F7FA';

  function applyZebra(sheet, rowStart, rowEnd, colCount) {
    for(let i=rowStart; i<=rowEnd; i++) {
      if ((i - rowStart) % 2 === 1) {
        const r = sheet.getRow(i);
        for(let c=1; c<=colCount; c++) {
          if(!r.getCell(c).fill) r.getCell(c).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_ZEBRA} };
        }
      }
    }
  }

  function hStyle() {
    return {
      font: { bold:true, color:{argb:'FFFFFFFF'}, name:'Arial', size:12 },
      fill: { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_HEADER} },
      alignment: { horizontal:'center', vertical:'middle', wrapText:true, readingOrder:2 },
      border: { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} }
    };
  }

  function dStyle() {
    return {
      font: { name:'Arial', size:11 },
      alignment: { horizontal:'center', vertical:'middle', readingOrder:2 },
      border: { top:{style:'hair', color:{argb:'FFD1DBE8'}}, bottom:{style:'hair', color:{argb:'FFD1DBE8'}}, left:{style:'hair', color:{argb:'FFD1DBE8'}}, right:{style:'hair', color:{argb:'FFD1DBE8'}} }
    };
  }

  // ── Sheet 1: الملخص ─────────────────────────────────────
  const s0 = wb.addWorksheet('الملخص والتفاصيل', { views: [{ rightToLeft:true }] });
  s0.columns = [{ width:45 }, { width:35 }];
  
  s0.mergeCells('A1:B1');
  s0.getCell('A1').value = 'ملخص معلومات المشروع';
  s0.getCell('A1').font = { bold:true, size:18, color:{argb:'FFFFFFFF'}, name:'Arial' };
  s0.getCell('A1').fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_HEADER} };
  s0.getCell('A1').alignment = { horizontal:'center', vertical:'middle', readingOrder:2 };
  s0.getRow(1).height = 40;

  s0.mergeCells('A2:B2');
  s0.getCell('A2').value = 'تاريخ الإرسال: ' + new Date(data.submittedAt||Date.now()).toLocaleString('ar-SA');
  s0.getCell('A2').font = { bold:true, size:11, color:{argb:'FF'+C_HEADER}, name:'Arial' };
  s0.getCell('A2').fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_ACCENT} };
  s0.getCell('A2').alignment = { horizontal:'center', readingOrder:2 };
  s0.getRow(2).height = 25;

  const summaryData = [
    ['فكرة المشروع', data.projectIdea||''],
    ['إجمالي التكاليف التأسيسية', data.summary?.foundingTotal||'$0'],
    ['إجمالي الإيرادات المتوقعة (سنوياً)', data.summary?.revenueAnnual||'$0'],
    ['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0'],
    ['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual||'$0'],
    ['الاهتلاك (سنوياً)', data.summary?.depreciation||'$0'],
    ['الربح الصافي (سنوياً)', data.summary?.netProfit||'$0'],
    ['عدد الموظفين', data.summary?.employees||'0'],
  ];

  summaryData.forEach(([label, val], i) => {
    const r = s0.getRow(i+4);
    r.height = 28;
    r.getCell(1).value = label;
    r.getCell(1).font  = { bold:true, name:'Arial', size:12, color:{argb:'FF1A3A5C'} };
    r.getCell(1).alignment = { horizontal:'right', vertical:'middle', readingOrder:2 };
    r.getCell(1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF9FAFC'} };
    
    r.getCell(2).value = String(val);
    r.getCell(2).alignment = { horizontal:'center', vertical:'middle', readingOrder:2 };
    r.getCell(2).font  = { bold: i===6, name:'Arial', size:12, color: i===6 ? {argb:'FF16A34A'} : {argb:'FF000000'} };
    
    [1,2].forEach(c => {
      r.getCell(c).border = { top:{style:'thin', color:{argb:'FFD1DBE8'}}, bottom:{style:'thin', color:{argb:'FFD1DBE8'}}, left:{style:'thin', color:{argb:'FFD1DBE8'}}, right:{style:'thin', color:{argb:'FFD1DBE8'}} };
    });
  });

  // ── Helper to build standard sheets ──────────────────────
  function buildDataSheet(name, columns, headers, rowsData, totalRowData) {
    const sheet = wb.addWorksheet(name, { views: [{ rightToLeft:true, state:'frozen', ySplit:1 }] });
    sheet.columns = columns.map(w => ({width:w}));
    
    const hr = sheet.getRow(1);
    hr.height = 30;
    headers.forEach((h,i) => { hr.getCell(i+1).value=h; Object.assign(hr.getCell(i+1), hStyle()); });
    sheet.autoFilter = { from: { row:1, column:1 }, to: { row:1, column:headers.length } };

    rowsData.forEach((rowVals, i) => {
      const row = sheet.getRow(i+2);
      row.height = 22;
      rowVals.forEach((val, c) => {
        row.getCell(c+1).value = val;
        Object.assign(row.getCell(c+1), dStyle());
        if(typeof val === 'number' && headers[c].includes('($)')) row.getCell(c+1).numFmt = '_("$"* #,##0.00_)';
      });
    });
    
    applyZebra(sheet, 2, rowsData.length+1, headers.length);

    if (totalRowData) {
      const tr = sheet.getRow(rowsData.length + 2);
      tr.height = 25;
      totalRowData.forEach((val, c) => {
        if(val !== null) {
          tr.getCell(c+1).value = val;
          tr.getCell(c+1).font = { bold:true, name:'Arial', size:12, color:{argb:'FF1A3A5C'} };
          tr.getCell(c+1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+C_TOTAL} };
          tr.getCell(c+1).alignment = { horizontal:'center', vertical:'middle', readingOrder:2 };
          if(typeof val === 'number') tr.getCell(c+1).numFmt = '_("$"* #,##0.00_)';
        }
      });
      sheet.mergeCells(`A${rowsData.length + 2}:${String.fromCharCode(64 + headers.length - 1)}${rowsData.length + 2}`);
      tr.getCell(1).alignment = { horizontal:'right', vertical:'middle' };
    }
  }

  // ── Sheet 2: التكاليف التأسيسية ────────────────────────
  if (data.foundingRows?.length) {
    let tot = 0;
    const rData = data.foundingRows.map((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot += v;
      return [i+1, r.cat, r.bayan, r.dep?'✓':'', parseFloat(r.qty)||0, r.notes, parseFloat(r.price)||0, v];
    });
    buildDataSheet('التكاليف التأسيسية', [8,18,28,10,10,22,22,22], 
      ['#','الصنف','البيان','اهتلاك','العدد','ملاحظات','التكلفة للواحدة ($)','التكلفة الإجمالية ($)'], 
      rData, ['الإجمالي المالي للتأسيس', null, null, null, null, null, null, tot]);
  }

  // ── Sheet 3: التكاليف الثابتة ─────────────────────────
  if (data.fixedRows?.length) {
    let tot = 0;
    const rData = data.fixedRows.map((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot += v;
      return [i+1, r.cat, r.bayan, parseFloat(r.qty)||0, r.notes, parseFloat(r.price)||0, v];
    });
    buildDataSheet('التكاليف الثابتة', [8,18,28,10,22,25,25], 
      ['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة ($)','التكلفة الشهرية الإجمالية ($)'], 
      rData, ['الإجمالي الشهري', null, null, null, null, null, tot]);
  }

  // ── Sheet 4: الموارد البشرية ──────────────────────────
  if (data.hrRows?.length) {
    let tot = 0;
    const rData = data.hrRows.map((r,i) => {
      const v = parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0;
      tot += v;
      return [i+1, r.position, r.type, parseFloat(r.qty)||0, r.reports, parseFloat(r.salary)||0, v];
    });
    buildDataSheet('الموارد البشرية', [8,22,18,10,22,25,25], 
      ['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي ($)','الراتب الشهري الإجمالي ($)'], 
      rData, ['إجمالي الرواتب الشهري', null, null, null, null, null, tot]);
  }

  return wb.xlsx.writeBuffer();
}

// ════════════════════════════════════════════════════════════
//  WORD GENERATOR (PREMIUM)
// ════════════════════════════════════════════════════════════
function sectionTitle(title) {
  return new Paragraph({
    bidirectional: true,
    heading: HeadingLevel.HEADING_1,
    spacing: { before:400, after:200 },
    border: { bottom: { color: "C8A84B", space: 1, value: "single", size: 12 } },
    children: [new TextRun({ text:title, bold:true, size:30, font:'Arial', color:'1A3A5C' })],
  });
}

function makeProTable(headers, rowsData) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: "1A3A5C" },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: "1A3A5C" },
      left: { style: BorderStyle.SINGLE, size: 2, color: "1A3A5C" },
      right: { style: BorderStyle.SINGLE, size: 2, color: "1A3A5C" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D1DBE8" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "D1DBE8" },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map(h => new TableCell({
          shading: { fill: "1A3A5C" },
          margins: { top: 120, bottom: 120, left: 100, right: 100 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(h), color: "FFFFFF", bold: true, font: "Arial", size: 22 })] })]
        }))
      }),
      ...rowsData.map((row, rIdx) => new TableRow({
        children: row.map(cell => new TableCell({
          shading: { fill: rIdx % 2 === 0 ? "F4F7FA" : "FFFFFF" },
          margins: { top: 100, bottom: 100, left: 100, right: 100 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(cell ?? ''), font: "Arial", size: 20 })] })]
        }))
      }))
    ]
  });
}

async function generateWord(data) {
  const pids = Object.keys(data.products||{}).filter(p=>data.products[p]?.name);
  const ch = [];

  // Cover Page
  ch.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2500, after: 400 }, children: [new TextRun({ text: "دراسة مشروع", bold: true, size: 64, color: "1A3A5C", font:"Arial" })] }));
  ch.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1200 }, children: [new TextRun({ text: data.projectIdea || 'غير محدد', bold: true, size: 36, color: "C8A84B", font:"Arial" })] }));
  ch.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "تاريخ الإرسال: " + new Date(data.submittedAt||Date.now()).toLocaleString('ar-SA'), size: 24, color: "555555", font:"Arial" })] }));
  ch.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "رقم الطلب: " + (data.id||"---"), size: 20, color: "888888", font:"Arial" })] }));
  ch.push(new Paragraph({ pageBreakBefore: true }));

  // Summary
  ch.push(sectionTitle('ملخص معلومات المشروع'));
  ch.push(makeProTable(['البند','القيمة'], [
    ['فكرة المشروع', data.projectIdea||''],
    ['إجمالي التكاليف التأسيسية', data.summary?.foundingTotal||'$0'],
    ['إجمالي الإيرادات السنوية', data.summary?.revenueAnnual||'$0'],
    ['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0'],
    ['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual||'$0'],
    ['الاهتلاك السنوي', data.summary?.depreciation||'$0'],
    ['الربح الصافي السنوي', data.summary?.netProfit||'$0'],
    ['عدد الموظفين', data.summary?.employees||'0'],
  ]));

  // Founding
  if (data.foundingRows?.length) {
    ch.push(sectionTitle('التكاليف التأسيسية'));
    ch.push(makeProTable(['#','الصنف','البيان','العدد','التكلفة للواحدة','الإجمالي'], 
      data.foundingRows.map((r,i)=>[i+1, r.cat, r.bayan, r.qty, r.price, r.total])
    ));
  }

  // Fixed
  if (data.fixedRows?.length) {
    ch.push(sectionTitle('التكاليف الثابتة (الشهرية)'));
    ch.push(makeProTable(['#','الصنف','البيان','العدد','الواحدة','الإجمالي'], 
      data.fixedRows.map((r,i)=>[i+1, r.cat, r.bayan, r.qty, r.price, r.total])
    ));
  }

  // HR
  if (data.hrRows?.length) {
    ch.push(sectionTitle('الموارد البشرية'));
    ch.push(makeProTable(['#','المنصب','النوع','العدد','الراتب','الإجمالي'], 
      data.hrRows.map((r,i)=>[i+1, r.position, r.type, r.qty, r.salary, r.total])
    ));
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial" }, paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { bottom: { color: "D1DBE8", space: 1, value: "single", size: 6 } },
              spacing: { after: 200 },
              children: [new TextRun({ text: "قسم المشاريع التنموية - نموذج طرح دراسة مشروع", color: "1A3A5C", bold: true, size: 20, font:"Arial" })]
            })
          ]
        })
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { color: "D1DBE8", space: 1, value: "single", size: 6 } },
              spacing: { before: 200 },
              children: [
                new TextRun({ text: "صفحة ", color: "888888", size: 18, font:"Arial" }),
                new TextRun({ children: [PageNumber.CURRENT], color: "888888", size: 18, font:"Arial" }),
                new TextRun({ text: " من ", color: "888888", size: 18, font:"Arial" }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], color: "888888", size: 18, font:"Arial" })
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
//  ROUTES & SERVER
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
