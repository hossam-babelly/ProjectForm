/**
 * ─────────────────────────────────────────────────────────────
 * نموذج طرح دراسة مشروع — Classic Word Template
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
  BorderStyle
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
//  WORD GENERATOR (CLASSIC EXACT TEMPLATE)
// ════════════════════════════════════════════════════════════

// دالة لإنشاء جداول بسيطة وكلاسيكية من اليمين لليسار
function createClassicTable(headers, rowsData) {
  return new Table({
    visuallyRightToLeft: true,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: "000000" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "000000" }, right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" }, insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map(h => new TableCell({
          margins: { top: 50, bottom: 50, left: 50, right: 50 },
          children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: String(h), bold: true, font: "Arial", size: 20 })] })]
        }))
      }),
      ...rowsData.map((row) => new TableRow({
        children: row.map(cell => new TableCell({
          margins: { top: 50, bottom: 50, left: 50, right: 50 },
          children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: String(cell ?? ''), font: "Arial", size: 20 })] })]
        }))
      }))
    ]
  });
}

function createParagraph(text, bold = false, size = 22) {
  return new Paragraph({
    bidirectional: true,
    alignment: AlignmentType.RIGHT,
    spacing: { after: 120 },
    children: [new TextRun({ text: text, bold: bold, font: "Arial", size: size })]
  });
}

async function generateWord(data) {
  const ch = [];

  // العناوين
  ch.push(createParagraph("قسم المشاريع التنموية", true, 28));
  ch.push(createParagraph("نموذج طرح دراسة مشروع", true, 28));
  ch.push(new Paragraph({ spacing: { after: 300 } })); // مسافة فارغة

  // 1. ملخص معلومات المشروع
  ch.push(createParagraph("ملخص معلومات المشروع", true, 24));
  ch.push(createClassicTable(['البيان','القيمة'], [
    ['فكرة المشروع', data.projectIdea||''],
    ['اسم مقدم المشروع', data.submitterName||''],
    ['رقم الهاتف', data.submitterPhone||''],
    ['إجمالي التكاليف التأسيسية', data.summary?.foundingTotal||'$0'],
    ['إجمالي الإيرادات المتوقعة (سنوياً)', data.summary?.revenueAnnual||'$0'],
    ['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0'],
    ['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual||'$0'],
    ['الاهتلاك (سنوياً)', data.summary?.depreciation||'$0'],
    ['الربح الصافي (سنوياً)', data.summary?.netProfit||'$0'],
    ['عدد الموظفين في المشروع', data.summary?.employees||'0'],
  ]));
  ch.push(new Paragraph({ spacing: { after: 300 } }));

  // 2. التكاليف التأسيسية
  ch.push(createParagraph("التكاليف التأسيسية", true, 24));
  const foundingRows = data.foundingRows?.map((r,i)=>[i+1, r.cat, r.bayan, r.dep?'✓':'', r.qty, r.notes, fmtMoney(r.price), r.total]) || [];
  const foundingTotal = data.foundingRows?.reduce((sum, r) => sum + (parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0), 0) || 0;
  foundingRows.push(['الإجمالي', '', '', '', '', '', '', fmtMoney(foundingTotal)]);
  ch.push(createClassicTable(['#','الصنف','البيان','الاهتلاك','العدد','ملاحظات','التكلفة للواحدة','التكلفة الإجمالية'], foundingRows));
  ch.push(new Paragraph({ spacing: { after: 300 } }));

  // 3. جدول المنتجات
  ch.push(createParagraph("جدول المنتجات", true, 24));
  const productsRows = [];
  Object.values(data.products||{}).forEach(p => {
    p.components.filter(c=>c).forEach((c, idx) => {
      productsRows.push([idx===0 ? p.name : '', idx===0 ? p.unit : '', c]);
    });
  });
  ch.push(createClassicTable(['البيان','الواحدة','المكونات'], productsRows));
  ch.push(new Paragraph({ spacing: { after: 300 } }));

  // 4. الإيرادات المتوقعة
  ch.push(createParagraph("الإيرادات المتوقعة", true, 24));
  const revRows = [];
  let totalRevPerMonth = Array(12).fill(0);
  Object.entries(data.products||{}).forEach(([pid, p]) => {
      const rev = data.revenueData?.[pid]||{};
      const qRow = [p.name]; 
      const pRow = ['سعر مبيع الواحدة']; 
      const tRow = ['سعر المبيع الإجمالي'];
      for(let i=0; i<12; i++) {
        const q = Number(rev[i]?.qty||0);
        const price = Number(rev[i]?.price||0);
        const t = q * price;
        qRow.push(q);
        pRow.push(price);
        tRow.push(t);
        totalRevPerMonth[i] += t;
      }
      revRows.push(qRow, pRow, tRow);
  });
  revRows.push(['الإجمالي', ...totalRevPerMonth]);
  ch.push(createClassicTable(['البيان', ...MONTHS], revRows));
  ch.push(new Paragraph({ spacing: { after: 300 } }));

  // 5. التكاليف التشغيلية
  ch.push(createParagraph("التكاليف التشغيلية", true, 24));
  const opsRows = [];
  let totalOpsPerMonth = Array(12).fill(0);
  Object.entries(data.products||{}).forEach(([pid, p]) => {
      let productSubtotalPerMonth = Array(12).fill(0);
      p.components.filter(c=>c).forEach((c, ci) => {
          const r = [ci===0 ? p.name : '', ci===0 ? p.unit : '', c];
          for(let i=0; i<12; i++) {
            const unitCost = Number(data.opsData?.[pid]?.[ci]?.[i]||0);
            const qty = Number(data.revenueData?.[pid]?.[i]?.qty||0);
            const v = unitCost * qty;
            r.push(v);
            productSubtotalPerMonth[i] += v;
          }
          opsRows.push(r);
      });
      opsRows.push([`إجمالي ${p.name}`, '', '', ...productSubtotalPerMonth]);
      for(let i=0; i<12; i++) totalOpsPerMonth[i] += productSubtotalPerMonth[i];
  });
  opsRows.push(['الإجمالي', '', '', ...totalOpsPerMonth]);
  ch.push(createClassicTable(['البيان','الواحدة','التفاصيل', ...MONTHS], opsRows));
  ch.push(new Paragraph({ spacing: { after: 300 } }));

  // 6. الموارد البشرية
  ch.push(createParagraph("الموارد البشرية", true, 24));
  const hrRows = data.hrRows?.map((r,i)=>[i+1, r.position, r.type, r.qty, r.reports, fmtMoney(r.salary), r.total]) || [];
  const hrTotal = data.hrRows?.reduce((sum, r) => sum + (parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0), 0) || 0;
  hrRows.push(['الإجمالي', '', '', '', '', '', fmtMoney(hrTotal)]);
  ch.push(createClassicTable(['#','المنصب','النوع','العدد','تابع لـ','الراتب الشهري الفردي','الراتب الشهري الإجمالي'], hrRows));
  ch.push(new Paragraph({ spacing: { after: 300 } }));

  // 7. التكاليف الثابتة
  ch.push(createParagraph("التكاليف الثابتة", true, 24));
  const fixedRows = data.fixedRows?.map((r,i)=>[i+1, r.cat, r.bayan, r.qty, r.notes, fmtMoney(r.price), r.total]) || [];
  const fixedTotal = data.fixedRows?.reduce((sum, r) => sum + (parseFloat(String(r.total||'0').replace(/[^0-9.]/g,''))||0), 0) || 0;
  fixedRows.push(['الإجمالي', '', '', '', '', '', fmtMoney(fixedTotal)]);
  ch.push(createClassicTable(['#','الصنف','البيان','العدد','ملاحظات','التكلفة الشهرية للواحدة','التكلفة الشهرية الإجمالية'], fixedRows));
  ch.push(new Paragraph({ spacing: { after: 300 } }));

  // 8. الاهتلاك
  ch.push(createParagraph("الاهتلاك", true, 24));
  const dRows = []; let dIdx=1; let depTotal = 0;
  data.foundingRows?.forEach((fr, i) => {
      if(fr.dep) {
          const pct = Number(data.depData?.[i+1]||0); 
          const depU = Number(fr.price||0)*(pct/100); 
          const depT = depU*Number(fr.qty||0);
          depTotal += depT;
          dRows.push([dIdx++, fr.cat, fr.bayan, pct+'%', Number(fr.qty||0), fr.notes||'', fmtMoney(depU), fmtMoney(depT)]);
      }
  });
  dRows.push(['إجمالي قيمة الاهتلاك', '', '', '', '', '', '', fmtMoney(depTotal)]);
  ch.push(createClassicTable(['#','الصنف','البيان','نسبة الاهتلاك','العدد','ملاحظات','قيمة الاهتلاك للواحدة','قيمة الاهتلاك الإجمالية'], dRows));

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial" }, paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT } } } },
    sections: [{
      properties: { page: { size: { orientation: PageOrientation.LANDSCAPE, width: 16838, height: 11906 }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
      children: ch
    }]
  });
  return Packer.toBuffer(doc);
}


// ════════════════════════════════════════════════════════════
//  EXCEL GENERATOR (بقي كما هو للحفاظ على العملية)
// ════════════════════════════════════════════════════════════
async function generateExcel(data) {
  const wb = new ExcelJS.Workbook();
  const s0 = wb.addWorksheet('البيانات', { views: [{ rightToLeft:true }] });
  s0.columns = [{ width:30 }, { width:30 }];
  s0.addRow(['اسم مقدم المشروع', data.submitterName||'']);
  s0.addRow(['رقم الهاتف', data.submitterPhone||'']);
  s0.addRow(['فكرة المشروع', data.projectIdea||'']);
  s0.addRow(['إجمالي التكاليف التأسيسية', data.summary?.foundingTotal||'$0']);
  s0.addRow(['إجمالي الإيرادات المتوقعة (سنوياً)', data.summary?.revenueAnnual||'$0']);
  s0.addRow(['إجمالي التكاليف التشغيلية (سنوياً)', data.summary?.opsAnnual||'$0']);
  s0.addRow(['إجمالي التكاليف الثابتة (سنوياً)', data.summary?.fixedAnnual||'$0']);
  s0.addRow(['الاهتلاك (سنوياً)', data.summary?.depreciation||'$0']);
  s0.addRow(['الربح الصافي (سنوياً)', data.summary?.netProfit||'$0']);
  s0.addRow(['عدد الموظفين في المشروع', data.summary?.employees||'0']);
  return wb.xlsx.writeBuffer();
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
    return res.json({ success:true, id });
  } catch(err) {
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
