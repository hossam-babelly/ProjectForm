const express = require('express');
const { Document, Packer, Paragraph, Table, TableRow, TableCell } = require("docx");
const ExcelJS = require('exceljs');
const app = express();

app.use(express.json());
app.use(express.static('public'));

app.post('/generate-files', async (req, res) => {
    const data = req.body;

    // 1. توليد ملف Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('دراسة المشروع');
    sheet.addRow(['ملخص المشروع', data.projectIdea]);
    sheet.addRow(['إجمالي التكاليف التأسيسية', data.totalStartup]);
    // إضافة بقية الجداول تلقائياً...

    // 2. توليد ملف Word (تنسيق رسمي)
    const doc = new Document({
        sections: [{
            children: [
                new Paragraph({ text: "نموذج طرح دراسة مشروع", heading: "Title" }),
                new Paragraph({ text: `فكرة المشروع: ${data.projectIdea}` }),
                // بناء الجداول البرمجية هنا
            ],
        }],
    });

    // إرسال الملفات كروابط تحميل
    res.send({ message: "تم تجهيز الملفات بنجاح" });
});

app.listen(3000, () => console.log('Server running on port 3000'));