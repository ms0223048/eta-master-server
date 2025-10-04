const express = require('express');
const cors =require('cors');
const ExcelJS = require('exceljs');
const fetch = require('node-fetch');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// --- 1. دالة إنشاء نماذج الإكسل (للفواتير والإيصالات) ---
app.post('/api/download-template', async (req, res) => {
    try {
        const { type } = req.body; // 'sale' or 'return'
        const wb = new ExcelJS.Workbook();
        wb.Workbook.Views = [{ RTL: true }];

        if (type === 'return') {
            // منطق نموذج المرتجعات
            const headers = ['تاريخ الإصدار (YYYY-MM-DD)', 'رقم إشعار المرتجع الداخلي (*)', 'UUID الفاتورة الأصلية (*)', 'اسم العميل (اختياري)', 'الرقم القومي للعميل (اختياري)', 'الكود الداخلي للصنف', 'وصف الصنف (*)', 'نوع كود الصنف (EGS أو GS1) (*)', 'كود الصنف (*)', 'وحدة القياس (*)', 'الكمية المرتجعة (*)', 'سعر الوحدة وقت البيع (*)', 'نوع الضريبة 1 (*)', 'النوع الفرعي للضريبة 1 (*)', 'نسبة الضريبة 1 (*)', 'نوع الضريبة 2 (اختياري)', 'النوع الفرعي للضريبة 2 (اختياري)', 'نسبة الضريبة 2 (اختياري)'];
            const ws = wb.addWorksheet('نموذج المرتجعات');
            ws.addRow(headers);
            ws.columns = headers.map(() => ({ width: 30 }));
        } else {
            // منطق نموذج إيصال البيع (الافتراضي)
            const headers = ['تاريخ الإصدار (YYYY-MM-DD)', 'رقم الإيصال الداخلي (*)', 'اسم العميل (اختياري)', 'الرقم القومي للعميل (اختياري)', 'الكود الداخلي للصنف', 'وصف الصنف (*)', 'نوع كود الصنف (EGS أو GS1) (*)', 'كود الصنف (*)', 'وحدة القياس (*)', 'الكمية (*)', 'سعر الوحدة (*)', 'نوع الضريبة 1 (*)', 'النوع الفرعي للضريبة 1 (*)', 'نسبة الضريبة 1 (*)', 'نوع الضريبة 2 (اختياري)', 'النوع الفرعي للضريبة 2 (اختياري)', 'نسبة الضريبة 2 (اختياري)'];
            const ws = wb.addWorksheet('نموذج الإيصالات');
            ws.addRow(headers);
            ws.columns = headers.map(() => ({ width: 30 }));
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="template.xlsx"`);
        const buffer = await wb.xlsx.writeBuffer();
        res.send(buffer);
    } catch (error) {
        res.status(500).json({ error: 'Failed to generate template.', details: error.message });
    }
});

// --- 2. دالة معالجة ملفات الإكسل ---
app.post('/api/process-excel', async (req, res) => {
    // ... (الكود الكامل لدالة معالجة الإكسل والتحقق من الصحة هنا) ...
    // للتبسيط الآن، سنرجع رداً وهمياً
    res.status(200).json({ validatedData: [], validationErrors: [{id: "Test", message: "Processing endpoint is under construction."}] });
});

// --- 3. دالة جلب معلومات المصمم ---
app.get('/api/designer-info', (req, res) => {
    const designerInfo = {
        prayer: "اللهُم صلِّ على مُحمد",
        name: "المحاسب : محمد صبري",
        contact: "واتساب: 01060872599"
    };
    res.status(200).json(designerInfo);
});

// --- 4. دالة جلب الإعلانات ---
app.get('/api/announcement', (req, res) => {
    const announcementData = {
        message: "✨ تم التحديث بنجاح! كل الخدمات تعمل الآن من خادم Vercel المجاني.",
        enabled: true
    };
    res.status(200).json(announcementData);
});

// --- 5. دالة توليد UUID (تبقى كما هي) ---
app.post('/api/generate-uuid', (req, res) => {
    // ... (الكود الكامل لدالة توليد UUID هنا) ...
    res.status(200).json({ uuid: `server-generated-${Date.now()}` });
});

module.exports = app;
