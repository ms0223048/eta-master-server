// -----------------------------------------------------------------------------
// |                                                                           |
// |                    ملف الخادم الموحد لـ Vercel (النسخة الكاملة)             |
// |                                                                           |
// |     تم تجميع كل الدوال والمسارات في هذا الملف الواحد.                      |
// |     الاعتماديات مجمعة في ملف package.json المرفق.                         |
// |                                                                           |
// -----------------------------------------------------------------------------

// 1. استدعاء جميع المكتبات المطلوبة من كل الملفات
const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const axios = require('axios');
const fetch = require('node-fetch');
const crypto = require('crypto');
const JSZip = require('jszip');

// 2. إعداد تطبيق Express الرئيسي
const app = express();

// 3. إعدادات CORS المتقدمة للسماح بالوصول من كل المصادر
const corsOptions = {
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
};
app.use(cors(corsOptions));
app.options('*', cors(corsOptions)); // للرد على الطلبات الاستباقية (preflight)

// 4. إعدادات إضافية للتطبيق
app.use(express.json({ limit: '50mb' })); // لتمكين قراءة JSON بحجم كبير

// -----------------------------------------------------------------------------
// |                     ✅ بداية المسارات (Routes)                              |
// -----------------------------------------------------------------------------

// المسار الرئيسي: يعرض رسالة ترحيب
app.get('/', (req, res) => {
    res.status(200).json({ message: "أهلاً بك في الخادم الموحد. كل الخدمات تعمل على المسارات المخصصة لها." });
});

// --- المسار 1: جلب الإعلانات (كان الملف رقم 1) ---
app.get('/get-announcement', (req, res) => {
    const announcementData = {
      "message": "﴿ إِنَّ اللَّهَ وَمَلَائِكَتَهُ يُصَلُّونَ عَلَى النَّبِيِّ ۚ يَا أَيُّهَا الَّذِينَ آمَنُوا صَلُّوا عَلَيْهِ وَسَلِّمُوا تَسْلِيمًا﴾ ",
      "enabled": true
    };
    res.status(200).json(announcementData);
});

// --- المسار 2: إنشاء قالب إكسل للفواتير (كان الملف رقم 2) ---
app.post('/generate-invoice-template', async (req, res) => {
    try {
        const { headers, data } = req.body;
        if (!headers || !data) return res.status(400).send('Missing headers or data');

        const wb = new ExcelJS.Workbook();
        const mainSheetName = "Invoices";
        const listsSheetName = "Lists";

        const listsData = {
            receiverTypes: [{ name: "أعمال", code: "B" }, { name: "شخصي", code: "P" }, { name: "أجنبي", code: "F" }],
            codeTypes: ["EGS", "GS1"],
            mainTaxTypes: [
                { name: "ضريبة القيمة المضافة", code: "T1" }, { name: "ضريبة الجدول (نسبية)", code: "T2" },
                { name: "ضريبة الجدول (النوعية)", code: "T3" }, { name: "الخصم تحت حساب الضريبة", code: "T4" },
                { name: "ضريبة الدمغة (نسبية)", code: "T5" }, { name: "ضريبة الدمغة (قطعية)", code: "T6" },
                { name: "ضريبة الملاهي", code: "T7" }, { name: "رسم تنمية الموارد", code: "T8" },
                { name: "رسم خدمة", code: "T9" }, { name: "رسم المحليات", code: "T10" },
                { name: "رسم التأمين الصحي", code: "T11" }, { name: "رسوم أخرى", code: "T12" }
            ],
            taxSubtypes: [
                { name: "تصدير للخارج", code: "V001", ref: "T1" }, { name: "تصدير مناطق حرة", code: "V002", ref: "T1" },
                { name: "سلعة أو خدمة معفاة", code: "V003", ref: "T1" }, { name: "سلعة أو خدمة غير خاضعة", code: "V004", ref: "T1" },
                { name: "إعفاءات دبلوماسيين", code: "V005", ref: "T1" }, { name: "إعفاءات دفاع وأمن قومي", code: "V006", ref: "T1" },
                { name: "إعفاءات اتفاقيات", code: "V007", ref: "T1" }, { name: "إعفاءات خاصة وأخرى", code: "V008", ref: "T1" },
                { name: "سلع عامة (14%)", code: "V009", ref: "T1" }, { name: "نسب ضريبة أخرى", code: "V010", ref: "T1" },
                { name: "ضريبة جدول (نسبية)", code: "Tbl01", ref: "T2" }, { name: "ضريبة جدول (نوعية)", code: "Tbl02", ref: "T3" },
                { name: "المقاولات", code: "W001", ref: "T4" }, { name: "التوريدات", code: "W002", ref: "T4" },
                { name: "المشتريات", code: "W003", ref: "T4" }, { name: "الخدمات", code: "W004", ref: "T4" },
                { name: "أتعاب مهنية", code: "W010", ref: "T4" }, { name: "دمغة (نسبية)", code: "ST01", ref: "T5" },
                { name: "دمغة (قطعية)", code: "ST02", ref: "T6" }, { name: "ملاهي (نسبة)", code: "Ent01", ref: "T7" },
                { name: "تنمية موارد (نسبة)", code: "RD01", ref: "T8" }, { name: "رسم خدمة (نسبة)", code: "SC01", ref: "T9" },
                { name: "محليات (نسبة)", code: "Mn01", ref: "T10" }, { name: "تأمين صحي (نسبة)", code: "MI01", ref: "T11" },
                { name: "رسوم أخرى", code: "OF01", ref: "T12" }
            ],
            unitTypes: [
                { name: "each (ST) ( ST )", code: "EA" }, { name: "ملليفولت", code: "2Z" }, { name: "مللي أمبير", code: "4K" },
                { name: "ميكروفاراد", code: "4O" }, { name: "جيجا أوم", code: "A87" }, { name: "جرام/متر مكعب", code: "A93" },
                { name: "جرام/سم مكعب", code: "A94" }, { name: "أمبير", code: "AMP" }, { name: "سنة", code: "ANN" },
                { name: "كيلو أمبير", code: "B22" }, { name: "كيلو أوم", code: "B49" }, { name: "ميج أوم", code: "B75" },
                { name: "ميجافولت", code: "B78" }, { name: "ميكرو أمبير", code: "B84" }, { name: "بار", code: "BAR" },
                { name: "برميل", code: "BBL" }, { name: "حقيبة", code: "BG" }, { name: "زجاجة", code: "BO" },
                { name: "صندوق", code: "BOX" }, { name: "مللي فاراد", code: "C10" }, { name: "نانو أمبير", code: "C39" },
                { name: "نانو فاراد", code: "C41" }, { name: "نانومتر", code: "C45" }, { name: "وحدة نشاط", code: "C62" },
                { name: "علبة", code: "CA" }, { name: "سم مربع", code: "CMK" }, { name: "سم مكعب", code: "CMQ" },
                { name: "سنتيمتر", code: "CMT" }, { name: "كرتونة", code: "CS" }, { name: "كرتون", code: "CT" },
                { name: "سنتيلتر", code: "CTL" }, { name: "سيمنز/متر", code: "D10" }, { name: "تسلا", code: "D33" },
                { name: "طن/متر مكعب", code: "D41" }, { name: "يوم", code: "DAY" }, { name: "ديسيمتر", code: "DMT" },
                { name: "أسطوانة", code: "DRM" }, { name: "فاراد", code: "FAR" }, { name: "قدم", code: "FOT" },
                { name: "قدم مربع", code: "FTK" }, { name: "قدم مكعب", code: "FTQ" }, { name: "ميكروسيمنز/سم", code: "G42" },
                { name: "جرام/لتر", code: "GL" }, { name: "جالون", code: "GLL" }, { name: "جرام/متر مربع", code: "GM" },
                { name: "جالون/ألف", code: "GPT" }, { name: "جرام", code: "GRM" }, { name: "ملليجرام/سم مربع", code: "H63" },
                { name: "قوة حصان هيدروليكي", code: "HHP" }, { name: "هكتولتر", code: "HLT" }, { name: "هرتز", code: "HTZ" },
                { name: "ساعة", code: "HUR" }, { name: "عدد الأشخاص", code: "IE" }, { name: "بوصة", code: "INH" },
                { name: "بوصة مربعة", code: "INK" }, { name: "وظيفة", code: "JOB" }, { name: "كيلوجرام", code: "KGM" },
                { name: "كيلوهرتز", code: "KHZ" }, { name: "كم/ساعة", code: "KMH" }, { name: "كم مربع", code: "KMK" },
                { name: "كجم/متر مكعب", code: "KMQ" }, { name: "كيلومتر", code: "KMT" }, { name: "كجم/متر مربع", code: "KSM" },
                { name: "كيلوفولت", code: "KVT" }, { name: "كيلووات", code: "KWT" }, { name: "رطل", code: "LB" },
                { name: "لتر", code: "LTR" }, { name: "مستوى", code: "LVL" }, { name: "متر", code: "M" },
                { name: "رجل", code: "MAN" }, { name: "ميجاوات", code: "MAW" }, { name: "ملليجرام", code: "MGM" },
                { name: "ميجاهرتز", code: "MHZ" }, { name: "دقيقة", code: "MIN" }, { name: "مم مربع", code: "MMK" },
                { name: "مم مكعب", code: "MMQ" }, { name: "ملليمتر", code: "MMT" }, { name: "شهر", code: "MON" },
                { name: "متر مربع", code: "MTK" }, { name: "متر مكعب", code: "MTQ" }, { name: "أوم", code: "OHM" },
                { name: "أونصة", code: "ONZ" }, { name: "باسكال", code: "PAL" }, { name: "بالتة", code: "PF" },
                { name: "عبوة", code: "PK" }, { name: "كيس", code: "SK" }, { name: "ميل", code: "SMI" },
                { name: "طن قصير", code: "ST" }, { name: "طن", code: "TNE" }, { name: "طن متري", code: "TON" },
                { name: "فولت", code: "VLT" }, { name: "أسبوع", code: "WEE" }, { name: "وات", code: "WTT" },
                { name: "متر/ساعة", code: "X03" }, { name: "ياردة مكعبة", code: "YDQ" }, { name: "ياردة", code: "YRD" },
                { name: "عدد العبوات", code: "NMP" }, { name: "لوح", code: "ST" }, { name: "قدم مكعب قياسي", code: "5I" },
                { name: "أمبير لكل متر", code: "AE" }, { name: "برميل إمبراطوري", code: "B4" }, { name: "صندوق أساسي", code: "BB" },
                { name: "لوح خشبي", code: "BD" }, { name: "حزمة", code: "BE" }, { name: "سلة", code: "BK" },
                { name: "بالة", code: "BL" }, { name: "حاوية", code: "CH" }, { name: "صندوق شحن", code: "CR" },
                { name: "ديكار", code: "DAA" }, { name: "ديسيتون", code: "DTN" }, { name: "دستة", code: "DZN" },
                { name: "رطل/قدم مربع", code: "FP" }, { name: "هكتومتر", code: "HMT" }, { name: "بوصة مكعبة", code: "INQ" },
                { name: "برميل صغير", code: "KG" }, { name: "كيلومتر", code: "KTM" }, { name: "دفعة", code: "LO" },
                { name: "ملليلتر", code: "MLT" }, { name: "بساط", code: "MT" }, { name: "ملليجرام/كيلوجرام", code: "NA" },
                { name: "عدد المواد", code: "NAR" }, { name: "سيارة", code: "NC" }, { name: "لتر صافي", code: "NE" },
                { name: "عدد الطرود", code: "NPL" }, { name: "مركبة", code: "NV" }, { name: "رزمة", code: "PA" },
                { name: "طبق", code: "PG" }, { name: "دلو", code: "PL" }, { name: "زوج", code: "PR" },
                { name: "باينت", code: "PT" }, { name: "بكرة", code: "RL" }, { name: "لفة", code: "RO" },
                { name: "طقم", code: "SET" }, { name: "سيجارة", code: "STK" }, { name: "ألف قطعة", code: "T3" },
                { name: "حمولة شاحنة", code: "TC" }, { name: "خزان مستطيل", code: "TK" }, { name: "صفيح", code: "TN" },
                { name: "عشرة آلاف سيجارة", code: "TTS" }, { name: "منفذ اتصالات", code: "UC" }, { name: "قارورة", code: "VI" },
                { name: "بالجملة", code: "VQ" }, { name: "ياردة مربعة", code: "YDK" }, { name: "برميل خشبي", code: "Z3" }
            ],
            countryCodes: [
                { name: "مصر", code: "EG" }, { name: "الإمارات", code: "AE" }, { name: "السعودية", code: "SA" },
                { name: "الكويت", code: "KW" }, { name: "قطر", code: "QA" }, { name: "البحرين", code: "BH" },
                { name: "عمان", code: "OM" }, { name: "الأردن", code: "JO" }, { name: "لبنان", code: "LB" },
                { name: "سوريا", code: "SY" }, { name: "العراق", code: "IQ" }, { name: "اليمن", code: "YE" },
                { name: "السودان", code: "SD" }, { name: "ليبيا", code: "LY" }, { name: "تونس", code: "TN" },
                { name: "الجزائر", code: "DZ" }, { name: "المغرب", code: "MA" }, { name: "أمريكا", code: "US" },
                { name: "بريطانيا", code: "GB" }, { name: "ألمانيا", code: "DE" }, { name: "فرنسا", code: "FR" },
                { name: "إيطاليا", code: "IT" }, { name: "الصين", code: "CN" }, { name: "اليابان", code: "JP" },
                { name: "الهند", code: "IN" }, { name: "تركيا", code: "TR" }
            ],
            currencyCodes: [
                { name: "جنيه مصري", code: "EGP" }, { name: "دولار أمريكي", code: "USD" }, { name: "يورو", code: "EUR" },
                { name: "جنيه استرليني", code: "GBP" }, { name: "ريال سعودي", code: "SAR" }, { name: "درهم إماراتي", code: "AED" }
            ]
        };

        const ws = wb.addWorksheet(mainSheetName);
        ws.views = [{ rightToLeft: true }];
        ws.addRows([headers, ...data]);
        ws.columns = headers.map(() => ({ width: 30 }));

        const lists_ws = wb.addWorksheet(listsSheetName);
        lists_ws.state = 'hidden';

        lists_ws.getColumn('A').values = ['ReceiverTypeNames', ...listsData.receiverTypes.map(t => t.name)];
        lists_ws.getColumn('B').values = ['ReceiverTypeCodes', ...listsData.receiverTypes.map(t => t.code)];
        lists_ws.getColumn('C').values = ['CodeTypes', ...listsData.codeTypes];
        lists_ws.getColumn('D').values = ['UnitTypeNames', ...listsData.unitTypes.map(t => t.name)];
        lists_ws.getColumn('E').values = ['UnitTypeCodes', ...listsData.unitTypes.map(t => t.code)];
        lists_ws.getColumn('F').values = ['CountryNames', ...listsData.countryCodes.map(t => t.name)];
        lists_ws.getColumn('G').values = ['CountryCodes', ...listsData.countryCodes.map(t => t.code)];
        lists_ws.getColumn('H').values = ['CurrencyNames', ...listsData.currencyCodes.map(t => t.name)];
        lists_ws.getColumn('I').values = ['CurrencyCodes', ...listsData.currencyCodes.map(t => t.code)];
        lists_ws.getColumn('J').values = ['TaxTypeNames', ...listsData.mainTaxTypes.map(t => t.name)];
        lists_ws.getColumn('K').values = ['TaxTypeCodes', ...listsData.mainTaxTypes.map(t => t.code)];

        listsData.mainTaxTypes.forEach((taxType, index) => {
            const col = lists_ws.getColumn(12 + (index * 2));
            const subtypes = listsData.taxSubtypes.filter(st => st.ref === taxType.code);
            col.values = [taxType.code, ...subtypes.map(st => st.name)];
            lists_ws.getColumn(13 + (index * 2)).values = [taxType.code, ...subtypes.map(st => st.code)];
        });

        wb.definedNames.add(`'${listsSheetName}'!$A$2:$A$${listsData.receiverTypes.length + 1}`, 'ReceiverTypeNames');
        wb.definedNames.add(`'${listsSheetName}'!$C$2:$C$${listsData.codeTypes.length + 1}`, 'CodeTypes');
        wb.definedNames.add(`'${listsSheetName}'!$D$2:$D$${listsData.unitTypes.length + 1}`, 'UnitTypeNames');
        wb.definedNames.add(`'${listsSheetName}'!$F$2:$F$${listsData.countryCodes.length + 1}`, 'CountryNames');
        wb.definedNames.add(`'${listsSheetName}'!$H$2:$H$${listsData.currencyCodes.length + 1}`, 'CurrencyNames');
        wb.definedNames.add(`'${listsSheetName}'!$J$2:$J$${listsData.mainTaxTypes.length + 1}`, 'TaxTypeNames');

        listsData.mainTaxTypes.forEach((taxType, index) => {
            const colLetter = String.fromCharCode('A'.charCodeAt(0) + 11 + (index * 2));
            const lastRow = (listsData.taxSubtypes.filter(st => st.ref === taxType.code).length || 1) + 1;
            wb.definedNames.add(`'${listsSheetName}'!$${colLetter}$2:$${colLetter}$${lastRow}`, `Sub_${taxType.code}`);
        });

        for (let i = 2; i <= data.length + 500; i++) {
            ws.getCell(`D${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=ReceiverTypeNames'] };
            ws.getCell(`E${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=CountryNames'] };
            ws.getCell(`K${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=CodeTypes'] };
            ws.getCell(`N${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=UnitTypeNames'] };
            ws.getCell(`Q${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=CurrencyNames'] };
            
            const taxCols = ['U', 'X', 'AA'];
            const subTypeCols = ['V', 'Y', 'AB'];
            taxCols.forEach((col, index) => {
                ws.getCell(`${col}${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=TaxTypeNames'] };
                const formula = `=INDIRECT("Sub_"&VLOOKUP(${col}${i},${listsSheetName}!$J$2:$K$20,2,FALSE))`;
                ws.getCell(`${subTypeCols[index]}${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [formula] };
            });
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="Template_Full_Arabic_v2.xlsx"');
        await wb.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error("Error generating invoice Excel file:", error);
        res.status(500).send("Error generating file.");
    }
});

// --- المسار 3: إنشاء قالب إكسل للإيصالات (كان الملف رقم 3) ---
app.post('/generate-receipt-template', async (req, res) => {
    try {
        const { headers, data } = req.body;
        if (!headers || !data) return res.status(400).send('Missing headers or data');

        const wb = new ExcelJS.Workbook();
        const mainSheetName = "Receipts";
        const listsSheetName = "Lists";

        const listsData = {
            codeTypes: ["EGS", "GS1"],
            mainTaxTypes: [
                { name: "ضريبة القيمة المضافة", code: "T1" }, { name: "ضريبة الجدول (نسبية)", code: "T2" },
                { name: "ضريبة الجدول (النوعية)", code: "T3" }, { name: "الخصم تحت حساب الضريبة", code: "T4" },
                { name: "ضريبة الدمغة (نسبية)", code: "T5" }, { name: "ضريبة الدمغة (قطعية)", code: "T6" },
                { name: "ضريبة الملاهي", code: "T7" }, { name: "رسم تنمية الموارد", code: "T8" },
                { name: "رسم خدمة", code: "T9" }, { name: "رسم المحليات", code: "T10" },
                { name: "رسم التأمين الصحي", code: "T11" }, { name: "رسوم أخرى", code: "T12" }
            ],
            taxSubtypes: [
                { name: "تصدير للخارج", code: "V001", ref: "T1" }, { name: "تصدير مناطق حرة", code: "V002", ref: "T1" },
                { name: "سلعة أو خدمة معفاة", code: "V003", ref: "T1" }, { name: "سلعة أو خدمة غير خاضعة", code: "V004", ref: "T1" },
                { name: "إعفاءات دبلوماسيين", code: "V005", ref: "T1" }, { name: "إعفاءات دفاع وأمن قومي", code: "V006", ref: "T1" },
                { name: "إعفاءات اتفاقيات", code: "V007", ref: "T1" }, { name: "إعفاءات خاصة وأخرى", code: "V008", ref: "T1" },
                { name: "سلع عامة (14%)", code: "V009", ref: "T1" }, { name: "نسب ضريبة أخرى", code: "V010", ref: "T1" },
                { name: "ضريبة جدول (نسبية)", code: "Tbl01", ref: "T2" }, { name: "ضريبة جدول (نوعية)", code: "Tbl02", ref: "T3" },
                { name: "المقاولات", code: "W001", ref: "T4" }, { name: "التوريدات", code: "W002", ref: "T4" },
                { name: "المشتريات", code: "W003", ref: "T4" }, { name: "الخدمات", code: "W004", ref: "T4" },
                { name: "أتعاب مهنية", code: "W010", ref: "T4" }, { name: "دمغة (نسبية)", code: "ST01", ref: "T5" },
                { name: "دمغة (قطعية)", code: "ST02", ref: "T6" }, { name: "ملاهي (نسبة)", code: "Ent01", ref: "T7" },
                { name: "تنمية موارد (نسبة)", code: "RD01", ref: "T8" }, { name: "رسم خدمة (نسبة)", code: "SC01", ref: "T9" },
                { name: "محليات (نسبة)", code: "Mn01", ref: "T10" }, { name: "تأمين صحي (نسبة)", code: "MI01", ref: "T11" },
                { name: "رسوم أخرى", code: "OF01", ref: "T12" }
            ],
            unitTypes: [
                { name: "قطعة", code: "EA" }, { name: "كرتونة", code: "CS" }, { name: "حقيبة", code: "BG" },
                { name: "عبوة", code: "PK" }, { name: "صندوق", code: "BX" }, { name: "كيلو جرام", code: "KGM" },
                { name: "متر", code: "M" }, { name: "لتر", code: "LTR" }, { name: "طن", code: "TNE" }
            ]
        };

        const ws = wb.addWorksheet(mainSheetName);
        ws.views = [{ rightToLeft: true }];
        ws.addRows([headers, ...data]);
        ws.columns = headers.map(() => ({ width: 30 }));

        const lists_ws = wb.addWorksheet(listsSheetName);
        lists_ws.state = 'hidden';

        lists_ws.getColumn('A').values = ['CodeTypes', ...listsData.codeTypes];
        lists_ws.getColumn('B').values = ['UnitTypeNames', ...listsData.unitTypes.map(t => t.name)];
        lists_ws.getColumn('C').values = ['UnitTypeCodes', ...listsData.unitTypes.map(t => t.code)];
        lists_ws.getColumn('D').values = ['TaxTypeNames', ...listsData.mainTaxTypes.map(t => t.name)];
        lists_ws.getColumn('E').values = ['TaxTypeCodes', ...listsData.mainTaxTypes.map(t => t.code)];

        listsData.mainTaxTypes.forEach((taxType, index) => {
            const col = lists_ws.getColumn(6 + (index * 2));
            const subtypes = listsData.taxSubtypes.filter(st => st.ref === taxType.code);
            col.values = [taxType.code, ...subtypes.map(st => st.name)];
            lists_ws.getColumn(7 + (index * 2)).values = [taxType.code, ...subtypes.map(st => st.code)];
        });

        wb.definedNames.add(`'${listsSheetName}'!$A$2:$A$${listsData.codeTypes.length + 1}`, 'CodeTypes');
        wb.definedNames.add(`'${listsSheetName}'!$B$2:$B$${listsData.unitTypes.length + 1}`, 'UnitTypeNames');
        wb.definedNames.add(`'${listsSheetName}'!$D$2:$D$${listsData.mainTaxTypes.length + 1}`, 'TaxTypeNames');

        listsData.mainTaxTypes.forEach((taxType, index) => {
            const colLetter = String.fromCharCode('A'.charCodeAt(0) + 5 + (index * 2));
            const lastRow = (listsData.taxSubtypes.filter(st => st.ref === taxType.code).length || 1) + 1;
            wb.definedNames.add(`'${listsSheetName}'!$${colLetter}$2:$${colLetter}$${lastRow}`, `Sub_${taxType.code}`);
        });

        for (let i = 2; i <= data.length + 500; i++) {
            ws.getCell(`G${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=CodeTypes'] };
            ws.getCell(`I${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=UnitTypeNames'] };
            
            const taxCols = ['L', 'O'];
            const subTypeCols = ['M', 'P'];
            taxCols.forEach((col, index) => {
                ws.getCell(`${col}${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=TaxTypeNames'] };
                const formula = `=INDIRECT("Sub_"&VLOOKUP(${col}${i},${listsSheetName}!$D$2:$E$20,2,FALSE))`;
                ws.getCell(`${subTypeCols[index]}${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [formula] };
            });
        }
// ... تتمة الكود من النقطة التي توقف عندها

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="Template_Receipts_Arabic.xlsx"');
        await wb.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error("Error generating receipt Excel file:", error);
        res.status(500).send("Error generating file.");
    }
});

// --- المسار 4: تصدير بيانات الفواتير إلى إكسل (كان الملف رقم 4) ---
app.post('/export-eta-data', async (req, res) => {
    try {
        const { detailData, summaryData, analytics, allDetails } = req.body;

        if (!summaryData && !detailData) {
            return res.status(400).send({ error: 'No data provided to generate the file.' });
        }

        const wb = new ExcelJS.Workbook();
        
        const headerStyle = {
            font: { name: 'Calibri', bold: true, color: { argb: 'FFFFFFFF' }, size: 12 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF44546A' } },
            alignment: { horizontal: 'center', vertical: 'middle' }
        };

        const processRowForExcel = (row) => {
            if (!row) return [];
            return row.map(cellValue => {
                if (cellValue === null || cellValue === undefined || String(cellValue).trim() === "") {
                    return null;
                }
                if (typeof cellValue === 'string' && cellValue.match(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
                    return new Date(cellValue);
                }
                if (typeof cellValue === 'string' && !isNaN(cellValue) && (cellValue.length >= 9 || cellValue.startsWith('0'))) {
                    return String(cellValue);
                }
                if (!isNaN(cellValue) && !isNaN(parseFloat(cellValue))) {
                    if (String(cellValue).trim() === '') return null;
                    return parseFloat(cellValue);
                }
                return String(cellValue);
            });
        };

        if (summaryData && summaryData.length > 1) {
            const summaryWS = wb.addWorksheet("البيانات الإجمالية", {
                views: [{ rightToLeft: true }]
            });
            
            const summaryHeaders = summaryData[0];
            const finalHeaders = ["مسلسل", "التفاصيل", ...summaryHeaders, "عرض الفاتورة أونلاين"];
            
            const headerRow = summaryWS.addRow(finalHeaders);
            headerRow.height = 22;
            headerRow.eachCell((cell) => {
                cell.style = headerStyle;
            });

            const originalBody = summaryData.slice(1);
            originalBody.forEach((row, rowIndex) => {
                const processedRow = processRowForExcel(row);

                const uuid = allDetails[rowIndex]?.uuid || "";
                const invoiceUrl = `https://invoicing.eta.gov.eg/documents/${uuid}`;
                const detailSheetRow = detailData && detailData.length > 1 ? rowIndex + 2 : 1;

                const finalRowData = [
                    rowIndex + 1,
                    { text: 'عرض التفاصيل', hyperlink: `#'البيانات التفصيلية'!A${detailSheetRow}` },
                    ...processedRow,
                    { text: 'اضغط هنا', hyperlink: invoiceUrl }
                ];
                const addedRow = summaryWS.addRow(finalRowData  );

                addedRow.eachCell((cell) => {
                    if (cell.value instanceof Date) {
                        cell.numFmt = 'dd/mm/yyyy';
                    }
                });
                
                addedRow.getCell(2).font = { color: { argb: 'FF0000FF' }, underline: true };
                addedRow.getCell(finalRowData.length).font = { color: { argb: 'FF0000FF' }, underline: true };
            });
            
            summaryWS.columns = finalHeaders.map(() => ({ width: 22 }));
            summaryWS.getColumn(1).width = 8;
            summaryWS.getColumn(2).width = 15;

            summaryWS.autoFilter = { from: 'A1', to: { row: 1, column: finalHeaders.length } };
        }

        if (detailData && detailData.length > 1) {
            const detailWS = wb.addWorksheet("البيانات التفصيلية", {
                views: [{ rightToLeft: true }]
            });

            const detailHeaderRow = detailWS.addRow(detailData[0]);
            detailHeaderRow.height = 22;
            detailHeaderRow.eachCell((cell) => {
                cell.style = headerStyle;
            });

            const detailBody = detailData.slice(1);
            detailBody.forEach(row => {
                const processedRow = processRowForExcel(row);
                const addedRow = detailWS.addRow(processedRow);
                addedRow.eachCell((cell) => {
                    if (cell.value instanceof Date) {
                        cell.numFmt = 'dd/mm/yyyy';
                    }
                });
            });

            detailWS.columns = detailData[0].map(() => ({ width: 22 }));
            detailWS.autoFilter = { from: 'A1', to: { row: 1, column: detailData[0].length } };
        }
        
        if (analytics && analytics.salesByCustomer && analytics.purchasesBySupplier && analytics.totalTaxes) {
            const dashboardWS = wb.addWorksheet("ملخص تحليلي", {
                views: [{ rightToLeft: true }]
            });
            const formatNumber = (num) => parseFloat((num || 0).toFixed(2));
            
            dashboardWS.addRow(["--- تحليل المبيعات (الفواتير الصادرة) ---"]);
            dashboardWS.addRow(["اسم العميل", "رقم التسجيل", "إجمالي القيمة"]);
            Object.entries(analytics.salesByCustomer).sort(([,a],[,b]) => b.total - a.total).forEach(([name, data]) => {
                dashboardWS.addRow([name, data.id ? String(data.id) : '', formatNumber(data.total)]);
            });
            
            dashboardWS.addRow([]);
            
            dashboardWS.addRow(["--- تحليل المشتريات (الفواتير المستلمة) ---"]);
            dashboardWS.addRow(["اسم المورد", "رقم التسجيل", "إجمالي القيمة"]);
            Object.entries(analytics.purchasesBySupplier).sort(([,a],[,b]) => b.total - a.total).forEach(([name, data]) => {
                dashboardWS.addRow([name, data.id ? String(data.id) : '', formatNumber(data.total)]);
            });

            dashboardWS.addRow([]);

            const netVat = analytics.totalTaxes.salesVat - analytics.totalTaxes.purchasesVat;
            dashboardWS.addRow(["--- ملخص الضرائب العام ---"]);
            dashboardWS.addRow(["البيان", "القيمة"]);
            dashboardWS.addRow(["إجمالي القيمة قبل الضريبة (مبيعات)", formatNumber(analytics.totalTaxes.salesSubTotal)]);
            dashboardWS.addRow(["إجمالي القيمة قبل الضريبة (مشتريات)", formatNumber(analytics.totalTaxes.purchasesSubTotal)]);
            dashboardWS.addRow(["إجمالي ضريبة القيمة المضافة (مبيعات)", formatNumber(analytics.totalTaxes.salesVat)]);
            dashboardWS.addRow(["إجمالي ضريبة القيمة المضافة (مشتريات)", formatNumber(analytics.totalTaxes.purchasesVat)]);
            dashboardWS.addRow(["ضريبة الخصم من المنبع (مبيعات)", formatNumber(analytics.totalTaxes.salesWithholdingTax)]);
            dashboardWS.addRow(["ضريبة الخصم من المنبع (مشتريات)", formatNumber(analytics.totalTaxes.purchasesWithholdingTax)]);
            dashboardWS.addRow(["صافي ضريبة القيمة المضافة المستحقة", formatNumber(netVat)]);

            dashboardWS.columns = [{ width: 45 }, { width: 25 }, { width: 20 }];
        }

        const buffer = await wb.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="ETA-Export.xlsx"');
        res.status(200).send(Buffer.from(buffer));

    } catch (error) {
        console.error('Error in Express Server:', error.stack);
        res.status(500).send({ error: 'An internal server error occurred.', details: error.message });
    }
});

// --- المسار 5: معالجة ملف إكسل الإيصالات (كان الملف رقم 5) ---
// دوال مساعدة خاصة بهذا المسار
async function fetchMyEGSCode(itemCode, token) {
    if (!token || !itemCode) return null;
    const url = `https://api-portal.invoicing.eta.gov.eg/api/v1/codetypes/codes/my?CodeTypeID=9&ItemCode=${itemCode}&Ps=1`;
    try {
        const response = await fetch(url, { headers: { "Authorization": `Bearer ${token}` } }  );
        if (response.ok) {
            const data = await response.json();
            return (data.result && data.result.length > 0) ? data.result[0] : null;
        }
        return null;
    } catch (error) {
        return null;
    }
}

async function fetchGS1Code(itemCode, token) {
    if (!token || !itemCode) return null;
    const url = `https://api-portal.invoicing.eta.gov.eg/api/v1/codetypes/2/codes?CodeLookupValue=${itemCode}&ApplyMinChoiceLevel=true&Ps=1`;
    try {
        const response = await fetch(url, { headers: { "Authorization": `Bearer ${token}` } }  );
        if (response.ok) {
            const data = await response.json();
            return (data.result && data.result.length > 0) ? data.result[0] : null;
        }
        return null;
    } catch (error) {
        return null;
    }
}

async function validateNID_API(nid, token) {
    if (!nid || nid.length !== 14 || !/^\d+$/.test(nid)) {
        return { valid: false, message: "يجب أن يتكون من 14 رقمًا." };
    }
    try {
        if (!token) return { valid: false, message: "خطأ مصادقة." };
        const response = await fetch(`https://api-portal.invoicing.eta.gov.eg/api/v1/person/${nid}`, { headers: { 'Authorization': `Bearer ${token}` } }  );
        if (response.status === 200) return { valid: true };
        if (response.status === 400) return { valid: false, message: "الرقم غير مسجل أو غير صحيح." };
        return { valid: false, message: `خطأ ${response.status} من الخادم.` };
    } catch (error) {
        return { valid: false, message: "فشل التحقق من الرقم." };
    }
}

async function validateAndEnrichReceiptData(receiptsMap, token) {
    const validationErrors = [];
    const validatedMap = new Map();
    const requiredItemFields = {
        'description': 'وصف الصنف', 'itemType': 'نوع كود الصنف',
        'itemCode': 'كود الصنف', 'quantity': 'الكمية', 'unitPrice': 'سعر الوحدة'
    };

    const validationPromises = Array.from(receiptsMap.entries()).map(async ([receiptNumber, items]) => {
        const enrichedItems = [];
        let receiptTotalAmount = 0;

        for (const [itemIndex, item] of items.entries()) {
            const enrichedItem = { ...item, officialCodeName: '' };
            receiptTotalAmount += (parseFloat(item.quantity) || 0) * (parseFloat(item.unitPrice) || 0);

            for (const key in requiredItemFields) {
                if (!enrichedItem[key] || String(enrichedItem[key]).trim() === '') {
                    validationErrors.push({ id: `${receiptNumber} (البند ${itemIndex + 1})`, field: requiredItemFields[key], value: 'فارغ', message: 'هذا الحقل إجباري.' });
                }
            }

            const itemCodeType = (enrichedItem.itemType || '').toUpperCase().trim();
            const itemCode = (enrichedItem.itemCode || '').toString().trim();
            if (itemCodeType && itemCode) {
                let codeData = null;
                if (itemCodeType === 'EGS') codeData = await fetchMyEGSCode(itemCode, token);
                else if (itemCodeType === 'GS1') codeData = await fetchGS1Code(itemCode, token);
                
                if (codeData) {
                    enrichedItem.officialCodeName = codeData.codeNameSecondaryLang || "!! اسم غير مسجل !!";
                } else {
                    validationErrors.push({ id: `${receiptNumber} (البند ${itemIndex + 1})`, field: `كود الصنف (${itemCodeType})`, value: itemCode, message: 'الكود غير صحيح أو غير مسجل.' });
                }
            }
            enrichedItems.push(enrichedItem);
        }

        const firstItem = items[0] || {};
        const buyerId = (firstItem.buyerId || '').toString().trim();

        if (receiptTotalAmount > 150000) {
            if (!buyerId) {
                validationErrors.push({ id: receiptNumber, field: 'الرقم القومي للعميل', value: 'فارغ', message: 'إجباري لأن الإجمالي يتجاوز 150,000 جنيه.' });
            } else {
                const nidResult = await validateNID_API(buyerId, token);
                if (!nidResult.valid) {
                    validationErrors.push({ id: receiptNumber, field: 'الرقم القومي للعميل', value: buyerId, message: nidResult.message });
                }
            }
        } else if (buyerId) {
            const nidResult = await validateNID_API(buyerId, token);
            if (!nidResult.valid) {
                validationErrors.push({ id: receiptNumber, field: 'الرقم القومي للعميل', value: buyerId, message: nidResult.message });
            }
        }
        validatedMap.set(receiptNumber, enrichedItems);
    });

    await Promise.all(validationPromises);
    const validatedArray = Array.from(validatedMap.entries());
    return { validatedData: validatedArray, validationErrors };
}

app.post('/process-receipts-excel', async (req, res) => {
    try {
        const { fileBase64, token, type } = req.body;
        if (!fileBase64 || !token) {
            return res.status(400).json({ error: "File data or token is missing." });
        }

        const buffer = Buffer.from(fileBase64, 'base64');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.worksheets[0];
        const allRows = worksheet.getSheetValues().slice(2).map(row => row.slice(1));

        const isReturn = type === 'return';
        const headerMapping = isReturn ? {
            'تاريخ الإصدار (YYYY-MM-DD)': 'dateTimeIssued', 'رقم إشعار المرتجع الداخلي (*)': 'receiptNumber', 'UUID الفاتورة الأصلية (*)': 'referenceUUID',
            'اسم العميل (اختياري)': 'buyerName', 'الرقم القومي للعميل (اختياري)': 'buyerId', 'الكود الداخلي للصنف': 'internalCode',
            'وصف الصنف (*)': 'description', 'نوع كود الصنف (EGS أو GS1) (*)': 'itemType', 'كود الصنف (*)': 'itemCode',
            'وحدة القياس (*)': 'unitType', 'الكمية المرتجعة (*)': 'quantity', 'سعر الوحدة وقت البيع (*)': 'unitPrice',
            'نوع الضريبة 1 (*)': 'taxType_1', 'النوع الفرعي للضريبة 1 (*)': 'taxSubType_1', 'نسبة الضريبة 1 (*)': 'taxRate_1',
            'نوع الضريبة 2 (اختياري)': 'taxType_2', 'النوع الفرعي للضريبة 2 (اختياري)': 'taxSubType_2', 'نسبة الضريبة 2 (اختياري)': 'taxRate_2'
        } : {
            'تاريخ الإصدار (YYYY-MM-DD)': 'dateTimeIssued', 'رقم الإيصال الداخلي (*)': 'receiptNumber', 'اسم العميل (اختياري)': 'buyerName',
            'الرقم القومي للعميل (اختياري)': 'buyerId', 'الكود الداخلي للصنف': 'internalCode', 'وصف الصنف (*)': 'description',
            'نوع كود الصنف (EGS أو GS1) (*)': 'itemType', 'كود الصنف (*)': 'itemCode', 'وحدة القياس (*)': 'unitType',
            'الكمية (*)': 'quantity', 'سعر الوحدة (*)': 'unitPrice', 'نوع الضريبة 1 (*)': 'taxType_1',
            'النوع الفرعي للضريبة 1 (*)': 'taxSubType_1', 'نسبة الضريبة 1 (*)': 'taxRate_1', 'نوع الضريبة 2 (اختياري)': 'taxType_2',
            'النوع الفرعي للضريبة 2 (اختياري)': 'taxSubType_2', 'نسبة الضريبة 2 (اختياري)': 'taxRate_2'
        };

        const mappedRows = allRows.map(row => {
            const newRow = {};
            Object.keys(headerMapping).forEach((header, index) => {
                const key = headerMapping[header];
                if (key) newRow[key] = row[index];
            });
            return newRow;
        });

        const receiptsMap = new Map();
        let lastReceiptNumber = '';
        let lastHeaderInfo = {};
        mappedRows.forEach(row => {
            const currentReceiptNumber = String(row.receiptNumber || lastReceiptNumber).trim();
            if (!currentReceiptNumber) return;
            if (currentReceiptNumber !== lastReceiptNumber) {
                lastHeaderInfo = { dateTimeIssued: row.dateTimeIssued, buyerName: row.buyerName, buyerId: row.buyerId, referenceUUID: row.referenceUUID };
                receiptsMap.set(currentReceiptNumber, []);
            }
            receiptsMap.get(currentReceiptNumber).push({ ...lastHeaderInfo, ...row });
            lastReceiptNumber = currentReceiptNumber;
        });

        const { validatedData, validationErrors } = await validateAndEnrichReceiptData(receiptsMap, token);
        res.status(200).json({ validatedData, validationErrors });

    } catch (error) {
        console.error("Error on server:", error);
        res.status(500).json({ error: "An internal server error occurred.", details: error.message });
    }
});

// --- المسار 6: التعامل مع منطق الإيصالات والمرتجعات (كان الملف رقم 6) ---
// دوال مساعدة خاصة بهذا المسار
function calculateReceiptDataOnServer(itemsData, sellerData, deviceSerial, activityCode) {
    const firstRow = itemsData[0];
    const header = {
        dateTimeIssued: new Date().toISOString().substring(0, 19) + "Z",
        receiptNumber: String(firstRow.receiptNumber || `RCPT_${Math.floor(Date.now() / 1000)}`),
        uuid: "",
        previousUUID: "",
        currency: "EGP",
        exchangeRate: 0.00,
    };

    let finalTotalSales = 0;
    const finalTaxTotalsMap = new Map();

    const calculatedItemData = itemsData.map(item => {
        const quantity = parseFloat(item.quantity) || 0;
        const unitPrice = parseFloat(item.unitPrice) || 0;
        const itemTotalSale = parseFloat((quantity * unitPrice).toFixed(5));
        const itemNetSale = itemTotalSale;
        const taxableItems = [];
        let totalTaxAmountForItem = 0;

        for (let i = 1; i <= 2; i++) {
            const taxType = item[`taxType_${i}`];
            const taxRate = parseFloat(item[`taxRate_${i}`]);
            if (taxType && !isNaN(taxRate) && taxRate > 0) {
                const taxAmount = parseFloat((itemNetSale * (taxRate / 100)).toFixed(5));
                taxableItems.push({
                    taxType: String(taxType),
                    amount: taxAmount,
                    subType: String(item[`taxSubType_${i}`]),
                    rate: taxRate
                });
                totalTaxAmountForItem += (taxType === 'T4' ? -taxAmount : taxAmount);
                finalTaxTotalsMap.set(String(taxType), (finalTaxTotalsMap.get(String(taxType)) || 0) + taxAmount);
            }
        }

        const itemTotal = parseFloat((itemNetSale + totalTaxAmountForItem).toFixed(5));
        finalTotalSales += itemTotalSale;

        return {
            internalCode: String(item.internalCode),
            description: String(item.description),
            itemType: String(item.itemType || 'EGS'),
            itemCode: String(item.itemCode),
            unitType: String(item.unitType || 'EA'),
            quantity: quantity,
            unitPrice: unitPrice,
            totalSale: itemTotalSale,
            netSale: itemNetSale,
            total: itemTotal,
            taxableItems: taxableItems
        };
    });

    return {
        header: header,
        documentType: { receiptType: "S", typeVersion: "1.2" },
        seller: sellerData,
        buyer: { type: "P", id: firstRow.buyerId || "", name: firstRow.buyerName || "عميل نقدي" },
        itemData: calculatedItemData,
        totalSales: parseFloat(finalTotalSales.toFixed(5)),
        netAmount: parseFloat(finalTotalSales.toFixed(5)),
        taxTotals: Array.from(finalTaxTotalsMap, ([taxType, amount]) => ({ taxType, amount: parseFloat(amount.toFixed(5)) })),
        totalAmount: parseFloat(calculatedItemData.reduce((sum, item) => sum + item.total, 0).toFixed(5)),
        paymentMethod: "C"
    };
}

function calculateReturnReceiptDataOnServer(itemsData, sellerData, deviceSerial, activityCode) {
    const firstRow = itemsData[0];
    const header = {
        dateTimeIssued: new Date().toISOString().substring(0, 19) + "Z",
        receiptNumber: String(firstRow.receiptNumber || `RTN_${Math.floor(Date.now() / 1000)}`),
        uuid: "",
        previousUUID: "",
        referenceUUID: String(firstRow.referenceUUID || ""),
        currency: "EGP",
        exchangeRate: 0.0,
    };

    let finalTotalSales = 0;
    const finalTaxTotalsMap = new Map();

    const calculatedItemData = itemsData.map(item => {
        const quantity = parseFloat(item.quantity) || 0;
        const unitPrice = parseFloat(item.unitPrice) || 0;
        const itemTotalSale = parseFloat((quantity * unitPrice).toFixed(5));
        const itemNetSale = itemTotalSale;
        let totalTaxAmountForItem = 0;
        const taxableItems = [];

        for (let i = 1; i <= 2; i++) {
            const taxType = item[`taxType_${i}`];
            const taxRate = parseFloat(item[`taxRate_${i}`]);
            if (taxType && !isNaN(taxRate) && taxRate > 0) {
                const taxAmount = parseFloat((itemNetSale * (taxRate / 100)).toFixed(5));
                taxableItems.push({ taxType: String(taxType), amount: taxAmount, subType: String(item[`taxSubType_${i}`]), rate: taxRate });
                totalTaxAmountForItem += (taxType === 'T4' ? -taxAmount : taxAmount);
                finalTaxTotalsMap.set(String(taxType), (finalTaxTotalsMap.get(String(taxType)) || 0) + taxAmount);
            }
        }

        const itemTotal = parseFloat((itemNetSale + totalTaxAmountForItem).toFixed(5));
        finalTotalSales += itemTotalSale;

        return {
            internalCode: String(item.internalCode),
            description: String(item.description),
            itemType: String(item.itemType || 'EGS'),
            itemCode: String(item.itemCode),
            unitType: String(item.unitType || 'EA'),
            quantity: quantity,
            unitPrice: unitPrice,
            totalSale: itemTotalSale,
            netSale: itemNetSale,
            total: itemTotal,
            taxableItems: taxableItems
        };
    });

    const totalAmount = parseFloat(calculatedItemData.reduce((sum, item) => sum + item.total, 0).toFixed(5));

    return {
        header: header,
        documentType: { receiptType: "R", typeVersion: "1.2" },
        seller: {
            rin: sellerData.id,
            companyTradeName: sellerData.name,
            branchCode: "0",
            branchAddress: { country: "EG", governate: sellerData.governate, regionCity: sellerData.regionCity, street: sellerData.street, buildingNumber: sellerData.buildingNumber },
            deviceSerialNumber: deviceSerial,
            activityCode: activityCode,
        },
        buyer: { type: "P", id: firstRow.buyerId || "", name: firstRow.buyerName || "عميل نقدي" },
        itemData: calculatedItemData,
        totalSales: parseFloat(finalTotalSales.toFixed(5)),
        netAmount: parseFloat(finalTotalSales.toFixed(5)),
        taxTotals: Array.from(finalTaxTotalsMap, ([taxType, amount]) => ({ taxType, amount: parseFloat(amount.toFixed(5)) })),
        totalAmount: totalAmount,
        paymentMethod: "C"
    };
}

app.post('/eta-logic-handler', async (req, res) => {
    const action = req.body.action;

    if (action === 'getJobs') {
        try {
            const binId = '68de58f9ae596e708f037bdc';
            const accessKey = '$2a$10$rXrBfSrwkJ60zqKQInt5.eVxCq14dTw9vQX8LXcpnWb7SJ5ZLNoKe';
            const response = await fetch(`https://api.jsonbin.io/v3/b/${binId}`, {
                method: 'GET',
                headers: { 'X-Access-Key': accessKey }
            }  );
            if (!response.ok) throw new Error(`Failed to fetch from jsonbin: ${response.statusText}`);
            const data = await response.json();
            res.status(200).json({ success: true, data: data.record?.jobs || data.jobs || [] });
        } catch (error) {
            res.status(500).json({ success: false, error: error.message });
        }
    }
    else if (action === 'calculateReceipt') {
        try {
            const { items, seller, deviceSerial, activityCode } = req.body;
            if (!items || !seller || !deviceSerial || !activityCode) throw new Error("Incomplete data for receipt.");
            const receiptPayload = calculateReceiptDataOnServer(items, seller, deviceSerial, activityCode);
            res.status(200).json({ success: true, data: receiptPayload });
        } catch (error) {
            res.status(500).json({ success: false, error: error.message });
        }
    }
    else if (action === 'calculateReturnReceipt') {
        try {
            const { items, seller, deviceSerial, activityCode } = req.body;
            if (!items || !seller || !deviceSerial || !activityCode) throw new Error("Incomplete data for return receipt.");
            const returnPayload = calculateReturnReceiptDataOnServer(items, seller, deviceSerial, activityCode);
            res.status(200).json({ success: true, data: returnPayload });
        } catch (error) {
            res.status(500).json({ success: false, error: error.message });
        }
    }
    else {
        res.status(400).json({ success: false, error: 'Invalid action specified.' });
    }
});

// --- المسار 7: تحميل الفواتير كملف PDF (كان الملف رقم 7) ---
app.post('/fetch-pdf', async (req, res) => {
    try {
        const { uuid, token } = req.body;

        if (!uuid || !token) {
            return res.status(400).send({ error: 'Missing UUID or token.' });
        }

        const url = `https://api-portal.invoicing.eta.gov.eg/api/v1.0/documents/${uuid}/pdf`;
        const response = await axios.get(url, {
            headers: { Authorization: `Bearer ${token}` },
            responseType: 'arraybuffer'
        }  );
        
        res.setHeader('Content-Type', 'application/pdf');
        return res.status(200).send(response.data);

    } catch (error) {
        console.error(`Failed to fetch PDF for ${uuid}:`, error.message);
        const status = error.response ? error.response.status : 500;
        const message = error.response ? error.response.data.toString() : error.message;
        return res.status(status).send({ error: `Failed to fetch PDF: ${message}` });
    }
});

// --- المسار 8: إنشاء قالب إكسل للمرتجعات (كان الملف رقم 8) ---
app.post('/generate-return-template', async (req, res) => {
    try {
        const wb = new ExcelJS.Workbook();
        const mainSheetName = "Returns";
        const listsSheetName = "Lists";

     // ... تكملة الكود من داخل المسار /generate-return-template

        const listsData = {
            codeTypes: ["EGS", "GS1"],
            mainTaxTypes: [
                { name: "ضريبة القيمة المضافة", code: "T1" }, { name: "ضريبة الجدول (نسبية)", code: "T2" },
                { name: "ضريبة الجدول (النوعية)", code: "T3" }, { name: "الخصم تحت حساب الضريبة", code: "T4" },
                { name: "ضريبة الدمغة (نسبية)", code: "T5" }, { name: "ضريبة الدمغة (قطعية)", code: "T6" },
                { name: "ضريبة الملاهي", code: "T7" }, { name: "رسم تنمية الموارد", code: "T8" },
                { name: "رسم خدمة", code: "T9" }, { name: "رسم المحليات", code: "T10" },
                { name: "رسم التأمين الصحي", code: "T11" }, { name: "رسوم أخرى", code: "T12" }
            ],
            taxSubtypes: [
                { name: "تصدير للخارج", code: "V001", ref: "T1" }, { name: "تصدير مناطق حرة", code: "V002", ref: "T1" },
                { name: "سلعة أو خدمة معفاة", code: "V003", ref: "T1" }, { name: "سلعة أو خدمة غير خاضعة", code: "V004", ref: "T1" },
                { name: "إعفاءات دبلوماسيين", code: "V005", ref: "T1" }, { name: "إعفاءات دفاع وأمن قومي", code: "V006", ref: "T1" },
                { name: "إعفاءات اتفاقيات", code: "V007", ref: "T1" }, { name: "إعفاءات خاصة وأخرى", code: "V008", ref: "T1" },
                { name: "سلع عامة (14%)", code: "V009", ref: "T1" }, { name: "نسب ضريبة أخرى", code: "V010", ref: "T1" },
                { name: "ضريبة جدول (نسبية)", code: "Tbl01", ref: "T2" }, { name: "ضريبة جدول (نوعية)", code: "Tbl02", ref: "T3" },
                { name: "المقاولات", code: "W001", ref: "T4" }, { name: "التوريدات", code: "W002", ref: "T4" },
                { name: "المشتريات", code: "W003", ref: "T4" }, { name: "الخدمات", code: "W004", ref: "T4" },
                { name: "أتعاب مهنية", code: "W010", ref: "T4" }, { name: "دمغة (نسبية)", code: "ST01", ref: "T5" },
                { name: "دمغة (قطعية)", code: "ST02", ref: "T6" }, { name: "ملاهي (نسبة)", code: "Ent01", ref: "T7" },
                { name: "تنمية موارد (نسبة)", code: "RD01", ref: "T8" }, { name: "رسم خدمة (نسبة)", code: "SC01", ref: "T9" },
                { name: "محليات (نسبة)", code: "Mn01", ref: "T10" }, { name: "تأمين صحي (نسبة)", code: "MI01", ref: "T11" },
                { name: "رسوم أخرى", code: "OF01", ref: "T12" }
            ],
            unitTypes: [
                { name: "قطعة", code: "EA" }, { name: "كرتونة", code: "CS" }, { name: "حقيبة", code: "BG" },
                { name: "عبوة", code: "PK" }, { name: "صندوق", code: "BX" }, { name: "كيلو جرام", code: "KGM" },
                { name: "متر", code: "M" }, { name: "لتر", code: "LTR" }, { name: "طن", code: "TNE" }
            ]
        };

        const ws = wb.addWorksheet(mainSheetName);
        ws.views = [{ rightToLeft: true }];

        const headers = [
            'تاريخ الإصدار (YYYY-MM-DD)', 'رقم إشعار المرتجع الداخلي (*)', 'UUID الفاتورة الأصلية (*)',
            'اسم العميل (اختياري)', 'الرقم القومي للعميل (اختياري)', 'الكود الداخلي للصنف',
            'وصف الصنف (*)', 'نوع كود الصنف (EGS أو GS1) (*)', 'كود الصنف (*)',
            'وحدة القياس (*)', 'الكمية المرتجعة (*)', 'سعر الوحدة وقت البيع (*)',
            'نوع الضريبة 1 (*)', 'النوع الفرعي للضريبة 1 (*)', 'نسبة الضريبة 1 (*)',
            'نوع الضريبة 2 (اختياري)', 'النوع الفرعي للضريبة 2 (اختياري)', 'نسبة الضريبة 2 (اختياري)'
        ];
        ws.getRow(1).values = headers;
        ws.getRow(2).values = [
            '2025-10-03', 'RTN-001', 'اكتب هنا UUID الفاتورة الأصلية', 'عميل نقدي', '', 'ITEM-01',
            'صنف مرتجع تجريبي', 'EGS', 'EG-xxxx-101', 'قطعة', 1, 100,
            'ضريبة القيمة المضافة', 'سلع عامة (14%)', 14
        ];
        ws.columns = headers.map(() => ({ width: 30 }));

        const lists_ws = wb.addWorksheet(listsSheetName);
        lists_ws.state = 'hidden';

        lists_ws.getColumn('A').values = ['CodeTypes', ...listsData.codeTypes];
        lists_ws.getColumn('B').values = ['UnitTypeNames', ...listsData.unitTypes.map(t => t.name)];
        lists_ws.getColumn('C').values = ['UnitTypeCodes', ...listsData.unitTypes.map(t => t.code)];
        lists_ws.getColumn('D').values = ['TaxTypeNames', ...listsData.mainTaxTypes.map(t => t.name)];
        lists_ws.getColumn('E').values = ['TaxTypeCodes', ...listsData.mainTaxTypes.map(t => t.code)];

        listsData.mainTaxTypes.forEach((taxType, index) => {
            const col = lists_ws.getColumn(6 + (index * 2));
            const subtypes = listsData.taxSubtypes.filter(st => st.ref === taxType.code);
            col.values = [taxType.code, ...subtypes.map(st => st.name)];
            lists_ws.getColumn(7 + (index * 2)).values = [taxType.code, ...subtypes.map(st => st.code)];
        });

        wb.definedNames.add(`'${listsSheetName}'!$A$2:$A$${listsData.codeTypes.length + 1}`, 'CodeTypes');
        wb.definedNames.add(`'${listsSheetName}'!$B$2:$B$${listsData.unitTypes.length + 1}`, 'UnitTypeNames');
        wb.definedNames.add(`'${listsSheetName}'!$D$2:$D$${listsData.mainTaxTypes.length + 1}`, 'TaxTypeNames');

        listsData.mainTaxTypes.forEach((taxType, index) => {
            const colLetter = String.fromCharCode('A'.charCodeAt(0) + 5 + (index * 2));
            const lastRow = (listsData.taxSubtypes.filter(st => st.ref === taxType.code).length || 1) + 1;
            wb.definedNames.add(`'${listsSheetName}'!$${colLetter}$2:$${colLetter}$${lastRow}`, `Sub_${taxType.code}`);
        });

        for (let i = 2; i <= 501; i++) {
            ws.getCell(`H${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=CodeTypes'] };
            ws.getCell(`J${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=UnitTypeNames'] };
            ws.getCell(`M${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=TaxTypeNames'] };
            ws.getCell(`N${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`=INDIRECT("Sub_"&VLOOKUP(M${i},${listsSheetName}!$D$2:$E$20,2,FALSE))`] };
            ws.getCell(`P${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['=TaxTypeNames'] };
            ws.getCell(`Q${i}`).dataValidation = { type: 'list', allowBlank: true, formulae: [`=INDIRECT("Sub_"&VLOOKUP(P${i},${listsSheetName}!$D$2:$E$20,2,FALSE))`] };
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="Template_Return_Receipts.xlsx"');
        await wb.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error("Error generating Excel file:", error);
        res.status(500).send("Error generating file.");
    }
});

// --- المسار 9: جلب معلومات المبرمج (كان الملف رقم 9) ---
app.get('/get-designer-info', (req, res) => {
    const designerInfo = {
        name: "محاسب قانوني: محمد صبري",
        phone: "01060872599"
    };
    res.status(200).json(designerInfo);
});

// --- المسار 10: حساب UUID (كان الملف رقم 10) ---
// دوال مساعدة خاصة بهذا المسار
async function sha256Hex(str) {
    const hash = crypto.createHash('sha256');
    hash.update(str);
    return hash.digest('hex');
}
function isWS(c) { return c === 0x20 || c === 0x0A || c === 0x0D || c === 0x09; }
function Serializer(src) { this.s = src; this.n = src.length; this.i = 0; this.out = []; }
Serializer.prototype.peek = function() { return this.i < this.n ? this.s.charCodeAt(this.i) : -1; }
Serializer.prototype.skip = function() { while (this.i < this.n && isWS(this.s.charCodeAt(this.i))) this.i++; }
Serializer.prototype.expect = function(ch) { if (this.s[this.i] !== ch) throw new Error('Expected ' + ch + ' at ' + this.i); this.i++; }
Serializer.prototype.readString = function() { let start = this.i; this.expect('"'); while (this.i < this.n) { const c = this.s.charCodeAt(this.i); if (c === 0x22) { this.i++; break; } if (c === 0x5C) { this.i += 2; } else { this.i++; } } return this.s.slice(start, this.i); }
Serializer.prototype.readNumber = function() { let start = this.i; if (this.s[this.i] === '-') this.i++; if (this.s[this.i] === '0') { this.i++; } else { if (!(this.s[this.i] >= '1' && this.s[this.i] <= '9')) throw new Error('num'); while (this.s[this.i] >= '0' && this.s[this.i] <= '9') this.i++; } if (this.s[this.i] === '.') { this.i++; if (!(this.s[this.i] >= '0' && this.s[this.i] <= '9')) throw new Error('frac'); while (this.s[this.i] >= '0' && this.s[this.i] <= '9') this.i++; } if (this.s[this.i] === 'e' || this.s[this.i] === 'E') { this.i++; if (this.s[this.i] === '+' || this.s[this.i] === '-') this.i++; if (!(this.s[this.i] >= '0' && this.s[this.i] <= '9')) throw new Error('exp'); while (this.s[this.i] >= '0' && this.s[this.i] <= '9') this.i++; } return this.s.slice(start, this.i); }
Serializer.prototype.readLiteral = function() { if (this.s.startsWith('true', this.i)) { this.i += 4; return 'true'; } if (this.s.startsWith('false', this.i)) { this.i += 5; return 'false'; } if (this.s.startsWith('null', this.i)) { this.i += 4; return 'null'; } throw new Error('literal@' + this.i); }
Serializer.prototype.emitKey = function(nameUpper) { this.out.push('"' + nameUpper + '"'); }
Serializer.prototype.emitScalar = function(lexeme) { this.out.push('"' + lexeme + '"'); }
Serializer.prototype.serializeObject = function(path, exclude) { this.skip(); this.expect('{'); this.skip(); let first = true; while (this.i < this.n && this.s[this.i] !== '}') { if (!first) { if (this.s[this.i] === ',') { this.i++; this.skip(); } } const keyLex = this.readString(); const key = JSON.parse(keyLex); const K = key.toUpperCase(); this.skip(); this.expect(':'); this.skip(); const cur = path ? path + '.' + key : key; const ex = exclude.indexOf(cur) !== -1; const c = this.peek(); if (c === 0x22) { const v = this.readString(); if (!ex) { this.emitKey(K); this.emitScalar(v.slice(1, -1)); } } else if (c === 0x7B) { if (!ex) { this.emitKey(K); } this.serializeObject(cur, exclude); } else if (c === 0x5B) { if (!ex) { this.emitKey(K); } this.serializeArray(cur, exclude, ex, K); } else if ((c === 0x2D) || (c >= 0x30 && c <= 0x39)) { const num = this.readNumber(); if (!ex) { this.emitKey(K); this.emitScalar(num); } } else { const lit = this.readLiteral(); if (!ex) { this.emitKey(K); this.emitScalar(lit); } } this.skip(); first = false; } this.expect('}'); }
Serializer.prototype.serializeArray = function(path, exclude, isExcluded, propNameUpper) { this.skip(); this.expect('['); this.skip(); let first = true; while (this.i < this.n && this.s[this.i] !== ']') { if (!first) { if (this.s[this.i] === ',') { this.i++; this.skip(); } } if (!isExcluded && propNameUpper) { this.emitKey(propNameUpper); } const c = this.peek(); if (c === 0x7B) { this.serializeObject(path, exclude); } else if (c === 0x22) { const v = this.readString(); if (!isExcluded) { this.emitScalar(v.slice(1, -1)); } } else if ((c === 0x2D) || (c >= 0x30 && c <= 0x39)) { const num = this.readNumber(); if (!isExcluded) { this.emitScalar(num); } } else { const lit = this.readLiteral(); if (!isExcluded) { this.emitScalar(lit); } } this.skip(); first = false; } this.expect(']'); }
function findFirstReceiptSlice(src) { const m = src.match(/"receipts"\s*:\s*\[/); if (!m) return src.trim(); let i = m.index + m[0].length; while (i < src.length && /\s/.test(src[i])) i++; if (src[i] !== '{') return src.trim(); let depth = 0, start = i; while (i < src.length) { const ch = src[i]; if (ch === '"') { i++; while (i < src.length) { if (src[i] === '\\') { i += 2; continue; } if (src[i] === '"') { i++; break; } i++; } continue; } if (ch === '{') { depth++; } if (ch === '}') { depth--; if (depth === 0) { i++; break; } } i++; } return src.slice(start, i); }
function getCanonicalFromRawText(raw) { const slice = findFirstReceiptSlice(raw); const ser = new Serializer(slice); ser.serializeObject('', []); return ser.out.join(''); }
async function computeUuidFromRawText(raw) { const canonical = getCanonicalFromRawText(raw); return await sha256Hex(canonical); }

app.post('/compute-uuid', async (req, res) => {
    const rawText = req.body.data;
    if (!rawText) {
        return res.status(400).send({ error: 'No data provided.' });
    }
    try {
        const uuid = await computeUuidFromRawText(rawText);
        res.status(200).json({ uuid: uuid });
    } catch (error) {
        res.status(500).send({ error: `Error: ${error.message}` });
    }
});

// --- المسار 11: جلب معلومات إضافية (كان الملف رقم 11) ---
app.get('/get-prayer-info', (req, res) => {
    const designerInfo = {
        prayer: "صلي علي محمد",
        name: "المحاسب : محمد صبري",
        contact: "واتساب: 01060872599"
    };
    res.status(200).json(designerInfo);
});

// --- المسار 12: معالجة وترجمة بيانات الإكسل (كان الملف رقم 12) ---
app.post('/process-excel-hybrid', async (req, res) => {
    try {
        const { rawData, enrichedData } = req.body;
        if (!rawData || !enrichedData) {
            return res.status(400).json({ error: "Missing required data." });
        }

        const reverseTranslationMaps = {
            receiverType: { "أعمال": "B", "شخصي": "P", "أجنبي": "F" },
            country: { "مصر": "EG", "الإمارات": "AE", "السعودية": "SA", "الكويت": "KW", "قطر": "QA", "البحرين": "BH", "عمان": "OM", "الأردن": "JO", "لبنان": "LB", "أمريكا": "US", "بريطانيا": "GB", "ألمانيا": "DE" },
            unitType: { "each (ST) ( ST )": "EA", "كرتونة": "CS", "حقيبة": "BG", "عبوة": "PK", "صندوق": "BX", "كيلو جرام": "KG", "متر": "M", "لتر": "L", "طن": "T" },
            currency: { "جنيه مصري": "EGP", "دولار أمريكي": "USD", "يورو": "EUR", "جنيه استرليني": "GBP", "ريال سعودي": "SAR", "درهم إماراتي": "AED" },
            taxType: { "ضريبة القيمة المضافة": "T1", "ضريبة الجدول (نسبية)": "T2", "ضريبة الجدول (النوعية)": "T3", "الخصم تحت حساب الضريبة": "T4", "ضريبة الدمغة (نسبية)": "T5", "ضريبة الدمغة (قطعية)": "T6", "ضريبة الملاهي": "T7", "رسم تنمية الموارد": "T8", "رسم خدمة": "T9", "رسم المحليات": "T10", "رسم التأمين الصحي": "T11", "رسوم أخرى": "T12" },
            taxSubtype: { "تصدير للخارج": "V001", "تصدير مناطق حرة": "V002", "سلعة أو خدمة معفاة": "V003", "سلعة أو خدمة غير خاضعة": "V004", "إعفاءات دبلوماسيين": "V005", "إعفاءات دفاع وأمن قومي": "V006", "إعفاءات اتفاقيات": "V007", "إعفاءات خاصة وأخرى": "V008", "سلع عامة (14%)": "V009", "نسب ضريبة أخرى": "V010", "ضريبة جدول (نسبية)": "Tbl01", "ضريبة جدول (نوعية)": "Tbl02", "المقاولات": "W001", "التوريدات": "W002", "المشتريات": "W003", "الخدمات": "W004", "أتعاب مهنية": "W010", "دمغة (نسبية)": "ST01", "دمغة (قطعية)": "ST02", "ملاهي (نسبة)": "Ent01", "تنمية موارد (نسبة)": "RD01", "رسم خدمة (نسبة)": "SC01", "محليات (نسبة)": "Mn01", "تأمين صحي (نسبة)": "MI01", "رسوم أخرى": "OF01" }
        };

        const columnMap = { 3: 'receiverType', 4: 'country', 13: 'unitType', 16: 'currency', 20: 'taxType', 21: 'taxSubtype', 23: 'taxType', 24: 'taxSubtype', 26: 'taxType', 27: 'taxSubtype' };
        const translatedData = rawData.map(row => {
            const newRow = [...row];
            for (const index in columnMap) {
                const mapName = columnMap[index];
                const cellValue = newRow[index]?.trim();
                if (cellValue && reverseTranslationMaps[mapName] && reverseTranslationMaps[mapName][cellValue]) {
                    newRow[index] = reverseTranslationMaps[mapName][cellValue];
                }
            }
            return newRow;
        });

        let lastInvoiceData = [];
        const processedData = translatedData.map((row, index) => {
            const hasInternalID = row[0] && String(row[0]).trim() !== '';
            if (hasInternalID) {
                lastInvoiceData = row.slice(0, 9);
                return [...lastInvoiceData, ...row.slice(9)];
            } else {
                if (index === 0 || !lastInvoiceData.length) {
                    lastInvoiceData = [`AUTO_${Date.now()}`, '12345678', 'عميل نقدي', 'P', 'EG', 'القاهرة', 'عابدين', '1', '1'];
                }
                return [...lastInvoiceData, ...row.slice(9)];
            }
        }).filter(row => row && (row[9] || row[11]));

        const finalData = processedData.map((row, index) => {
            const enrichedRow = enrichedData[index] || {};
            row[2] = enrichedRow[2] || row[2];
            row[5] = enrichedRow[5] || row[5];
            row[6] = enrichedRow[6] || row[6];
            row[7] = enrichedRow[7] || row[7];
            row[8] = enrichedRow[8] || row[8];
            row[37] = enrichedRow[37] || '';
            return row;
        });

        res.status(200).json({ validatedData: finalData, validationErrors: [] });

    } catch (error) {
        console.error("Error on Excel processing server:", error);
        res.status(500).json({ error: "An internal server error occurred.", details: error.message });
    }
});

// --- المسار 13: إنشاء وتحديث المسودات (كان الملف رقم 13) ---
app.post('/handle-draft', async (req, res) => {
    try {
        const { action, payload, draftId, accessToken } = req.body;
        if (!action || !payload || !accessToken) {
            return res.status(400).json({ success: false, error: "Missing required data." });
        }
        let targetUrl, method;
        if (action === 'create') {
            targetUrl = "https://api-portal.invoicing.eta.gov.eg/api/v1/documents/drafts";
            method = 'POST';
        } else if (action === 'update' && draftId  ) {
            targetUrl = `https://api-portal.invoicing.eta.gov.eg/api/v1/documents/drafts/${draftId}`;
            method = 'PUT';
        } else {
            return res.status(400  ).json({ success: false, error: "Invalid action." });
        }
        const etaResponse = await fetch(targetUrl, {
            method: method,
            headers: { "Content-Type": "application/json", "Authorization": `Bearer ${accessToken}` },
            body: JSON.stringify(payload)
        });
        const responseData = await etaResponse.json();
        if (etaResponse.ok) {
            res.status(200).json({ success: true, data: responseData });
        } else {
            const errorMessage = responseData.error?.details?.[0]?.message || responseData.error?.message || JSON.stringify(responseData);
            res.status(etaResponse.status).json({ success: false, error: errorMessage });
        }
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// -----------------------------------------------------------------------------
// |                       ✅ نهاية المسارات                                    |
// -----------------------------------------------------------------------------
// 5. تشغيل الخادم
const port = process.env.PORT || 8080;
app.listen(port, () => {
    console.log(`Unified Server is running and listening on port ${port}`);
});


// 6. تصدير التطبيق ليعمل مع Vercel
// 6. تصدير التطبيق ليعمل مع Vercel
module.exports = app;
