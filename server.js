// server.js - نسخة Vercel
require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

// CORS
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    next();
});

// =============================================
// المصادقة مع Google Sheets
// =============================================
function getAuth() {
    const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
    const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n');

    return new google.auth.JWT({
        email: clientEmail,
        key: privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
}

// =============================================
// 1. جلب البيانات الأساسية
// =============================================
async function fetchBasicHospitals() {
    try {
        const spreadsheetId = process.env.SPREADSHEET_ID;
        const auth = getAuth();
        const sheets = google.sheets({ version: 'v4', auth });

        console.log('📊 جاري قراءة البيانات الأساسية...');

        const info = await sheets.spreadsheets.get({ spreadsheetId });
        const sheetNames = info.data.sheets.map(s => s.properties?.title);
        const firstSheetName = sheetNames[0] || 'Sheet1';

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `${firstSheetName}!A1:E`,
        });

        const rows = response.data.values;
        if (!rows || rows.length < 2) {
            throw new Error('لا توجد بيانات');
        }

        const headers = rows[0];
        const colGov = headers.indexOf('المحافظة');
        const colName = headers.indexOf('اسم_المستشفى');
        const colLat = headers.indexOf('خط_العرض');
        const colLng = headers.indexOf('خط_الطول');
        const colSheetId = headers.indexOf('sheet_id');

        const governoratesMap = {};

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const govName = row[colGov]?.trim();
            if (!govName) continue;

            if (!governoratesMap[govName]) {
                governoratesMap[govName] = {
                    اسم: govName,
                    مركز: [0, 0],
                    مستشفيات: []
                };
            }

            const hospital = {
                اسم: row[colName]?.trim() || 'غير محدد',
                خط_العرض: parseFloat(row[colLat]) || 0,
                خط_الطول: parseFloat(row[colLng]) || 0,
                sheet_id: row[colSheetId]?.trim() || '',
            };

            governoratesMap[govName].مستشفيات.push(hospital);
        }

        Object.values(governoratesMap).forEach(gov => {
            const hospitals = gov.مستشفيات;
            if (hospitals.length > 0) {
                const validHospitals = hospitals.filter(h => h.خط_العرض !== 0 && h.خط_الطول !== 0);
                if (validHospitals.length > 0) {
                    gov.مركز = [
                        validHospitals.reduce((sum, h) => sum + h.خط_العرض, 0) / validHospitals.length,
                        validHospitals.reduce((sum, h) => sum + h.خط_الطول, 0) / validHospitals.length
                    ];
                }
            }
        });

        return { محافظات: Object.values(governoratesMap) };

    } catch (error) {
        console.error('❌ خطأ:', error.message);
        throw error;
    }
}

// =============================================
// 2. جلب التفاصيل
// =============================================
async function fetchHospitalDetails(sheetId) {
    try {
        const auth = getAuth();
        const sheets = google.sheets({ version: 'v4', auth });

        console.log(`📊 جاري جلب التفاصيل: ${sheetId}`);

        const info = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
        const sheetNames = info.data.sheets.map(s => s.properties?.title);

        const allData = {};

        for (const sheetName of sheetNames) {
            try {
                const response = await sheets.spreadsheets.values.get({
                    spreadsheetId: sheetId,
                    range: `${sheetName}!A1:Z`,
                });
                const rows = response.data.values;
                if (rows && rows.length > 0) {
                    allData[sheetName] = rows;
                }
            } catch (error) {
                console.warn(`⚠️ ${sheetName}:`, error.message);
            }
        }

        const result = {};

        // الاعتمادات
        if (allData['وصف المستشفى']) {
            const rows = allData['وصف المستشفى'];
            let found = false;
            const credits = [];
            for (const row of rows) {
                if (!found && row.some(cell => cell?.includes('الاعتماد') || cell?.includes('GAHAR'))) {
                    found = true;
                    continue;
                }
                if (found && row.length > 1) {
                    credits.push(row);
                }
                if (found && row.some(cell => cell?.trim() === '')) break;
            }
            if (credits.length > 0) result.الاعتمادات = credits;
        }

        // الهيكل الإداري
        for (const [name, rows] of Object.entries(allData)) {
            if (name.includes('الهيكل') || name.includes('ادارى')) {
                const table = [];
                let start = false;
                for (const row of rows) {
                    if (!start && row.some(cell => cell?.includes('الوظيفة') || cell?.includes('مسمى'))) {
                        start = true;
                        continue;
                    }
                    if (start && row.length > 1 && row.some(cell => cell?.trim())) {
                        table.push(row);
                    }
                }
                if (table.length > 0) {
                    result.الهيكل_الاداري = table;
                    break;
                }
            }
        }

        // رؤساء الأقسام
        for (const [name, rows] of Object.entries(allData)) {
            if (name.includes('رؤساء') || name.includes('اقسام')) {
                const table = [];
                let start = false;
                for (const row of rows) {
                    if (!start && row.some(cell => cell?.includes('البيان') || cell?.includes('القسم'))) {
                        start = true;
                        continue;
                    }
                    if (start && row.length > 1 && row.some(cell => cell?.trim())) {
                        table.push(row);
                    }
                }
                if (table.length > 0) {
                    result.رؤساء_الاقسام = table;
                    break;
                }
            }
        }

        // اللجان
        for (const [name, rows] of Object.entries(allData)) {
            if (name.includes('لجان')) {
                const table = [];
                let start = false;
                for (const row of rows) {
                    if (!start && row.some(cell => cell?.includes('اسم اللجنة'))) {
                        start = true;
                        continue;
                    }
                    if (start && row.length > 1 && row.some(cell => cell?.trim())) {
                        table.push(row);
                    }
                }
                if (table.length > 0) {
                    result.اللجان = table;
                    break;
                }
            }
        }

        // أنظمة التشغيل
        for (const [name, rows] of Object.entries(allData)) {
            if (name.includes('تشغيل') || name.includes('كهروميكانيكال')) {
                const table = [];
                let start = false;
                for (const row of rows) {
                    if (!start && row.some(cell => cell?.includes('البيان'))) {
                        start = true;
                        continue;
                    }
                    if (start && row.length > 1 && row.some(cell => cell?.trim())) {
                        table.push(row);
                    }
                }
                if (table.length > 0) {
                    result.انظمة_التشغيل = table;
                    break;
                }
            }
        }

        // القوة البشرية
        for (const [name, rows] of Object.entries(allData)) {
            if (name.includes('القوة البشرية') || name.includes('البشرية')) {
                const table = [];
                let start = false;
                for (const row of rows) {
                    if (!start && row.some(cell => cell?.includes('البيان') || cell?.includes('العدد'))) {
                        start = true;
                        continue;
                    }
                    if (start && row.length > 1 && row.some(cell => cell?.trim())) {
                        table.push(row);
                    }
                }
                if (table.length > 0) {
                    result.القوة_البشرية = table;
                    break;
                }
            }
        }

        // التعاقدات
        for (const [name, rows] of Object.entries(allData)) {
            if (name.includes('تعاقدات')) {
                const table = [];
                let start = false;
                for (const row of rows) {
                    if (!start && row.some(cell => cell?.includes('البيان') || cell?.includes('يوجد'))) {
                        start = true;
                        continue;
                    }
                    if (start && row.length > 1 && row.some(cell => cell?.trim())) {
                        table.push(row);
                    }
                }
                if (table.length > 0) {
                    result.التعاقدات = table;
                    break;
                }
            }
        }

        // شبكة الغازات
        for (const [name, rows] of Object.entries(allData)) {
            if (name.includes('شبكة الغازات')) {
                const table = [];
                let start = false;
                for (const row of rows) {
                    if (!start && row.some(cell => cell?.includes('البيان') || cell?.includes('تانك'))) {
                        start = true;
                        continue;
                    }
                    if (start && row.length > 1 && row.some(cell => cell?.trim())) {
                        table.push(row);
                    }
                }
                if (table.length > 0) {
                    result.شبكة_الغازات = table;
                    break;
                }
            }
        }

        return result;

    } catch (error) {
        console.error('❌ خطأ:', error.message);
        return { error: 'فشل في جلب التفاصيل', details: error.message };
    }
}

// =============================================
// API Endpoints
// =============================================

app.get('/api/hospitals', async (req, res) => {
    try {
        const data = await fetchBasicHospitals();
        res.json(data);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.get('/api/hospital/:sheetId', async (req, res) => {
    try {
        const { sheetId } = req.params;
        const details = await fetchHospitalDetails(sheetId);
        res.json(details);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.get('/health', (req, res) => {
    res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// =============================================
// تشغيل السيرفر (لـ Vercel)
// =============================================
if (require.main === module) {
    app.listen(PORT, () => {
        console.log('✅ السيرفر شغال على http://localhost:' + PORT);
    });
}

module.exports = app;