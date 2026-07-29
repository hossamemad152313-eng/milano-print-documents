// /api/xlsx-to-pdf.js
// السيرفر البسيط اللي بياخد ملف الإكسل من الموقع، يرفعه مؤقتًا على Google
// Drive وهو بيتحول تلقائي لـ Google Sheet، بعدين يصدّره كـ PDF (بيحترم
// اتجاه/مقاس/منطقة الطباعة بتاعة كل شيت زي ما هي في الملف الأصلي)،
// وبعد كده يمسح النسخة المؤقتة من درايف، ويرجّع الـ PDF للموقع.
//
// إعداد مطلوب قبل النشر (مرة واحدة بس):
// 1) اعمل مشروع على https://console.cloud.google.com وفعّل:
//    - Google Drive API
//    - Google Sheets API
// 2) اعمل Service Account (APIs & Services → Credentials → Create
//    Credentials → Service Account) ونزّل مفتاحه كـ JSON.
// 3) شارك أي فولدر في Google Drive بتاعك مع إيميل الـ Service Account
//    (موجود جوه ملف الـ JSON، شكله xxx@xxx.iam.gserviceaccount.com)
//    وديله صلاحية Editor — ده الفولدر اللي هنرفع فيه الملفات المؤقتة.
// 4) في مشروع Vercel بتاعك: Settings → Environment Variables
//    - GOOGLE_SERVICE_ACCOUNT_JSON = محتوى ملف الـ JSON كامل (كنص واحد)
//    - GOOGLE_DRIVE_FOLDER_ID = الـ ID بتاع الفولدر اللي شاركته (من رابطه)
//    وأعد النشر (Redeploy).
// 5) لازم تضيف باكدج "googleapis" لمشروعك (npm install googleapis).

import { google } from 'googleapis';
import { Readable } from 'stream';

function bufferToStream(buffer) {
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);
  return stream;
}

function getAuth() {
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!raw) {
    throw new Error('مفيش GOOGLE_SERVICE_ACCOUNT_JSON متظبط على السيرفر. ضيفه من إعدادات Vercel وأعد النشر.');
  }
  let credentials;
  try {
    credentials = JSON.parse(raw);
  } catch (e) {
    throw new Error('محتوى GOOGLE_SERVICE_ACCOUNT_JSON مش JSON صحيح — انسخ ملف المفتاح كامل من غير أي تعديل.');
  }
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/spreadsheets.readonly'
    ]
  });
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'الطريقة دي مش مسموحة، لازم POST' });
    return;
  }

  const { fileBase64, filename } = req.body || {};
  if (!fileBase64) {
    res.status(400).json({ error: 'مفيش ملف اتبعت مع الطلب' });
    return;
  }
  const approxBytes = fileBase64.length * 0.75;
  if (approxBytes > 14 * 1024 * 1024) {
    res.status(413).json({ error: 'الملف أكبر من الحد المسموح — قسّمه لملفين أصغر.' });
    return;
  }

  let auth;
  try {
    auth = getAuth();
  } catch (err) {
    res.status(500).json({ error: err.message });
    return;
  }

  const drive = google.drive({ version: 'v3', auth });
  const sheetsApi = google.sheets({ version: 'v4', auth });
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

  let fileId = null;
  try {
    // 1) ارفع الملف وحوّله لـ Google Sheet في نفس الخطوة
    const createRes = await drive.files.create({
      requestBody: {
        name: (filename || 'workbook').replace(/\.[^.]+$/, ''),
        mimeType: 'application/vnd.google-apps.spreadsheet',
        ...(folderId ? { parents: [folderId] } : {})
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: bufferToStream(Buffer.from(fileBase64, 'base64'))
      },
      fields: 'id'
    });
    fileId = createRes.data.id;

    // 2) اجيب كل الشيتات الظاهرة (مش المخفية) بترتيبها زي ما هي في الملف
    const meta = await sheetsApi.spreadsheets.get({
      spreadsheetId: fileId,
      fields: 'sheets.properties'
    });
    const visibleSheets = (meta.data.sheets || [])
      .map(s => s.properties)
      .filter(p => !p.hidden);

    if (!visibleSheets.length) {
      throw new Error('الملف مفيهوش أي شيتات ظاهرة تتصدّر');
    }

    // 3) صدّر الملف كله كـ PDF — تكرار gid لكل شيت بيخلي جوجل يحطهم كلهم
    // في نفس الـ PDF بترتيبهم، وكل شيت بياخد إعدادات الطباعة (اتجاه/مقاس/
    // منطقة طباعة) المتخزنة فيه هو نفسه، بالظبط زي لو فتحتيه وطبعتيه يدوي.
    const gidParams = visibleSheets.map(p => `gid=${p.sheetId}`).join('&');
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${fileId}/export` +
      `?format=pdf&size=A4&fitw=true&gridlines=true&printtitle=false` +
      `&sheetnames=false&pagenumbers=false&${gidParams}`;

    const accessToken = await auth.getAccessToken();
    const pdfResp = await fetch(exportUrl, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!pdfResp.ok) {
      let details = '';
      try { details = await pdfResp.text(); } catch (e) { /* ignore */ }
      throw new Error(`فشل تصدير الـ PDF من جوجل (خطأ ${pdfResp.status}). ${details.slice(0, 300)}`);
    }

    const pdfBuffer = Buffer.from(await pdfResp.arrayBuffer());
    res.setHeader('Content-Type', 'application/pdf');
    res.status(200).send(pdfBuffer);
  } catch (err) {
    console.error('فشل التحويل عبر جوجل:', err);
    res.status(500).json({ error: (err && err.message) || 'حصل خطأ غير متوقع أثناء التحويل' });
  } finally {
    // 4) امسح النسخة المؤقتة من درايف دايمًا، حتى لو التصدير فشل
    if (fileId) {
      try {
        await drive.files.delete({ fileId });
      } catch (delErr) {
        console.error('تعذر حذف النسخة المؤقتة من درايف:', delErr);
      }
    }
  }
}

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '14mb'
    }
  }
};
