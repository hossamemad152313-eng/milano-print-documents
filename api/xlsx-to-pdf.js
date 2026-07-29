// /api/xlsx-to-pdf.js
// السيرفر البسيط اللي بياخد ملف الإكسل من الموقع، يرفعه مؤقتًا على Google
// Drive وهو بيتحول تلقائي لـ Google Sheet، بعدين يصدّره كـ PDF (بيحترم
// اتجاه/مقاس/منطقة الطباعة بتاعة كل شيت زي ما هي في الملف الأصلي)،
// وبعد كده يمسح النسخة المؤقتة من درايف، ويرجّع الـ PDF للموقع.
//
// بيستخدم OAuth 2.0 بحساب Google شخصي عادي (مش Service Account) عشان
// الملفات تتحسب على مساحة تخزين حساب حقيقي (الـ Service Accounts
// مساحتها صفر على حسابات Gmail العادية).
//
// إعداد مطلوب قبل النشر (مرة واحدة بس):
// 1) اعمل مشروع على https://console.cloud.google.com وفعّل:
//    - Google Drive API
//    - Google Sheets API
// 2) اعمل OAuth Client (Web application) واحصل على refresh token عن طريق
//    https://developers.google.com/oauthplayground (Authorize APIs بصلاحيات
//    drive و spreadsheets، بعدين Exchange authorization code for tokens).
// 3) في مشروع Vercel بتاعك: Settings → Environment Variables ضيفي:
//    - GOOGLE_OAUTH_CLIENT_ID
//    - GOOGLE_OAUTH_CLIENT_SECRET
//    - GOOGLE_OAUTH_REFRESH_TOKEN
//    - GOOGLE_DRIVE_FOLDER_ID = الـ ID بتاع فولدر في درايفك هيتحطوا فيه الملفات المؤقتة
//    وأعد النشر (Redeploy).
// 4) لازم تضيف باكدج "googleapis" لمشروعك (npm install googleapis).

import { google } from 'googleapis';
import { Readable } from 'stream';
import { PDFDocument } from 'pdf-lib';

function bufferToStream(buffer) {
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);
  return stream;
}

function getAuth() {
  const clientId = process.env.GOOGLE_OAUTH_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_OAUTH_CLIENT_SECRET;
  const refreshToken = process.env.GOOGLE_OAUTH_REFRESH_TOKEN;
  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('مفيش GOOGLE_OAUTH_CLIENT_ID / GOOGLE_OAUTH_CLIENT_SECRET / GOOGLE_OAUTH_REFRESH_TOKEN متظبطين على السيرفر. ضيفهم من إعدادات Vercel وأعد النشر.');
  }
  const oauth2Client = new google.auth.OAuth2(clientId, clientSecret);
  oauth2Client.setCredentials({ refresh_token: refreshToken });
  return oauth2Client;
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

    // 3) صدّر كل شيت لوحده كـ PDF منفصل (صفحة واحدة).
    // ملحوظة مهمة: رابط تصدير جوجل شيتس بيتجاهل تكرار gid= في نفس اللينك
    // وبياخد شيت واحد بس مهما كررنا الباراميتر ده — ده كان سبب المشكلة
    // اللي كل النماذج مش بتتصدّر. الحل: طلب منفصل لكل شيت بترتيبه، وكل
    // طلب بياخد إعدادات الطباعة (اتجاه/مقاس/منطقة طباعة) المتخزنة في
    // الشيت نفسه، بالظبط زي لو فتحتيه وطبعتيه يدوي.
    //
    // مهم: لو طلبنا كل الشيتات مرة واحدة بالتوازي، جوجل بيرفض بعضها بخطأ
    // 429 (Too Many Requests). عشان كده بنحدد أقصى عدد طلبات شغالة في نفس
    // اللحظة (CONCURRENCY)، وأي طلب يرجع بـ 429 بنعيده تاني بعد انتظار
    // قصير (retryFetch) بدل ما نفشل على طول.
    const tokenResult = await auth.getAccessToken();
    const accessToken = (typeof tokenResult === 'string') ? tokenResult : tokenResult.token;

    async function retryFetch(url, attempts = 4) {
      for (let i = 0; i < attempts; i++) {
        const resp = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
        if (resp.status !== 429) return resp;
        // كل محاولة فاشلة بتستنى أكتر من اللي قبلها (400ms, 800ms, 1600ms...)
        await new Promise(r => setTimeout(r, 400 * Math.pow(2, i)));
      }
      // آخر محاولة، أيًا كانت نتيجتها بنرجعها زي ما هي عشان تتحقق من status بعدين
      return fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    }

    const CONCURRENCY = 3;
    const sheetPdfBuffers = new Array(visibleSheets.length);
    let nextIndex = 0;

    async function worker() {
      while (nextIndex < visibleSheets.length) {
        const i = nextIndex++;
        const sheetProps = visibleSheets[i];
        const exportUrl =
          `https://docs.google.com/spreadsheets/d/${fileId}/export` +
          `?format=pdf&size=A4&fitw=true&gridlines=true&printtitle=false` +
          `&sheetnames=false&pagenumbers=false&gid=${sheetProps.sheetId}`;

        const pdfResp = await retryFetch(exportUrl);
        if (!pdfResp.ok) {
          let details = '';
          try { details = await pdfResp.text(); } catch (e) { /* ignore */ }
          throw new Error(`فشل تصدير شيت "${sheetProps.title}" من جوجل (خطأ ${pdfResp.status}). ${details.slice(0, 300)}`);
        }
        sheetPdfBuffers[i] = await pdfResp.arrayBuffer();
      }
    }

    await Promise.all(Array.from({ length: Math.min(CONCURRENCY, visibleSheets.length) }, worker));

    const mergedPdf = await PDFDocument.create();
    for (const sheetPdfBytes of sheetPdfBuffers) {
      const sheetPdf = await PDFDocument.load(sheetPdfBytes);
      const copiedPages = await mergedPdf.copyPages(sheetPdf, sheetPdf.getPageIndices());
      copiedPages.forEach(page => mergedPdf.addPage(page));
    }

    // 4) نلزّق كل الصفحات دي في ملف نهائي واحد ونرجّعه للموقع، بنفس ترتيب
    // الشيتات، عشان التقطيع بمكتبة pdf.js في المتصفح يطابق أسامي الشيتات.
    const mergedBytes = await mergedPdf.save();
    res.setHeader('Content-Type', 'application/pdf');
    res.status(200).send(Buffer.from(mergedBytes));
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
  },
  maxDuration: 60
  // ملحوظة: الرقم ده شغال عادي على خطة Hobby كمان (مش لازم Pro) لإن
  // Vercel بقت مفعّلة Fluid Compute افتراضيًا، وده بيرفع أقصى مدة تنفيذ
  // على Hobby لحد 300 ثانية. سبب أي فشل سريع مش هيبقى الوقت، هيبقى الـ
  // 429 من جوجل لو الشيتات كتير قوي ومحتاجة إعادة محاولة أكتر.
};
