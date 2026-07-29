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

    // 3) صدّر الملف كله كـ PDF — تكرار gid لكل شيت بيخلي جوجل يحطهم كلهم
    // في نفس الـ PDF بترتيبهم، وكل شيت بياخد إعدادات الطباعة (اتجاه/مقاس/
    // منطقة طباعة) المتخزنة فيه هو نفسه، بالظبط زي لو فتحتيه وطبعتيه يدوي.
    const gidParams = visibleSheets.map(p => `gid=${p.sheetId}`).join('&');
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${fileId}/export` +
      `?format=pdf&size=A4&fitw=true&gridlines=true&printtitle=false` +
      `&sheetnames=false&pagenumbers=false&${gidParams}`;

    const tokenResult = await auth.getAccessToken();
    const accessToken = (typeof tokenResult === 'string') ? tokenResult : tokenResult.token;
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
