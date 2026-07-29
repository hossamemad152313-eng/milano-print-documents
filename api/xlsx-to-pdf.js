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

    // 3) صدّر الملف كله كـ PDF بطلب واحد بس وسريع.
    // اكتشفنا إن رابط تصدير جوجل شيتس لو *ماتكتبش gid خالص* (مش تكرره
    // ومش تسيبه فاضي)، بيصدّر كل الشيتات الظاهرة مرة واحدة في نفس ملف
    // الـ PDF، كل شيت بإعدادات الطباعة (اتجاه/مقاس/منطقة طباعة) بتاعته
    // هو، وده أسرع بكتير من إننا نطلب كل شيت لوحده. المشكلة الأصلية
    // كانت إننا كنا بنكرر gid= لكل شيت في نفس اللينك، وده اللي كان
    // بيخلي جوجل ياخد شيت واحد بس.
    //
    // احتياطًا (لو التصرف ده اتغيّر من جوجل أو ملف معين اتصرف بشكل
    // مختلف)، بعد ما ناخد الـ PDF بنتاكد إن عدد صفحاته يطابق عدد
    // الشيتات الظاهرة. لو مطابق، نرجعه على طول. لو مش مطابق، نرجع
    // لطريقة الاحتياط الأبطأ: نصدّر كل شيت لوحده (بالتتابع مع فواصل
    // زمنية وإعادة محاولة لو حصل 429) ونلزقهم بمكتبة pdf-lib.
    const tokenResult = await auth.getAccessToken();
    const accessToken = (typeof tokenResult === 'string') ? tokenResult : tokenResult.token;

    function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

    async function retryFetch(url, attempts = 6) {
      for (let i = 0; i < attempts; i++) {
        const resp = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
        if (resp.status !== 429) return resp;
        if (i < attempts - 1) {
          // انتظار متزايد بين كل محاولة والتانية (1s, 2s, 4s, 8s, 16s...)
          await sleep(1000 * Math.pow(2, i));
        } else {
          return resp; // خلصت المحاولات، رجّع آخر رد زي ما هو
        }
      }
    }

    let finalPdfBytes = null;

    // --- المحاولة السريعة: طلب واحد بدون gid خالص ---
    try {
      const allSheetsUrl =
        `https://docs.google.com/spreadsheets/d/${fileId}/export` +
        `?format=pdf&size=A4&fitw=true&fith=true&scale=4&gridlines=true&printtitle=false` +
        `&sheetnames=false&pagenumbers=false`;
      const pdfResp = await retryFetch(allSheetsUrl);
      if (pdfResp.ok) {
        const bytes = await pdfResp.arrayBuffer();
        const pdf = await PDFDocument.load(bytes);
        if (pdf.getPageCount() === visibleSheets.length) {
          finalPdfBytes = bytes;
        }
      }
    } catch (fastPathErr) {
      console.error('المحاولة السريعة (بدون gid) فشلت، هنرجع للطريقة الاحتياطية:', fastPathErr);
    }

    // --- الطريقة الاحتياطية: شيت شيت لوحده ثم لزق بمكتبة pdf-lib ---
    if (!finalPdfBytes) {
      const GAP_BETWEEN_REQUESTS_MS = 600; // فترة ثابتة بعد كل طلب ناجح قبل ما نبعت اللي بعده
      const sheetPdfBuffers = [];

      for (const sheetProps of visibleSheets) {
        const exportUrl =
          `https://docs.google.com/spreadsheets/d/${fileId}/export` +
          `?format=pdf&size=A4&fitw=true&fith=true&scale=4&gridlines=true&printtitle=false` +
          `&sheetnames=false&pagenumbers=false&gid=${sheetProps.sheetId}`;

        const pdfResp = await retryFetch(exportUrl);
        if (!pdfResp.ok) {
          if (pdfResp.status === 429) {
            throw new Error(`جوجل رفض تصدير شيت "${sheetProps.title}" بسبب كثرة الطلبات (429) حتى بعد 6 محاولات وانتظار متزايد بينهم. لو الملف فيه شيتات كتير جدًا، جرّبي تقسّميه لملفين أصغر وابعتيهم على مرتين.`);
          }
          let details = '';
          try { details = await pdfResp.text(); } catch (e) { /* ignore */ }
          const cleanDetails = details.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
          throw new Error(`فشل تصدير شيت "${sheetProps.title}" من جوجل (خطأ ${pdfResp.status}). ${cleanDetails.slice(0, 200)}`);
        }
        sheetPdfBuffers.push(await pdfResp.arrayBuffer());
        await sleep(GAP_BETWEEN_REQUESTS_MS);
      }

      const mergedPdf = await PDFDocument.create();
      for (const sheetPdfBytes of sheetPdfBuffers) {
        const sheetPdf = await PDFDocument.load(sheetPdfBytes);
        const copiedPages = await mergedPdf.copyPages(sheetPdf, sheetPdf.getPageIndices());
        copiedPages.forEach(page => mergedPdf.addPage(page));
      }
      finalPdfBytes = await mergedPdf.save();
    }

    res.setHeader('Content-Type', 'application/pdf');
    res.status(200).send(Buffer.from(finalPdfBytes));
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
