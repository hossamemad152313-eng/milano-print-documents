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
import JSZip from 'jszip';

function bufferToStream(buffer) {
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);
  return stream;
}

// بيقرا إعدادات الطباعة الحقيقية (اتجاه/مقاس ورق/هل مفعّل "ملائمة الصفحة")
// المخزّنة جوه ملف الإكسل الأصلي نفسه لكل شيت، بدل ما نفترض إعدادات ثابتة.
// ملف الـ xlsx هو zip فيه ملفات XML؛ إعدادات الطباعة موجودة في
// xl/worksheets/sheetN.xml جوه تاج <pageSetup>.
async function extractPageSetups(fileBuffer) {
  const result = {}; // sheetName -> { portrait, sizeCode, fitToPage }
  try {
    const zip = await JSZip.loadAsync(fileBuffer);

    const workbookXmlFile = zip.file('xl/workbook.xml');
    const relsXmlFile = zip.file('xl/_rels/workbook.xml.rels');
    if (!workbookXmlFile || !relsXmlFile) return result;

    const workbookXml = await workbookXmlFile.async('string');
    const relsXml = await relsXmlFile.async('string');

    // اسم الشيت -> r:id بتاعه (من ترتيب <sheet .../> جوه workbook.xml)
    const nameToRid = {};
    for (const tag of workbookXml.match(/<sheet\b[^>]*\/>/g) || []) {
      const name = (tag.match(/name="([^"]*)"/) || [])[1];
      const rid = (tag.match(/r:id="([^"]*)"/) || [])[1];
      if (name && rid) nameToRid[name] = rid;
    }

    // r:id -> اسم ملف الـ XML بتاع الشيت (worksheets/sheetX.xml)
    const ridToTarget = {};
    for (const tag of relsXml.match(/<Relationship\b[^>]*\/>/g) || []) {
      const id = (tag.match(/Id="([^"]*)"/) || [])[1];
      const target = (tag.match(/Target="([^"]*)"/) || [])[1];
      if (id && target) ridToTarget[id] = target;
    }

    // مطابقة أكواد مقاس الورق بتاعة Excel (ECMA-376) بأكواد مقاس التصدير
    // بتاعة جوجل شيتس (0=Letter,2=Legal,6=A3,7=A4,8=A5)
    const paperSizeMap = { 1: 0, 5: 2, 8: 6, 9: 7, 11: 8 };

    for (const [sheetName, rid] of Object.entries(nameToRid)) {
      const target = ridToTarget[rid];
      if (!target) continue;
      const normalizedPath = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
      const sheetFile = zip.file(normalizedPath) || zip.file(target);
      if (!sheetFile) continue;

      const sheetXml = await sheetFile.async('string');
      const pageSetupTag = (sheetXml.match(/<pageSetup\b[^>]*\/>/) || [])[0] || '';
      const fitToPageTag = sheetXml.match(/<pageSetUpPr\b[^>]*fitToPage="(1|true)"[^>]*\/>/);

      const orientation = (pageSetupTag.match(/orientation="([^"]*)"/) || [])[1] || 'portrait';
      const paperSizeCodeRaw = (pageSetupTag.match(/paperSize="([^"]*)"/) || [])[1];

      result[sheetName] = {
        portrait: orientation !== 'landscape',
        sizeCode: paperSizeMap[parseInt(paperSizeCodeRaw, 10)] ?? 7, // افتراضي A4
        fitToPage: !!fitToPageTag
      };
    }
  } catch (err) {
    console.error('تعذر قراءة إعدادات الطباعة من الملف الأصلي، هنستخدم إعدادات افتراضية (A4 طولي، ملائمة صفحة):', err);
  }
  return result;
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
    const fileBuffer = Buffer.from(fileBase64, 'base64');
    const pageSetups = await extractPageSetups(fileBuffer);
    const DEFAULT_CONFIG = { portrait: true, sizeCode: 7, fitToPage: true };

    // 1) ارفع الملف وحوّله لـ Google Sheet في نفس الخطوة
    const createRes = await drive.files.create({
      requestBody: {
        name: (filename || 'workbook').replace(/\.[^.]+$/, ''),
        mimeType: 'application/vnd.google-apps.spreadsheet',
        ...(folderId ? { parents: [folderId] } : {})
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: bufferToStream(fileBuffer)
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

    // 3) صدّر الملف كـ PDF، باستخدام إعدادات الطباعة الحقيقية لكل شيت
    // (الاتجاه/مقاس الورق/هل الشيت مظبوط "ملائمة صفحة") اللي قرأناها من
    // الملف الأصلي فوق. لو كل الشيتات الظاهرة بنفس الإعدادات بالظبط،
    // بنعمل طلب واحد سريع للملف كله (بدون gid — جوجل بيصدّر وقتها كل
    // الشيتات الظاهرة مرة واحدة). لو الإعدادات مختلفة من شيت لشيت (زي لو
    // شيت طولي وشيت عرضي في نفس الملف)، لازم نصدّر كل شيت لوحده بإعداداته
    // الصح، لإن الطلب الواحد مش هيقدر يطبّق أكتر من اتجاه/مقاس في نفس
    // الوقت.
    const sheetConfigs = visibleSheets.map(p => pageSetups[p.title] || DEFAULT_CONFIG);
    const allSameConfig = sheetConfigs.every(c =>
      c.portrait === sheetConfigs[0].portrait &&
      c.sizeCode === sheetConfigs[0].sizeCode &&
      c.fitToPage === sheetConfigs[0].fitToPage
    );

    function buildExportUrl(cfg, gid) {
      const fitParams = cfg.fitToPage ? 'fitw=true&fith=true&scale=4' : 'scale=1';
      // هوامش شبه معدومة (بوصة) عشان المحتوى ياخد أكبر مساحة ممكنة من
      // الورقة بدل ما يفضل فراغ أبيض كبير حوالين الحواف.
      const margins = 'top_margin=0.05&bottom_margin=0.05&left_margin=0.05&right_margin=0.05';
      return `https://docs.google.com/spreadsheets/d/${fileId}/export` +
        `?format=pdf&size=${cfg.sizeCode}&portrait=${cfg.portrait}&${fitParams}&${margins}` +
        `&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false` +
        (gid !== undefined ? `&gid=${gid}` : '');
    }

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

    // --- المحاولة السريعة: طلب واحد بدون gid (بس بس لو كل الشيتات متطابقة الإعدادات) ---
    if (allSameConfig) {
      try {
        const allSheetsUrl = buildExportUrl(sheetConfigs[0]);
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
    }

    // --- الطريقة الاحتياطية: شيت شيت لوحده بإعداداته الخاصة ثم لزق بمكتبة pdf-lib ---
    if (!finalPdfBytes) {
      const GAP_BETWEEN_REQUESTS_MS = 600; // فترة ثابتة بعد كل طلب ناجح قبل ما نبعت اللي بعده
      const sheetPdfBuffers = [];

      for (let i = 0; i < visibleSheets.length; i++) {
        const sheetProps = visibleSheets[i];
        const cfg = sheetConfigs[i];
        const exportUrl = buildExportUrl(cfg, sheetProps.sheetId);

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
