// /api/xlsx-to-pdf.js
// السيرفر البسيط اللي بياخد ملف الإكسل من الموقع، ويبعته لخدمة CloudConvert
// (اللي بتستخدم محرك إكسل حقيقي للتحويل)، ويرجّع PDF واحد فيه كل الشيتات
// مرتبة زي ما هي في الملف الأصلي، كل شيت بمقاسه واتجاهه الحقيقي.
//
// إعداد مطلوب قبل النشر (مرة واحدة بس):
// 1) اعمل حساب مجاني على https://cloudconvert.com
// 2) من "API Keys" اعمل مفتاح جديد بصلاحيات task.read + task.write
// 3) في مشروع Vercel بتاعك: Settings → Environment Variables
//    أضف متغير اسمه CLOUDCONVERT_API_KEY وقيمته المفتاح ده، وأعد النشر (Redeploy)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'الطريقة دي مش مسموحة، لازم POST' });
    return;
  }

  const apiKey = process.env.CLOUDCONVERT_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: 'مفيش مفتاح CLOUDCONVERT_API_KEY متظبط على السيرفر. ضيفه من إعدادات Vercel وأعد النشر.' });
    return;
  }

  const { fileBase64, filename } = req.body || {};
  if (!fileBase64) {
    res.status(400).json({ error: 'مفيش ملف اتبعت مع الطلب' });
    return;
  }
  // نفس الحد بتاع 10 ميجا اللي CloudConvert نفسها بتنصح بيه لطريقة الرفع دي (base64 مباشر داخل الطلب)
  const approxBytes = fileBase64.length * 0.75;
  if (approxBytes > 10 * 1024 * 1024) {
    res.status(413).json({ error: 'الملف أكبر من 10 ميجا — قسّمه لملفين أصغر.' });
    return;
  }

  try {
    const jobPayload = {
      tasks: {
        'import-file': {
          operation: 'import/base64',
          file: fileBase64,
          filename: filename || 'workbook.xlsx'
        },
        'convert-file': {
          operation: 'convert',
          input: 'import-file',
          input_format: 'xlsx',
          output_format: 'pdf',
          engine: 'libreoffice'
        },
        'export-file': {
          operation: 'export/url',
          input: 'convert-file'
        }
      },
      redirect: true
    };

    const ccResponse = await fetch('https://sync.api.cloudconvert.com/v2/jobs', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(jobPayload),
      redirect: 'follow'
    });

    if (!ccResponse.ok) {
      let details = '';
      try { details = await ccResponse.text(); } catch (e) { /* ignore */ }
      res.status(502).json({
        error: 'فشل التحويل عند CloudConvert (خطأ ' + ccResponse.status + '). ممكن الملف فيه مشكلة أو المفتاح غلط.',
        details
      });
      return;
    }

    const pdfArrayBuffer = await ccResponse.arrayBuffer();
    res.setHeader('Content-Type', 'application/pdf');
    res.status(200).send(Buffer.from(pdfArrayBuffer));
  } catch (err) {
    res.status(500).json({ error: 'حصل خطأ غير متوقع أثناء التحويل: ' + (err && err.message) });
  }
}

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '14mb'
    }
  }
};
