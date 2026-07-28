// /api/xlsx-to-pdf.js
// السيرفر البسيط اللي بياخد ملف الإكسل من الموقع، يصلّح مشكلة اتجاه النص
// العربي (النقطتين والقوسين اللي بيتقلبوا) تلقائيًا، وبعدين يبعت الملف
// المُصلّح لخدمة CloudConvert (اللي بتستخدم محرك إكسل حقيقي للتحويل)،
// ويرجّع PDF واحد فيه كل الشيتات مرتبة زي ما هي في الملف الأصلي، كل شيت
// بمقاسه واتجاهه الحقيقي.
//
// إعداد مطلوب قبل النشر (مرة واحدة بس):
// 1) اعمل حساب مجاني على https://cloudconvert.com
// 2) من "API Keys" اعمل مفتاح جديد بصلاحيات task.read + task.write
// 3) في مشروع Vercel بتاعك: Settings → Environment Variables
//    أضف متغير اسمه CLOUDCONVERT_API_KEY وقيمته المفتاح ده، وأعد النشر (Redeploy)
// 4) لازم تضيف باكدج "jszip" لمشروعك (npm install jszip) — دي بتفتح ملف
//    الإكسل كأرشيف zip عشان نصلّح النصوص من غير ما نلمس أي تنسيق أو شكل.

import JSZip from 'jszip';

const RLM = '\u200F'; // Right-to-Left Mark — بيقول للبرنامج "الرمز اللي جنبي RTL" صراحةً
const ARABIC = '\\u0600-\\u06FF';

// الرموز "المحايدة" اللي بتتقلب لما تكون ملاصقة لنص عربي. مقسّمة لمجموعتين
// لأن اتجاه الإصلاح بيختلف حسب نوع الرمز:
const OPEN_BRACKETS  = '\\(\\[\\{';           // محتاجة RLM بعدها لو اللي جاي بعدها عربي
const CLOSE_BRACKETS = '\\)\\]\\}';           // محتاجة RLM قبلها لو اللي قبلها عربي
const SYMMETRIC       = ':;,./\\\\%+\\-"\''; // بتتصلح من أي ناحية بتلمس فيها عربي: : ؛ , . / \\ % + - " '

// بيدور على أي رمز محايد (نقطتين، قوس، سلاش، شرطة، علامة تنصيص... الخ)
// ملاصق لنص عربي، ويحط RLM جنبه من ناحية العربي فقط — بالظبط زي الحل اليدوي
// اللي جربناه بالفورمولا، بس هنا بيتطبق تلقائي على كل النصوص في الملف مرة
// واحدة، وعلى مجموعة أوسع من الرموز مش بس النقطتين والقوس المدور.
// أي رقم أو تاريخ أو وقت (زي 10:30 أو 12/07/2026) بيفضل زي ما هو، لأن
// الشرط بيتحقق من إن الجار الفعلي حرف عربي مش رقم.
function fixArabicNeutralChars(text) {
  if (!text || typeof text !== 'string') return text;
  if (!new RegExp(`[${ARABIC}]`).test(text)) return text; // مفيش عربي، سيبه زي ما هو

  let result = text;

  // قاعدة "قبل الرمز": عربي + مسافات (لو فيه) + رمز محايد  ->  حط RLM قبل الرمز.
  // بتغطي: الأقواس القافلة ) ] } وكل الرموز المتماثلة : ؛ , . / \ % + - " '
  const beforeClass = CLOSE_BRACKETS + SYMMETRIC;
  result = result.replace(
    new RegExp(`([${ARABIC}])([ \\t]*)(?!${RLM})([${beforeClass}])`, 'g'),
    `$1$2${RLM}$3`
  );

  // قاعدة "بعد الرمز": رمز محايد + مسافات (لو فيه) + عربي  ->  حط RLM بعد الرمز.
  // بتغطي: الأقواس الفاتحة ( [ { ونفس الرموز المتماثلة (عشان تتصلح من
  // الناحية التانية لو هي اللي لامسة العربي).
  const afterClass = OPEN_BRACKETS + SYMMETRIC;
  result = result.replace(
    new RegExp(`([${afterClass}])(?!${RLM})([ \\t]*)([${ARABIC}])`, 'g'),
    `$1${RLM}$2$3`
  );

  return result;
}

// بيصلّح كل الـ <t>...</t> اللي جوه ملف XML واحد (شيت أو shared strings)،
// من غير ما يلمس أي حاجة تانية في الملف (تنسيق، ألوان، حدود، الخ).
function fixXmlTextNodes(xml) {
  return xml.replace(/(<t[^>]*>)([\s\S]*?)(<\/t>)/g, (match, open, inner, close) => {
    const fixed = fixArabicNeutralChars(inner);
    return `${open}${fixed}${close}`;
  });
}

async function fixArabicPunctuationInWorkbook(fileBuffer) {
  const zip = await JSZip.loadAsync(fileBuffer);

  // shared strings هو المكان اللي بيتخزن فيه معظم النصوص في ملف الإكسل
  const sharedStringsPath = 'xl/sharedStrings.xml';
  if (zip.file(sharedStringsPath)) {
    const xml = await zip.file(sharedStringsPath).async('string');
    zip.file(sharedStringsPath, fixXmlTextNodes(xml));
  }

  // بعض الخلايا بتتخزن كـ "inline string" جوه ملف الشيت نفسه بدل shared strings
  const sheetFiles = Object.keys(zip.files).filter(name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));
  for (const sheetPath of sheetFiles) {
    const xml = await zip.file(sheetPath).async('string');
    zip.file(sheetPath, fixXmlTextNodes(xml));
  }

  return zip.generateAsync({ type: 'nodebuffer' });
}

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

  let fixedFileBase64 = fileBase64;
  // TEMP DEBUG BUILD: دالة تصحيح النص العربي متعطلة عمدًا في النسخة دي
  // عشان نختبر لو هي سبب مشكلة تراكب النص. النسخة دي بس للاختبار —
  // لو التحويل طلع سليم كده، يبقى المشكلة فعلاً في الدالة دي وهنصلّحها
  // من غير ما نلغيها خالص.
  const ARABIC_FIX_ENABLED = false;
  if (ARABIC_FIX_ENABLED) {
    try {
      const originalBuffer = Buffer.from(fileBase64, 'base64');
      const fixedBuffer = await fixArabicPunctuationInWorkbook(originalBuffer);
      fixedFileBase64 = fixedBuffer.toString('base64');
    } catch (err) {
      // لو الإصلاح فشل لأي سبب (مثلاً ملف مش xlsx عادي)، نكمل بالملف الأصلي
      // بدل ما نوقف التحويل بالكامل
      console.error('تعذر تصحيح اتجاه النص العربي تلقائيًا، هيتم التحويل بالملف الأصلي:', err);
    }
  }

  try {
    const jobPayload = {
      tasks: {
        'import-file': {
          operation: 'import/base64',
          file: fixedFileBase64,
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
