import { put } from '@vercel/blob';

// Accepts a base64 data URL and uploads it to Vercel Blob storage, returning
// the public URL. The front-end stores that URL in the shared JSON state
// (see /api/state.js) instead of the raw base64 data, which is what used to
// make the state blob grow huge and eventually trip the request-size limit.
export const config = {
  api: {
    bodyParser: {
      sizeLimit: '15mb'
    }
  }
};

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    let body = req.body;
    if (typeof body === 'string') {
      try { body = JSON.parse(body); } catch (e) { body = {}; }
    }
    body = body || {};

    const dataUrl = body.dataUrl;
    if (!dataUrl || typeof dataUrl !== 'string') {
      return res.status(400).json({ error: 'Missing dataUrl' });
    }

    const match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
    if (!match) {
      return res.status(400).json({ error: 'dataUrl must be a base64 image data URL' });
    }
    const mimeType = match[1];
    const base64Data = match[2];
    const buffer = Buffer.from(base64Data, 'base64');

    const ext = (mimeType.split('/')[1] || 'jpg').replace(/[^a-z0-9]/gi, '') || 'jpg';
    const filename = `sheets/${Date.now()}_${Math.random().toString(36).slice(2, 8)}.${ext}`;

    const blob = await put(filename, buffer, {
      access: 'public',
      contentType: mimeType
    });

    return res.status(200).json({ url: blob.url });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Server error' });
  }
}
