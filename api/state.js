import { kv } from '@vercel/kv';

const KEY = 'milano:state';

const DEFAULT_STATE = {
  deletedFiles: [],
  deletedSheets: [],
  customFiles: {},
  customSheets: {},
  fileOverrides: {},
  sheetOverrides: {}
};

export default async function handler(req, res) {
  // Allow the static front-end to call this from the same site without CORS issues,
  // and keep it simple if it's ever called from elsewhere.
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    if (req.method === 'GET') {
      const state = await kv.get(KEY);
      return res.status(200).json(state || DEFAULT_STATE);
    }

    if (req.method === 'POST') {
      let body = req.body;
      if (typeof body === 'string') {
        try { body = JSON.parse(body); } catch (e) { body = {}; }
      }
      body = body || {};

      // Basic shape guard so a malformed request can't corrupt the store.
      const state = {
        deletedFiles: Array.isArray(body.deletedFiles) ? body.deletedFiles : [],
        deletedSheets: Array.isArray(body.deletedSheets) ? body.deletedSheets : [],
        customFiles: (body.customFiles && typeof body.customFiles === 'object') ? body.customFiles : {},
        customSheets: (body.customSheets && typeof body.customSheets === 'object') ? body.customSheets : {},
        fileOverrides: (body.fileOverrides && typeof body.fileOverrides === 'object') ? body.fileOverrides : {},
        sheetOverrides: (body.sheetOverrides && typeof body.sheetOverrides === 'object') ? body.sheetOverrides : {}
      };

      await kv.set(KEY, state);
      return res.status(200).json({ ok: true });
    }

    return res.status(405).json({ error: 'Method not allowed' });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Server error' });
  }
}
