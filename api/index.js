import express from 'express';
import fetch from 'node-fetch';
import 'dotenv/config';

// Serverless-compatible Express app for Vercel
const app = express();
app.use(express.json({ limit: '1mb' }));

// Health check
app.get('/health', (req, res) => {
  res.status(200).json({ success: true, data: { status: 'ok', time: new Date().toISOString() } });
});

// Helper: get access token via client credentials
async function getAccessToken() {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error('Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET');
  }

  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Token request failed (${resp.status}): ${text}`);
  }

  const json = await resp.json();
  return json.access_token;
}

// Helper: Build Graph base URL for a workbook
function buildWorkbookBase({ driveId, itemId }) {
  if (!driveId || !itemId) {
    throw new Error('driveId and itemId are required');
  }
  return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook`;
}

// Helper: Graph fetch
async function graphFetch(url, options = {}) {
  const token = await getAccessToken();
  const resp = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });

  const contentType = resp.headers.get('content-type') || '';
  const isJson = contentType.includes('application/json');
  const data = isJson ? await resp.json() : await resp.text();
  if (!resp.ok) {
    const msg = typeof data === 'string' ? data : JSON.stringify(data);
    throw new Error(`Graph error (${resp.status}): ${msg}`);
  }
  return data;
}

// POST /excel/read
// Body: { driveId, itemId, sheetName, range }
app.post('/excel/read', async (req, res) => {
  try {
    const { driveId, itemId, sheetName, range } = req.body || {};
    if (!driveId || !itemId || !sheetName || !range) {
      return res.status(400).json({ success: false, error: 'Missing body. Required: driveId, itemId, sheetName, range' });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(range)}')`;
    const data = await graphFetch(url, { method: 'GET' });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// POST /excel/write
// Body: { driveId, itemId, sheetName, range, values } where values is 2D array
app.post('/excel/write', async (req, res) => {
  try {
    const { driveId, itemId, sheetName, range, values } = req.body || {};
    if (!driveId || !itemId || !sheetName || !range || !Array.isArray(values)) {
      return res.status(400).json({ success: false, error: 'Missing body. Required: driveId, itemId, sheetName, range, values(2D array)' });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(range)}')`;
    const data = await graphFetch(url, {
      method: 'PATCH',
      body: JSON.stringify({ values }),
    });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// POST /excel/add-sheet
// Body: { driveId, itemId, sheetName }
app.post('/excel/add-sheet', async (req, res) => {
  try {
    const { driveId, itemId, sheetName } = req.body || {};
    if (!driveId || !itemId || !sheetName) {
      return res.status(400).json({ success: false, error: 'Missing body. Required: driveId, itemId, sheetName' });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets/add`;
    const data = await graphFetch(url, {
      method: 'POST',
      body: JSON.stringify({ name: sheetName }),
    });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// POST /excel/delete
// Clears data in a range
// Body: { driveId, itemId, sheetName, range, applyTo } // applyTo defaults to 'contents' (other options: formats, hyperLinks, etc.)
app.post('/excel/delete', async (req, res) => {
  try {
    const { driveId, itemId, sheetName, range, applyTo = 'contents' } = req.body || {};
    if (!driveId || !itemId || !sheetName || !range) {
      return res.status(400).json({ success: false, error: 'Missing body. Required: driveId, itemId, sheetName, range' });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(sheetName)}')/range(address='${encodeURIComponent(range)}')/clear`;
    const data = await graphFetch(url, {
      method: 'POST',
      body: JSON.stringify({ applyTo }),
    });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// Important for Vercel: export the app as default
export default app;
