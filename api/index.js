import express from "express";
import fetch from "node-fetch";
import "dotenv/config";

// Serverless-compatible Express app for Vercel
const app = express();
app.use(express.json({ limit: "1mb" }));

// Health check
app.get("/health", (req, res) => {
  res.status(200).json({
    success: true,
    data: { status: "ok", time: new Date().toISOString() },
  });
});

// Helper: get access token via client credentials
let cachedToken = null;
let tokenExpiresAt = 0; // epoch ms
let refreshPromise = null; // to avoid concurrent refreshes

async function getAccessToken() {
  // Prefer AZURE_* envs, fallback to legacy names if present
  const TENANT_ID = process.env.AZURE_TENANT_ID || process.env.TENANT_ID;
  const CLIENT_ID = process.env.AZURE_CLIENT_ID || process.env.CLIENT_ID;
  const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || process.env.CLIENT_SECRET;

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error(
      "Missing required environment variables: AZURE_TENANT_ID/AZURE_CLIENT_ID/AZURE_CLIENT_SECRET (or TENANT_ID/CLIENT_ID/CLIENT_SECRET)"
    );
  }

  const now = Date.now();
  const safetyWindowMs = 60_000; // refresh 60s before expiry
  if (cachedToken && now < tokenExpiresAt - safetyWindowMs) {
    return cachedToken;
  }

  if (refreshPromise) {
    // Another request is already refreshing the token; await it
    return refreshPromise;
  }

  refreshPromise = (async () => {
    const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    console.log(`[Auth] Fetching new Graph token for tenant ${TENANT_ID}, client ${CLIENT_ID}.`);
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });

    if (!resp.ok) {
      const text = await resp.text();
      console.error(`[Auth] Token request failed (${resp.status}).`);
      throw new Error(`Token request failed (${resp.status}): ${text}`);
    }

    const json = await resp.json();
    const expiresInSec = Number(json.expires_in) || 3600;
    cachedToken = json.access_token;
    tokenExpiresAt = Date.now() + expiresInSec * 1000;
    console.log(`[Auth] Token acquired. Expires in ~${expiresInSec}s.`);
    return cachedToken;
  })();

  try {
    const token = await refreshPromise;
    return token;
  } finally {
    // Ensure we clear the promise so future refreshes can occur
    refreshPromise = null;
  }
}

// Helper: Build Graph base URL for a workbook
function buildWorkbookBase({ driveId, itemId }) {
  if (!driveId || !itemId) {
    throw new Error("driveId and itemId are required");
  }
  return `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    driveId
  )}/items/${encodeURIComponent(itemId)}/workbook`;
}

// Helper: Graph fetch with auto token handling and single retry on 401
async function graphFetch(url, options = {}) {
  const makeRequest = async () => {
    const token = await getAccessToken();
    return fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });
  };

  let resp = await makeRequest();
  if (resp.status === 401) {
    // Clear cache and retry once
    console.warn(`[Auth] Received 401 from Graph. Clearing cached token and retrying once...`);
    cachedToken = null;
    tokenExpiresAt = 0;
    resp = await makeRequest();
  }

  const contentType = resp.headers.get("content-type") || "";
  const isJson = contentType.includes("application/json");
  const data = isJson ? await resp.json() : await resp.text();
  if (!resp.ok) {
    const msg = typeof data === "string" ? data : JSON.stringify(data);
    throw new Error(`Graph error (${resp.status}): ${msg}`);
  }
  return data;
}

// Helpers to resolve driveId/itemId from names
async function resolveSiteId() {
  const hostname = process.env.SHAREPOINT_HOSTNAME;
  const siteName = process.env.SHAREPOINT_SITE_NAME;
  if (!hostname || !siteName) {
    throw new Error(
      "Missing SHAREPOINT_HOSTNAME or SHAREPOINT_SITE_NAME env for site resolution"
    );
  }
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(
    hostname
  )}:/sites/${encodeURIComponent(siteName)}?$select=id`;
  return graphFetch(url, { method: "GET" });
}

async function resolveDriveIdByName(driveName) {
  const site = await resolveSiteId();
  const siteId = site.id;
  const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(
    siteId
  )}/drives`;
  const data = await graphFetch(url, { method: "GET" });
  const match = (data.value || []).find((d) => d.name === driveName);
  if (!match) {
    return null;
  }
  return match.id;
}

async function resolveItemIdByName(driveId, itemName) {
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(
    driveId
  )}/root/children?$select=id,name&$top=999`;
  const data = await graphFetch(url, { method: "GET" });
  const match = (data.value || []).find((it) => it.name === itemName);
  if (!match) {
    return null;
  }
  return match.id;
}

function parseSheetAndAddress(range) {
  const str = String(range || "");
  const idx = str.indexOf("!");
  if (idx > 0) {
    return { sheetName: str.slice(0, idx), address: str.slice(idx + 1) };
  }
  return { sheetName: null, address: str };
}

// POST /excel/read
// Body: supports either { driveId, itemId, sheetName, range } or { driveName, itemName, range (optionally Sheet!A1:B2) }
app.post("/excel/read", async (req, res) => {
  try {
    let { driveId, itemId, driveName, itemName, sheetName, range } = req.body || {};
    if ((!driveId || !itemId) && (driveName && itemName)) {
      const resolvedDriveId = await resolveDriveIdByName(driveName);
      if (!resolvedDriveId) {
        return res.status(404).json({ success: false, error: `Drive not found: ${driveName}` });
      }
      const resolvedItemId = await resolveItemIdByName(resolvedDriveId, itemName);
      if (!resolvedItemId) {
        return res.status(404).json({ success: false, error: `Item not found: ${itemName}` });
      }
      driveId = resolvedDriveId;
      itemId = resolvedItemId;
    }

    if (!driveId || !itemId || !range) {
      return res.status(400).json({
        success: false,
        error: "Missing body. Required: (driveId+itemId) or (driveName+itemName), and range",
      });
    }

    // Support Sheet!A1:B2
    const parsed = parseSheetAndAddress(range);
    if (parsed.sheetName && !sheetName) sheetName = parsed.sheetName;
    const address = parsed.address;
    if (!sheetName) {
      return res.status(400).json({ success: false, error: "sheetName is required (or prefix range as Sheet!A1:B2)" });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(
      sheetName
    )}')/range(address='${encodeURIComponent(address)}')`;
    const data = await graphFetch(url, { method: "GET" });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// POST /excel/write
// Body: supports ids or names; range may be sheet-qualified; values is 2D array
app.post("/excel/write", async (req, res) => {
  try {
    let { driveId, itemId, driveName, itemName, sheetName, range, values } = req.body || {};
    if ((!driveId || !itemId) && (driveName && itemName)) {
      const resolvedDriveId = await resolveDriveIdByName(driveName);
      if (!resolvedDriveId) {
        return res.status(404).json({ success: false, error: `Drive not found: ${driveName}` });
      }
      const resolvedItemId = await resolveItemIdByName(resolvedDriveId, itemName);
      if (!resolvedItemId) {
        return res.status(404).json({ success: false, error: `Item not found: ${itemName}` });
      }
      driveId = resolvedDriveId;
      itemId = resolvedItemId;
    }

    const parsed = parseSheetAndAddress(range);
    if (parsed.sheetName && !sheetName) sheetName = parsed.sheetName;
    const address = parsed.address;

    if (!driveId || !itemId || !sheetName || !address || !Array.isArray(values)) {
      return res.status(400).json({
        success: false,
        error:
          "Missing body. Required: (driveId+itemId) or (driveName+itemName), sheetName (or prefix range), range, values(2D array)",
      });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(
      sheetName
    )}')/range(address='${encodeURIComponent(address)}')`;
    const data = await graphFetch(url, {
      method: "PATCH",
      body: JSON.stringify({ values }),
    });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// POST /excel/add-sheet
// Body: supports ids or names
app.post("/excel/add-sheet", async (req, res) => {
  try {
    let { driveId, itemId, driveName, itemName, sheetName } = req.body || {};
    if ((!driveId || !itemId) && (driveName && itemName)) {
      const resolvedDriveId = await resolveDriveIdByName(driveName);
      if (!resolvedDriveId) {
        return res.status(404).json({ success: false, error: `Drive not found: ${driveName}` });
      }
      const resolvedItemId = await resolveItemIdByName(resolvedDriveId, itemName);
      if (!resolvedItemId) {
        return res.status(404).json({ success: false, error: `Item not found: ${itemName}` });
      }
      driveId = resolvedDriveId;
      itemId = resolvedItemId;
    }

    if (!driveId || !itemId || !sheetName) {
      return res.status(400).json({
        success: false,
        error: "Missing body. Required: driveId, itemId, sheetName",
      });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets/add`;
    const data = await graphFetch(url, {
      method: "POST",
      body: JSON.stringify({ name: sheetName }),
    });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// POST /excel/delete
// Clears data in a range
// Body: supports ids or names; applyTo defaults to 'contents'
app.post("/excel/delete", async (req, res) => {
  try {
    let {
      driveId,
      itemId,
      driveName,
      itemName,
      sheetName,
      range,
      applyTo = "contents",
    } = req.body || {};

    if ((!driveId || !itemId) && (driveName && itemName)) {
      const resolvedDriveId = await resolveDriveIdByName(driveName);
      if (!resolvedDriveId) {
        return res.status(404).json({ success: false, error: `Drive not found: ${driveName}` });
      }
      const resolvedItemId = await resolveItemIdByName(resolvedDriveId, itemName);
      if (!resolvedItemId) {
        return res.status(404).json({ success: false, error: `Item not found: ${itemName}` });
      }
      driveId = resolvedDriveId;
      itemId = resolvedItemId;
    }

    if (!driveId || !itemId || !sheetName || !range) {
      return res.status(400).json({
        success: false,
        error: "Missing body. Required: driveId, itemId, sheetName, range",
      });
    }

    const base = buildWorkbookBase({ driveId, itemId });
    const url = `${base}/worksheets('${encodeURIComponent(
      sheetName
    )}')/range(address='${encodeURIComponent(range)}')/clear`;
    const data = await graphFetch(url, {
      method: "POST",
      body: JSON.stringify({ applyTo }),
    });
    return res.json({ success: true, data });
  } catch (err) {
    return res.status(500).json({ success: false, error: err.message });
  }
});

// Important for Vercel: export the app as default
export default app;
