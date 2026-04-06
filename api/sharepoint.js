const https = require('https');

const CONFIG = {
  tenantId: '81c593c6-a9c4-4bee-bf8b-3c2d75e6675a',
  clientId: 'b0862ae0-fa71-4b2c-ae5e-5a3d0a31746a',
  clientSecret: '1Dx8Q~q3nBod8xHQ~ZZgjdIdVQfMnTlDzqtgvchD',
  siteUrl: 'https://ashalholding.sharepoint.com/sites/DFSales',
  resource: 'https://ashalholding.sharepoint.com'
};

async function fetchJson(url, options = {}) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const reqOptions = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: options.method || 'GET',
      headers: options.headers || {}
    };
    const req = https.request(reqOptions, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve({ ok: res.statusCode < 300, status: res.statusCode, data: JSON.parse(data) }); }
        catch(e) { resolve({ ok: res.statusCode < 300, status: res.statusCode, data: {} }); }
      });
    });
    req.on('error', reject);
    if (options.body) req.write(options.body);
    req.end();
  });
}

async function getToken() {
  const url = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: CONFIG.clientId,
    client_secret: CONFIG.clientSecret,
    scope: `${CONFIG.resource}/.default`
  }).toString();
  const res = await fetchJson(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  return res.data.access_token;
}

async function getDigest(token) {
  const res = await fetchJson(`${CONFIG.siteUrl}/_api/contextinfo`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json;odata=verbose',
      'Content-Length': '0'
    }
  });
  return res.data?.d?.GetContextWebInformation?.FormDigestValue || '';
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const token = await getToken();
    const { action, list, filter, id, item } = req.body || {};

    if (action === 'get') {
      let url = `${CONFIG.siteUrl}/_api/lists/getbytitle('${encodeURIComponent(list)}')/items?$top=500`;
      if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
      const result = await fetchJson(url, {
        headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json;odata=verbose' }
      });
      return res.status(200).json({ ok: true, data: result.data?.d?.results || [] });
    }

    if (action === 'post') {
      const metaRes = await fetchJson(`${CONFIG.siteUrl}/_api/lists/getbytitle('${encodeURIComponent(list)}')`, {
        headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json;odata=verbose' }
      });
      const type = metaRes.data?.d?.ListItemEntityTypeFullName;
      const digest = await getDigest(token);
      const body = JSON.stringify({ __metadata: { type }, ...item });
      const result = await fetchJson(`${CONFIG.siteUrl}/_api/lists/getbytitle('${encodeURIComponent(list)}')/items`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': digest
        },
        body
      });
      return res.status(200).json({ ok: result.ok, data: result.data });
    }

    if (action === 'patch') {
      const metaRes = await fetchJson(`${CONFIG.siteUrl}/_api/lists/getbytitle('${encodeURIComponent(list)}')`, {
        headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json;odata=verbose' }
      });
      const type = metaRes.data?.d?.ListItemEntityTypeFullName;
      const digest = await getDigest(token);
      const body = JSON.stringify({ __metadata: { type }, ...item });
      const result = await fetchJson(`${CONFIG.siteUrl}/_api/lists/getbytitle('${encodeURIComponent(list)}')/items(${id})`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': digest,
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body
      });
      return res.status(200).json({ ok: result.ok });
    }

    res.status(400).json({ ok: false, error: 'Unknown action' });
  } catch(e) {
    res.status(500).json({ ok: false, error: e.message });
  }
};
