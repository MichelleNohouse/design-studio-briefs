// Cloudflare Worker v4
// Handles:
// 1. GET /?q=name  — Azure AD directory search
// 2. POST /upload  — Upload file to SharePoint Brief Attachments library

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

const SHAREPOINT_SITE = 'https://asiansourcinglink.sharepoint.com/sites/LinkCreativeServices';
const SHAREPOINT_LIBRARY = 'Brief Attachments';

export default {
  async fetch(request, env) {

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: CORS_HEADERS });
    }

    const url = new URL(request.url);

    // ── ROUTE 1: POST /upload — upload file to SharePoint ──
    if (request.method === 'POST' && url.pathname === '/upload') {
      try {
        const body = await request.json();
        const { jobId, fileName, fileContent, fileType } = body;

        if (!jobId || !fileName || !fileContent) {
          return new Response(JSON.stringify({ error: 'Missing jobId, fileName or fileContent' }), {
            status: 400, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
          });
        }

        // Get token
        const token = await getToken(env);
        if (!token) {
          return new Response(JSON.stringify({ error: 'Auth failed' }), {
            status: 401, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
          });
        }

        // Get SharePoint site ID
        const siteResp = await fetch(
          `https://graph.microsoft.com/v1.0/sites/asiansourcinglink.sharepoint.com:/sites/LinkCreativeServices`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        const siteData = await siteResp.json();
        const siteId = siteData.id;

        // Get drive ID for Brief Attachments library
        const drivesResp = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        const drivesData = await drivesResp.json();
        const drive = drivesData.value.find(d => d.name === SHAREPOINT_LIBRARY);

        if (!drive) {
          return new Response(JSON.stringify({ error: 'Brief Attachments library not found' }), {
            status: 404, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
          });
        }

        // Create folder named after Job ID if it doesn't exist
        await fetch(
          `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/children`,
          {
            method: 'POST',
            headers: {
              Authorization: `Bearer ${token}`,
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              name: jobId,
              folder: {},
              '@microsoft.graph.conflictBehavior': 'rename'
            })
          }
        );

        // Upload file to Job ID folder
        const fileBytes = Uint8Array.from(atob(fileContent), c => c.charCodeAt(0));
        const uploadResp = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${jobId}/${fileName}:/content`,
          {
            method: 'PUT',
            headers: {
              Authorization: `Bearer ${token}`,
              'Content-Type': fileType || 'application/octet-stream'
            },
            body: fileBytes
          }
        );

        const uploadData = await uploadResp.json();

        return new Response(JSON.stringify({
          ok: true,
          fileUrl: uploadData.webUrl || '',
          fileName: fileName,
          jobId: jobId
        }), {
          headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
        });

      } catch (err) {
        return new Response(JSON.stringify({ error: 'Upload error', detail: err.message }), {
          status: 500, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
        });
      }
    }

    // ── ROUTE 2: GET /?q=name — Azure AD directory search ──
    if (request.method === 'GET') {
      const query = url.searchParams.get('q');
      if (!query || query.length < 2) {
        return new Response(JSON.stringify({ value: [] }), {
          headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
        });
      }

      try {
        const token = await getToken(env);
        if (!token) {
          return new Response(JSON.stringify({ error: 'Auth failed' }), {
            status: 401, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
          });
        }

        const graphUrl = `https://graph.microsoft.com/v1.0/users?$search="displayName:${query}"&$select=displayName,mail,jobTitle&$top=8&$orderby=displayName&ConsistencyLevel=eventual`;

        const graphResp = await fetch(graphUrl, {
          headers: {
            Authorization: `Bearer ${token}`,
            ConsistencyLevel: 'eventual'
          }
        });

        const graphData = await graphResp.json();
        const users = (graphData.value || []).map(u => ({
          displayName: u.displayName || '',
          mail: u.mail || '',
          jobTitle: u.jobTitle || ''
        }));

        return new Response(JSON.stringify({ value: users }), {
          headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
        });

      } catch (err) {
        return new Response(JSON.stringify({ error: 'Search error', detail: err.message }), {
          status: 500, headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
        });
      }
    }

    return new Response('Not found', { status: 404, headers: CORS_HEADERS });
  }
};

async function getToken(env) {
  try {
    const resp = await fetch(
      `https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: env.CLIENT_ID,
          client_secret: env.CLIENT_SECRET,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials'
        }).toString()
      }
    );
    const data = await resp.json();
    return data.access_token || null;
  } catch (e) {
    return null;
  }
}
