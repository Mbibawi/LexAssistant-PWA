import { Document, Packer, Paragraph, TextRun } from 'docx';
const endPoint = (path) => `https://graph.microsoft.com/v1.0/me/drive/root:/${path}`;
export const proxyHandler = async (req, res) => {
    const path = req.path;
    const method = req.method;
    try {
        if (path === '/api/proxy/docx' && method === 'PUT') {
            return await createAndUploadDocx(req, res);
        }
        if (path === '/api/proxy/save' && method === 'POST') {
            return await saveFileToOneDrive(req, res);
        }
        if (path === '/api/proxy/list' && method === 'GET') {
            return await listOneDriveItems(req, res);
        }
        if (path === '/api/proxy/delete' && method === 'DELETE') {
            return await deleteOneDriveItem(req, res);
        }
        if (path === '/api/proxy/fetch' && method === 'GET') {
            return await fetchFileFromOneDrive(req, res);
        }
        res.status(404).json({ error: 'Endpoint not found or method mismatch' });
    }
    catch (error) {
        console.error('Global Proxy Error:', error);
        res.status(500).json({ error: 'Internal Server Error', message: error.message });
    }
};
/**
 * 1. Create and Upload DOCX
 */
async function createAndUploadDocx(req, res) {
    const text = req.body;
    const folder = req.headers['x-folder-path'];
    const name = req.headers['x-file-name'] || `Doc_${Date.now()}.docx`;
    const doc = new Document({
        sections: [{ children: [new Paragraph({ children: [new TextRun(text)] })] }],
    });
    const buffer = await Packer.toBuffer(doc);
    const fullPath = folder ? `${folder}/${name}` : name;
    const response = await fetch(endPoint(`${fullPath}:/content`), await request(req, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer));
    return await finalize(response, res);
}
/**
 * 2. Save General File
 */
async function saveFileToOneDrive(req, res) {
    const path = req.headers['x-path'];
    const mime = req.headers['x-mime-type'];
    const response = await fetch(endPoint(`${path}:/content`), await request(req, mime, req.body));
    return await finalize(response, res);
}
/**
 * 3. List Folder Items
 */
async function listOneDriveItems(req, res) {
    const path = req.headers['x-path'];
    const response = await fetch(endPoint(`${path}:/children`), await request(req));
    return await finalize(response, res);
}
/**
 * 4. Delete Item
 */
async function deleteOneDriveItem(req, res) {
    const path = req.headers['x-path'];
    const response = await fetch(endPoint(path), await request(req));
    if (response.status === 204)
        return res.status(200).json({ message: 'Deleted' });
    return await finalize(response, res);
}
/**
 * 5. Fetch File
 */
async function fetchFileFromOneDrive(req, res) {
    const filePath = req.headers['x-path'];
    const response = await fetch(endPoint(`${filePath}:/content`), await request(req));
    if (!response.ok)
        return await finalize(response, res);
    const buffer = await response.arrayBuffer();
    res.setHeader('Content-Type', response.headers.get('content-type') || 'application/octet-stream');
    return res.send(Buffer.from(buffer));
}
/**
 * Helpers
 */
async function request(req, contentType, body) {
    const userOid = req.header['x-user'];
    const accessToken = await getMicrosoftAccessToken(userOid);
    if (!accessToken)
        throw new Error('Microsoft access token not found');
    const headers = {
        'Authorization': `Bearer ${accessToken}`,
    };
    if (contentType)
        headers['Content-Type'] = contentType;
    const init = {
        method: req.method,
        headers,
    };
    if (body)
        init.body = body;
    return init;
}
async function finalize(response, res) {
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        return res.status(response.status).json({
            error: 'OneDrive API Error',
            message: data.error?.message || 'Request failed'
        });
    }
    return res.status(200).json(data);
}
async function getMicrosoftAccessToken(userOid) {
    const { CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN } = process.env;
    const params = new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        grant_type: 'refresh_token',
        refresh_token: REFRESH_TOKEN,
        scope: 'openid profile offline_access Files.ReadWrite.All',
    });
    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString()
    });
    const data = await response.json();
    if (!response.ok)
        throw new Error(data.error_description || 'Auth failed');
    if (userOid !== (data.id_token ? JSON.parse(Buffer.from(data.id_token.split('.')[1], 'base64').toString()).oid : 'Unknown'))
        throw new Error('The user name does not match');
    return data.access_token;
}
//# sourceMappingURL=onedrive-proxy.js.map