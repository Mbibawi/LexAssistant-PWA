import { Document, Packer, Paragraph, TextRun } from 'docx';
const endPoint = (path: string) => `https://graph.microsoft.com/v1.0/me/drive/root:/${path}`

export const proxyHandler = async (req: any, res: any) => {
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
    } catch (error: any) {
        console.error('Global Proxy Error:', error);
        res.status(500).json({ error: 'Internal Server Error', message: error.message });
    }
};

/**
 * 1. Create and Upload DOCX
 */
async function createAndUploadDocx(req: any, res: any) {
    const text = req.body;
    const folder = req.headers['x-folder-path'];
    const name = req.headers['x-file-name'] || `Doc_${Date.now()}.docx`;

    const doc = new Document({
        sections: [{ children: [new Paragraph({ children: [new TextRun(text)] })] }],
    });
    const buffer = await Packer.toBuffer(doc);
    const fullPath = folder ? `${folder}/${name}` : name;
    const response = await fetch(
        endPoint(`${fullPath}:/content`),
        await request(req, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer)
    );
    return await finalize(response, res);
}

/**
 * 2. Save General File
 */
async function saveFileToOneDrive(req: any, res: any) {
    const path = req.headers['x-path'];
    const mime = req.headers['x-mime-type'];
    const response = await fetch(
        endPoint(`${path}:/content`),
        await request(req, mime, req.body)
    );
    return await finalize(response, res);
}

/**
 * 3. List Folder Items
 */
async function listOneDriveItems(req: any, res: any) {
    const path = req.headers['x-path'];
    const response = await fetch(
        endPoint(`${path}:/children`),
        await request(req));
    return await finalize(response, res);
}


/**
 * 4. Delete Item
 */
async function deleteOneDriveItem(req: any, res: any) {
    const path = req.headers['x-path'];
    const response = await fetch(
        endPoint(path),
        await request(req));
    if (response.status === 204) return res.status(200).json({ message: 'Deleted' });
    return await finalize(response, res);
}

/**
 * 5. Fetch File
 */
async function fetchFileFromOneDrive(req: any, res: any) {
    const filePath = req.headers['x-path'];
    const response = await fetch(
        endPoint(`${filePath}:/content`),
        await request(req));
    if (!response.ok) return await finalize(response, res);
    const buffer = await response.arrayBuffer();
    res.setHeader('Content-Type', response.headers.get('content-type') || 'application/octet-stream');
    return res.send(Buffer.from(buffer));
}

/**
 * Helpers
 */
async function request(req: any, contentType?: string, body?: any) {
    const accessToken = await getMicrosoftAccessToken();
    const headers: any = {
        'Authorization': `Bearer ${accessToken}`,
    };
    if (contentType) headers['Content-Type'] = contentType;

    const init: RequestInit = {
        method: req.method,
        headers,
    }
    if (body) init.body = body;
    return init;
}


async function finalize(response: Response, res: any) {
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
        return res.status(response.status).json({
            error: 'OneDrive API Error',
            message: data.error?.message || 'Request failed'
        });
    }
    return res.status(200).json(data);
}

async function getMicrosoftAccessToken(): Promise<string> {
    const { CLIENT_ID, CLIENT_SECRET, REFRESH_TOKEN } = process.env;
    const params = new URLSearchParams({
        client_id: CLIENT_ID!,
        client_secret: CLIENT_SECRET!,
        grant_type: 'refresh_token',
        refresh_token: REFRESH_TOKEN!,
        scope: 'offline_access Files.ReadWrite.All',
    });

    const response = await fetch('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString()
    });

    const data: any = await response.json();
    if (!response.ok) throw new Error(data.error_description || 'Auth failed');
    return data.access_token;
}