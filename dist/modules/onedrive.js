/**
 * onedrive.ts — Microsoft Graph API, path-based, class hierarchy.
 *
 * class Configuration  — localStorage config read/write
 * class OneDriveAuth   — MSAL auth + raw Graph/proxy fetch
 * class Folders        — base folder/file/JSON ops (extends OneDriveAuth)
 * class Scenario       — shared logic: skills, UI helpers, OD connection
 * class Cases          — dossiers scenario (extends Scenario)
 * class Library        — bibliothèque scenario (extends Scenario)
 *
 * OneDrive structure:
 *   <root>/
 *     _Skills/                       ← .md/.txt skills
 *     Affaires/<Folder>/
 *       _meta.json                   ← CaseMeta
 *       _notes.json                  ← PermanentNote[]
 *       _conversation.json           ← ChatMessage[]
 *       _kb_YYYY-MM-DD_HHmm.md      ← versioned knowledge bases
 *       <file>
 *     Bibliotheque/<Domain>/
 *       _meta.json                   ← LibDomainMeta
 *       _conversation.json           ← LibConversationMessage[]
 *       _kb_YYYY-MM-DD_HHmm.md      ← versioned knowledge bases
 *       <file>
 *     Bibliotheque/_conversation.json ← "all" domain conversation
 */
import { downloadBlob as download, byID, toast, spinnerEl as spinner, el, toggle, uid, formatDate, formatDateTime, qs, qsa, setActive, } from './ui.js';
import { isSupported, mimeLabel, mimeIcon, formatSize, makeCaseDocMeta, makeLibDocMeta, guessKind, kindLabel } from './ingest.js';
import { oneDrive, ids } from '../main.js';
import { renderMarkdown } from './markdown.js';
import { ClaudeAPI } from './api.js';
import { generateDocx } from './docxgen.js';
// ─── Constants ────────────────────────────────────────────────────────────────
const LS_CONFIG = 'lex_onedrive_config';
const APP_ROOT = "Legal/Mon Cabinet d'Avocat/_LexAssistant";
const FOLDER_SKILLS = '_Skills';
const APP_CONFIG_FILE = '_config.json';
const CLAUDE = new ClaudeAPI();
// ─── Configuration ────────────────────────────────────────────────────────────
class Configuration {
    config;
    constructor() {
        this.config = this.getConfig() || this.initiateConfig() || null;
    }
    getConfig() {
        const raw = localStorage.getItem(LS_CONFIG);
        try {
            return raw ? JSON.parse(raw) : null;
        }
        catch {
            return null;
        }
    }
    setConfig(cfg) {
        localStorage.setItem(LS_CONFIG, JSON.stringify(cfg));
    }
    initiateConfig() {
        return this.config; //!We are desactivating this for now since we are using the google function onedrive-proxy
        const config = {
            clientId: prompt('Provide the OneDrive Client ID') || '',
            tenantId: prompt('Provide the OneDrive tenant ID') || '',
            rootFolder: APP_ROOT,
        };
        this.setConfig(config);
        return config;
    }
    clearConfig() { localStorage.removeItem(LS_CONFIG); }
    isConfigured() { return Boolean(this.config?.clientId && this.config?.rootFolder); }
    root() { return this.config?.rootFolder ?? APP_ROOT; }
}
class MSAL {
    _clientId = "9cb553c1-8473-4b2a-91d4-fef8b7cd7bff";
    _tenantID = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
    _redirectUri = "https://mbibawi.github.io/LexAssistant-PWA/"; //!must be the same domain as the app;
    _app = new msal.PublicClientApplication(this.msalConfig()); //!this must come after clientId and redirectUri are delcared
    loginRequest = { scopes: [''] };
    constructor(scopes = ["Files.ReadWrite"]) {
        this.loginRequest.scopes = scopes;
    }
    get msalApp() { return this._app; }
    ;
    msalConfig() {
        return {
            auth: {
                clientId: this._clientId,
                authority: "https://login.microsoftonline.com/common",
                redirectUri: this._redirectUri,
            },
            cache: {
                cacheLocation: "localStorage",
                storeAuthStateInCookie: true
            }
        };
    }
    ;
    async init() {
        await this._app.initialize(); // required in MSAL browser v3+
        const response = await this._app.handleRedirectPromise();
        if (response?.account) {
            this._app.setActiveAccount(response.account);
        }
    }
    // Function to check existing authentication context
    async acquireToken() {
        await this.init();
        try {
            const account = this._app.getAllAccounts()[0];
            if (account) {
                this._app.setActiveAccount(account);
                return await this.acquireTokenSilently(account);
            }
            else {
                return await this.loginWithPopup();
            }
        }
        catch (error) {
            console.error("Failed to acquire token from acquireToken(): ", error);
        }
    }
    // Function to get access token silently
    async acquireTokenSilently(account) {
        try {
            const tokenRequest = {
                account: account,
                scopes: this.loginRequest.scopes, // OneDrive scopes
            };
            const tokenResponse = await this._app.acquireTokenSilent(tokenRequest);
            if (!tokenResponse || !tokenResponse.accessToken)
                return null;
            console.log("Token acquired silently :", tokenResponse.accessToken);
            return { token: tokenResponse.accessToken, account };
        }
        catch (error) {
            //if (error instanceof this.msalInstance.InteractionRequiredAuthError)// Silent failed, fall back to popup
            console.error("Token silent acquisition error:", error);
            return await this.loginWithPopup() || null;
        }
    }
    ;
    async loginWithPopup() {
        try {
            const loginResponse = await this._app.loginPopup(this.loginRequest);
            console.log('loginResponse = ', loginResponse);
            const account = loginResponse.account;
            if (!account)
                return null;
            this._app.setActiveAccount(account);
            const tokenResponse = await this._app.acquireTokenSilent({
                account: account,
                scopes: ["Files.ReadWrite"]
            });
            console.log("Token acquired from loginWithPopup: ", tokenResponse.accessToken);
            return { token: tokenResponse.accessToken, account };
        }
        catch (error) {
            console.error("Error acquiring token from loginWithPopup(): ", error);
            //@ts-ignore
            // if (error instanceof InteractionRequiredAuthError) { }
            // Fallback to popup if silent token acquisition fails
            const response = await this._app.acquireTokenPopup({
                scopes: ["Files.ReadWrite"]
            });
            console.log("Token acquired via popup:", response.accessToken);
            return { token: response.accessToken, account: this._app.getActiveAccount() };
        }
    }
    async credentitalsToken(tenantId) {
        const msalConfig = {
            auth: {
                clientId: this._clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                //clientSecret: clientSecret,
            }
        };
        //@ts-ignore
        const cca = new msal.application.ConfidentialClientApplication(msalConfig);
        const tokenRequest = {
            scopes: ["Files.ReadWrite"],
        };
        try {
            const response = await cca.acquireTokenByClientCredential(tokenRequest);
            return response.accessToken;
        }
        catch (error) {
            console.log('Error acquiring Token: ', error);
            return null;
        }
    }
    async getOfficeToken() {
        try {
            //@ts-ignore
            return await OfficeRuntime.auth.getAccessToken();
        }
        catch (error) {
            console.log("Error : ", error);
        }
    }
    async getTokenWithSSO(email, tenantId) {
        const msalConfig = {
            auth: {
                clientId: this._clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                redirectUri: this._redirectUri,
                navigateToLoginRequestUrl: true,
            },
            cache: {
                cacheLocation: "ExcelAddIn",
                storeAuthStateInCookie: true
            }
        };
        try {
            //@ts-ignore
            const response = await this._app.ssoSilent({
                scopes: ["Files.ReadWrite"],
                //scopes: ["https://graph.microsoft.com/.default"],
                loginHint: email // Forces MSAL to recognize the signed-in user
            });
            console.log("Token acquired via SSO:", response.accessToken);
            return response.accessToken;
        }
        catch (error) {
            console.error("SSO silent authentication failed:", error);
            return null;
        }
    }
    openLoginWindow() {
        const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${this._clientId}&response_type=token&redirect_uri=${this._redirectUri}&scope=https://graph.microsoft.com/.default`;
        // Open in a new window (only works if triggered by user action)
        const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");
        if (!authWindow) {
            console.error("Popup blocked! Please allow popups.");
        }
    }
    // Function to handle login and acquire token
    async loginAndGetToken() {
        const msalConfig = {
            auth: {
                clientId: this._clientId,
                authority: "https://login.microsoftonline.com/common",
                redirectUri: this._redirectUri
            },
            cache: {
                cacheLocation: "ExcelInvoicing", // Specify cache location
                storeAuthStateInCookie: true // Set this to true for IE 11
            }
        };
        return await acquire(this);
        async function acquire(this$) {
            try {
                const response = await this$._app.handleRedirectPromise();
                if (response !== null) {
                    console.log("Login successful:", response);
                    return response.accessToken;
                }
                const accounts = this$._app.getAllAccounts();
                if (accounts.length > 0) {
                    const tokenResponse = await this$._app.acquireTokenSilent({
                        account: accounts[0],
                        scopes: ["https://graph.microsoft.com/.default"]
                    });
                    console.log("Token acquired silently:", tokenResponse.accessToken);
                    return tokenResponse.accessToken;
                }
            }
            catch (error) {
                console.error("Error acquiring token:", error);
                //@ts-ignore
                if (error instanceof msal.InteractionRequiredAuthError) {
                    this$._app.acquireTokenRedirect({
                        scopes: ["https://graph.microsoft.com/.default"]
                    });
                }
            }
        }
        // Function to handle redirect response
        async function handleRedirectResponse(this$) {
            try {
                const authResult = await this$._app.handleRedirectPromise();
                if (authResult && authResult.accessToken) {
                    console.log("Access token:", authResult.accessToken);
                    return authResult.accessToken;
                }
            }
            catch (error) {
                console.error("Redirect handling error:", error);
            }
            return undefined;
        }
    }
}
// ─── OneDriveAuth — MSAL auth + raw Graph/proxy fetch ────────────────────────
export class OneDriveAuth {
    GRAPH = 'https://graph.microsoft.com/v1.0/me/drive/root:/';
    CLAUDE_PROXY = 'https://claude-ai-proxy-428231091257.europe-west1.run.app/api/proxy/';
    cfg = new Configuration();
    _scopes = ['Files.ReadWrite', 'User.Read', 'openid', 'profile'];
    _MSAL = new MSAL();
    _token = null;
    _account = null;
    _userOid = null;
    get account() { return this._account; }
    get userOid() { return this._userOid; }
    get token() { return this._token; }
    get config() { return this.cfg.config; }
    get isConfigured() { return this.cfg.isConfigured(); }
    get root() { return this.cfg.root(); }
    setConfig(cfg) { this.cfg.setConfig(cfg); }
    /**
     * getAccessToken - Get the access token for the current user.
     * @returns {Promise<string>} The access token.
     */
    async getAccessToken() {
        return this.signIn();
    }
    /**
     * signIn - Sign in to OneDrive.
     * @returns {Promise<void>}
     */
    async signIn() {
        const acquired = await this._MSAL.acquireToken();
        this._token = acquired?.token || null;
        this._account = acquired?.account || null;
        this._token ? alert(`Signed in successfully` + this._token) : alert(`Failed to sign in`);
        return this._token;
    }
    /**
     * signOut - Sign out of OneDrive.
     * @returns {Promise<void>}
     */
    async signOut() {
        this._account = null;
        this._token = null;
        sessionStorage.clear();
    }
    /**
     * isSignedIn - Check if the user is signed in.
     * @returns {Promise<string | null>} The signed-in user or null.
     */
    async isSignedIn() {
        try {
            const accounts = this._MSAL.msalApp.getAllAccounts();
            if (accounts.length) {
                this._account = accounts[0];
                return this._account.name ?? this._account.username;
            }
            return null;
        }
        catch {
            return null;
        }
    }
    /**
     * oneDriveProxy - Proxy for OneDrive operations.
     * @param root The root folder for the OneDrive operations.
     * @param method The HTTP method for the request.
     * @param payload The payload for the request.
     * @returns {Promise<any>} The response from the OneDrive operations.
     */
    async oneDriveProxy(root, method, payload) {
        if (!this.account)
            await this.getAccessToken();
        console.log('user oid = ', this.account?.idTokenClaims.oid);
        const url = `https://onedrive-proxy-428231091257.europe-west1.run.app/api/proxy/${root}`;
        const { body, path, mimeType } = payload;
        const headers = {
            'x-path': path || '',
            'x-mime-type': mimeType || 'application/octet-stream',
            'x-user': this.account?.idTokenClaims.oid.toLowerCase() || '',
        };
        const response = await fetch(url, {
            method: method,
            headers: headers,
            // data is sent as the raw binary body
            body: body || null
        });
        if (root === 'fetch' && response.ok)
            return await response.blob();
        const result = await response.json();
        if (!response.ok)
            throw new Error(result.message || 'Proxy Error');
        return result;
    }
    async callClaudeProxy(api, body, anthropicVersion) {
        return await this.gFetch(`${this.CLAUDE_PROXY}${api}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'anthropic-version': anthropicVersion,
            },
            body: body,
        }, false);
    }
    /**
     * Universal fetch for both Graph API and external URLs (GCF proxy).
     * - If url starts with 'https://' it is used verbatim (external call).
     * - Otherwise it is appended to the Graph base URL.
     * - rawBody=true skips automatic Content-Type injection for binary/proxy calls.
     */
    async gFetch(path, opts = {}, rawBody = false) {
        if (!path.startsWith(this.CLAUDE_PROXY) && !this._token)
            await this.getAccessToken();
        const url = path.startsWith(this.CLAUDE_PROXY) ? path : `${this.GRAPH}${path}`;
        const headers = {
            Authorization: `Bearer ${this._token}`,
            ...(opts.headers ?? {}),
        };
        if (!rawBody && opts.body && typeof opts.body === 'string') {
            headers['Content-Type'] = 'application/json';
        }
        const resp = await fetch(url, { ...opts, headers });
        if (!resp.ok) {
            let msg = resp.statusText;
            try {
                const e = (await resp.json());
                msg = e.error?.message ?? msg;
            }
            catch { }
            throw new Error(`Fetch ${resp.status}: ${msg}`);
        }
        return resp;
    }
    encode(odPath) {
        return odPath.split('/').map((seg) => encodeURIComponent(seg)).join('/');
    }
}
// ─── Folders — base file/folder/JSON operations ───────────────────────────────
class Folders extends OneDriveAuth {
    get appConfigPath() {
        return `${this.root}/${APP_CONFIG_FILE}`;
    }
    mainFolder = null;
    constructor(mainFolder) {
        super();
        this.mainFolder = mainFolder;
    }
    /**
     * Ensure a folder exists.
     * @param folderPath The path to the folder.
     */
    async ensureFolder(folderPath) {
        try {
            await this.gFetch(this.encode(folderPath));
            return;
        }
        catch { }
        const parts = folderPath.split('/');
        const name = parts.pop();
        const parentPath = parts.join('/');
        const parentEndpoint = parentPath ? `${this.encode(parentPath)}/children` : 'children';
        await this.gFetch(parentEndpoint, {
            method: 'POST',
            body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }),
        });
    }
    /**
     * List all items in a folder.
     * @param folderPath The path to the folder.
     */
    async listAllFolderItems(folderPath) {
        return this.oneDriveProxy('list', 'GET', { path: folderPath });
        const resp = await this.gFetch(`${this.encode(folderPath)}/children?$select=name,size,file,folder,webUrl,lastModifiedDateTime&$top=500`);
        const data = (await resp.json());
        return data.value ?? [];
    }
    /**
     * Lists immediate subfolders of a path relative to root.
     * Used by both Cases (list dossiers) and Library (list domains).
     * @param parentRelPath The relative path to the parent folder.
     */
    async listSubFolders(parentRelPath) {
        try {
            const items = await this.listAllFolderItems(`${this.root}/${parentRelPath}`);
            return items.filter((f) => f.folder).map((i) => i.name);
        }
        catch {
            return [];
        }
    }
    /**
     * Lists non-underscore files in a folder.
     * Used by both Cases and Library to enumerate documents.
     * @param folderAbsPath The absolute path to the folder.
     */
    async listFiles(folderAbsPath) {
        const items = await this.listAllFolderItems(folderAbsPath);
        return items.filter((i) => i.file);
    }
    /**
     * Reads a JSON file from the given file path.
     * @param filePath The path to the JSON file.
     * @returns The parsed JSON data, or null if the file cannot be read.
     */
    async readJson(filePath) {
        try {
            const buf = await this.readFilePath(filePath);
            const text = new TextDecoder().decode(buf);
            return JSON.parse(text);
        }
        catch {
            return null;
        }
    }
    /**
     * Reads the content of a file from OneDrive.
     * @param filePath The path to the file to read.
     * @returns A Promise that resolves to the content of the file as an ArrayBuffer.
     */
    async readFilePath(filePath) {
        return this.oneDriveProxy('fetch', 'GET', { path: filePath });
        const resp = await this.gFetch(`${this.encode(filePath)}:/content`);
        if (!resp.ok)
            throw new Error(`Read ${filePath}: ${resp.status}`);
        return resp.arrayBuffer();
    }
    /**
      * Writes JSON data to a file.
     * @param filePath The path to the file to write.
     * @param data The data to write to the file.
     */
    async writeJson(filePath, data) {
        await this.writeFilePath(filePath, JSON.stringify(data, null, 2), 'application/json');
    }
    /**
     * Writes a file to OneDrive.
     * @param filePath The path to the file to write.
     * @param data The data to write to the file.
     * @param mimeType The MIME type of the file.
     */
    async writeFileLarge(filePath, data, mimeType) {
        if (data.byteLength <= 4 * 1024 * 1024) {
            await this.writeFilePath(filePath, data, mimeType);
            return;
        }
        const sessResp = await this.oneDriveProxy('save', 'POST', {
            path: `${filePath}/createUploadSession`,
            body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
            mimeType
        });
        /*const sessResp = await this.gFetch(`${this.encode(filePath)}/createUploadSession`, {
          method: 'POST',
          body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
        });*/
        const { uploadUrl } = (await sessResp.json());
        const chunk = 10 * 1024 * 1024;
        for (let off = 0; off < data.byteLength; off += chunk) {
            const end = Math.min(off + chunk, data.byteLength);
            const r = await fetch(uploadUrl, {
                method: 'PUT',
                headers: { 'Content-Range': `bytes ${off}-${end - 1}/${data.byteLength}` },
                body: data.slice(off, end),
            });
            if (!r.ok && r.status !== 202)
                throw new Error(`Chunk upload failed at ${off}`);
        }
    }
    /**
     * Writes a file to OneDrive.
     * @param filePath The path to the file to write.
     * @param data The data to write to the file.
     * @param mimeType The MIME type of the file.
     */
    async writeFilePath(filePath, data, mimeType) {
        const body = typeof data === 'string' ? new TextEncoder().encode(data) : data;
        return this.oneDriveProxy('save', 'POST', { path: filePath, body, mimeType });
        await this.gFetch(`${this.encode(filePath)}/content`, {
            method: 'PUT',
            headers: { Authorization: `Bearer ${this.token}`, 'Content-Type': mimeType },
            body,
        }, true);
    }
    async deleteFilePath(filePath) {
        return this.oneDriveProxy('delete', 'DELETE', { path: filePath });
        await this.gFetch(`${this.encode(filePath)}`, { method: 'DELETE' });
    }
    async readAppConfig() {
        return this.readJson(this.appConfigPath);
    }
    async writeAppConfig(cfg) {
        await this.writeJson(this.appConfigPath, cfg);
    }
    async initRootStructure() {
        const r = this.root;
        await this.ensureFolder(r);
        await this.ensureFolder(`${r}/${FOLDER_SKILLS}`);
        await this.ensureFolder(`${r}/${this.mainFolder}`);
        const existing = await this.readAppConfig();
        if (!existing)
            await this.writeAppConfig({});
    }
}
// ─── Scenario — shared: skills, OD status, UI helpers ────────────────────────
class Common extends Folders {
    claude = CLAUDE;
    //readonly config  = oneDrive.config;
    activeMode = 'analyse';
    skills = [];
    user = () => oneDrive.userOid;
    // ─── Skills ───────────────────────────────────────────────────────────────
    /**
     * Fetches all .md/.txt files from the _Skills folder and loads them.
     * Called once after OneDrive connection. Used by both Cases and Library.
     */
    async fetchSkills() {
        const path = `${this.root}/${FOLDER_SKILLS}`;
        try {
            const items = await this.listAllFolderItems(path);
            const skills = [];
            for (const item of items.filter((i) => !i.folder)) {
                const ext = item.name.split('.').pop()?.toLowerCase() ?? '';
                if (!['md', 'txt'].includes(ext))
                    continue;
                try {
                    const buf = await this.readFilePath(`${path}/${item.name}`);
                    skills.push({ name: item.name, content: new TextDecoder().decode(buf) });
                }
                catch { }
            }
            this.skills = skills;
            this.updateSkillIndicator();
            return skills;
        }
        catch {
            return [];
        }
    }
    // ─── OneDrive connection ──────────────────────────────────────────────────
    updateODStatus() {
        const statusEl = byID(ids.oneDriveStatus);
        if (!statusEl)
            return;
        const userName = oneDrive.isSignedIn();
        if (userName) {
            statusEl.textContent = `☁ ${userName}`;
            statusEl.className = 'od-status od-status--connected';
        }
        else {
            statusEl.textContent = '☁ Non connecté';
            statusEl.className = 'od-status od-status--disconnected';
        }
    }
    /**
     *
     * @param msg
     * @returns void
     */
    appendMsg(msg) {
        const area = byID(ids.chatArea);
        if (!area)
            return;
        area.querySelector('.empty-state')?.remove();
        area.appendChild(msg);
        area.scrollTop = area.scrollHeight;
    }
    /**
     *
     * @returns
     */
    setupInputArea() {
        const ta = byID(ids.userInput);
        const sendBtn = byID(ids.sendBtn);
        if (!ta || !sendBtn)
            return;
        ta.addEventListener('input', () => this.autoResize(ta));
        ta.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            this.sendMessage(ta, sendBtn);
        } });
        sendBtn.onclick = () => this.sendMessage(ta, sendBtn);
        this.updateQuickPrompts(ta);
    }
    // ─── Settings modal (OneDrive only — no API key) ──────────────────────────
    openSettingsModal() {
        byID(ids.settingsOverlay)?.remove();
        const userName = oneDrive.isSignedIn();
        const overlay = el('div', { className: 'modal-overlay', id: ids.settingsOverlay });
        const dialog = el('div', { className: 'modal-dialog modal-dialog--settings' });
        dialog.innerHTML = `
      <h2 class="modal-title">⚙ Paramètres</h2>
      <h3 class="settings-section-title">☁ Microsoft OneDrive (Graph API)</h3>
      <p class="settings-hint"><strong>portal.azure.com</strong> → App registrations → New registration<br>
      Type : SPA — Redirect URI : <code>${window.location.origin}</code><br>
      Permissions : <code>Files.ReadWrite</code> + <code>User.Read</code></p>
      <label class="form-label">Application (Client) ID <span class="required">*</span></label>
      <input class="form-input" id="s-client" type="text" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" value="${this.config?.clientId ?? ''}"/>
      <label class="form-label">Tenant ID</label>
      <input class="form-input" id="s-tenant" type="text" placeholder="common" value="${this.config?.tenantId ?? 'common'}"/>
      <label class="form-label">Dossier racine OneDrive</label>
      <input class="form-input" id="s-root" type="text" placeholder="LexAssistant" value="${this.config?.rootFolder ?? 'LexAssistant'}"/>
      <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap">
        <button class="btn btn--primary btn--sm" id="s-od-save">Enregistrer</button>
        <button class="btn btn--secondary btn--sm" id="s-od-init">Initialiser structure OneDrive</button>
        <button class="btn btn--secondary btn--sm" id="s-od-signout">Déconnecter</button>
      </div>
      ${userName ? `<p class="od-connected-label">✓ Connecté : ${userName}</p>` : ''}
      <div class="modal-btns" style="margin-top:24px">
        <button class="btn btn--secondary" id="s-close">Fermer</button>
      </div>`;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        overlay.addEventListener('click', (e) => { if (e.target === overlay)
            overlay.remove(); });
        qs('#s-close', dialog).onclick = () => overlay.remove();
        qs('#s-od-save', dialog).onclick = () => {
            const clientId = qs('#s-client', dialog).value.trim();
            const tenantId = qs('#s-tenant', dialog).value.trim() || 'common';
            const rootFolder = qs('#s-root', dialog).value.trim() || 'LexAssistant';
            if (!clientId) {
                toast('Client ID requis.', 'error');
                return;
            }
            oneDrive.setConfig({ clientId, tenantId, rootFolder });
            toast('Configuration OneDrive enregistrée.', 'success');
        };
        qs('#s-od-init', dialog).onclick = async () => {
            if (!oneDrive.isConfigured) {
                toast('Sauvegardez la configuration d\'abord.', 'error');
                return;
            } //!might need to be change to if(!oneDrive.userOid)
            try {
                //if (!oneDrive.userOid) { await oneDrive.signIn(); oneDrive.userOid = oneDrive.getSignedInUser(); this.updateODStatus(); }
                await this.initRootStructure();
                toast('Structure initialisée avec succès.', 'success');
            }
            catch (err) {
                toast('Erreur : ' + err.message, 'error');
            }
        };
        qs('#s-od-signout', dialog).onclick = async () => {
            await oneDrive.signOut();
            this.updateODStatus();
            toast('Déconnecté.', 'info');
            overlay.remove();
        };
    }
    // ─── Skill indicator ──────────────────────────────────────────────────────
    updateSkillIndicator() {
        const n = this.skills.length;
        for (const id of [ids.skillsBadge, ids.libSkillsBadge]) {
            const badge = byID(id);
            if (!badge)
                continue;
            badge.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''}` : '';
            toggle(badge, n > 0);
        }
        const top = byID(ids.libSkillsTop);
        if (top)
            top.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''} actif${n > 1 ? 's' : ''}` : '';
    }
    // ─── Shared UI helpers ────────────────────────────────────────────────────
    onClick(btn, action) {
        if (btn)
            btn.onclick = action;
    }
    autoResize(ta) {
        ta.style.height = 'auto';
        ta.style.height = Math.min(ta.scrollHeight, 160) + 'px';
    }
    sanitiseFolder(name) {
        return name.replace(/[/\\:*?"<>|]/g, '_').replace(/\s+/g, '_').slice(0, 60);
    }
    showNotConnected() {
        const area = byID(ids.chatArea);
        if (!area)
            return;
        area.innerHTML = '';
        area.appendChild(el('div', { className: 'empty-state' }, el('div', { className: 'empty-icon', textContent: '☁' }), el('h2', { textContent: 'OneDrive non connecté' }), el('p', { textContent: 'Configurez votre App Registration Azure et connectez-vous.' }), (() => {
            const b = el('button', { className: 'btn btn--primary', textContent: '☁ Configurer OneDrive' });
            b.onclick = () => this.openSettingsModal();
            return b;
        })()));
    }
    // ─── Shared typing indicator ──────────────────────────────────────────────
    appendTypingTo(areaId, label) {
        const area = byID(areaId);
        const typing = el('div', { className: 'msg msg--assistant', id: ids.typing });
        const bubble = el('div', { className: 'msg_bubble' });
        bubble.append(spinner(), el('span', { textContent: ` ${label}` }));
        typing.append(el('div', { className: 'msg_label', textContent: 'Lex Assistant' }), bubble);
        area.appendChild(typing);
        area.scrollTop = area.scrollHeight;
        return typing;
    }
}
// ─── Cases — dossiers scenario ────────────────────────────────────────────────
export class Cases extends Common {
    _saving = false;
    _docFilter = null;
    _foldersMeta = [];
    _activeCase = null;
    _caseNotes = [];
    _caseMessages = [];
    _caseKb = null; // cached knowledge base content
    // ─── Path helpers ─────────────────────────────────────────────────────────
    casePath = (f) => `${this.root}/${this.mainFolder}/${f}`;
    metaPath = (f) => `${this.casePath(f)}/_meta.json`;
    notesPath = (f) => `${this.casePath(f)}/_notes.json`;
    convPath = (f) => `${this.casePath(f)}/_conversation.json`;
    // ─── Show UI ──────────────────────────────────────────────────────────────
    async showUI() {
        this.buildUI();
        // Wire all scenario-specific UI
        this.setupBarsBtns();
        if (this.user()) {
            this.updateODStatus();
            this.setupInputArea();
            this.renderChat();
            this.renderDocList();
            this.updateSkillIndicator();
            await this.loadAllSubFolders();
            await this.fetchSkills();
        }
        else {
            this.showNotConnected();
        }
    }
    buildUI() {
        const content = byID(ids.content);
        content.innerHTML = '';
        content.className = 'dossiers-view';
        const main = el('div', { id: ids.mainLayout });
        content.appendChild(main);
        const aside = el('aside', { id: ids.sidebar });
        const wSpace = el('div', { id: ids.workspace });
        main.append(aside, wSpace);
        // Sidebar
        aside.append(el('div', { className: 'sidebar_section-title', innerText: 'Dossiers' }), el('div', { id: ids.caseList }), el('button', { id: ids.btnNewSidebar, className: 'btn btn--ghost btn--dashed', innerText: '+ Nouveau dossier' }), el('div', { className: 'sidebar_divider' }), el('div', { className: 'sidebar_section-title', innerHTML: "Pièces <span id='doc-count' class='doc-count'>0 pièces</span>" }));
        const filterTabs = el('div', { className: 'doc-filter-tabs' });
        const types = {
            all: 'Tout',
            piece: 'Pièces',
            jurisprudence: 'Jurisprudence',
            doctrine: 'Doctrine',
            redige: 'Rédigés'
        };
        filterTabs.append(...['all', 'piece', 'jurisprudence', 'doctrine', 'redige'].map((f) => el('button', {
            className: `doc-filter-tab${f === 'all' ? ' active' : ''}`, 'data-filter': f,
            innerText: types[f]
        })));
        const fileInput = el('input', {
            type: 'file', id: ids.fileInput, multiple: true,
            accept: '.pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt,.txt,.md,.rtf', style: { display: 'none' }
        });
        const upload = el('div', { className: 'sidebar_upload-row' });
        upload.append(fileInput, el('button', { id: ids.btnUpload, className: 'btn btn--ghost btn--sm', innerText: '⬆ Upload' }), el('button', { id: ids.btnOdSync, className: 'btn btn--ghost btn--sm', innerText: '☁ Sync' }), el('button', { id: ids.btnNotesOpen, className: 'btn btn--ghost btn--sm', innerText: '📌 Notes' }));
        aside.append(filterTabs, el('div', { id: ids.docList, className: 'doc-list' }), upload);
        // Workspace
        const mode = el('div', { id: ids.modeBar });
        const inputArea = el('div', { id: ids.inputArea });
        const prompts = el('div', { id: ids.quickPrompts });
        wSpace.append(mode, el('div', { id: ids.noteBar, className: 'note-bar', style: { display: 'none' } }), el('div', { id: ids.chatArea, className: 'chat-area', role: 'log', 'aria-live': 'polite' }), prompts, inputArea);
        mode.append(el('span', { className: 'mode-bar_label', innerText: 'Mode :' }), ...['analyse', 'redaction', 'modification', 'note'].map((m, i) => el('button', {
            className: `mode-btn${i === 0 ? ' active' : ''}`, 'data-mode': m,
            innerText: ({ analyse: 'Analyse', redaction: 'Rédaction', modification: 'Modification', note: 'Note permanente' })[m]
        })), el('div', { className: 'mode-bar_spacer' }), el('button', { id: ids.btnCaseSummary, className: 'btn btn--ghost btn--sm', innerText: 'Point dossier ↗' }), el('button', { id: ids.btnBuildKb, className: 'btn btn--ghost btn--sm', innerText: '🧠 Base de connaissance' }));
        prompts.append(...[
            ['Risques du dossier', 'Analyse les risques juridiques et fiscaux du dossier et liste les points d\'attention prioritaires.'],
            ['Mise en demeure', 'Rédige une mise en demeure formelle à la partie adverse sur la base des pièces du dossier.'],
            ['Chronologie des faits', 'Fais une synthèse chronologique des faits pertinents issus des pièces du dossier.'],
            ['Analyse chiffrée', 'Analyse les données chiffrées des tableaux Excel et leurs implications juridiques et fiscales.'],
            ['Mémorandum juridique', 'Rédige un mémorandum juridique complet sur le point de droit central avec jurisprudence applicable.'],
        ].map(([label, prompt]) => el('button', { className: 'quick-btn', 'data-prompt': prompt, innerText: label })));
        inputArea.append(el('textarea', {
            id: ids.userInput, rows: 2,
            placeholder: 'Posez une question, demandez la rédaction d\'un acte, ou donnez une instruction…', 'aria-label': 'Message'
        }), el('button', {
            id: ids.sendBtn, className: 'btn btn--primary', 'aria-label': 'Envoyer',
            innerHTML: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>'
        }));
    }
    // ─── OneDrive CRUD ────────────────────────────────────────────────────────
    readCaseMeta(f) { return this.readJson(this.metaPath(f)); }
    async writeCaseMeta(f, m) {
        await this.ensureFolder(this.casePath(f));
        await this.writeJson(this.metaPath(f), m);
    }
    async readNotes(f) { return (await this.readJson(this.notesPath(f)))?.notes ?? []; }
    writeNotes(f, n) { return this.writeJson(this.notesPath(f), { notes: n }); }
    async readConversation(f) { return (await this.readJson(this.convPath(f)))?.messages ?? []; }
    writeConversation(f, m) { return this.writeJson(this.convPath(f), { messages: m }); }
    readCaseFile(folderName, fileName) {
        return this.readFilePath(`${this.casePath(folderName)}/${fileName}`);
    }
    async writeCaseFile(folderName, fileName, data, mimeType) {
        await this.ensureFolder(this.casePath(folderName));
        await this.writeFileLarge(`${this.casePath(folderName)}/${fileName}`, data, mimeType);
    }
    // ─── Persist helpers ──────────────────────────────────────────────────────
    async saveMeta() {
        if (!this._activeCase || this._saving)
            return;
        this._saving = true;
        const meta = {
            name: this._activeCase.name,
            folderName: this._activeCase.folderName,
            domain: this._activeCase.domain,
            status: this._activeCase.status,
            createdAt: this._activeCase.createdAt,
            updatedAt: Date.now(),
            documents: this._activeCase.documents,
        };
        this._activeCase.updatedAt = meta.updatedAt;
        try {
            await this.writeCaseMeta(this._activeCase.folderName, meta);
        }
        finally {
            this._saving = false;
        }
    }
    saveNotes() { return this._activeCase ? this.writeNotes(this._activeCase.folderName, this._caseNotes) : Promise.resolve(); }
    saveConversation() { return this._activeCase ? this.writeConversation(this._activeCase.folderName, this._caseMessages) : Promise.resolve(); }
    // ─── Load all cases ───────────────────────────────────────────────────────
    async loadAllSubFolders() {
        const subFolders = await this.listSubFolders(this.mainFolder);
        this._foldersMeta = [];
        await Promise.all(subFolders.map(async (folderName) => {
            const meta = await this.readCaseMeta(folderName);
            if (!meta)
                return;
            this._foldersMeta.push({
                folderName,
                name: meta.name,
                domain: meta.domain,
                status: meta.status,
                createdAt: meta.createdAt,
                updatedAt: meta.updatedAt,
                documents: meta.documents ?? []
            });
        }));
        this._foldersMeta.sort((a, b) => b.updatedAt - a.updatedAt);
        this.renderCaseList();
        if (this._foldersMeta.length > 0)
            await this.selectCase(this._foldersMeta[0].folderName);
        else
            this.showEmptyState();
    }
    // ─── Select case ──────────────────────────────────────────────────────────
    async selectCase(folderName) {
        const c = this._foldersMeta.find((x) => x.folderName === folderName);
        if (!c)
            return;
        this._activeCase = c;
        this._caseKb = null;
        this._caseNotes = await this.readNotes(folderName);
        this._caseMessages = await this.readConversation(folderName);
        // Try to load the latest knowledge base silently
        const kb = await this.claude.loadLatestCaseKb(this.casePath(folderName), this.listAllFolderItems, this.readFilePath);
        if (kb)
            this._caseKb = kb.content;
        qsa('.case-item').forEach((el) => el.classList.toggle('active', el.dataset.folder === folderName));
        const nameEl = byID(ids.topBarCaseName);
        const domainEl = byID(ids.topBarCaseDomain);
        if (nameEl)
            nameEl.textContent = c.name;
        if (domainEl)
            domainEl.textContent = c.domain;
        this.renderNoteBar();
        this.renderDocList();
        this.renderChat();
        this.updateDocCount();
    }
    async refreshCaseFromOneDrive() {
        if (!this._activeCase)
            return;
        try {
            const items = await this.listFiles(this.casePath(this._activeCase.folderName));
            let added = 0;
            for (const item of items) {
                if (!item.file || !isSupported(item.name))
                    continue;
                if (this._activeCase.documents.some((d) => d.name === item.name))
                    continue;
                this._activeCase.documents.push({
                    name: item.name, kind: guessKind(item.name),
                    mimeType: item.file.mimeType || 'application/octet-stream',
                    sizeBytes: item.size ?? 0, addedAt: Date.now(),
                });
                added++;
            }
            if (added > 0) {
                await this.saveMeta();
                this.renderDocList();
                this.updateDocCount();
            }
            toast(`${added} nouveau(x) document(s) indexé(s) depuis OneDrive.`, 'success');
        }
        catch (err) {
            toast('Erreur sync : ' + err.message, 'error');
        }
    }
    // ─── Knowledge base ───────────────────────────────────────────────────────
    async buildKnowledgeBase(appendMode = false) {
        if (!this._activeCase)
            return;
        const btn = byID(ids.btnBuildKb);
        if (btn) {
            btn.disabled = true;
            btn.textContent = '🧠 Génération…';
        }
        try {
            const markdown = await this.claude.buildCaseKnowledgeBase(this._activeCase, this.readCaseFile, [], // TODO: pass existing KB docs fingerprints from _meta if tracked
            appendMode);
            const filename = `_kb${this._activeCase.name}_${this.claude.kbTimestamp()}.md`;
            const filePath = `${this.casePath(this._activeCase.folderName)}/${filename}`;
            await this.writeFilePath(filePath, markdown, 'text/markdown');
            this._caseKb = markdown;
            toast('Base de connaissance générée et sauvegardée.', 'success');
        }
        catch (err) {
            toast('Erreur KB : ' + err.message, 'error');
        }
        finally {
            if (btn) {
                btn.disabled = false;
                btn.textContent = '🧠 Base de connaissance';
            }
        }
    }
    // ─── Render: case list ────────────────────────────────────────────────────
    renderCaseList() {
        const list = byID(ids.caseList);
        if (!list)
            return;
        list.innerHTML = '';
        const labels = { active: 'En cours', closed: 'Clôturé', suspended: 'Suspendu' };
        for (const c of this._foldersMeta) {
            const item = el('div', { className: 'case-item' + (c.folderName === this._activeCase?.folderName ? ' active' : '') });
            item.dataset.folder = c.folderName;
            item.append(el('span', { className: 'case-item_name', textContent: c.name }), el('span', { className: 'case-item_domain', textContent: c.domain }), el('span', { className: `case-item_status case-item_status--${c.status}`, textContent: labels[c.status] }));
            item.onclick = () => this.selectCase(c.folderName);
            item.addEventListener('contextmenu', (e) => { e.preventDefault(); this.openCaseContextMenu(c, e.clientX, e.clientY); });
            list.appendChild(item);
        }
    }
    // ─── Render: doc list ─────────────────────────────────────────────────────
    renderDocList() {
        const list = byID(ids.docList);
        if (!list || !this._activeCase)
            return;
        list.innerHTML = '';
        const docs = !this._docFilter
            ? this._activeCase.documents
            : this._activeCase.documents.filter((d) => d.kind === this._docFilter);
        if (!docs.length) {
            list.appendChild(el('div', { className: 'doc-empty', textContent: 'Aucun document. Uploadez ou synchronisez OneDrive.' }));
            return;
        }
        for (const doc of [...docs].sort((a, b) => b.addedAt - a.addedAt)) {
            const item = el('div', { className: 'doc-item' });
            const info = el('div', { className: 'doc-info' });
            info.append(el('div', { className: 'doc-name', textContent: doc.name }), el('div', { className: 'doc-meta', textContent: `${formatDate(doc.addedAt)} · ${mimeLabel(doc.mimeType)} · ${formatSize(doc.sizeBytes)}` }));
            const del = el('button', { className: 'doc-delete', textContent: '×', title: 'Retirer du dossier' });
            del.onclick = async (e) => {
                e.stopPropagation();
                const ok = await confirm(`Retirer "${doc.name}" ?\n(Fichier OneDrive conservé, seul l'index local est supprimé.)`);
                if (!ok)
                    return;
                this._activeCase.documents = this._activeCase.documents.filter((d) => d.name !== doc.name);
                await this.saveMeta();
                this.renderDocList();
                this.updateDocCount();
                toast('Document retiré de l\'index.', 'info');
            };
            item.append(el('span', { className: 'doc-icon', textContent: mimeIcon(doc.mimeType) }), info, el('span', { className: `doc-kind doc-kind--${doc.kind}`, textContent: kindLabel(doc.kind) }), del);
            list.appendChild(item);
        }
    }
    updateDocCount() {
        const countEl = byID(ids.docCount);
        if (countEl && this._activeCase)
            countEl.textContent = `${this._activeCase.documents.length} pièce${this._activeCase.documents.length !== 1 ? 's' : ''}`;
    }
    // ─── Render: note bar ─────────────────────────────────────────────────────
    renderNoteBar() {
        if (!this._caseNotes.length)
            return;
        const bar = byID(ids.noteBar);
        if (!bar)
            return;
        bar.innerHTML = '';
        toggle(bar, this._caseNotes.length > 0);
        const n = this._caseNotes.length;
        bar.append(el('span', { className: 'note-badge', textContent: String(n) }), el('span', {
            className: 'note-bar_text',
            textContent: `note${n > 1 ? 's' : ''} active${n > 1 ? 's' : ''} · ` + this._caseNotes.map((x) => x.content.slice(0, 40) + '…').join(' — ')
        }), (() => { const b = el('button', { className: 'note-bar_manage', textContent: 'Gérer' }); b.onclick = () => this.openNotesModal(); return b; })());
    }
    // ─── Render: chat ─────────────────────────────────────────────────────────
    renderChat() {
        const area = byID(ids.chatArea);
        if (!area)
            return;
        area.innerHTML = '';
        if (!this._caseMessages.length) {
            const c = this._activeCase;
            if (!c)
                return;
            area.appendChild(el('div', { className: 'msg msg--assistant' }, el('div', { className: 'msg_label', textContent: 'Lex Assistant' }), el('div', { className: 'msg_bubble' }, el('p', {
                innerHTML: `Dossier <strong>${c.name}</strong>. ${c.documents.length} pièce(s), ${this._caseNotes.length} note(s)${this.skills.length ? `, ${this.skills.length} skill(s)` : ''}${this._caseKb ? ' · <em>Base de connaissance chargée</em>' : ''}.`
            }), el('p', { textContent: 'Que souhaitez-vous faire ?' }))));
            return;
        }
        for (const msg of this._caseMessages)
            area.appendChild(this.buildMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    buildMsgEl(msg) {
        const wrap = el('div', { className: `msg msg--${msg.role}` });
        const bubble = el('div', { className: 'msg_bubble' });
        bubble.innerHTML = renderMarkdown(msg.content);
        wrap.append(el('div', { className: 'msg_label', textContent: msg.role === 'user' ? 'Vous' : 'Lex Assistant' }), bubble);
        if (msg.role === 'assistant') {
            const actions = el('div', { className: 'msg_actions' });
            const copy = el('button', { className: 'msg-action-btn', textContent: 'Copier' });
            copy.onclick = () => { navigator.clipboard.writeText(msg.content); toast('Copié.', 'info', 1500); };
            actions.appendChild(copy);
            if (msg.mode === 'redaction' || msg.generatedDocName) {
                const dl = el('button', { className: 'msg-action-btn msg-action-btn--primary', textContent: '⬇ Télécharger .docx' });
                dl.onclick = async () => {
                    try {
                        const blob = await generateDocx({ title: msg.generatedDocName ?? 'Document', content: msg.content, caseRef: this._activeCase?.name ?? '' });
                        download(blob, (msg.generatedDocName ?? 'document').replace(/[^a-z0-9_\- ]/gi, '_') + '.docx');
                    }
                    catch (err) {
                        toast('Erreur DOCX : ' + err.message, 'error');
                    }
                };
                const odSave = el('button', { className: 'msg-action-btn', textContent: '☁ Sauver sur OneDrive' });
                odSave.onclick = async () => {
                    if (!this._activeCase)
                        return;
                    try {
                        const blob = await generateDocx({ title: msg.generatedDocName ?? 'Document', content: msg.content, caseRef: this._activeCase.name });
                        const ab = await blob.arrayBuffer();
                        const fname = (msg.generatedDocName ?? 'document').replace(/[^a-z0-9_\- ]/gi, '_') + '.docx';
                        const mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
                        await this.writeCaseFile(this._activeCase.folderName, fname, ab, mime);
                        if (!this._activeCase.documents.some((d) => d.name === fname)) {
                            this._activeCase.documents.push({ name: fname, kind: 'redige', mimeType: mime, sizeBytes: ab.byteLength, addedAt: Date.now() });
                            await this.saveMeta();
                            this.renderDocList();
                            this.updateDocCount();
                        }
                        toast(`"${fname}" sauvegardé.`, 'success');
                    }
                    catch (err) {
                        toast('Erreur OneDrive : ' + err.message, 'error');
                    }
                };
                actions.append(dl, odSave);
            }
            wrap.appendChild(actions);
        }
        return wrap;
    }
    // ─── Send message ─────────────────────────────────────────────────────────
    async sendMessage(ta, sendBtn) {
        if (!ta || !this._activeCase)
            return;
        const text = ta.value.trim();
        if (!text)
            return;
        ta.value = '';
        this.autoResize(ta);
        const userMsg = { id: uid(), role: 'user', content: text, timestamp: Date.now(), mode: this.activeMode };
        this._caseMessages.push(userMsg);
        this.appendMsg(this.buildMsgEl(userMsg));
        if (this.activeMode === 'note') {
            const note = { id: uid(), content: text, createdAt: Date.now(), updatedAt: Date.now() };
            this._caseNotes.push(note);
            await this.saveNotes();
            this.renderNoteBar();
        }
        const typing = this.appendTypingTo(ids.chatArea, 'Analyse en cours…');
        if (sendBtn)
            sendBtn.disabled = true;
        try {
            const response = await this.claude.callClaudeCase(this._activeCase.folderName, {
                caseName: this._activeCase.name,
                caseDomain: this._activeCase.domain,
                notes: this._caseNotes,
                docs: this._activeCase.documents,
                skills: this.skills,
                mode: this.activeMode,
                userMessage: text,
                knowledgeBase: this._caseKb ?? undefined,
                readFile: this.readCaseFile,
            });
            typing.remove();
            let docName;
            if (this.activeMode === 'redaction') {
                const first = response.split('\n')[0].replace(/^#+\s*/, '').trim();
                docName = first.length > 0 && first.length < 100 ? first : 'Document rédigé';
            }
            const asst = { id: uid(), role: 'assistant', content: response, timestamp: Date.now(), mode: this.activeMode, generatedDocName: docName };
            this._caseMessages.push(asst);
            this.appendMsg(this.buildMsgEl(asst));
            await this.saveConversation();
        }
        catch (err) {
            typing.remove();
            toast(err.message, 'error', 6000);
            this._caseMessages.pop();
        }
        finally {
            if (sendBtn)
                sendBtn.disabled = false;
            ta.focus();
        }
    }
    // ─── UI setup ─────────────────────────────────────────────────────────────
    setupBarsBtns() {
        const btn = (id) => byID(id);
        this.onClick(btn(ids.btnNewSidebar), () => this.openCaseFormModal(null));
        this.onClick(btn('btn-new-item-top'), () => this.openCaseFormModal(null));
        this.onClick(btn('btn-settings'), () => this.openSettingsModal());
        this.onClick(btn('btn-onedrive'), async () => { await this.refreshCaseFromOneDrive(); });
        this.onClick(btn(ids.btnNotesOpen), () => this.openNotesModal());
        this.onClick(btn(ids.btnOdSync), () => this.refreshCaseFromOneDrive());
        this.onClick(btn(ids.btnUpload), () => btn(ids.fileInput)?.click());
        this.onClick(btn(ids.btnBuildKb), () => this.buildKnowledgeBase());
        this.onClick(btn(ids.btnCaseSummary), async () => {
            const ta = byID(ids.userInput);
            if (!ta)
                return;
            ta.value = 'Fais un point complet sur ce dossier : enjeux principaux, risques identifiés, actions restantes, points d\'attention prioritaires.';
            await this.sendMessage(byID(ids.toast), byID(ids.sendBtn));
        });
        qsa('.doc-filter-tab').forEach((tab) => this.onClick(tab, () => {
            setActive(qsa('.doc-filter-tab'), tab, 'active');
            this._docFilter = (tab.dataset.filter ?? null);
            this.renderDocList();
        }));
        qsa('.mode-btn[data-mode]').forEach((b) => this.onClick(b, () => {
            setActive(qsa('.mode-btn[data-mode]'), b, 'active');
            this.activeMode = b.dataset.mode;
            const hints = {
                analyse: 'Posez une question, demandez une analyse du dossier…',
                redaction: 'Précisez l\'acte à rédiger (courrier, assignation, conclusions, contrat…)',
                modification: 'Indiquez le document à modifier et les changements souhaités…',
                note: 'Rédigez une correction → sauvegardée dans _notes.json…',
            };
            const ta = byID(ids.userInput);
            if (ta)
                ta.placeholder = hints[this.activeMode];
        }));
    }
    updateQuickPrompts(ta) {
        qsa('.quick-btn').forEach((b) => {
            b.onclick = () => { ta.value = b.dataset.prompt ?? ''; this.autoResize(ta); ta.focus(); };
        });
    }
    setupFileUpload() {
        const fi = byID(ids.fileInput);
        if (!fi)
            return;
        fi.onchange = async () => {
            if (!fi.files?.length || !this._activeCase)
                return;
            for (const file of Array.from(fi.files)) {
                if (!isSupported(file.name)) {
                    toast(`Format non supporté : ${file.name}`, 'error');
                    continue;
                }
                try {
                    const ab = await file.arrayBuffer();
                    const meta = makeCaseDocMeta(file);
                    await this.writeCaseFile(this._activeCase.folderName, file.name, ab, meta.mimeType);
                    if (!this._activeCase.documents.some((d) => d.name === file.name)) {
                        this._activeCase.documents.push(meta);
                        await this.saveMeta();
                    }
                    toast(`"${file.name}" ajouté au dossier.`, 'success');
                }
                catch (err) {
                    toast(`Erreur : ${err.message}`, 'error');
                }
            }
            fi.value = '';
            this.renderDocList();
            this.updateDocCount();
        };
    }
    // ─── Modals ───────────────────────────────────────────────────────────────
    openCaseFormModal(existing) {
        if (!oneDrive.userOid) {
            this.openSettingsModal();
            toast('Connectez OneDrive d\'abord.', 'error');
            return;
        }
        const isEdit = !!existing;
        const overlay = el('div', { className: 'modal-overlay' });
        document.body.appendChild(overlay);
        const dialog = el('div', { className: 'modal-dialog modal-dialog--form' });
        overlay.appendChild(dialog);
        dialog.innerHTML = `
      <h2 class="modal-title">${isEdit ? 'Modifier le dossier' : 'Nouveau dossier'}</h2>
      <label class="form-label">Intitulé <span class="required">*</span></label>
      <input class="form-input" id="f-name" type="text" value="${existing?.name ?? ''}"/>
      <label class="form-label">Domaine juridique</label>
      <input class="form-input" id="f-domain" type="text" placeholder="Droit commercial…" value="${existing?.domain ?? ''}"/>
      <label class="form-label">Nom du dossier OneDrive <span style="font-weight:400;color:var(--c-gray-400)">(auto si vide)</span></label>
      <input class="form-input" id="f-folder" type="text" value="${existing?.folderName ?? ''}"/>
      <label class="form-label">Statut</label>
      <select class="form-select" id="f-status">
        <option value="active"    ${!existing || existing.status === 'active' ? 'selected' : ''}>En cours</option>
        <option value="suspended" ${existing?.status === 'suspended' ? 'selected' : ''}>Suspendu</option>
        <option value="closed"    ${existing?.status === 'closed' ? 'selected' : ''}>Clôturé</option>
      </select>
      <div class="modal-btns">
        <button class="btn btn--secondary" id="f-cancel">Annuler</button>
        <button class="btn btn--primary"   id="f-save">${isEdit ? 'Enregistrer' : 'Créer'}</button>
      </div>`;
        const nameInput = qs('#f-name', dialog);
        const folderInput = qs('#f-folder', dialog);
        nameInput.addEventListener('input', () => {
            if (!isEdit && !folderInput.value)
                folderInput.placeholder = this.sanitiseFolder(nameInput.value) || 'DOSSIER_NOM';
        });
        qs('#f-cancel', dialog).onclick = () => overlay.remove();
        qs('#f-save', dialog).onclick = async () => {
            const name = nameInput.value.trim();
            const domain = qs('#f-domain', dialog).value.trim() || 'Droit général';
            const raw = folderInput.value.trim();
            const folder = raw ? this.sanitiseFolder(raw) : this.sanitiseFolder(name);
            const status = qs('#f-status', dialog).value;
            if (!name || !folder) {
                toast(!name ? 'Nom obligatoire.' : 'Dossier invalide.', 'error');
                return;
            }
            const btn = qs('#f-save', dialog);
            btn.disabled = true;
            btn.textContent = 'Création…';
            const now = Date.now();
            const meta = { name, folderName: folder, domain, status, createdAt: existing?.createdAt ?? now, updatedAt: now, documents: existing?.documents ?? [] };
            try {
                await this.writeCaseMeta(folder, meta);
                overlay.remove();
                const c = { ...meta };
                if (isEdit) {
                    const idx = this._foldersMeta.findIndex((x) => x.folderName === existing.folderName);
                    if (idx >= 0)
                        this._foldersMeta[idx] = c;
                    if (this._activeCase?.folderName === existing.folderName)
                        this._activeCase = c;
                }
                else {
                    this._foldersMeta.unshift(c);
                }
                this.renderCaseList();
                await this.selectCase(folder);
                toast(isEdit ? 'Dossier modifié.' : 'Dossier créé.', 'success');
            }
            catch (err) {
                toast('Erreur : ' + err.message, 'error');
                btn.disabled = false;
                btn.textContent = isEdit ? 'Enregistrer' : 'Créer';
            }
        };
        nameInput.focus();
    }
    openCaseContextMenu(c, x, y) {
        byID(ids.contextMenu)?.remove();
        const menu = el('div', { className: 'context-menu', id: ids.contextMenu });
        menu.style.left = `${x}px`;
        menu.style.top = `${y}px`;
        const items = [
            { label: 'Modifier', action: () => this.openCaseFormModal(c) },
            { label: '☁ Sync OneDrive', action: () => this.refreshCaseFromOneDrive() },
            { label: '🧠 Compléter base KB', action: () => this.buildKnowledgeBase(true) },
            { label: 'Effacer conversation', action: () => this.clearConversation(c.folderName) },
            { label: 'Supprimer le dossier', action: () => this.deleteCaseIndex(c.folderName), danger: true },
        ];
        for (const item of items) {
            const btn = el('button', { className: 'context-menu_item' + (item.danger ? ' context-menu_item--danger' : ''), textContent: item.label });
            btn.onclick = () => { menu.remove(); item.action(); };
            menu.appendChild(btn);
        }
        document.body.appendChild(menu);
        const dismiss = (e) => { if (!menu.contains(e.target)) {
            menu.remove();
            document.removeEventListener('click', dismiss);
        } };
        setTimeout(() => document.addEventListener('click', dismiss), 0);
    }
    async clearConversation(folderName) {
        if (!await confirm('Effacer tout l\'historique de conversation de ce dossier ?'))
            return;
        this._caseMessages = [];
        await this.writeConversation(folderName, []);
        if (this._activeCase?.folderName === folderName)
            this.renderChat();
        toast('Conversation effacée.', 'info');
    }
    async deleteCaseIndex(folderName) {
        if (!await confirm('Supprimer ce dossier ? Les fichiers OneDrive sont conservés, seuls les fichiers Lex Assistant (_meta, _notes, _conversation, _kb_*) sont supprimés.'))
            return;
        const items = await this.listAllFolderItems(this.casePath(folderName));
        const toDelete = items.filter((i) => i.file && (i.name.startsWith('_meta') || i.name.startsWith('_notes') || i.name.startsWith('_conversation') || i.name.startsWith('_kb_')));
        await Promise.all(toDelete.map((i) => this.deleteFilePath(`${this.casePath(folderName)}/${i.name}`).catch(() => { })));
        this._foldersMeta = this._foldersMeta.filter((c) => c.folderName !== folderName);
        this.renderCaseList();
        if (this._activeCase?.folderName === folderName) {
            this._activeCase = null;
            if (this._foldersMeta.length > 0)
                await this.selectCase(this._foldersMeta[0].folderName);
            else
                this.showEmptyState();
        }
        toast('Dossier supprimé.', 'success');
    }
    openNotesModal() {
        const overlay = el('div', { className: 'modal-overlay' });
        const dialog = el('div', { className: 'modal-dialog modal-dialog--notes' });
        const refresh = () => {
            const list = qs('#notes-list', dialog);
            list.innerHTML = '';
            if (!this._caseNotes.length) {
                list.appendChild(el('p', { className: 'note-empty', textContent: 'Aucune note. Ajoutez-en une ci-dessous ou utilisez le mode "Note permanente".' }));
                return;
            }
            for (const note of [...this._caseNotes].sort((a, b) => b.createdAt - a.createdAt)) {
                const row = el('div', { className: 'note-row' });
                const del = el('button', { className: 'btn btn--sm btn--danger', textContent: 'Supprimer' });
                del.onclick = async () => {
                    this._caseNotes = this._caseNotes.filter((n) => n.id !== note.id);
                    await this.saveNotes();
                    this.renderNoteBar();
                    refresh();
                };
                row.append(el('p', { className: 'note-content', textContent: note.content }), el('span', { className: 'note-meta', textContent: formatDateTime(note.createdAt) }), del);
                list.appendChild(row);
            }
        };
        dialog.innerHTML = `
      <h2 class="modal-title">Notes permanentes du dossier</h2>
      <p class="modal-subtitle">Sauvegardées dans <code>_notes.json</code>. Réinjectées dans chaque appel Claude (priorité absolue).</p>
      <div id="notes-list"></div><hr>
      <label class="form-label">Ajouter une correction</label>
      <textarea class="form-textarea" id="new-note" rows="3" placeholder="Ex. : Les pénalités CGI art.1727 ne sont pas justifiées — contester systématiquement."></textarea>
      <div class="modal-btns">
        <button class="btn btn--secondary" id="n-close">Fermer</button>
        <button class="btn btn--primary"   id="n-add">Ajouter</button>
      </div>`;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        refresh();
        qs('#n-close', dialog).onclick = () => overlay.remove();
        qs('#n-add', dialog).onclick = async () => {
            const val = qs('#new-note', dialog).value.trim();
            if (!val)
                return;
            const note = { id: uid(), content: val, createdAt: Date.now(), updatedAt: Date.now() };
            this._caseNotes.push(note);
            await this.saveNotes();
            this.renderNoteBar();
            refresh();
            qs('#new-note', dialog).value = '';
            toast('Note sauvegardée dans _notes.json.', 'success');
        };
    }
    showEmptyState() {
        const area = byID(ids.chatArea);
        if (!area)
            return;
        area.innerHTML = '';
        area.appendChild(el('div', { className: 'empty-state' }, el('div', { className: 'empty-icon', textContent: '⚖️' }), el('h2', { textContent: 'Bienvenue dans Lex Assistant' }), el('p', { textContent: 'Connectez OneDrive et créez votre premier dossier.' }), (() => { const b = el('button', { className: 'btn btn--primary', textContent: '+ Nouveau dossier' }); b.onclick = () => this.openCaseFormModal(null); return b; })()));
    }
}
// ─── Library — bibliothèque scenario ─────────────────────────────────────────
export class Library extends Common {
    activeDomain = 'all';
    domainDocs = new Map();
    libMessages = [];
    _libKb = null; // cached knowledge base for active domain
    DOMAINS = [
        { id: 'commercial', label: 'Commercial', icon: '🏢' },
        { id: 'fiscal', label: 'Fiscal', icon: '💰' },
        { id: 'social', label: 'Social', icon: '👥' },
        { id: 'civil', label: 'Civil', icon: '⚖️' },
        { id: 'penal', label: 'Pénal', icon: '🔒' },
        { id: 'immobilier', label: 'Immobilier', icon: '🏠' },
        { id: 'international', label: 'International', icon: '🌐' },
        { id: 'autre', label: 'Autre', icon: '📚' },
    ];
    // ─── Boot ─────────────────────────────────────────────────────────────────
    async showUI() {
        const content = byID(ids.content);
        content.innerHTML = '';
        content.className = 'bibliotheque-view';
        content.appendChild(this.buildUI());
        if (this.user()) {
            this.updateODStatus();
            this.setupInputArea();
            this.renderDomainPills();
            this.libMessages = await this.readLibConversation().catch(() => []);
            this.renderChat();
            this.renderDocList();
            this.setupFileUpload();
            await this.fetchSkills();
            this.updateSkillIndicator();
        }
        else {
            this.showNotConnected();
        }
    }
    // ─── UI builder ───────────────────────────────────────────────────────────
    buildUI() {
        const wrap = el('div', { className: 'lib-layout' });
        const sidebar = el('div', { className: 'lib-sidebar' });
        const hdr = el('div', { className: 'lib-sidebar_header' });
        hdr.append(el('span', { className: 'lib-sidebar_title', textContent: 'Bibliothèque juridique' }));
        const skillBadge = el('span', { className: 'lib-skill-badge', id: ids.libSkillsBadge });
        toggle(skillBadge, false);
        hdr.appendChild(skillBadge);
        const fi = el('input', { type: 'file', id: ids.fileInput, multiple: true, accept: '.pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt,.txt,.md', style: { display: 'none' } });
        const upBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: '⬆ Ajouter', id: ids.btnLibUpload });
        const synBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: '⟳ Sync OneDrive', id: ids.btnLibSync });
        const kbBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: '🧠 Base KB', id: ids.btnBuildKb });
        upBtn.onclick = () => fi.click();
        synBtn.onclick = async () => await this.syncCurrentDomain();
        kbBtn.onclick = () => this.buildKnowledgeBase();
        const acts = el('div', { className: 'lib-action-row' });
        acts.append(fi, upBtn, synBtn, kbBtn);
        sidebar.append(hdr, el('div', { className: 'lib-domain-pills', id: ids.libDomainPills }), el('div', { className: 'lib-doc-list', id: ids.libDocList }), acts);
        const main = el('div', { className: 'lib-main' });
        const topbar = el('div', { className: 'lib-topbar' });
        const domLbl = el('span', { className: 'lib-topbar_domain', id: ids.libActiveDomain, textContent: 'Tous domaines' });
        const skillLbl = el('span', { className: 'lib-topbar_skills', id: ids.libSkillsTop });
        const clrBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: 'Effacer conversation' });
        clrBtn.onclick = () => this.clearLibConv();
        topbar.append(domLbl, skillLbl, clrBtn);
        const chatArea = el('div', { className: 'chat-area', id: ids.libChatArea, role: 'log' });
        chatArea.setAttribute('aria-live', 'polite');
        const quickArea = el('div', { className: 'lib-quick-prompts', id: ids.quickPrompts });
        const inputArea = el('div', { className: 'lib-input-area' });
        const ta = el('textarea', { id: ids.userInput, rows: 2, placeholder: 'Interrogez la bibliothèque…' });
        const sendBtn = el('button', { className: 'btn btn--primary', id: ids.sendBtn });
        sendBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>';
        inputArea.append(ta, sendBtn);
        main.append(topbar, quickArea, chatArea, inputArea);
        wrap.append(sidebar, main);
        return wrap;
    }
    domainLabel(id) {
        if (id === 'all')
            return 'Tous domaines';
        return this.DOMAINS.find((d) => d.id === id)?.label ?? id;
    }
    // ─── Path helpers ─────────────────────────────────────────────────────────
    libDomainPath(domain) {
        return `${this.root}/${this.mainFolder}/${this.domainLabel(domain)}`;
    }
    libMetaPath(domain) {
        return `${this.libDomainPath(domain)}/_meta.json`;
    }
    getConvPath() {
        return this.activeDomain === 'all'
            ? `${this.root}/${this.mainFolder}/_conversation.json`
            : `${this.libDomainPath(this.activeDomain)}/_conversation.json`;
    }
    // ─── OneDrive CRUD ────────────────────────────────────────────────────────
    async readLibMeta(domain) {
        return this.readJson(this.libMetaPath(domain));
    }
    async writeLibMeta(meta) {
        await this.ensureFolder(this.libDomainPath(meta.domain));
        await this.writeJson(this.libMetaPath(meta.domain), meta);
    }
    async readLibConversation() {
        return (await this.readJson(this.getConvPath()))?.messages ?? [];
    }
    writeLibConversation(messages) {
        return this.writeJson(this.getConvPath(), { messages });
    }
    async listLibFiles(domain) {
        return this.listFiles(this.libDomainPath(domain));
    }
    async readLibFile(domain, fileName) {
        return this.readFilePath(`${this.libDomainPath(domain)}/${fileName}`);
    }
    async writeLibFile(domain, fileName, data, mimeType) {
        await this.ensureFolder(this.libDomainPath(domain));
        await this.writeFileLarge(`${this.libDomainPath(domain)}/${fileName}`, data, mimeType);
    }
    // ─── Domain meta ──────────────────────────────────────────────────────────
    async loadDomainMeta(domain) {
        if (this.domainDocs.has(domain))
            return this.domainDocs.get(domain);
        const docs = (await this.readLibMeta(domain).catch(() => null))?.documents ?? [];
        this.domainDocs.set(domain, docs);
        return docs;
    }
    async saveDomainMeta(domain, docs) {
        this.domainDocs.set(domain, docs);
        await this.writeLibMeta({ domain, documents: docs });
    }
    // ─── Init ─────────────────────────────────────────────────────────────────
    async initRootStructure() {
        await super.initRootStructure();
        for (const d of this.DOMAINS) {
            await this.ensureFolder(this.libDomainPath(d.id));
        }
    }
    // ─── Sync from OneDrive ───────────────────────────────────────────────────
    async syncDomainFromOneDrive(domain) {
        if (!oneDrive.account)
            await oneDrive.signIn();
        //if (!oneDrive.userOid) await oneDrive.signIn(this.msal);
        const items = await this.listLibFiles(domain);
        const existing = await this.loadDomainMeta(domain);
        let added = 0;
        for (const item of items) {
            if (!item.file || !isSupported(item.name))
                continue;
            if (existing.some((d) => d.name === item.name))
                continue;
            existing.push({ name: item.name, mimeType: item.file.mimeType || 'application/octet-stream', sizeBytes: item.size ?? 0, addedAt: Date.now(), tags: [] });
            added++;
        }
        if (added > 0)
            await this.saveDomainMeta(domain, existing);
        return added;
    }
    async syncCurrentDomain() {
        const btn = byID(ids.btnLibSync);
        if (btn) {
            btn.disabled = true;
            btn.textContent = '⟳ Sync…';
        }
        try {
            let total = 0;
            if (this.activeDomain === 'all') {
                for (const d of this.DOMAINS)
                    total += await this.syncDomainFromOneDrive(d.id);
            }
            else {
                total = await this.syncDomainFromOneDrive(this.activeDomain);
            }
            this.renderDocList();
            this.renderDomainPills();
            toast(`${total} nouveau(x) document(s) indexé(s).`, 'success');
        }
        catch (err) {
            toast('Erreur sync : ' + err.message, 'error');
        }
        finally {
            if (btn) {
                btn.disabled = false;
                btn.textContent = '⟳ Sync OneDrive';
            }
        }
    }
    // ─── Knowledge base ───────────────────────────────────────────────────────
    async buildKnowledgeBase(appendMode = false) {
        if (this.activeDomain === 'all') {
            toast('Sélectionnez un domaine spécifique pour générer une base de connaissance.', 'info');
            return;
        }
        const btn = byID(ids.btnBuildKb);
        if (btn) {
            btn.disabled = true;
            btn.textContent = '🧠 Génération…';
        }
        try {
            const docs = this.domainDocs.get(this.activeDomain) ?? [];
            const markdown = await this.claude.buildLibKnowledgeBase(this.activeDomain, docs, (name) => this.readLibFile(this.activeDomain, name), [], appendMode);
            // Save versioned markdown to OneDrive
            const filename = `_kb_${this.activeDomain}_${this.claude.kbTimestamp}.md`;
            const filePath = `${this.libDomainPath(this.activeDomain)}/${filename}`;
            await this.writeFilePath(filePath, markdown, 'text/markdown');
            this._libKb = markdown;
            toast('Base de connaissance bibliothèque générée et sauvegardée.', 'success');
        }
        catch (err) {
            toast('Erreur KB : ' + err.message, 'error');
        }
        finally {
            if (btn) {
                btn.disabled = false;
                btn.textContent = '🧠 Base de connaissance';
            }
        }
    }
    // ─── Domain navigation ────────────────────────────────────────────────────
    renderDomainPills() {
        const container = byID(ids.libDomainPills);
        if (!container)
            return;
        container.innerHTML = '';
        const all = el('button', { className: `lib-pill${this.activeDomain === 'all' ? ' active' : ''}`, textContent: 'Tous' });
        all.onclick = () => this.switchDomain('all');
        container.appendChild(all);
        for (const d of this.DOMAINS) {
            const docs = this.domainDocs.get(d.id) ?? [];
            const pill = el('button', { className: `lib-pill${this.activeDomain === d.id ? ' active' : ''}` });
            pill.textContent = `${d.icon} ${d.label}`;
            if (docs.length)
                pill.appendChild(el('span', { className: 'lib-pill_count', textContent: String(docs.length) }));
            pill.onclick = () => this.switchDomain(d.id);
            container.appendChild(pill);
        }
    }
    async switchDomain(domain) {
        this.activeDomain = domain;
        this._libKb = null;
        if (domain !== 'all') {
            await this.loadDomainMeta(domain);
            // Try to load the latest KB for this domain
            const kb = await this.claude.loadLatestCaseKb(this.libDomainPath(domain), (p) => this.listAllFolderItems(p), (p) => this.readFilePath(p));
            if (kb)
                this._libKb = kb.content;
        }
        this.libMessages = await this.readLibConversation().catch(() => []);
        this.renderDomainPills();
        this.renderDocList();
        this.renderChat();
        const lbl = byID(ids.libActiveDomain);
        if (lbl)
            lbl.textContent = this.domainLabel(domain);
        this.updateQuickPrompts(byID(ids.toast));
    }
    // ─── Render: lib doc list ─────────────────────────────────────────────────
    renderDocList() {
        const list = byID(ids.libDocList);
        if (!list)
            return;
        list.innerHTML = '';
        let docs = [];
        if (this.activeDomain === 'all') {
            for (const [, d] of this.domainDocs)
                docs.push(...d);
        }
        else {
            docs = this.domainDocs.get(this.activeDomain) ?? [];
        }
        docs = docs.sort((a, b) => b.addedAt - a.addedAt);
        if (!docs.length) {
            list.appendChild(el('div', { className: 'lib-doc-empty', textContent: 'Aucun document. Ajoutez ou synchronisez.' }));
            return;
        }
        for (const doc of docs) {
            const item = el('div', { className: 'lib-doc-item' });
            const info = el('div', { className: 'doc-info' });
            info.append(el('div', { className: 'doc-name', textContent: doc.name }), el('div', { className: 'doc-meta', textContent: `${mimeLabel(doc.mimeType)} · ${formatSize(doc.sizeBytes)} · ${formatDate(doc.addedAt)}` }));
            const del = el('button', { className: 'doc-delete', textContent: '×', title: 'Retirer de la bibliothèque' });
            del.onclick = async (e) => {
                e.stopPropagation();
                if (!await confirm(`Retirer "${doc.name}" de la bibliothèque ? (Fichier OneDrive conservé.)`))
                    return;
                for (const [dom, list] of this.domainDocs) {
                    const idx = list.findIndex((d) => d.name === doc.name);
                    if (idx >= 0) {
                        list.splice(idx, 1);
                        await this.saveDomainMeta(dom, list);
                        break;
                    }
                }
                this.renderDocList();
                this.renderDomainPills();
                toast('Document retiré de la bibliothèque.', 'info');
            };
            item.append(el('span', { className: 'doc-icon', textContent: mimeIcon(doc.mimeType) }), info, del);
            list.appendChild(item);
        }
    }
    // ─── Render: lib chat ─────────────────────────────────────────────────────
    renderChat() {
        const area = byID(ids.libChatArea);
        if (!area)
            return;
        area.innerHTML = '';
        if (!this.libMessages.length) {
            area.appendChild(el('div', { className: 'empty-state' }, el('div', { className: 'empty-icon', textContent: '📚' }), el('h2', { textContent: 'Bibliothèque juridique' }), el('p', { textContent: 'Sélectionnez un domaine, synchronisez vos documents OneDrive, puis posez votre question.' })));
            return;
        }
        for (const msg of this.libMessages)
            area.appendChild(this.buildMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    buildMsgEl(msg) {
        const wrap = el('div', { className: `msg msg--${msg.role}` });
        const bubble = el('div', { className: 'msg_bubble' });
        bubble.innerHTML = renderMarkdown(msg.content);
        wrap.append(el('div', { className: 'msg_label', textContent: msg.role === 'user' ? 'Vous' : 'Lex Assistant' }), bubble);
        if (msg.role === 'assistant') {
            const acts = el('div', { className: 'msg_actions' });
            const copy = el('button', { className: 'msg-action-btn', textContent: 'Copier' });
            copy.onclick = () => { navigator.clipboard.writeText(msg.content); toast('Copié.', 'info', 1500); };
            acts.appendChild(copy);
            wrap.appendChild(acts);
        }
        return wrap;
    }
    async clearLibConv() {
        if (!await confirm('Effacer l\'historique de la bibliothèque pour ce domaine ?'))
            return;
        this.libMessages = [];
        await this.writeLibConversation([]);
        this.renderChat();
        toast('Conversation effacée.', 'info');
    }
    // ─── Send message ─────────────────────────────────────────────────────────
    async sendMessage(ta, sendBtn) {
        if (!ta || !sendBtn)
            return;
        const text = ta.value.trim();
        if (!text)
            return;
        ta.value = '';
        let docs = [];
        if (this.activeDomain === 'all') {
            for (const [, d] of this.domainDocs)
                docs.push(...d);
            for (const d of this.DOMAINS) {
                if (!this.domainDocs.has(d.id))
                    docs.push(...(await this.loadDomainMeta(d.id)));
            }
        }
        else {
            docs = await this.loadDomainMeta(this.activeDomain);
        }
        const userMsg = { id: uid(), role: 'user', content: text, timestamp: Date.now(), domain: this.activeDomain };
        this.libMessages.push(userMsg);
        this.appendMsg(this.buildMsgEl(userMsg));
        const typing = this.appendTypingTo(ids.libChatArea, 'Consultation de la bibliothèque…');
        sendBtn.disabled = true;
        const history = this.libMessages.slice(-21, -1).map((m) => ({ role: m.role, content: m.content }));
        try {
            const response = await this.claude.callClaudeLib(this.activeDomain, {
                domain: this.activeDomain,
                docs,
                skills: this.skills,
                userMessage: text,
                history,
                knowledgeBase: this._libKb ?? undefined,
                readFile: this.readLibFile,
            });
            typing.remove();
            const asstMsg = { id: uid(), role: 'assistant', content: response, timestamp: Date.now(), domain: this.activeDomain };
            this.libMessages.push(asstMsg);
            this.appendMsg(this.buildMsgEl(asstMsg));
            await this.writeLibConversation(this.libMessages);
        }
        catch (err) {
            typing.remove();
            toast(err.message, 'error', 6000);
            this.libMessages.pop();
        }
        finally {
            sendBtn.disabled = false;
            ta.focus();
        }
    }
    // ─── Quick prompts ────────────────────────────────────────────────────────
    updateQuickPrompts(ta) {
        const area = byID(ids.quickPrompts);
        if (!area)
            return;
        area.innerHTML = '';
        const prompts = {
            commercial: ['Jurisprudence récente sur la responsabilité du dirigeant pour insuffisance d\'actif.', 'Conditions de validité d\'une clause de non-concurrence en droit commercial français.', 'Règles applicables à la cession de fonds de commerce.'],
            fiscal: ['Analyse la jurisprudence sur l\'abus de droit fiscal (LPF art. L.64).', 'Conditions d\'application de l\'acte anormal de gestion.', 'Jurisprudence récente sur la déductibilité des charges en IS.'],
            social: ['Conditions de validité du licenciement pour motif économique.', 'Analyse jurisprudentielle du harcèlement moral au travail.', 'Règles applicables au transfert du contrat de travail (L.1224-1 CT).'],
            civil: ['Jurisprudence récente sur la responsabilité délictuelle.', 'Conditions de la résolution pour inexécution (C.civ. art. 1224).', 'Évolutions de la jurisprudence sur le préjudice moral.'],
            penal: ['Éléments constitutifs de l\'abus de biens sociaux.', 'Jurisprudence sur la complicité en droit pénal des affaires.', 'Conditions de mise en cause de la responsabilité pénale des personnes morales.'],
            immobilier: ['Régime des baux commerciaux : droit au renouvellement et indemnité d\'éviction.', 'Conditions de l\'action en garantie des vices cachés en droit immobilier.', 'Jurisprudence sur la responsabilité du promoteur immobilier.'],
            international: ['Conditions d\'applicabilité des conventions fiscales bilatérales.', 'Jurisprudence sur le centre des intérêts vitaux (CGI art. 4 B).', 'Règles de conflit de lois en matière successorale (Règl. UE 650/2012).'],
            all: ['Quels sont les documents disponibles dans la bibliothèque ?', 'Synthèse des principales règles jurisprudentielles sur la responsabilité civile.', 'Analyse comparative des régimes de responsabilité civile et pénale du dirigeant.'],
        };
        const list = prompts[this.activeDomain] ?? prompts.all;
        for (const p of list) {
            const btn = el('button', { className: 'quick-btn', textContent: p });
            btn.onclick = () => { {
                ta.value = p;
                ta.focus();
            } };
            area.appendChild(btn);
        }
    }
    setupFileUpload() {
        const fi = byID(ids.fileInput);
        if (!fi)
            return;
        fi.onchange = async () => {
            if (!fi.files?.length)
                return;
            const domain = await this.pickDomainModal();
            if (!domain)
                return;
            for (const file of Array.from(fi.files)) {
                if (!isSupported(file.name)) {
                    toast(`Format non supporté : ${file.name}`, 'error');
                    continue;
                }
                try {
                    const ab = await file.arrayBuffer();
                    const meta = makeLibDocMeta(file);
                    await this.writeLibFile(domain, file.name, ab, meta.mimeType);
                    const existing = await this.loadDomainMeta(domain);
                    if (!existing.some((d) => d.name === file.name)) {
                        existing.push(meta);
                        await this.saveDomainMeta(domain, existing);
                    }
                    toast(`"${file.name}" ajouté à ${this.domainLabel(domain)}.`, 'success');
                }
                catch (err) {
                    toast(`Erreur : ${err.message}`, 'error');
                }
            }
            fi.value = '';
            this.renderDocList();
            this.renderDomainPills();
        };
    }
    // ─── Domain picker modal ──────────────────────────────────────────────────
    pickDomainModal() {
        return new Promise((resolve) => {
            const overlay = el('div', { className: 'modal-overlay' });
            const dialog = el('div', { className: 'modal-dialog' });
            dialog.innerHTML = '<h2 class="modal-title">Domaine juridique</h2><p class="modal-subtitle">Dans quel domaine classer ce(s) document(s) ?</p>';
            const grid = el('div', { className: 'domain-grid' });
            for (const d of this.DOMAINS) {
                const btn = el('button', { className: 'domain-btn' });
                btn.innerHTML = `<span class="domain-btn_icon">${d.icon}</span><span>${d.label}</span>`;
                btn.onclick = () => { overlay.remove(); resolve(d.id); };
                grid.appendChild(btn);
            }
            dialog.appendChild(grid);
            const cancel = el('button', { className: 'btn btn--secondary', textContent: 'Annuler' });
            cancel.style.marginTop = '16px';
            cancel.onclick = () => { overlay.remove(); resolve(null); };
            dialog.appendChild(cancel);
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
        });
    }
}
//# sourceMappingURL=onedrive.js.map