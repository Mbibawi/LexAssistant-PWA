
import {
  downloadBlob as download, byID, toast, spinnerEl as spinner, el, toggle, uid,
  formatDate, formatDateTime, qs, qsa, setActive,
} from './ui.js';
import { mimeLabel, mimeIcon, formatSize } from './ingest.js';
import { ids, odSingleton } from '../main.js';
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
  readonly config: OneDriveConfig | null;
  constructor() {
    this.config = this.getConfig() || this.initiateConfig() || null;
  }
  get root(): string { return this.config?.rootFolder ?? APP_ROOT; }
  getConfig(): OneDriveConfig | null {
    const raw = localStorage.getItem(LS_CONFIG);
    try { return raw ? (JSON.parse(raw) as OneDriveConfig) : null; } catch { return null; }
  }
  setConfig(cfg: OneDriveConfig): void {
    localStorage.setItem(LS_CONFIG, JSON.stringify(cfg));
  }
  initiateConfig(): OneDriveConfig | null {
    return this.config//!We are desactivating this for now since we are using the google function onedrive-proxy
    const config: OneDriveConfig = {
      clientId: prompt('Provide the OneDrive Client ID') || '',
      tenantId: prompt('Provide the OneDrive tenant ID') || '',
      rootFolder: APP_ROOT,
    };
    this.setConfig(config);
    return config;
  }
  clearConfig(): void { localStorage.removeItem(LS_CONFIG); }
  isConfigured(): boolean { return Boolean(this.config?.clientId && this.config?.rootFolder); }
}



class MSAL {
  private readonly _clientId: string = "9cb553c1-8473-4b2a-91d4-fef8b7cd7bff";
  private readonly _tenantID: string = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
  private readonly _redirectUri: string = "https://mbibawi.github.io/ExcelInvoicingAddIn"; //!must be the same domain as the app;
  private _app: MsalApp = new msal.PublicClientApplication(this.msalConfig());
  private _initialized: boolean = false;
  private loginRequest = { scopes: [''] };

  constructor(scopes: string[] = ["Files.ReadWrite"]) {
    this.loginRequest.scopes = scopes;
  }

  get msalApp() { return this._app };


  private msalConfig(): msalConfig {
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
    }
  };

  private async init(): Promise<void> {
    if (this._initialized) return;
    this._app = new msal.PublicClientApplication(this.msalConfig());
    await this._app.initialize(); // required in MSAL browser v3+
    this._initialized = true;
    const response = await this._app.handleRedirectPromise();
    if (response?.account) {
      this._app.setActiveAccount(response.account);
    }
  }


  // Function to check existing authentication context
  async acquireToken(): Promise<{ token: string; account: MsalAccount } | null | void> {
    //await this.init();
    try {
      const account = this._app.getAllAccounts()[0];
      if (account) {
        this._app.setActiveAccount(account);
        return await this.acquireTokenSilently(account);
      } else {
        return await this.loginWithPopup();
      }
    } catch (error) {
      console.error("Failed to acquire token from acquireToken(): ", error);
    }
  }
  // Function to get access token silently
  private async acquireTokenSilently(account: any): Promise<{ token: string; account: MsalAccount } | null> {
    try {
      const tokenRequest = {
        account: account,
        scopes: this.loginRequest.scopes, // OneDrive scopes
      };

      const tokenResponse = await this._app.acquireTokenSilent(tokenRequest);
      if (!tokenResponse || !tokenResponse.accessToken) return null;
      console.log("Token acquired silently :", tokenResponse.accessToken);
      return { token: tokenResponse.accessToken, account }

    } catch (error) {
      //if (error instanceof this.msalInstance.InteractionRequiredAuthError)// Silent failed, fall back to popup
      console.error("Token silent acquisition error:", error);
      return await this.loginWithPopup() || null;
    }
  };

  private async loginWithPopup() {
    try {
      const loginResponse = await this._app.loginPopup(this.loginRequest);
      console.log('loginResponse = ', loginResponse);
      const account = loginResponse.account;
      if (!account) return null;
      this._app.setActiveAccount(account);

      const tokenResponse = await this._app.acquireTokenSilent({
        account: account,
        scopes: ["Files.ReadWrite"]
      });

      console.log("Token acquired from loginWithPopup: ", tokenResponse.accessToken);
      return { token: tokenResponse.accessToken, account }
    } catch (error) {
      console.error("Error acquiring token from loginWithPopup(): ", error);
      return null
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

  private async credentitalsToken(tenantId: string) {
    const msalConfig = {
      auth: {
        clientId: this._clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        //clientSecret: clientSecret,
      }
    }
    //@ts-ignore
    const cca = new msal.application.ConfidentialClientApplication(msalConfig);

    const tokenRequest = {
      scopes: ["Files.ReadWrite"],
    }

    try {
      const response = await cca.acquireTokenByClientCredential(tokenRequest);
      return response.accessToken;
    } catch (error) {
      console.log('Error acquiring Token: ', error)
      return null

    }

  }

  private async getOfficeToken() {
    try {
      //@ts-ignore
      return await OfficeRuntime.auth.getAccessToken()

    } catch (error) {
      console.log("Error : ", error)

    }

  }

  private async getTokenWithSSO(email: string, tenantId: string) {
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
    } catch (error) {
      console.error("SSO silent authentication failed:", error);
      return null;
    }
  }

  private openLoginWindow() {
    const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${this._clientId}&response_type=token&redirect_uri=${this._redirectUri}&scope=https://graph.microsoft.com/.default`;

    // Open in a new window (only works if triggered by user action)
    const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");

    if (!authWindow) {
      console.error("Popup blocked! Please allow popups.");
    }
  }

  // Function to handle login and acquire token
  private async loginAndGetToken(): Promise<string | null | undefined> {
    const msalConfig = {
      auth: {
        clientId: this._clientId,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: this._redirectUri
      },

      cache: {
        cacheLocation: "ExcelInvoicing", // Specify cache location
        storeAuthStateInCookie: true  // Set this to true for IE 11
      }
    };

    return await acquire(this);
    async function acquire(this$: MSAL) {
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
      } catch (error) {
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
    async function handleRedirectResponse(this$: MSAL): Promise<string | undefined> {
      try {
        const authResult = await this$._app.handleRedirectPromise();
        if (authResult && authResult.accessToken) {
          console.log("Access token:", authResult.accessToken);
          return authResult.accessToken;
        }
      } catch (error) {
        console.error("Redirect handling error:", error);
      }
      return undefined;
    }
  }

}


// ─── OneDriveAuth — MSAL auth + raw Graph/proxy fetch ────────────────────────

export class OneDriveAuth {
  readonly GRAPH = 'https://graph.microsoft.com/v1.0/me/drive/root:/';
  readonly cfg: Configuration = new Configuration();
  private readonly CLAUDE_PROXY = 'https://claude-ai-proxy-428231091257.europe-west1.run.app/api/proxy/';
  private _scopes: string[] = ['Files.ReadWrite', 'User.Read', 'openid', 'profile'];
  private readonly _MSAL: MSAL = new MSAL();
  private _token: string | null = null;
  private _account: MsalAccount | null = null;
  private _userOid: string | null = null;


  get account() { return this._account; }
  get userName() { return this._account?.name; }
  get token() { return this._token; }
  get config() { return this.cfg.config; }
  get isConfigured() { return this.cfg.isConfigured(); }
  get root() { return this.cfg.root; }
  setConfig(cfg: OneDriveConfig) { this.cfg.setConfig(cfg); }


  /**
   * signIn - Sign in to OneDrive.
   * @returns {Promise<void>}
   */
  async signIn(): Promise<string | null> {
    const acquired = await this._MSAL.acquireToken();
    this._token = acquired?.token || null;
    this._account = acquired?.account || null;
    this._token ? alert(`Signed in successfully, user: ${this.userName}...`) : alert(`Failed to sign in`);
    return this._token;
  }

  /**
   * signOut - Sign out of OneDrive.
   * @returns {Promise<void>}
   */
  async signOut(): Promise<void> {
    this._account = null;
    this._token = null;
    sessionStorage.clear();
  }

  /**
   * isSignedIn - Check if the user is signed in.
   * @returns {Promise<string | null>} The signed-in user or null.
   */
  async isSignedIn(): Promise<string | null> {
    try {
      const accounts = this._MSAL.msalApp.getAllAccounts();
      if (accounts.length) {
        this._account = accounts[0];
        return this._account.name ?? this._account.username;
      }
      return null;
    } catch { return null; }
  }



  /**
   * oneDriveProxy - Proxy for OneDrive operations.
   * @param roote The root folder for the OneDrive operations.
   * @param method The HTTP method for the request.
   * @param payload The payload for the request.
   * @returns {Promise<any>} The response from the OneDrive operations.
   */
  async oneDriveProxy(roote: string, method: string, payload: { path: string, body?: Uint8Array<ArrayBuffer> | ArrayBuffer | string, mimeType?: string }) {
    if (!this.userName) await this.signIn();
    console.log('user oid = ', odSingleton._userOid)
    const url = `https://onedrive-proxy-428231091257.europe-west1.run.app/api/proxy/${roote}`;

    const { body, path, mimeType } = payload;

    const headers: HeadersInit = {
      'x-path': path || '',
      'x-mime-type': mimeType || 'application/octet-stream',
      'x-user': this._userOid || '',
      'x-token': this._token || '',
    };

    const response = await fetch(url, {
      method: method,
      headers: headers,
      // data is sent as the raw binary body
      body: body || null
    });

    if (roote === 'fetch' && response.ok) return await response.blob();

    const result = await response.json();
    if (!response.ok) throw new Error(result.message || 'Proxy Error');
    return result;
  }


  /**
   * Universal fetch for both Graph API and external URLs (GCF proxy).
   * - If url starts with 'https://' it is used verbatim (external call).
   * - Otherwise it is appended to the Graph base URL.
   * - rawBody=true skips automatic Content-Type injection for binary/proxy calls.
   */
  async gFetch(
    path: string,
    opts: RequestInit = {},
    rawBody = false,
  ): Promise<Response> {
    const token = this._token ?? await this.signIn();

    const url = `${this.GRAPH}${path}`;
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      ...(opts.headers as Record<string, string> ?? {}),
    };


    if (!rawBody && opts.body && typeof opts.body === 'string') {
      headers['Content-Type'] = 'application/json';
    }

    const resp = await fetch(url, { ...opts, headers });
    if (!resp.ok) {
      let msg = resp.statusText;
      try {
        const e = (await resp.json()) as { error?: { message?: string } };
        msg = e.error?.message ?? msg;
      } catch { }
      console.log(`Fetch ${resp.status}: ${msg}`);
      //throw new Error(`Fetch ${resp.status}: ${msg}`);
    }
    return resp;
  }

  encode(odPath: string): string {
    return odPath.split('/').map((seg) => encodeURIComponent(seg)).join('/');
  }
}

// ─── Folders — base file/folder/JSON operations ───────────────────────────────

export class Folders {
  private readonly SUPPORTED_EXTS: Record<string, string> = {
    pdf: 'application/pdf',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc: 'application/msword',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls: 'application/vnd.ms-excel',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppt: 'application/vnd.ms-powerpoint',
    txt: 'text/plain',
    md: 'text/markdown',
    rtf: 'application/rtf',
  }
  private readonly od = odSingleton;
  protected readonly gFetch = this.od.gFetch;
  protected readonly encode = this.od.encode;
  protected readonly isSignedIn = this.od.isSignedIn;
  protected readonly setConfig = this.od.setConfig;
  protected readonly signIn = this.od.signIn;
  protected readonly signOut = this.od.signOut;
  protected readonly root = this.od.root;
  protected readonly GRAPH = this.od.GRAPH;
  protected readonly isConfigured = this.od.isConfigured;
  protected readonly config = this.od.config;
  protected get _token() { return this.od.token };
  private get appConfigPath(): string {
    return `${this.root}/${APP_CONFIG_FILE}`;
  }
  protected mainFolder: string | null = null;

  constructor(mainFolder: string) {
    this.mainFolder = mainFolder;
  }
  /**
   * Ensure a folder exists.
   * @param folderPath The path to the folder.
   */
  async ensureFolder(folderPath: string): Promise<void> {
    try {
      const resp = await this.gFetch(folderPath);
      if (!resp.ok) await this.createFolder(folderPath);
    } catch {
    }
  }

  async createFolder(folderPath: string) {
    const parts = folderPath.split('/');
    const name = parts.pop()!;
    const parentPath = parts.join('/');
    const parentEndpoint = parentPath ? `${parentPath}:/children` : ':children';
    const resp = await this.gFetch(parentEndpoint, {
      method: 'POST',
      body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }),
    });
    if (!resp.ok) alert('Failed to create a new folder at: \n' + folderPath);
  }


  /**
   * List all items in a folder.
   * @param folderPath The path to the folder.
   */
  protected async listFolderItems(folderPath: string): Promise<GraphDriveItem[]> {
    //return this.oneDriveProxy('list', 'GET', { path: folderPath });
    const resp = await this.gFetch(
      `${folderPath}:/children?$select=name,size,file,folder,webUrl,lastModifiedDateTime&$top=500`,
    );
    if (!resp.ok) return [];
    const data = (await resp.json()) as { value: GraphDriveItem[] };
    return data.value ?? [];
  }


  /**
   * Lists immediate subfolders of a path relative to root.
   * Used by both Cases (list dossiers) and Library (list domains).
   * @param parentRelPath The relative path to the parent folder.
   */
  protected async listSubFolders(parentRelPath: string): Promise<string[]> {
    try {
      const items = await this.listFolderItems(`${this.root}/${parentRelPath}`);
      return items.filter((f) => f.folder).map((i) => i.name);
    } catch { return []; }
  }

  /**
   * Lists non-underscore files in a folder.
   * Used by both Cases and Library to enumerate documents.
   * @param folderAbsPath The absolute path to the folder.
   */
  protected async listFiles(folderAbsPath: string): Promise<DocumentMeta[]> {
    const items = await this.listFolderItems(folderAbsPath);
    return items.filter((i) => i.file).map(file => this.makeDocMeta(file));
  }


  protected makeDocMeta(item: GraphDriveItem): DocumentMeta {
    return {
      id: item.id ?? null,
      name: item.name,
      kind: this.guessKind(item.name),
      mimeType: this.guessMime(item.name, item.file!.mimeType),
      sizeBytes: item.size,
      addedAt: Date.now(),
      size: item.size,
      tags: [],
    }
  }

  private guessKind(name: string): DocKind {
    const l = name.toLowerCase();
    if (/jurisp|arrêt|arret|décision|cass|Cass|conseil.d.état/.test(l)) return 'jurisprudence';
    if (/doctrine|article|revue|doctr/.test(l)) return 'doctrine';
    return 'piece';
  }
  private guessMime(fileName: string, typehint = ''): string {
    if (typehint) return typehint;
    const ext = fileName.split('.').pop()?.toLowerCase() ?? '';
    return this.SUPPORTED_EXTS[ext] ?? 'application/octet-stream';
  }

  protected isSupported(name: string): boolean {
    const ext = name.split('.').pop()?.toLowerCase() ?? '';
    return ext in this.SUPPORTED_EXTS;
  }

  /**
   * Reads a JSON file from the given file path.
   * @param filePath The path to the JSON file.
   * @returns The parsed JSON data, or null if the file cannot be read.
   */
  async readJson<T>(filePath: string): Promise<T | null> {
    try {
      const buf = await this.readFilePath(filePath);
      const text = new TextDecoder().decode(buf);
      return JSON.parse(text) as T;
    } catch { return null; }
  }

  /**
   * Reads the content of a file from OneDrive.
   * @param filePath The path to the file to read.
   * @returns A Promise that resolves to the content of the file as an ArrayBuffer.
   */
  protected async readFilePath(filePath: string): Promise<ArrayBuffer> {
    //return this.oneDriveProxy('fetch', 'GET', { path: filePath });
    const resp = await this.gFetch(`${filePath}:/content`);
    if (!resp.ok) throw new Error(`Read ${filePath}: ${resp.status}`);
    return resp.arrayBuffer();
  }

  /**
    * Writes JSON data to a file.
   * @param filePath The path to the file to write.
   * @param data The data to write to the file.
   */
  async writeJson(filePath: string, data: unknown): Promise<void> {
    await this.writeFilePath(filePath, JSON.stringify(data, null, 2), 'application/json');
  }

  /**
   * Writes a file to OneDrive.
   * @param filePath The path to the file to write.
   * @param data The data to write to the file.
   * @param mimeType The MIME type of the file.
   */
  protected async writeFileLarge(filePath: string, data: ArrayBuffer, mimeType: string): Promise<void> {
    if (data.byteLength <= 4 * 1024 * 1024) {
      await this.writeFilePath(filePath, data, mimeType);
      return;
    }
    /*const sessResp = await this.oneDriveProxy('save', 'POST', {
      path: `${filePath}/createUploadSession`,
      body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
      mimeType
    });*/
    const sessResp = await this.gFetch(
      `${filePath}/createUploadSession`, {
      method: 'POST',
      body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
    });
    const { uploadUrl } = (await sessResp.json()) as { uploadUrl: string };
    const chunk = 10 * 1024 * 1024;
    for (let off = 0; off < data.byteLength; off += chunk) {
      const end = Math.min(off + chunk, data.byteLength);
      const r = await fetch(uploadUrl, {
        method: 'PUT',
        headers: { 'Content-Range': `bytes ${off}-${end - 1}/${data.byteLength}` },
        body: data.slice(off, end),
      });
      if (!r.ok && r.status !== 202) throw new Error(`Chunk upload failed at ${off}`);
    }
  }

  /**
   * Writes a file to OneDrive.
   * @param filePath The path to the file to write.
   * @param data The data to write to the file.
   * @param mimeType The MIME type of the file.
   */
  protected async writeFilePath(filePath: string, data: ArrayBuffer | string, mimeType: string): Promise<void> {
    const body = typeof data === 'string' ? new TextEncoder().encode(data) : data;
    //return this.oneDriveProxy('save', 'POST', { path: `${filePath}:/content`, body, mimeType });
    await this.gFetch(`${filePath}:/content`, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${this._token}`, 'Content-Type': mimeType },
      body,
    }, true);
  }

  protected async deleteFilePath(filePath: string): Promise<void> {
    //return this.oneDriveProxy('delete', 'DELETE', { path: filePath });
    await this.gFetch(`${filePath}`, { method: 'DELETE' });
  }

  protected async readAppConfig(): Promise<{ apiKey?: string } | null> {
    return this.readJson<{ apiKey?: string }>(this.appConfigPath);
  }

  protected async writeAppConfig(cfg: Record<string, unknown>): Promise<void> {
    await this.writeJson(this.appConfigPath, cfg);
  }

  async initRootStructure(): Promise<void> {
    const r = this.root;
    await this.ensureFolder(r);
    await this.ensureFolder(`${r}/${FOLDER_SKILLS}`);
    await this.ensureFolder(`${r}/${this.mainFolder}`);
    const existing = await this.readAppConfig();
    if (!existing) await this.writeAppConfig({});
  }
}

// ─── Scenario — shared: skills, OD status, UI helpers ────────────────────────

abstract class Common extends Folders {
  readonly claude = CLAUDE;
  //readonly config  = oneDrive.config;
  protected activeMode: WorkMode = 'analyse';
  protected _skills: Array<{ name: string; content: string }> = [];
  protected _kb: string | null = null;  // cached knowledge base content
  protected _notes: PermanentNote[] = [];
  protected _foldersMeta: FolderMeta[] = [];
  protected _conversation: ClaudeMessage[] = [];
  protected _messages: ClaudeMessage[] = [];


  // ─── Abstract Methods ───────────────────────────────────────────────────────────────
  abstract showUI(): void;
  protected abstract buildUI(content: HTMLElement): { sendBtn: HTMLButtonElement, userInput: HTMLTextAreaElement };
  protected abstract renderDocList(): void;
  protected abstract buildMsgEl(msg: ClaudeMessage): HTMLElement;
  protected abstract sendMessage(ta: HTMLTextAreaElement, sendBtn: HTMLButtonElement): Promise<void>;
  protected abstract updateQuickPrompts(chatInput: HTMLTextAreaElement): void;
  protected abstract setupFileUpload(): void;
  protected abstract mainPath(f: string): string;
  protected readonly abstract _chatBody: HTMLDivElement;



  protected folderMeta = async (folderName: string | LibDomain) => await this.readMeta<FolderMeta>(folderName);
  protected metaPath = (f: string) => `${this.mainPath(f)}/_meta.json`;
  protected notesPath = (f: string) => `${this.mainPath(f)}/_notes.json`;
  protected convPath = (f: string) => `${this.mainPath(f)}/_conversation.json`;
  protected async readMeta<T>(f: string): Promise<T | null> {
    return await this.readJson<T>(this.metaPath(f))
  }
  protected async writeMeta(name: string, meta: DocumentMeta | FolderMeta): Promise<void> {
    await this.ensureFolder(this.mainPath(name));
    await this.writeJson(this.metaPath(name), meta);
  }
  protected async readConversation<T>(path: string): Promise<T[]> {
    const messages = await this.readJson<{ messages: T[] }>(this.convPath(path));
    return messages?.messages ?? [];
  }
  protected async writeConversation(f: string, m: ClaudeMessage[]) {
    if (!f) return Promise.resolve();
    return await this.writeJson(this.convPath(f), { messages: m } satisfies ConversationFile);
  }

  async readFile(folderName: string, fileName: string): Promise<ArrayBuffer> {
    return await this.readFilePath(`${this.mainPath(folderName)}/${fileName}`);
  }
  protected async processResponse(response: string[], folder: FolderMeta, messages: ClaudeMessage[], mode?: WorkMode, docName?: string) {
    const asst: ClaudeMessage = {
      id: uid(),
      role: 'assistant',
      content: { type: 'text', text: response.join('\n') },
      timestamp: Date.now(),
      domain: folder.domain as LibDomain,
      mode: mode,
      generatedDocName: docName
    };
    messages.push(asst);
    this.appendMsg(this.buildMsgEl(asst));
    await this.writeConversation(folder!.folderName, messages);
  }

  protected renderChat(messages: ClaudeMessage[], chatBody: HTMLElement): void {
    if (!messages.length) return;
    const area = byID(ids.chatArea);
    if (!area) return;
    area.innerHTML = '';
    area.appendChild(chatBody);
    for (const msg of messages) area.appendChild(this.buildMsgEl(msg));
    (area as HTMLElement).scrollTop = (area as HTMLElement).scrollHeight;
  }


  // ─── Skills ───────────────────────────────────────────────────────────────

  /**
   * Fetches all .md/.txt files from the _Skills folder and loads them.
   * Called once after OneDrive connection. Used by both Cases and Library.
   */
  async fetchSkills(): Promise<Array<{ name: string; content: string }>> {
    const path = `${this.root}/${FOLDER_SKILLS}`;
    try {
      const items = await this.listFolderItems(path);
      const skills: Array<{ name: string; content: string }> = [];
      for (const item of items.filter((i) => !i.folder)) {
        const ext = item.name.split('.').pop()?.toLowerCase() ?? '';
        if (!['md', 'txt'].includes(ext)) continue;
        try {
          const buf = await this.readFilePath(`${path}/${item.name}`);
          skills.push({ name: item.name, content: new TextDecoder().decode(buf) });
        } catch { }
      }
      this._skills = skills;
      this.updateSkillIndicator();
      return skills;
    } catch { return []; }
  }

  /**
   * Builds a knowledge base for the selected case folder or library domain.
   * @param caller 
   * @param activeFolder 
   * @param appendMode 
   * @returns 
   */
  async buildKB(caller: Cases | Library, activeFolder: FolderMeta | null, appendMode: Boolean = false) {
    const type = caller instanceof Cases ? 'dossier' : caller instanceof Library ? 'domaine' : null;
    if (!activeFolder || activeFolder.domain === 'all') {
      if (type) toast(`Vous devez sélectionner un ${type} spécifique pour générer une base de connaissance pour ce ${type}`, 'info');
      return;
    }
    const btn = byID(ids.btnBuildKb) as HTMLButtonElement | null;
    if (btn) { btn.disabled = true; btn.textContent = '🧠 Génération…'; }
    try {
      const markdown =
        caller instanceof Cases ?
          await this.claude.buildCaseKnowledgeBase(
        activeFolder,
        caller,
        [], // TODO: pass existing KB docs fingerprints from _meta if tracked
        appendMode,
          )
          : caller instanceof Library ? await this.claude.buildLibKnowledgeBase(activeFolder, caller, activeFolder.documents, appendMode)
            : null;
      if (!markdown) return;

      const filename = `_kb${activeFolder.name ?? activeFolder.domain}_${this.claude.kbTimestamp()}.md`;
      const filePath = `${this.mainPath(activeFolder.folderName)}/${filename}`;
      await this.writeFilePath(filePath, markdown, 'text/markdown');
      this._kb = markdown;
      toast('Base de connaissance générée et sauvegardée.', 'success');
    } catch (err) {
      toast('Erreur KB : ' + (err as Error).message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = '🧠 Base de connaissance'; }
    }

  }

  // ─── OneDrive connection ──────────────────────────────────────────────────

  updateODStatus(): void {
    const statusEl = byID(ids.oneDriveStatus);
    if (!statusEl) return;
    if (odSingleton.userName) {
      statusEl.textContent = `☁ ${odSingleton.userName}`;
      statusEl.className = 'od-status od-status--connected';
    } else {
      statusEl.textContent = '☁ Non connecté';
      statusEl.className = 'od-status od-status--disconnected';
    }
  }

  /**
   * 
   * @param msg 
   * @returns void
   */
  protected appendMsg(msg: HTMLElement): void {
    const area = byID(ids.chatArea);
    if (!area) return;
    area.querySelector('.empty-state')?.remove();
    area.appendChild(msg);
    (area as HTMLElement).scrollTop = area.scrollHeight;
  }

  /**
   * 
   * @returns 
   */
  protected setupInputArea(userInput: HTMLTextAreaElement, sendBtn: HTMLButtonElement): void {
    if (!userInput || !sendBtn) return;
    const send = () => this.sendMessage(userInput, sendBtn);
    sendBtn.onclick = send;
    userInput.addEventListener('input', () => this.autoResize(userInput));
    userInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        send();
      }
    });
    this.updateQuickPrompts(userInput);
  }
  // ─── Settings modal (OneDrive only — no API key) ──────────────────────────

  openSettingsModal(): void {
    byID(ids.settingsOverlay)?.remove();
    const userName = this.isSignedIn();
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
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });

    qs<HTMLButtonElement>('#s-close', dialog).onclick = () => overlay.remove();
    qs<HTMLButtonElement>('#s-od-save', dialog).onclick = () => {
      const clientId = qs<HTMLInputElement>('#s-client', dialog).value.trim();
      const tenantId = qs<HTMLInputElement>('#s-tenant', dialog).value.trim() || 'common';
      const rootFolder = qs<HTMLInputElement>('#s-root', dialog).value.trim() || 'LexAssistant';
      if (!clientId) { toast('Client ID requis.', 'error'); return; }
      this.setConfig({ clientId, tenantId, rootFolder });
      toast('Configuration OneDrive enregistrée.', 'success');
    };
    qs<HTMLButtonElement>('#s-od-init', dialog).onclick = async () => {
      if (!this.isConfigured) { toast('Sauvegardez la configuration d\'abord.', 'error'); return; } //!might need to be change to if(!oneDrive.userOid)
      try {
        //if (!oneDrive.userOid) { await oneDrive.signIn(); oneDrive.userOid = oneDrive.getSignedInUser(); this.updateODStatus(); }
        await this.initRootStructure();
        toast('Structure initialisée avec succès.', 'success');
      } catch (err) { toast('Erreur : ' + (err as Error).message, 'error'); }
    };
    qs<HTMLButtonElement>('#s-od-signout', dialog).onclick = async () => {
      await this.signOut();
      this.updateODStatus();
      toast('Déconnecté.', 'info');
      overlay.remove();
    };
  }


  // ─── Skill indicator ──────────────────────────────────────────────────────

  updateSkillIndicator(): void {
    const n = this._skills.length;
    for (const id of [ids.skillsBadge, ids.libSkillsBadge]) {
      const badge = byID(id);
      if (!badge) continue;
      badge.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''}` : '';
      toggle(badge as HTMLElement, n > 0);
    }
    const top = byID(ids.libSkillsTop);
    if (top) top.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''} actif${n > 1 ? 's' : ''}` : '';
  }

  // ─── Shared UI helpers ────────────────────────────────────────────────────

  protected onClick(btn: HTMLButtonElement | null, action: () => void) {
    if (btn) btn.onclick = action;
  }

  protected autoResize(ta: HTMLTextAreaElement): void {
    ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight, 160) + 'px';
  }

  protected sanitiseFolder(name: string): string {
    return name.replace(/[/\\:*?"<>|]/g, '_').replace(/\s+/g, '_').slice(0, 60);
  }

  showNotConnected(): void {
    const area = byID(ids.chatArea);
    if (!area) return;
    area.innerHTML = '';
    area.appendChild(el(
      'div', { className: 'empty-state' },
      el('div', { className: 'empty-icon', textContent: '☁' }),
      el('h2', { textContent: 'OneDrive non connecté' }),
      el('p', { textContent: 'Configurez votre App Registration Azure et connectez-vous.' }),
      (() => {
        const b = el('button', { className: 'btn btn--primary', textContent: '☁ Configurer OneDrive' });
        b.onclick = () => this.openSettingsModal();
        return b;
      })(),
    ));
  }

  // ─── Shared typing indicator ──────────────────────────────────────────────

  protected appendTypingTo(areaId: string, label: string): HTMLElement {
    const area = byID(areaId)!;
    const typing = el('div', { className: 'msg msg--assistant', id: ids.typing });
    const bubble = el('div', { className: 'msg_bubble' });
    bubble.append(spinner(), el('span', { textContent: ` ${label}` }));
    typing.append(el('div', { className: 'msg_label', textContent: 'Lex Assistant' }), bubble);
    area.appendChild(typing);
    (area as HTMLElement).scrollTop = area.scrollHeight;
    return typing;
  }
}


// ─── Cases — dossiers scenario ────────────────────────────────────────────────

export class Cases extends Common {
  private _saving = false;
  private _docFilter: DocKind | null = null;
  private _activeCase: FolderMeta | null = null;

  protected readonly _chatBody = el('div', { className: 'msg msg--assistant' },
    el('div', { className: 'msg_label', textContent: 'Lex Assistant' }),
    el('div', { className: 'msg_bubble' },
      el('p', {
        innerHTML: `Dossier <strong>${this._activeCase?.name}</strong>. ${this._activeCase?.documents.length} pièce(s), ${this._notes.length} note(s)${this._skills.length ? `, ${this._skills.length} skill(s)` : ''}${this._kb ? ' · <em>Base de connaissance chargée</em>' : ''}.`
      }),
      el('p', { textContent: 'Que souhaitez-vous faire ?' }),
    ));

  // ─── Path helpers ─────────────────────────────────────────────────────────

  protected mainPath = (f: string) => `${this.root}/${this.mainFolder}/${f}`;


  // ─── Show UI ──────────────────────────────────────────────────────────────
  async showUI(): Promise<void> {
    const content = byID(ids.content)!;
    const { userInput, sendBtn } = this.buildUI(content);
    // Wire all scenario-specific UI
    this.setupBarsBtns();
    this.setupInputArea(userInput, sendBtn);
    if (!odSingleton.userName) await odSingleton.signIn();
    if (!odSingleton.userName) return this.showNotConnected();
    this.updateODStatus();
    await this.loadAllSubFolders();//!this must come before renderChat(), because it sets this._activeCase
    await this.fetchSkills();
    this.updateSkillIndicator();
  }
  protected buildUI(content: HTMLElement): { userInput: HTMLTextAreaElement, sendBtn: HTMLButtonElement } {
    content.innerHTML = '';
    content.className = 'dossiers-view';
    const main = el('div', { id: ids.mainLayout });
    content.appendChild(main);
    const aside = el('aside', { id: ids.sidebar });
    const wSpace = el('div', { id: ids.workspace });
    main.append(aside, wSpace);

    // Sidebar
    aside.append(
      el('div', { className: 'sidebar_section-title', innerText: 'Dossiers' }),
      el('div', { id: ids.caseList }),
      el('button', { id: ids.btnNewSidebar, className: 'btn btn--ghost btn--dashed', innerText: '+ Nouveau dossier' }),
      el('div', { className: 'sidebar_divider' }),
      el('div', { className: 'sidebar_section-title', innerHTML: "Pièces <span id='doc-count' class='doc-count'>0 pièces</span>" }),
    );

    const filterTabs = el('div', { className: ids.docFilter });
    const types = {
      all: 'Tout',
      piece: 'Pièces',
      jurisprudence: 'Jurisprudence',
      doctrine: 'Doctrine',
      redige: 'Rédigés'
    }
    filterTabs.append(
      ...(['all', 'piece', 'jurisprudence', 'doctrine', 'redige'] as const).map((f) =>
        el('button', {
          className: `${ids.docFilter} ${f === 'all' ? ' active' : ''}`, 'data-filter': f,
          innerText: types[f]
        }),
      ),
    );

    const fileInput = el('input', {
      type: 'file', id: ids.fileInput, multiple: true,
      accept: '.pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt,.txt,.md,.rtf', style: { display: 'none' }
    });
    const upload = el('div', { className: 'sidebar_upload-row' });
    upload.append(
      fileInput,
      el('button', { id: ids.btnUpload, className: 'btn btn--ghost btn--sm', innerText: '⬆ Upload' }),
      el('button', { id: ids.btnOdSync, className: 'btn btn--ghost btn--sm', innerText: '☁ Sync' }),
      el('button', { id: ids.btnNotesOpen, className: 'btn btn--ghost btn--sm', innerText: '📌 Notes' }),
    );
    aside.append(filterTabs, el('div', { id: ids.docList, className: 'doc-list' }), upload);

    // Workspace
    const mode = el('div', { id: ids.modeBar });
    const chatInput = el('div', { id: ids.inputArea });
    const prompts = el('div', { id: ids.quickPrompts });
    wSpace.append(
      mode,
      el('div', { id: ids.noteBar, className: 'note-bar', style: { display: 'none' } }),
      el('div', { id: ids.chatArea, className: 'chat-area', role: 'log', 'aria-live': 'polite' }),
      prompts,
      chatInput,
    );

    mode.append(
      el('span', { className: 'mode-bar_label', innerText: 'Mode :' }),
      ...(['analyse', 'redaction', 'modification', 'note'] as WorkMode[]).map((m, i) =>
        el('button', {
          className: `mode-btn${i === 0 ? ' active' : ''}`, 'data-mode': m,
          innerText: ({ analyse: 'Analyse', redaction: 'Rédaction', modification: 'Modification', note: 'Note permanente' })[m]
        }),
      ),
      el('div', { className: 'mode-bar_spacer' }),
      el('button', { id: ids.btnCaseSummary, className: 'btn btn--ghost btn--sm', innerText: 'Point dossier ↗' }),
      el('button', { id: ids.btnBuildKb, className: 'btn btn--ghost btn--sm', innerText: '🧠 Base de connaissance' }),
    );

    prompts.append(
      ...([
        ['Risques du dossier', 'Analyse les risques juridiques et fiscaux du dossier et liste les points d\'attention prioritaires.'],
        ['Mise en demeure', 'Rédige une mise en demeure formelle à la partie adverse sur la base des pièces du dossier.'],
        ['Chronologie des faits', 'Fais une synthèse chronologique des faits pertinents issus des pièces du dossier.'],
        ['Analyse chiffrée', 'Analyse les données chiffrées des tableaux Excel et leurs implications juridiques et fiscales.'],
        ['Mémorandum juridique', 'Rédige un mémorandum juridique complet sur le point de droit central avec jurisprudence applicable.'],
      ] as [string, string][]).map(([label, prompt]) => el('button', { className: 'quick-btn', 'data-prompt': prompt, innerText: label })),
    );

    const sendBtn = el('button', {
      id: ids.sendBtn, className: 'btn btn--primary', 'aria-label': 'Envoyer',
      innerHTML: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>'
    });
    const userInput = el('textarea', {
      id: ids.userInput,
      rows: 2,
      placeholder: 'Posez une question, demandez la rédaction d\'un acte, ou donnez une instruction…', 'aria-label': 'Message'
    });
    chatInput.append(
      userInput,
      sendBtn,
    );
    return { userInput, sendBtn };
  }

  // ─── OneDrive CRUD ────────────────────────────────────────────────────────


  private async readNotes(f: string) { return (await this.readJson<NotesFile>(this.notesPath(f)))?.notes ?? []; }
  private writeNotes(f: string, n: PermanentNote[]) { return this.writeJson(this.notesPath(f), { notes: n } satisfies NotesFile); }


  private async writeCaseFile(folderName: string, fileName: string, data: ArrayBuffer, mimeType: string): Promise<void> {
    await this.ensureFolder(this.mainPath(folderName));
    await this.writeFileLarge(`${this.mainPath(folderName)}/${fileName}`, data, mimeType);
  }

  // ─── Persist helpers ──────────────────────────────────────────────────────

  private async saveMeta(): Promise<void> {
    if (!this._activeCase || this._saving) return;
    this._saving = true;
    const meta: FolderMeta = {
      name: this._activeCase.name,
      folderName: this._activeCase.folderName,
      domain: this._activeCase.domain,
      status: this._activeCase.status,
      createdAt: this._activeCase.createdAt,
      updatedAt: Date.now(),
      documents: this._activeCase.documents,
    };
    this._activeCase.updatedAt = meta.updatedAt;
    try { await this.writeMeta(meta.folderName, meta); }
    finally { this._saving = false; }
  }

  private saveNotes() { return this._activeCase ? this.writeNotes(this._activeCase.folderName, this._notes) : Promise.resolve(); }




  // ─── Load all cases ───────────────────────────────────────────────────────

  protected async loadAllSubFolders(): Promise<void> {
    const subFolders = await this.listSubFolders(this.mainFolder!);
    this._foldersMeta = [];
    await Promise.all(
      subFolders.map(async (folderName) => {
        const meta = await this.folderMeta(folderName);
      if (!meta) return;
        this._foldersMeta.push(meta);
    }));
    if (!this._foldersMeta.length) return;
    this._foldersMeta.sort((a, b) => b!.updatedAt - a!.updatedAt);
    const folderMeta = this.folderMeta;
    const caseMeta = await findCaseMeta(this._foldersMeta);
    if (!caseMeta) return alert('We could not find a case with the folder name  you provided');
    await this.selectCase(caseMeta);//!this must come before the case list is rendered
    this.renderCaseList();

    async function findCaseMeta(metas: FolderMeta[]) {
      const folderName = prompt("Enter the folder name of the case you want to select") || '';
      return metas.find(meta => meta?.folderName === folderName) ?? await folderMeta(folderName) ?? null;
    }
  }

  // ─── Select case ──────────────────────────────────────────────────────────

  private async selectCase(caseMeta: FolderMeta): Promise<void> {
    if (!caseMeta) return;
    const { folderName, name, domain } = caseMeta;
    this._activeCase = caseMeta;
    this._kb = null;
    this._notes = await this.readNotes(folderName);
    this._messages = await this.readConversation(folderName);
    this._activeCase.documents = await this.listFiles(this.mainPath(this._activeCase.folderName));

    // Try to load the latest knowledge base silently
    this._kb = await this.claude.loadLatestCaseKb(
      folderName,
      await this.listFolderItems(folderName),
      this,
    );

    qsa<HTMLElement>('.case-item').forEach((el) =>
      el.classList.toggle('active', el.dataset.folder === caseMeta.folderName));
    const nameEl = byID(ids.topBarCaseName);
    const domainEl = byID(ids.topBarCaseDomain);
    if (nameEl) nameEl.textContent = name;
    if (domainEl) domainEl.textContent = domain || null;
    this.renderNoteBar();
    this.renderDocList();
    this.renderChat(this._messages, this._chatBody);
    this.updateDocCount();
  }

  async refreshCaseFromOneDrive(): Promise<void> {
    if (!this._activeCase) return await this.loadAllSubFolders();
    try {
      const files = await this.listFiles(this.mainPath(this._activeCase.folderName));
      let added = 0;
      for (const file of files) {
        if (!this.isSupported(file.name)) continue;
        if (this._activeCase.documents.some((d) => d.name === file.name)) continue;
        this._activeCase.documents.push(file);
        added++;
      }
      if (added > 0) { await this.saveMeta(); this.renderDocList(); this.updateDocCount(); }
      toast(`${added} nouveau(x) document(s) indexé(s) depuis OneDrive.`, 'success');
    } catch (err) { toast('Erreur sync : ' + (err as Error).message, 'error'); }
  }

  // ─── Knowledge base ───────────────────────────────────────────────────────

  // ─── Render: case list ────────────────────────────────────────────────────

  private renderCaseList(): void {
    const list = byID(ids.caseList);
    if (!list) return;
    list.innerHTML = '';
    const labels = { active: 'En cours', closed: 'Clôturé', suspended: 'Suspendu' };
    this._foldersMeta
      .forEach(meta => {
        const { folderName, status, domain, name } = meta;
        const item = el('div', { className: 'case-item' + (folderName === this._activeCase?.folderName ? ' active' : '') });
        item.dataset.folder = folderName;
      item.append(
        el('span', { className: 'case-item_name', textContent: name }),
        el('span', { className: 'case-item_domain', textContent: domain }),
        el('span', { className: `case-item_status case-item_status--${status}`, textContent: labels[status] }),
      );
        item.onclick = () => this.selectCase(meta);
        item.addEventListener('contextmenu', (e) => { e.preventDefault(); this.openCaseContextMenu(meta, e.clientX, e.clientY); });
      list.appendChild(item);
      })
  }

  // ─── Render: doc list ─────────────────────────────────────────────────────

  protected renderDocList(): void {
    const list = byID(ids.docList);
    if (!list || !this._activeCase) return;
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
      info.append(
        el('div', { className: 'doc-name', textContent: doc.name }),
        el('div', { className: 'doc-meta', textContent: `${formatDate(doc.addedAt)} · ${mimeLabel(doc.mimeType)} · ${formatSize(doc.sizeBytes)}` }),
      );
      const del = el('button', { className: 'doc-delete', textContent: '×', title: 'Retirer du dossier' });
      del.onclick = async (e) => {
        e.stopPropagation();
        const ok = await confirm(`Retirer "${doc.name}" ?\n(Fichier OneDrive conservé, seul l'index local est supprimé.)`);
        if (!ok) return;
        this._activeCase!.documents = this._activeCase!.documents.filter((d) => d.name !== doc.name);
        await this.saveMeta();
        this.renderDocList();
        this.updateDocCount();
        toast('Document retiré de l\'index.', 'info');
      };
      item.append(
        el('span', { className: 'doc-icon', textContent: mimeIcon(doc.mimeType) }),
        info,
        el('span', { className: `doc-kind doc-kind--${doc.kind}`, textContent: this.kindLabel(doc.kind) }), del);
      list.appendChild(item);
    }
  }

  private updateDocCount(): void {
    const countEl = byID(ids.docCount);
    if (countEl && this._activeCase)
      countEl.textContent = `${this._activeCase.documents.length} pièce${this._activeCase.documents.length !== 1 ? 's' : ''}`;
  }

  // ─── Render: note bar ─────────────────────────────────────────────────────

  private renderNoteBar(): void {
    if (!this._notes.length) return;
    const bar = byID(ids.noteBar);
    if (!bar) return;
    bar.innerHTML = '';
    toggle(bar, this._notes.length > 0);
    const n = this._notes.length;
    bar.append(
      el('span', { className: 'note-badge', textContent: String(n) }),
      el('span', {
        className: 'note-bar_text',
        textContent: `note${n > 1 ? 's' : ''} active${n > 1 ? 's' : ''} · ` + this._notes.map((x) => x.content.slice(0, 40) + '…').join(' — ')
      }),
      (() => { const b = el('button', { className: 'note-bar_manage', textContent: 'Gérer' }); b.onclick = () => this.openNotesModal(); return b; })(),
    );
  }

  // ─── Render: chat ─────────────────────────────────────────────────────────
  protected buildMsgEl(msg: ClaudeMessage): HTMLElement {
    const wrap = el('div', { className: `msg msg--${msg.role}` });
    const bubble = el('div', { className: 'msg_bubble' });
    bubble.innerHTML = renderMarkdown((msg.content as ChatBlock).text);
    wrap.append(el('div', { className: 'msg_label', textContent: msg.role === 'user' ? 'Vous' : 'Lex Assistant' }), bubble);

    if (msg.role === 'assistant') {
      const actions = el('div', { className: 'msg_actions' });
      const copy = el('button', { className: 'msg-action-btn', textContent: 'Copier' });
      copy.onclick = () => { navigator.clipboard.writeText(((msg.content as ChatBlock).text)); toast('Copié.', 'info', 1500); };
      actions.appendChild(copy);

      if (msg.mode === 'redaction' || msg.generatedDocName) {
        const dl = el('button', { className: 'msg-action-btn msg-action-btn--primary', textContent: '⬇ Télécharger .docx' });
        dl.onclick = async () => {
          try {
            const blob = await generateDocx({ title: msg.generatedDocName ?? 'Document', content: (msg.content as ChatBlock).text, caseRef: this._activeCase?.name ?? '' });
            download(blob, (msg.generatedDocName ?? 'document').replace(/[^a-z0-9_\- ]/gi, '_') + '.docx');
          } catch (err) { toast('Erreur DOCX : ' + (err as Error).message, 'error'); }
        };
        const odSave = el('button', { className: 'msg-action-btn', textContent: '☁ Sauver sur OneDrive' });
        odSave.onclick = async () => {
          if (!this._activeCase) return;
          try {
            const blob = await generateDocx({ title: msg.generatedDocName ?? 'Document', content: (msg.content as ChatBlock).text, caseRef: this._activeCase.name });
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
          } catch (err) { toast('Erreur OneDrive : ' + (err as Error).message, 'error'); }
        };
        actions.append(dl, odSave);
      }
      wrap.appendChild(actions);
    }
    return wrap;
  }

  // ─── Send message ─────────────────────────────────────────────────────────

  protected async sendMessage(chatInput: HTMLTextAreaElement, sendBtn: HTMLButtonElement): Promise<void> {
    if (!chatInput || !this._activeCase) return;
    const text = chatInput.value.trim();
    if (!text) return;

    chatInput.value = '';
    this.autoResize(chatInput);

    const userMsg: ClaudeMessage = {
      id: uid(),
      role: 'user',
      content: { type: 'text', text: text },
      timestamp: Date.now(),
      mode: this.activeMode
    };
    this._messages.push(userMsg);
    this.appendMsg(this.buildMsgEl(userMsg));

    if (this.activeMode === 'note') {
      const note: PermanentNote = {
        id: uid(),
        content: text,
        createdAt: Date.now(),
        updatedAt: Date.now()
      };
      this._notes.push(note);
      await this.saveNotes();
      this.renderNoteBar();
    }

    const typing = this.appendTypingTo(ids.chatArea, 'Analyse en cours…');
    if (sendBtn) sendBtn.disabled = true;

    try {
      const { folderName, name, domain, documents } = this._activeCase;
      const response = await this.claude.callClaudeCase(
        folderName, {
        caseName: name,
          caseDomain: domain!,
          notes: this._notes,
          docs: documents,
          skills: this._skills,
        mode: this.activeMode,
        userMessage: text,
          knowledgeBase: this._kb ?? undefined,
          caller: this
      });

      typing.remove();
      let docName: string | undefined;
      if (this.activeMode === 'redaction') {
        const first = response[0].split('\n')[0].replace(/^#+\s*/, '').trim();
        docName = first.length > 0 && first.length < 100 ? first : 'Document rédigé';
      }
      await this.processResponse(response, this._activeCase, this._messages, this.activeMode, docName);
    } catch (err) {
      typing.remove();
      toast((err as Error).message, 'error', 6000);
      this._messages.pop();
    } finally {
      if (sendBtn) sendBtn.disabled = false;
      chatInput.focus();
    }
  }

  // ─── UI setup ─────────────────────────────────────────────────────────────

  private setupBarsBtns(): void {
    const btn = (id: string) => byID(id) as HTMLButtonElement | null;
    this.onClick(btn(ids.btnNewSidebar), () => this.openCaseFormModal(null));
    this.onClick(btn(ids.btnOdSync), () => this.refreshCaseFromOneDrive());
    this.onClick(btn('btn-new-item-top'), () => this.openCaseFormModal(null));
    this.onClick(btn(ids.settings), () => this.openSettingsModal());
    this.onClick(btn(ids.btnOneDrive), async () => await odSingleton.signIn());
    this.onClick(btn(ids.btnNotesOpen), () => this.openNotesModal());
    this.onClick(btn(ids.btnUpload), () => btn(ids.fileInput)?.click());
    this.onClick(btn(ids.btnBuildKb), () => this.buildKB(this, this._activeCase));
    this.onClick(btn(ids.btnCaseSummary), async () => {
      const chatInput = byID(ids.userInput) as HTMLTextAreaElement | null;
      if (!chatInput) return;
      chatInput.value = 'Fais un point complet sur ce dossier : enjeux principaux, risques identifiés, actions restantes, points d\'attention prioritaires.';
      //await this.sendMessage(chatInput, byID(ids.sendBtn) as HTMLButtonElement);
    });

    qsa<HTMLButtonElement>(`.${ids.docFilter}`).forEach((tab) =>
      this.onClick(tab, () => {
        setActive(qsa(`.${ids.docFilter}`), tab, 'active');
        this._docFilter = (tab.dataset.filter ?? null) as typeof this._docFilter;
        this.renderDocList();
      }),
    );

    qsa<HTMLButtonElement>('.mode-btn[data-mode]').forEach((b) =>
      this.onClick(b, () => {
        setActive(qsa('.mode-btn[data-mode]'), b, 'active');
        this.activeMode = b.dataset.mode as WorkMode;
        const hints: Record<WorkMode, string> = {
          analyse: 'Posez une question, demandez une analyse du dossier…',
          redaction: 'Précisez l\'acte à rédiger (courrier, assignation, conclusions, contrat…)',
          modification: 'Indiquez le document à modifier et les changements souhaités…',
          note: 'Rédigez une correction → sauvegardée dans _notes.json…',
        };
        const ta = byID(ids.userInput) as HTMLTextAreaElement | null;
        if (ta) ta.placeholder = hints[this.activeMode];
      }),
    );
  }

  protected updateQuickPrompts(ta: HTMLTextAreaElement) {
    qsa<HTMLButtonElement>('.quick-btn').forEach((b) => {
      b.onclick = () => {
        ta.value = b.dataset.prompt ?? '';
        this.autoResize(ta);
        ta.focus();
      };
    });
  }

  protected setupFileUpload(): void {
    const fi = byID(ids.fileInput) as HTMLInputElement | null;
    if (!fi) return;
    fi.onchange = async () => {
      if (!fi.files?.length || !this._activeCase) return;
      for (const file of Array.from(fi.files)) {
        if (!this.isSupported(file.name)) { toast(`Format non supporté : ${file.name}`, 'error'); continue; }
        try {
          const ab = await file.arrayBuffer();
          const meta = this.makeDocMeta(file);
          await this.writeCaseFile(this._activeCase.folderName, file.name, ab, meta.mimeType);
          if (!this._activeCase.documents.some((d) => d.name === file.name)) {
            this._activeCase.documents.push(meta);
            await this.saveMeta();
          }
          toast(`"${file.name}" ajouté au dossier.`, 'success');
        } catch (err) { toast(`Erreur : ${(err as Error).message}`, 'error'); }
      }
      fi.value = '';
      this.renderDocList();
      this.updateDocCount();
    };
  }

  // ─── Modals ───────────────────────────────────────────────────────────────

  private async openCaseFormModal(existing: FolderMeta | null): Promise<void> {
    if (!odSingleton.userName) await odSingleton.signIn();
    //if (!oneDrive.account) { this.openSettingsModal(); toast('Connectez OneDrive d\'abord.', 'error'); return; }
    const isEdit = !!existing;
    const overlay = el('div', { className: 'modal-overlay' });
    document.body.appendChild(overlay);
    const dialog = el('div', { className: 'modal-dialog modal-dialog--form' });
    overlay.appendChild(dialog);
    const nameInput = el("input", { className: "form-input", id: "f-name", type: "text", value: `${existing?.name ?? ''}` });
    const folderInput = el("input", { className: "form-input", id: "f-folder", type: "text", value: `${existing?.folderName ?? ''}` });
    const cancel = el("button", { className: "btn btn--secondary", id: "f-cancel", innerText: "Annuler" });
    const save = el("button", { className: "btn btn--primary", id: "f-save", innerText: isEdit ? 'Enregistrer' : 'Créer' });
    const selectStatuts = el("select", { className: "form-select", id: "f-status" });
    const fdomain = el("input", { className: "form-input", id: "f-domain", type: "text", placeholder: "Droit commercial…", value: `${existing?.domain ?? ''}` });
    selectStatuts.append(
      el("option", { value: "active", innerText: "En cours" }),
      el("option", { value: "suspended", innerText: "Suspendu" }),
      el("option", { value: "closed", innerText: "Clôturé" })
    );

    dialog.append(
      el("h2", { className: "modal-title", innerText: `${isEdit ? 'Modifier le dossier' : 'Nouveau dossier'}` }),
      el("label", { className: "form-label", innerHTML: `Intitulé <span class="required">*</span>` }),
      nameInput,
      el("label", { className: "form-label", innerText: `Domaine juridique` }),
      fdomain,
      el("label", { className: "form-label", innerHTML: `Nom du dossier OneDrive <span style="font-weight:400;color:var(--c-gray-400)">(auto si vide)</span></label>` }),
      folderInput,
      el("label", { className: "form-label", innerText: `Statut` }),
      selectStatuts,
      el("div", { className: "modal-btns" },
        cancel,
        save
      )
    );

    nameInput.addEventListener('input', () => {
      if (!isEdit && !folderInput.value) folderInput.placeholder = this.sanitiseFolder(nameInput.value) || 'DOSSIER_NOM';
    });
    cancel.onclick = () => overlay.remove();
    save.onclick = async () => {
      const name = nameInput.value.trim();
      const domain = fdomain.value.trim() || 'Droit général';
      const raw = folderInput.value.trim();
      const folder = raw ? this.sanitiseFolder(raw) : this.sanitiseFolder(name);
      const status = selectStatuts.value as FolderMeta['status'];
      if (!name || !folder) { toast(!name ? 'Nom obligatoire.' : 'Dossier invalide.', 'error'); return; }
      save.disabled = true;
      save.textContent = 'Création…';
      const now = Date.now();
      const meta: FolderMeta = {
        name,
        folderName: folder,
        domain,
        status,
        createdAt: existing?.createdAt ?? now,
        updatedAt: now,
        documents: existing?.documents ?? []
      };
      try {
        await this.writeMeta(folder, meta);
        overlay.remove();
        const c: FolderMeta = { ...meta };
        if (isEdit) {
          const idx = this._foldersMeta.findIndex((x) => x.folderName === existing!.folderName);
          if (idx >= 0) this._foldersMeta[idx] = c;
          if (this._activeCase?.folderName === existing!.folderName) this._activeCase = c;
        } else { this._foldersMeta.unshift(c); }
        this.renderCaseList();
        await this.selectCase(meta);
        toast(isEdit ? 'Dossier modifié.' : 'Dossier créé.', 'success');
      } catch (err) { toast('Erreur : ' + (err as Error).message, 'error'); save.disabled = false; save.textContent = isEdit ? 'Enregistrer' : 'Créer'; }
    };
    nameInput.focus();
  }

  private openCaseContextMenu(c: FolderMeta, x: number, y: number): void {
    if (!c) return;
    byID(ids.contextMenu)?.remove();
    const menu = el('div', { className: 'context-menu', id: ids.contextMenu });
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    const items = [
      { label: 'Modifier', action: () => this.openCaseFormModal(c) },
      { label: '☁ Sync OneDrive', action: () => this.refreshCaseFromOneDrive() },
      { label: '🧠 Compléter base KB', action: () => this.buildKB(this, this._activeCase, true) },
      { label: 'Effacer conversation', action: () => this.clearConversation(c.folderName) },
      { label: 'Supprimer le dossier', action: () => this.deleteCaseIndex(c.folderName), danger: true },
    ];
    for (const item of items) {
      const btn = el('button', { className: 'context-menu_item' + (item.danger ? ' context-menu_item--danger' : ''), textContent: item.label });
      btn.onclick = () => { menu.remove(); item.action(); };
      menu.appendChild(btn);
    }
    document.body.appendChild(menu);
    const dismiss = (e: MouseEvent) => { if (!menu.contains(e.target as Node)) { menu.remove(); document.removeEventListener('click', dismiss); } };
    setTimeout(() => document.addEventListener('click', dismiss), 0);
  }

  private async clearConversation(folderName: string): Promise<void> {
    if (!await confirm('Effacer tout l\'historique de conversation de ce dossier ?')) return;
    this._messages = [];
    await this.writeConversation(folderName, []);
    if (this._activeCase?.folderName === folderName) this.renderChat(this._messages, this._chatBody);
    toast('Conversation effacée.', 'info');
  }

  private async deleteCaseIndex(folderName: string): Promise<void> {
    if (!await confirm('Supprimer ce dossier ? Les fichiers OneDrive sont conservés, seuls les fichiers Lex Assistant (_meta, _notes, _conversation, _kb_*) sont supprimés.')) return;
    const items = await this.listFolderItems(this.mainPath(folderName));
    const toDelete = items.filter((i) => i.file && (i.name.startsWith('_meta') || i.name.startsWith('_notes') || i.name.startsWith('_conversation') || i.name.startsWith('_kb_')));
    await Promise.all(toDelete.map((i) => this.deleteFilePath(`${this.mainPath(folderName)}/${i.name}`).catch(() => { })));
    this._foldersMeta = this._foldersMeta.filter((c) => c.folderName !== folderName);
    this.renderCaseList();
    if (this._activeCase?.folderName === folderName) {
      this._activeCase = null;
      if (this._foldersMeta.length) await this.selectCase(this._foldersMeta[0]);
      else this.showEmptyState();
    }
    toast('Dossier supprimé.', 'success');
  }

  private openNotesModal(): void {
    const overlay = el('div', { className: 'modal-overlay' });
    const dialog = el('div', { className: 'modal-dialog modal-dialog--notes' });
    const refresh = () => {
      const list = qs<HTMLElement>('#notes-list', dialog);
      list.innerHTML = '';
      if (!this._notes.length) {
        list.appendChild(el('p', { className: 'note-empty', textContent: 'Aucune note. Ajoutez-en une ci-dessous ou utilisez le mode "Note permanente".' }));
        return;
      }
      for (const note of [...this._notes].sort((a, b) => b.createdAt - a.createdAt)) {
        const row = el('div', { className: 'note-row' });
        const del = el('button', { className: 'btn btn--sm btn--danger', textContent: 'Supprimer' });
        del.onclick = async () => {
          this._notes = this._notes.filter((n) => n.id !== note.id);
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
    qs<HTMLButtonElement>('#n-close', dialog).onclick = () => overlay.remove();
    qs<HTMLButtonElement>('#n-add', dialog).onclick = async () => {
      const val = qs<HTMLTextAreaElement>('#new-note', dialog).value.trim();
      if (!val) return;
      const note: PermanentNote = { id: uid(), content: val, createdAt: Date.now(), updatedAt: Date.now() };
      this._notes.push(note);
      await this.saveNotes();
      this.renderNoteBar();
      refresh();
      qs<HTMLTextAreaElement>('#new-note', dialog).value = '';
      toast('Note sauvegardée dans _notes.json.', 'success');
    };
  }

  private showEmptyState(): void {
    const area = byID(ids.chatArea);
    if (!area) return;
    area.innerHTML = '';
    area.appendChild(el('div', { className: 'empty-state' },
      el('div', { className: 'empty-icon', textContent: '⚖️' }),
      el('h2', { textContent: 'Bienvenue dans Lex Assistant' }),
      el('p', { textContent: 'Connectez OneDrive et créez votre premier dossier.' }),
      (() => { const b = el('button', { className: 'btn btn--primary', textContent: '+ Nouveau dossier' }); b.onclick = () => this.openCaseFormModal(null); return b; })(),
    ));
  }

  private kindLabel(kind: DocKind | undefined): string {
    if (!kind) return '';
    const m: Record<DocKind, string> = { piece: 'Pièce', jurisprudence: 'Jurisprudence', doctrine: 'Doctrine', redige: 'Rédigé' };
    return kind ? m[kind] : 'Inconnu';
  }

}

// ─── Library — bibliothèque scenario ─────────────────────────────────────────

export class Library extends Common {
  private _activeDomain: FolderMeta = undefined!;
  private _domainDocs = new Map<string, DocumentMeta[]>();

  readonly DOMAINS: Array<{ id: LibDomain; label: string; icon: string }> = [
    { id: 'commercial', label: 'Commercial', icon: '🏢' },
    { id: 'fiscal', label: 'Fiscal', icon: '💰' },
    { id: 'social', label: 'Social', icon: '👥' },
    { id: 'civil', label: 'Civil', icon: '⚖️' },
    { id: 'penal', label: 'Pénal', icon: '🔒' },
    { id: 'immobilier', label: 'Immobilier', icon: '🏠' },
    { id: 'international', label: 'International', icon: '🌐' },
    { id: 'autre', label: 'Autre', icon: '📚' },
  ];

  private get _domain() { return this._activeDomain?.domain ?? "all" };
  // ─── Path helpers ─────────────────────────────────────────────────────────

  protected mainPath(domain: string): string {
    return `${this.root}/${this.mainFolder}/${this.domainLabel(domain)}`;
  }


  // ─── Boot ─────────────────────────────────────────────────────────────────

  async showUI(): Promise<void> {
    const content = byID(ids.content)!;
    const { userInput, sendBtn } = this.buildUI(content);
    if (!odSingleton.userName) return this.showNotConnected();

    this.updateODStatus();
    this.setupInputArea(userInput, sendBtn);
    this.renderDomainPills();
    this._messages = await this.readConversation<ClaudeMessage>(this._domain).catch(() => []);
      this.renderChat(this._messages, this._chatBody);
      this.renderDocList();
      this.setupFileUpload();
      await this.fetchSkills();
    this.updateSkillIndicator();
  }

  // ─── UI builder ───────────────────────────────────────────────────────────

  protected buildUI(content: HTMLElement): { userInput: HTMLTextAreaElement, sendBtn: HTMLButtonElement } {
    content.innerHTML = '';
    content.className = 'bibliotheque-view';
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
    upBtn.onclick = () => (fi as HTMLInputElement).click();
    synBtn.onclick = async () => await this.syncCurrentDomain();
    kbBtn.onclick = () => this.buildKB(this, this._activeDomain);
    const acts = el('div', { className: 'lib-action-row' });
    acts.append(fi, upBtn, synBtn, kbBtn);

    sidebar.append(
      hdr,
      el('div', { className: 'lib-domain-pills', id: ids.libDomainPills }),
      el('div', { className: 'lib-doc-list', id: ids.libDocList }),
      acts,
    );

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
    const chatInput = el('div', { id: ids.inputArea, className: 'lib-input-area' });
    const userInput = el('textarea', { id: ids.userInput, rows: 2, placeholder: 'Interrogez la bibliothèque…' });
    const sendBtn = el('button', { className: 'btn btn--primary', id: ids.sendBtn });
    sendBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>';
    chatInput.append(userInput, sendBtn);
    main.append(topbar, quickArea, chatArea, chatInput);
    wrap.append(sidebar, main);
    return { userInput, sendBtn };
  }

  domainLabel(id: LibDomain | string): string {
    if (id === 'all') return 'Tous domaines';
    return this.DOMAINS.find((d) => d.id === id)?.label ?? id;
  }


  // ─── OneDrive CRUD ────────────────────────────────────────────────────────
  async writeLibFile(domain: LibDomain, fileName: string, data: ArrayBuffer, mimeType: string): Promise<void> {
    await this.ensureFolder(this.mainPath(domain));
    await this.writeFileLarge(`${this.mainPath(domain)}/${fileName}`, data, mimeType);
  }

  // ─── Domain meta ──────────────────────────────────────────────────────────

  async saveDomainMeta(domain: string, meta: FolderMeta | null): Promise<void> {
    if (!meta) return;
    this._domainDocs.set(domain, meta.documents);
    await this.writeMeta(domain, meta);
  }

  // ─── Init ─────────────────────────────────────────────────────────────────

  async initRootStructure(): Promise<void> {
    await super.initRootStructure();
    for (const d of this.DOMAINS) {
      await this.ensureFolder(this.mainPath(d.id));
    }
  }

  // ─── Sync from OneDrive ───────────────────────────────────────────────────

  async syncDomainFromOneDrive(domain: string): Promise<number> {
    if (!odSingleton.userName) await odSingleton.signIn();
    let added = 0;
    const files = await this.listFiles(this.mainPath(domain));
    const meta = await this.folderMeta(domain);
    if (!meta) return added;
    const existing = meta.documents;
    for (const file of files) {
      if (existing.some((d) => d.name === file.name)) continue;
      existing.push({ name: file.name, mimeType: file.mimeType || 'application/octet-stream', sizeBytes: file.size ?? 0, addedAt: Date.now(), tags: [] });
      added++;
    }

    if (added > 0) await this.saveDomainMeta(domain, meta);
    return added;
  }

  async syncCurrentDomain(): Promise<void> {
    const btn = byID(ids.btnLibSync) as HTMLButtonElement | null;
    if (btn) { btn.disabled = true; btn.textContent = '⟳ Sync…'; }
    try {
      let total = 0;
      if (this._domain === 'all') {
        for (const d of this.DOMAINS) total += await this.syncDomainFromOneDrive(d.id);
      } else {
        total = await this.syncDomainFromOneDrive(this._domain);
      }
      this.renderDocList();
      this.renderDomainPills();
      toast(`${total} nouveau(x) document(s) indexé(s).`, 'success');
    } catch (err) { toast('Erreur sync : ' + (err as Error).message, 'error'); }
    finally { if (btn) { btn.disabled = false; btn.textContent = '⟳ Sync OneDrive'; } }
  }

  // ─── Knowledge base ───────────────────────────────────────────────────────


  // ─── Domain navigation ────────────────────────────────────────────────────

  renderDomainPills(): void {
    const container = byID(ids.libDomainPills) as HTMLElement;
    if (!container) return;
    container.innerHTML = '';
    const all = el('button', { className: `lib-pill${this._domain === 'all' ? ' active' : ''}`, textContent: 'Tous' });
    all.onclick = () => this.switchDomain('all');
    container.appendChild(all);
    for (const d of this.DOMAINS) {
      const docs = this._domainDocs.get(d.id) ?? [];
      const pill = el('button', { className: `lib-pill${this._domain === d.id ? ' active' : ''}` });
      pill.textContent = `${d.icon} ${d.label}`;
      if (docs.length) pill.appendChild(el('span', { className: 'lib-pill_count', textContent: String(docs.length) }));
      pill.onclick = () => this.switchDomain(d.id);
      container.appendChild(pill);
    }
  }

  async switchDomain(domain: LibDomain | 'all'): Promise<void> {
    this._activeDomain.domain = domain;
    this._kb = null;
    if (domain !== 'all') {
      await this.folderMeta(domain);
      // Try to load the latest KB for this domain
      this._kb = await this.claude.loadLatestCaseKb(
        domain,
        await this.listFolderItems(domain),
        this,
      );
    }
    this._messages = await this.readConversation<ClaudeMessage>(domain).catch(() => []);
    this.renderDomainPills();
    this.renderDocList();
    this.renderChat(this._messages, this._chatBody);
    const lbl = byID(ids.libActiveDomain);
    if (lbl) lbl.textContent = this.domainLabel(domain);
    this.updateQuickPrompts(qs(ids.userInput) as HTMLTextAreaElement);
  }

  // ─── Render: lib doc list ─────────────────────────────────────────────────
  private getDocsMeta(domain: LibDomain): DocumentMeta[] {
    if (domain === 'all')
      return this.DOMAINS.map(d => this._domainDocs.get(d.id)).flat().filter(docs => docs !== undefined);
    return this._domainDocs.get(domain) ?? [];
  }
  protected renderDocList(): void {
    const list = byID(ids.libDocList);
    if (!list) return;
    list.innerHTML = '';
    let docs: DocumentMeta[] = this.getDocsMeta(this._activeDomain.domain as LibDomain);
    docs = docs.sort((a, b) => b.addedAt - a.addedAt);
    if (!docs.length) { list.appendChild(el('div', { className: 'lib-doc-empty', textContent: 'Aucun document. Ajoutez ou synchronisez.' })); return; }
    for (const doc of docs) {
      const item = el('div', { className: 'lib-doc-item' });
      const info = el('div', { className: 'doc-info' });
      info.append(
        el('div', { className: 'doc-name', textContent: doc.name }),
        el('div', { className: 'doc-meta', textContent: `${mimeLabel(doc.mimeType)} · ${formatSize(doc.sizeBytes)} · ${formatDate(doc.addedAt)}` }),
      );
      const del = el('button', { className: 'doc-delete', textContent: '×', title: 'Retirer de la bibliothèque' });
      del.onclick = async (e) => {
        e.stopPropagation();
        if (!confirm(`Retirer "${doc.name}" de la bibliothèque ? (Fichier OneDrive conservé.)`)) return;
        for (const [dom, list] of this._domainDocs) {
          const idx = list.findIndex((d) => d.name === doc.name);
          if (idx >= 0) {
            list.splice(idx, 1);
            const meta = await this.folderMeta(dom);
            if (!meta) return;
            meta.documents = list;
            await this.saveDomainMeta(dom as LibDomain, meta); break;
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

  private findDocument(name: string) {
    return Array.from(this._domainDocs.values()).flat().find((d) => d.name === name);
  }

  // ─── Render: lib chat ─────────────────────────────────────────────────────
  protected readonly _chatBody = el('div', { className: 'empty-state' },
    el('div', { className: 'empty-icon', textContent: '📚' }),
    el('h2', { textContent: 'Bibliothèque juridique' }),
    el('p', { textContent: 'Sélectionnez un domaine, synchronisez vos documents OneDrive, puis posez votre question.' })
  );

  protected buildMsgEl(msg: ClaudeMessage): HTMLElement {
    const wrap = el('div', { className: `msg msg--${msg.role}` });
    const bubble = el('div', { className: 'msg_bubble' });
    bubble.innerHTML = renderMarkdown((msg.content as ChatBlock).text);
    wrap.append(el('div', { className: 'msg_label', textContent: msg.role === 'user' ? 'Vous' : 'Lex Assistant' }), bubble);
    if (msg.role === 'assistant') {
      const acts = el('div', { className: 'msg_actions' });
      const copy = el('button', { className: 'msg-action-btn', textContent: 'Copier' });
      copy.onclick = () => { navigator.clipboard.writeText((msg.content as ChatBlock).text); toast('Copié.', 'info', 1500); };
      acts.appendChild(copy);
      wrap.appendChild(acts);
    }
    return wrap;
  }


  private async clearLibConv(): Promise<void> {
    if (!await confirm('Effacer l\'historique de la bibliothèque pour ce domaine ?')) return;
    const path = this._domain === 'all'
      ? `${this.root}/${this.mainFolder}/_conversation`
      : this.convPath(this._domain);
    this._messages = [];
    await this.writeJson(path, this._messages);//!this must be writeJson not writeConversation
    this.renderChat(this._messages, this._chatBody);
    toast('Conversation effacée.', 'info');
  }

  // ─── Send message ─────────────────────────────────────────────────────────

  protected async sendMessage(chatInput: HTMLTextAreaElement, sendBtn: HTMLButtonElement): Promise<void> {
    if (!chatInput || !sendBtn) return;
    const text = chatInput.value.trim();
    if (!text) return;
    chatInput.value = '';
    const docs: DocumentMeta[] = this._activeDomain.domain === 'all' ? [] : this.getDocsMeta(this._activeDomain.domain as LibDomain);

    const userMsg: ClaudeMessage = {
      id: uid(), role: 'user',
      content: { type: 'text', text: text },
      timestamp: Date.now(),
      domain: this._activeDomain.domain!
    };
    this._messages.push(userMsg);
    this.appendMsg(this.buildMsgEl(userMsg));

    const typing = this.appendTypingTo(ids.libChatArea, 'Consultation de la bibliothèque…');
    sendBtn.disabled = true;

    const history = this._messages.slice(-21, -1);

    try {
      const response = await this.claude.callClaudeLib(
        {
          domain: this._activeDomain.domain as LibDomain,
        docs,
          skills: this._skills,
        userMessage: text,
        history,
          knowledgeBase: this._kb ?? undefined,
          caller: this
      });
      typing.remove();
      await this.processResponse(response, this._activeDomain, this._messages);
    } catch (err) {
      typing.remove();
      toast((err as Error).message, 'error', 6000);
      this._messages.pop();
    } finally {
      sendBtn.disabled = false;
      chatInput.focus();
    }
  }

  // ─── Quick prompts ────────────────────────────────────────────────────────
  protected updateQuickPrompts(ta: HTMLTextAreaElement): void {
    const area = byID(ids.quickPrompts);
    if (!area) return;
    area.innerHTML = '';
    const prompts: Record<string, string[]> = {
      commercial: ['Jurisprudence récente sur la responsabilité du dirigeant pour insuffisance d\'actif.', 'Conditions de validité d\'une clause de non-concurrence en droit commercial français.', 'Règles applicables à la cession de fonds de commerce.'],
      fiscal: ['Analyse la jurisprudence sur l\'abus de droit fiscal (LPF art. L.64).', 'Conditions d\'application de l\'acte anormal de gestion.', 'Jurisprudence récente sur la déductibilité des charges en IS.'],
      social: ['Conditions de validité du licenciement pour motif économique.', 'Analyse jurisprudentielle du harcèlement moral au travail.', 'Règles applicables au transfert du contrat de travail (L.1224-1 CT).'],
      civil: ['Jurisprudence récente sur la responsabilité délictuelle.', 'Conditions de la résolution pour inexécution (C.civ. art. 1224).', 'Évolutions de la jurisprudence sur le préjudice moral.'],
      penal: ['Éléments constitutifs de l\'abus de biens sociaux.', 'Jurisprudence sur la complicité en droit pénal des affaires.', 'Conditions de mise en cause de la responsabilité pénale des personnes morales.'],
      immobilier: ['Régime des baux commerciaux : droit au renouvellement et indemnité d\'éviction.', 'Conditions de l\'action en garantie des vices cachés en droit immobilier.', 'Jurisprudence sur la responsabilité du promoteur immobilier.'],
      international: ['Conditions d\'applicabilité des conventions fiscales bilatérales.', 'Jurisprudence sur le centre des intérêts vitaux (CGI art. 4 B).', 'Règles de conflit de lois en matière successorale (Règl. UE 650/2012).'],
      all: ['Quels sont les documents disponibles dans la bibliothèque ?', 'Synthèse des principales règles jurisprudentielles sur la responsabilité civile.', 'Analyse comparative des régimes de responsabilité civile et pénale du dirigeant.'],
    };
    const list = prompts[this._activeDomain?.domain ?? 'all'];
    for (const p of list) {
      const btn = el('button', { className: 'quick-btn', textContent: p });
      btn.onclick = () => { { ta.value = p; ta.focus(); } };
      area.appendChild(btn);
    }
  }

  protected setupFileUpload(): void {
    const fi = byID(ids.fileInput) as HTMLInputElement | null;
    if (!fi) return;
    fi.onchange = async () => {
      if (!fi.files?.length) return;
      const domain = await this.pickDomainModal();
      if (!domain) return;
      const folderMeta = await this.folderMeta(domain);
      for (const file of Array.from(fi.files)) {
        if (!this.isSupported(file.name)) { toast(`Format non supporté : ${file.name}`, 'error'); continue; }
        try {
          const buffer = await file.arrayBuffer();
          const fileMeta = this.makeDocMeta(file);
          await this.writeLibFile(domain, file.name, buffer, fileMeta.mimeType);
          if (!folderMeta) return;
          if (!folderMeta.documents.some((d) => d.name === file.name)) {
            folderMeta.documents.push(fileMeta);
            await this.saveDomainMeta(domain, folderMeta);
            toast(`"${file.name}" ajouté à ${this.domainLabel(domain)}.`, 'success');
          }
        } catch (err) { toast(`Erreur : ${(err as Error).message}`, 'error'); }
      }
      fi.value = '';
      this.renderDocList();
      this.renderDomainPills();
    };
  }

  // ─── Domain picker modal ──────────────────────────────────────────────────

  private pickDomainModal(): Promise<LibDomain | null> {
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

