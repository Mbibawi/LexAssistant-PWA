/**
 * onedrive.ts — Microsoft Graph API, path-based, class hierarchy.
 *
 * class Configuration  — localStorage config read/write
 * class oneDrive       — MSAL auth + raw Graph fetch
 * class Folders        — base folder/file/JSON ops (extends oneDrive)
 * class Cases          — dossiers scenario (extends Folders)
 * class Library        — bibliothèque scenario (extends Cases)
 *
 * OneDrive structure:
 *   <root>/
 *     _Skills/                       ← .md/.txt skills
 *     Affaires/<Folder>/
 *       _meta.json                   ← CaseMeta
 *       _notes.json                  ← PermanentNote[]
 *       _conversation.json           ← ChatMessage[]
 *       <file>
 *     Bibliotheque/<Domain>/
 *       _meta.json                   ← LibDomainMeta
 *       _conversation.json           ← LibConversationMessage[]
 *       <file>
 *     Bibliotheque/_conversation.json ← "all" domain conversation
 */
import { oneDrive } from "../main.js";
import { byID, toast, spinnerEl, el, toggle, uid, formatDate, formatDateTime, qs, qsa, setActive, } from "./ui.js";
import { renderMarkdown } from './markdown.js';
import { callClaudeLib, callClaudeCase, getStoredKey, setStoredKey, clearStoredKey } from './api.js';
import { isSupported, mimeLabel, mimeIcon, formatSize, makeLibDocMeta, guessKind, kindLabel } from './ingest.js';
import { downloadBlob, generateDocx } from './docxgen.js';
// ─── Constants ────────────────────────────────────────────────────────────────
const LS_CONFIG = 'lex_onedrive_config';
const MSAL_CDN = "https://cdn.jsdelivr.net/npm/@azure/msal-browser@5.8.0/lib/msal-browser.min.js";
const GRAPH = "https://graph.microsoft.com/v1.0/me/drive/root:/";
const APP_ROOT = "Legal/Mon Cabinet d'Avocat/_LexAssistant";
const FOLDER_LIBRARY = "Bibliotheque";
const FOLDER_CASES = "Affaires";
const FOLDER_SKILLS = "_Skills";
const APP_CONFIG_FILE = "_config.json";
// ─── Configuration ────────────────────────────────────────────────────────────
class Configuration {
    config;
    constructor() {
        this.config = this.getConfig() || this.initateConfig();
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
    initateConfig() {
        const config = {
            clientId: prompt("Provide the OneDrive Client ID") || "",
            tenantId: prompt("Provide the OneDrive tenant ID") || "",
            rootFolder: APP_ROOT,
        };
        this.setConfig(config);
        return config;
    }
    clearConfig() {
        localStorage.removeItem(LS_CONFIG);
    }
    isConfigured() {
        return Boolean(this.config?.clientId && this.config?.rootFolder);
    }
    root() {
        return this.config?.rootFolder ?? APP_ROOT;
    }
}
// ─── oneDrive — MSAL auth + raw Graph fetch ───────────────────────────────────
export class OneDriveAuth {
    config = new Configuration();
    _scopes = ["Files.ReadWrite", "User.Read"];
    _msal = null;
    _loading = null;
    account = null;
    userName = null;
    get isConfigured() {
        return this.config.isConfigured();
    }
    setConfig(cfg) {
        this.config.setConfig(cfg);
    }
    getMsal() {
        if (this._msal)
            return Promise.resolve(this._msal);
        if (this._loading)
            return this._loading;
        this._loading = new Promise((res, rej) => {
            const init = async () => {
                const cfg = this.config.config;
                if (!cfg)
                    throw new Error("OneDrive non configuré");
                const app = new msal.PublicClientApplication({
                    auth: {
                        clientId: cfg.clientId,
                        authority: `https://login.microsoftonline.com/${cfg.tenantId ?? "common"}`,
                        redirectUri: window.location.origin,
                    },
                    cache: {
                        cacheLocation: "localStorage",
                        storeAuthStateInCookie: false,
                    },
                });
                await app.handleRedirectPromise().catch(() => null);
                this._msal = app;
                return app;
            };
            if (msal) {
                init().then(res).catch(rej);
                return;
            }
            const s = document.createElement("script");
            s.src = MSAL_CDN;
            s.onload = () => init().then(res).catch(rej);
            s.onerror = () => rej(new Error("Impossible de charger MSAL"));
            document.head.appendChild(s);
        });
        return this._loading;
    }
    async getAccessToken() {
        const msal = await this.getMsal();
        if (!this.account) {
            const accounts = msal.getAllAccounts();
            if (accounts.length)
                this.account = accounts[0];
        }
        if (this.account) {
            try {
                return (await msal.acquireTokenSilent({
                    scopes: this._scopes,
                    account: this.account,
                })).accessToken;
            }
            catch {
                /* fall through to popup */
            }
        }
        const r = await msal.loginPopup({ scopes: this._scopes });
        this.account = msal.getAllAccounts()[0] ?? null;
        return r.accessToken;
    }
    async signIn() {
        const msal = await this.getMsal();
        await msal.loginPopup({ scopes: this._scopes });
        this.account = msal.getAllAccounts()[0] ?? null;
        if (this.account)
            this.userName = this.getSignedInUser();
    }
    async signOut() {
        this.account = null;
        this._msal = null;
        this._loading = null;
        sessionStorage.clear();
    }
    async isSignedIn() {
        try {
            const msal = await this.getMsal();
            const accounts = msal.getAllAccounts();
            if (accounts.length) {
                this.account = accounts[0];
                return this.getSignedInUser();
            }
            return null;
        }
        catch {
            return null;
        }
    }
    getSignedInUser() {
        return this.account?.name ?? this.account?.username ?? null;
    }
}
// ─── Folders — base file/folder/JSON operations ───────────────────────────────
class Folders {
    get _rootFolder() {
        return oneDrive.config.root();
    }
    _token = null;
    // Encode a path for Graph: "a/b/c" → "/me/drive/root:/a/b/c:"
    encode(odPath) {
        const encoded = odPath
            .split("/")
            .map((seg) => encodeURIComponent(seg))
            .join("/");
        return `${encoded}:`;
    }
    async gFetch(path, opts = {}, rawBody = false) {
        if (!this._token)
            this._token = await oneDrive.getAccessToken();
        const headers = opts.headers ?? {
            Authorization: `Bearer ${this._token}`,
        };
        if (!rawBody && opts.body && typeof opts.body === "string") {
            headers["Content-Type"] =
                "application/json";
        }
        const resp = await fetch(`${GRAPH}${path}`, { ...opts, headers });
        if (!resp.ok) {
            let msg = resp.statusText;
            try {
                const e = (await resp.json());
                msg = e.error?.message ?? msg;
            }
            catch { }
            throw new Error(`Graph ${resp.status}: ${msg}`);
        }
        return resp;
    }
    async ensureFolder(folderPath) {
        try {
            await this.gFetch(this.encode(folderPath));
            return;
        }
        catch { }
        const parts = folderPath.split("/");
        const name = parts.pop();
        const parentPath = parts.join("/");
        const parentEndpoint = parentPath
            ? `${this.encode(parentPath)}/children`
            : `children`;
        await this.gFetch(parentEndpoint, {
            method: "POST",
            body: JSON.stringify({
                name,
                folder: {},
                "@microsoft.graph.conflictBehavior": "rename",
            }),
        });
    }
    async listFolder(folderPath) {
        const resp = await this.gFetch(`${this.encode(folderPath)}/children?$select=name,size,file,folder,webUrl,lastModifiedDateTime&$top=500`);
        const data = (await resp.json());
        return data.value ?? [];
    }
    async readFilePath(filePath) {
        const resp = await this.gFetch(`${this.encode(filePath)}/content, `);
        if (!resp.ok)
            throw new Error(`Read ${filePath}: ${resp.status}`);
        return resp.arrayBuffer();
    }
    async writeJson(filePath, data) {
        await this.writeFilePath(filePath, JSON.stringify(data, null, 2), "application/json");
    }
    async writeFileLarge(filePath, data, mimeType) {
        if (data.byteLength <= 4 * 1024 * 1024) {
            await this.writeFilePath(filePath, data, mimeType);
            return;
        }
        const sessResp = await this.gFetch(`${this.encode(filePath)}/createUploadSession`, {
            method: "POST",
            body: JSON.stringify({
                item: { "@microsoft.graph.conflictBehavior": "replace" },
            }),
        });
        const { uploadUrl } = (await sessResp.json());
        const chunk = 10 * 1024 * 1024;
        for (let off = 0; off < data.byteLength; off += chunk) {
            const end = Math.min(off + chunk, data.byteLength);
            const r = await fetch(uploadUrl, {
                method: "PUT",
                headers: {
                    "Content-Range": `bytes ${off}-${end - 1}/${data.byteLength}`,
                },
                body: data.slice(off, end),
            });
            if (!r.ok && r.status !== 202)
                throw new Error(`Chunk upload failed at ${off}`);
        }
    }
    async writeFilePath(filePath, data, mimeType) {
        if (!this._token)
            this._token = await oneDrive.getAccessToken();
        const body = typeof data === "string" ? new TextEncoder().encode(data) : data;
        const resp = await this.gFetch(`${this.encode(filePath)}/content`, {
            method: "PUT",
            headers: {
                Authorization: `Bearer ${this._token}`,
                "Content-Type": mimeType,
            },
            body,
        });
        if (!resp.ok) {
            let msg = resp.statusText;
            try {
                const e = (await resp.json());
                msg = e.error?.message ?? msg;
            }
            catch { }
            throw new Error(`Write ${filePath}: ${msg}`);
        }
    }
    async deleteFilePath(filePath) {
        await this.gFetch(`${this.encode(filePath)}`, { method: "DELETE" });
    }
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
    get appConfigPath() {
        return `${this._rootFolder}/${APP_CONFIG_FILE}`;
    }
    async readAppConfig() {
        return this.readJson(this.appConfigPath);
    }
    async writeAppConfig(cfg) {
        await this.writeJson(this.appConfigPath, cfg);
    }
    async loadApiKey() {
        try {
            const cfg = await this.readAppConfig();
            if (cfg?.apiKey)
                setStoredKey(cfg.apiKey);
        }
        catch { }
    }
    async initRootStructure() {
        const r = this._rootFolder;
        await this.ensureFolder(r);
        await this.ensureFolder(`${r}/${FOLDER_SKILLS}`);
        await this.ensureFolder(`${r}/${FOLDER_CASES}`);
        await this.ensureFolder(`${r}/${FOLDER_LIBRARY}`);
        for (const d of [
            "Commercial",
            "Fiscal",
            "Social",
            "Civil",
            "Penal",
            "Immobilier",
            "International",
        ]) {
            await this.ensureFolder(`${r}/${FOLDER_LIBRARY}/${d}`);
        }
        const existing = await this.readAppConfig();
        if (!existing)
            await this.writeAppConfig({});
    }
}
// ─── Cases — dossiers scenario ────────────────────────────────────────────────
class Scenario extends Folders {
    config = oneDrive.config.config;
    activeMode = "analyse";
    mainFolder = FOLDER_CASES;
    skills = [];
    // ─── Path helpers ─────────────────────────────────────────────────────────────
    userName = () => oneDrive.userName;
    // ─── Skills ───────────────────────────────────────────────────────────────────
    /** Called by main.ts after fetchSkills() to load skills into both modules */
    async fetchSkills() {
        const path = `${this._rootFolder}/${FOLDER_SKILLS}`;
        let items;
        try {
            items = await this.listFolder(path);
            items = items.filter((item) => !item.folder);
        }
        catch {
            return [];
        }
        const skills = [];
        for (const item of items) {
            const ext = item.name.split(".").pop()?.toLowerCase() ?? "";
            if (!["md", "txt"].includes(ext))
                continue;
            try {
                const buf = await this.readFilePath(`${path}/${item.name}`);
                skills.push({
                    name: item.name,
                    content: new TextDecoder().decode(buf),
                });
            }
            catch { }
        }
        this.skills = skills;
        this.updateSkillIndicator();
        return skills;
    }
    // ─── OneDrive connection ──────────────────────────────────────────────────────
    updateODStatus() {
        const statusEl = byID("onedrive-status");
        if (!statusEl)
            return;
        if (oneDrive.userName) {
            statusEl.textContent = `☁ ${oneDrive.userName}`;
            statusEl.className = "od-status od-status--connected";
        }
        else {
            statusEl.textContent = oneDrive.isConfigured
                ? "☁ Non connecté"
                : "☁ Non configuré";
            statusEl.className = "od-status od-status--disconnected";
        }
    }
    async connectOneDrive(loadSubfolders) {
        try {
            await oneDrive.signIn();
            oneDrive.userName = oneDrive.getSignedInUser();
            this.updateODStatus();
            await this.initRootStructure();
            await this.loadApiKey();
            if (loadSubfolders)
                await loadSubfolders();
            await this.fetchSkills();
            toast(`Connecté : ${oneDrive.userName}`, "success");
        }
        catch (err) {
            toast("Erreur connexion OneDrive : " + err.message, "error");
        }
    }
    // ─── Skill indicator ──────────────────────────────────────────────────────────
    updateSkillIndicator() {
        const badge = byID("skills-badge");
        const n = this.skills.length;
        if (!badge)
            return;
        badge.textContent = n > 0 ? `${n} skill${n > 1 ? "s" : ""}` : "";
        toggle(badge, n > 0);
    }
    openSettingsModal() {
        byID("settings-overlay")?.remove();
        const overlay = el("div", {
            className: "modal-overlay",
            id: "settings-overlay",
        });
        const dialog = el("div", {
            className: "modal-dialog modal-dialog--settings",
        });
        dialog.innerHTML = `
      <h2 class="modal-title">⚙ Paramètres</h2>
      <h3 class="settings-section-title">Clé API Claude (Anthropic)</h3>
      <p class="settings-hint">Stockée dans <code>_config.json</code> sur OneDrive. Transmise uniquement à api.anthropic.com.</p>
      <input class="form-input" id="s-api" type="password" placeholder="sk-ant-api03-…" value="${getStoredKey()}" autocomplete="off"/>
      <div style="display:flex;gap:8px;margin-top:6px">
        <button class="btn btn--primary btn--sm" id="s-api-save">Enregistrer</button>
        <button class="btn btn--secondary btn--sm" id="s-api-clear">Effacer</button>
      </div>
      <hr style="margin:20px 0">
      <h3 class="settings-section-title">☁ Microsoft OneDrive (Graph API)</h3>
      <p class="settings-hint"><strong>portal.azure.com</strong> → App registrations → New registration<br>
      Type : SPA — Redirect URI : <code>${window.location.origin}</code><br>
      Permissions : <code>Files.ReadWrite</code> + <code>User.Read</code></p>
      <label class="form-label">Application (Client) ID <span class="required">*</span></label>
      <input class="form-input" id="s-client" type="text" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" value="${this.config?.clientId ?? ""}"/>
      <label class="form-label">Tenant ID</label>
      <input class="form-input" id="s-tenant" type="text" placeholder="common" value="${this.config?.tenantId ?? "common"}"/>
      <label class="form-label">Dossier racine OneDrive</label>
      <input class="form-input" id="s-root" type="text" placeholder="LexAssistant" value="${this.config?.rootFolder ?? "LexAssistant"}"/>
      <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap">
      <button class="btn btn--primary btn--sm" id="s-od-save">Enregistrer</button>
        <button class="btn btn--secondary btn--sm" id="s-od-init">Initialiser structure OneDrive</button>
        <button class="btn btn--secondary btn--sm" id="s-od-signout">Déconnecter</button>
      </div>
      ${oneDrive.userName ? `<p class="od-connected-label">✓ Connecté : ${oneDrive.userName}</p>` : ""}
      <div class="modal-btns" style="margin-top:24px">
        <button class="btn btn--secondary" id="s-close">Fermer</button>
      </div>`;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        overlay.addEventListener("click", (e) => {
            if (e.target === overlay)
                overlay.remove();
        });
        qs("#s-close", dialog).onclick = () => overlay.remove();
        qs("#s-api-save", dialog).onclick = async () => {
            const v = qs("#s-api", dialog).value.trim();
            if (!v) {
                toast("Clé vide.", "error");
                return;
            }
            setStoredKey(v);
            if (oneDrive.userName) {
                try {
                    const existing = (await this.readAppConfig()) ?? {};
                    await this.writeAppConfig({ ...existing, apiKey: v });
                    toast("Clé API enregistrée sur OneDrive.", "success");
                }
                catch {
                    toast("Clé mémorisée mais non sauvegardée sur OneDrive (erreur).", "error");
                }
            }
            else {
                toast("Clé mémorisée pour cette session. Connectez OneDrive pour la sauvegarder définitivement.", "info");
            }
        };
        qs("#s-api-clear", dialog).onclick = async () => {
            clearStoredKey();
            qs("#s-api", dialog).value = "";
            if (oneDrive.userName) {
                try {
                    const existing = (await this.readAppConfig()) ?? {};
                    await this.writeAppConfig({ ...existing, apiKey: "" });
                }
                catch { }
            }
            toast("Clé effacée.", "info");
        };
        qs("#s-od-save", dialog).onclick = () => {
            const clientId = qs("#s-client", dialog).value.trim();
            const tenantId = qs("#s-tenant", dialog).value.trim() || "common";
            const rootFolder = qs("#s-root", dialog).value.trim() || "LexAssistant";
            if (!clientId) {
                toast("Client ID requis.", "error");
                return;
            }
            oneDrive.setConfig({ clientId, tenantId, rootFolder });
            toast("Configuration OneDrive enregistrée.", "success");
        };
        qs("#s-od-init", dialog).onclick = async () => {
            if (!oneDrive.isConfigured) {
                toast("Sauvegardez la configuration d'abord.", "error");
                return;
            }
            try {
                if (!oneDrive.userName) {
                    await oneDrive.signIn();
                    oneDrive.userName = oneDrive.getSignedInUser();
                    this.updateODStatus();
                }
                await this.initRootStructure();
                toast("Structure initialisée avec succès.", "success");
            }
            catch (err) {
                toast("Erreur : " + err.message, "error");
            }
        };
        qs("#s-od-signout", dialog).onclick = async () => {
            await oneDrive.signOut();
            oneDrive.userName = null;
            this.updateODStatus();
            toast("Déconnecté.", "info");
            overlay.remove();
        };
    }
    onClick(btn, action) {
        if (!btn)
            return;
        btn.onclick = () => action();
    }
    autoResize(ta) {
        ta.style.height = "auto";
        ta.style.height = Math.min(ta.scrollHeight, 160) + "px";
    }
    showEmptyState() {
        const area = byID("chat-area");
        if (!area)
            return;
        area.innerHTML = "";
        area.appendChild(el("div", { className: "empty-state" }, el("div", { className: "empty-icon", textContent: "⚖️" }), el("h2", { textContent: "Bienvenue dans Lex Assistant" }), el("p", {
            textContent: "Connectez OneDrive et créez votre premier dossier.",
        }), (() => {
            const b = el("button", {
                className: "btn btn--primary",
                textContent: "+ Nouveau dossier",
            });
            b.onclick = () => this.openCaseFormModal(null);
            return b;
        })()));
    }
    showNotConnected() {
        const area = byID("chat-area");
        if (!area)
            return;
        area.innerHTML = "";
        area.appendChild(el("div", { className: "empty-state" }, el("div", { className: "empty-icon", textContent: "☁" }), el("h2", { textContent: "OneDrive non connecté" }), el("p", {
            textContent: "Configurez votre App Registration Azure et connectez-vous.",
        }), (() => {
            const b = el("button", {
                className: "btn btn--primary",
                textContent: "☁ Configurer OneDrive",
            });
            b.onclick = () => this.openSettingsModal();
            return b;
        })()));
    }
    sanitiseFolder(name) {
        return name
            .replace(/[/\\:*?"<>|]/g, "_")
            .replace(/\s+/g, "_")
            .slice(0, 60);
    }
}
export class Cases extends Scenario {
    saving = false;
    _docFilter = null;
    _foldersMeta = [];
    _activeCase = null;
    _caseNotes = [];
    _caseMessages = [];
    // ─── Path helpers ─────────────────────────────────────────────────────────────
    casePath = (f) => `${this._rootFolder}/${FOLDER_CASES}/${f}`;
    metaPath = (f) => `${this.casePath(f)}/_meta.json`;
    notesPath = (f) => `${this.casePath(f)}/_notes.json`;
    convPath = (f) => `${this.casePath(f)}/_conversation.json`;
    // ─── CRUD ─────────────────────────────────────────────────────────────────────
    async readCaseMeta(folderName) {
        return this.readJson(this.metaPath(folderName));
    }
    async writeCaseMeta(subFolderName, meta) {
        await this.ensureFolder(this.casePath(subFolderName));
        await this.writeJson(this.metaPath(subFolderName), meta);
    }
    async readNotes(folderName) {
        const f = await this.readJson(this.notesPath(folderName));
        return f?.notes ?? [];
    }
    async writeNotes(folderName, notes) {
        await this.writeJson(this.notesPath(folderName), {
            notes,
        });
    }
    async readConversation(folderName) {
        const f = await this.readJson(this.convPath(folderName));
        return f?.messages ?? [];
    }
    async writeConversation(folderName, messages) {
        await this.writeJson(this.convPath(folderName), {
            messages,
        });
    }
    async listSubFolders() {
        try {
            const items = await this.listFolder(`${this._rootFolder}/${this.mainFolder}`);
            return items.filter((i) => i.folder).map((i) => i.name);
        }
        catch {
            return [];
        }
    }
    async listCaseFiles(folderName) {
        const items = await this.listFolder(this.casePath(folderName));
        return items.filter((i) => i.file && !i.name.startsWith("_"));
    }
    async writeCaseFile(folderName, fileName, data, mimeType) {
        await this.ensureFolder(this.casePath(folderName));
        await this.writeFileLarge(`${this.casePath(folderName)}/${fileName}`, data, mimeType);
    }
    async readCaseFile(folderName, fileName) {
        return this.readFilePath(`${this.casePath(folderName)}/${fileName}`);
    }
    // ─── Persist helpers ──────────────────────────────────────────────────────────
    async saveMeta() {
        if (!this._activeCase || this.saving)
            return;
        this.saving = true;
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
            this.saving = false;
        }
    }
    async saveNotes() {
        if (!this._activeCase)
            return;
        await this.writeNotes(this._activeCase.folderName, this._caseNotes);
    }
    async saveConversation() {
        if (!this._activeCase)
            return;
        await this.writeConversation(this._activeCase.folderName, this._caseMessages);
    }
    // ─── Skills ───────────────────────────────────────────────────────────────────
    /** Called by main.ts after fetchSkills() to load skills into both modules */
    async fetchSkills() {
        const path = `${this._rootFolder}/${FOLDER_SKILLS}`;
        let items;
        try {
            items = await this.listFolder(path);
            items = items.filter((item) => !item.folder);
        }
        catch {
            return [];
        }
        const skills = [];
        for (const item of items) {
            const ext = item.name.split(".").pop()?.toLowerCase() ?? "";
            if (!["md", "txt"].includes(ext))
                continue;
            try {
                const buf = await this.readFilePath(`${path}/${item.name}`);
                skills.push({
                    name: item.name,
                    content: new TextDecoder().decode(buf),
                });
            }
            catch { }
        }
        this.skills = skills;
        this.updateSkillIndicator();
        return skills;
    }
    // ─── Load all cases ───────────────────────────────────────────────────────────
    async loadAllSubFolders() {
        await this.loadApiKey();
        const subFolders = await this.listSubFolders();
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
                documents: meta.documents ?? [],
            });
        }));
        this._foldersMeta.sort((a, b) => b.updatedAt - a.updatedAt);
        this.renderCaseList();
        if (this._foldersMeta.length > 0)
            await this.selectCase(this._foldersMeta[0].folderName);
        else
            this.showEmptyState();
    }
    // ─── Select case ──────────────────────────────────────────────────────────────
    async selectCase(folderName) {
        const c = this._foldersMeta.find((x) => x.folderName === folderName);
        if (!c)
            return;
        this._activeCase = c;
        this._caseNotes = await this.readNotes(folderName);
        this._caseMessages = await this.readConversation(folderName);
        qsa(".case-item").forEach((el) => el.classList.toggle("active", el.dataset.folder === folderName));
        const nameEl = byID("topbar-case-name");
        const domainEl = byID("topbar-case-domain");
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
            const items = await this.listCaseFiles(this._activeCase.folderName);
            let added = 0;
            for (const item of items) {
                if (!item.file || !isSupported(item.name))
                    continue;
                if (this._activeCase.documents.some((d) => d.name === item.name))
                    continue;
                this._activeCase.documents.push({
                    name: item.name,
                    kind: guessKind(item.name),
                    mimeType: item.file.mimeType || "application/octet-stream",
                    sizeBytes: item.size ?? 0,
                    addedAt: Date.now(),
                });
                added++;
            }
            if (added > 0) {
                await this.saveMeta();
                this.renderDocList();
                this.updateDocCount();
            }
            toast(`${added} nouveau(x) document(s) indexé(s) depuis OneDrive.`, "success");
        }
        catch (err) {
            toast("Erreur sync : " + err.message, "error");
        }
    }
    // ─── Skill indicator ──────────────────────────────────────────────────────────
    updateSkillIndicator() {
        const badge = byID("skills-badge");
        const n = this.skills.length;
        if (!badge)
            return;
        badge.textContent = n > 0 ? `${n} skill${n > 1 ? "s" : ""}` : "";
        toggle(badge, n > 0);
    }
    // ─── Render: case list ────────────────────────────────────────────────────────
    renderCaseList() {
        const list = byID("case-list");
        if (!list)
            return;
        list.innerHTML = "";
        const labels = {
            active: "En cours",
            closed: "Clôturé",
            suspended: "Suspendu",
        };
        for (const c of this._foldersMeta) {
            const item = el("div", {
                className: "case-item" +
                    (c.folderName === this._activeCase?.folderName ? " active" : ""),
            });
            item.dataset.folder = c.folderName;
            item.append(el("span", { className: "case-item__name", textContent: c.name }), el("span", { className: "case-item__domain", textContent: c.domain }), el("span", {
                className: `case-item__status case-item__status--${c.status}`,
                textContent: labels[c.status],
            }));
            item.onclick = () => this.selectCase(c.folderName);
            item.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                this.openCaseContextMenu(c, e.clientX, e.clientY);
            });
            list.appendChild(item);
        }
    }
    // ─── Render: doc list ─────────────────────────────────────────────────────────
    renderDocList() {
        const list = byID("doc-list");
        if (!list || !this._activeCase)
            return;
        list.innerHTML = "";
        const docs = !this._docFilter
            ? this._activeCase.documents
            : this._activeCase.documents.filter((d) => d.kind === this._docFilter);
        if (!docs.length) {
            list.appendChild(el("div", {
                className: "doc-empty",
                textContent: "Aucun document. Uploadez ou synchronisez OneDrive.",
            }));
            return;
        }
        for (const doc of [...docs].sort((a, b) => b.addedAt - a.addedAt)) {
            const item = el("div", { className: "doc-item" });
            const info = el("div", { className: "doc-info" });
            info.append(el("div", { className: "doc-name", textContent: doc.name }), el("div", {
                className: "doc-meta",
                textContent: `${formatDate(doc.addedAt)} · ${mimeLabel(doc.mimeType)} · ${formatSize(doc.sizeBytes)}`,
            }));
            const del = el("button", {
                className: "doc-delete",
                textContent: "×",
                title: "Retirer du dossier",
            });
            del.onclick = async (e) => {
                e.stopPropagation();
                const ok = await confirm(`Retirer "${doc.name}" ?\n(Fichier OneDrive conservé, seul l'index local est supprimé.)`);
                if (!ok)
                    return;
                this._activeCase.documents = this._activeCase.documents.filter((d) => d.name !== doc.name);
                await this.saveMeta();
                this.renderDocList();
                this.updateDocCount();
                toast("Document retiré de l'index.", "info");
            };
            item.append(el("span", {
                className: "doc-icon",
                textContent: mimeIcon(doc.mimeType),
            }), info, el("span", {
                className: `doc-kind doc-kind--${doc.kind}`,
                textContent: kindLabel(doc.kind),
            }), del);
            list.appendChild(item);
        }
    }
    updateDocCount() {
        const countEl = byID("doc-count");
        if (countEl && this._activeCase) {
            countEl.textContent = `${this._activeCase.documents.length} pièce${this._activeCase.documents.length !== 1 ? "s" : ""}`;
        }
    }
    // ─── Render: note bar ─────────────────────────────────────────────────────────
    renderNoteBar() {
        const bar = byID("note-bar");
        if (!bar)
            return;
        bar.innerHTML = "";
        toggle(bar, this._caseNotes.length > 0);
        if (!this._caseNotes.length)
            return;
        const n = this._caseNotes.length;
        bar.append(el("span", { className: "note-badge", textContent: String(n) }), el("span", {
            className: "note-bar__text",
            textContent: `note${n > 1 ? "s" : ""} active${n > 1 ? "s" : ""} · ` +
                this._caseNotes.map((x) => x.content.slice(0, 40) + "…").join(" — "),
        }), (() => {
            const b = el("button", {
                className: "note-bar__manage",
                textContent: "Gérer",
            });
            b.onclick = () => this.openNotesModal();
            return b;
        })());
    }
    // ─── Render: chat ─────────────────────────────────────────────────────────────
    renderChat() {
        const area = byID("chat-area");
        if (!area)
            return;
        area.innerHTML = "";
        if (!this._caseMessages.length) {
            const c = this._activeCase;
            if (!c)
                return;
            area.appendChild(el("div", { className: "msg msg--assistant" }, el("div", { className: "msg__label", textContent: "Lex Assistant" }), el("div", { className: "msg__bubble" }, el("p", {
                innerHTML: `Dossier <strong>${c.name}</strong>. ${c.documents.length} pièce(s), ${this._caseNotes.length} note(s)${this.skills.length ? `, ${this.skills.length} skill(s)` : ""}.`,
            }), el("p", { textContent: "Que souhaitez-vous faire ?" }))));
            return;
        }
        for (const msg of this._caseMessages)
            area.appendChild(this.buildMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    buildMsgEl(msg) {
        const wrap = el("div", { className: `msg msg--${msg.role}` });
        const bubble = el("div", { className: "msg__bubble" });
        bubble.innerHTML = renderMarkdown(msg.content);
        wrap.append(el("div", {
            className: "msg__label",
            textContent: msg.role === "user" ? "Vous" : "Lex Assistant",
        }), bubble);
        if (msg.role === "assistant") {
            const actions = el("div", { className: "msg__actions" });
            const copy = el("button", {
                className: "msg-action-btn",
                textContent: "Copier",
            });
            copy.onclick = () => {
                navigator.clipboard.writeText(msg.content);
                toast("Copié.", "info", 1500);
            };
            actions.appendChild(copy);
            if (msg.mode === "redaction" || msg.generatedDocName) {
                const dl = el("button", {
                    className: "msg-action-btn msg-action-btn--primary",
                    textContent: "⬇ Télécharger .docx",
                });
                dl.onclick = async () => {
                    try {
                        const blob = await generateDocx({
                            title: msg.generatedDocName ?? "Document",
                            content: msg.content,
                            caseRef: this._activeCase?.name ?? "",
                        });
                        downloadBlob(blob, (msg.generatedDocName ?? "document").replace(/[^a-z0-9_\- ]/gi, "_") + ".docx");
                    }
                    catch (err) {
                        toast("Erreur DOCX : " + err.message, "error");
                    }
                };
                actions.appendChild(dl);
                const odSave = el("button", {
                    className: "msg-action-btn",
                    textContent: "☁ Sauver sur OneDrive",
                });
                odSave.onclick = async () => {
                    if (!this._activeCase)
                        return;
                    try {
                        const blob = await generateDocx({
                            title: msg.generatedDocName ?? "Document",
                            content: msg.content,
                            caseRef: this._activeCase.name,
                        });
                        const ab = await blob.arrayBuffer();
                        const fname = (msg.generatedDocName ?? "document").replace(/[^a-z0-9_\- ]/gi, "_") + ".docx";
                        const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        await this.writeCaseFile(this._activeCase.folderName, fname, ab, mime);
                        if (!this._activeCase.documents.some((d) => d.name === fname)) {
                            this._activeCase.documents.push({
                                name: fname,
                                kind: "redige",
                                mimeType: mime,
                                sizeBytes: ab.byteLength,
                                addedAt: Date.now(),
                            });
                            await this.saveMeta();
                            this.renderDocList();
                            this.updateDocCount();
                        }
                        toast(`"${fname}" sauvegardé.`, "success");
                    }
                    catch (err) {
                        toast("Erreur OneDrive : " + err.message, "error");
                    }
                };
                actions.appendChild(odSave);
            }
            wrap.appendChild(actions);
        }
        return wrap;
    }
    appendMsg(msg) {
        const area = byID("chat-area");
        if (!area)
            return;
        area.querySelector(".empty-state")?.remove();
        area.appendChild(this.buildMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    appendTyping() {
        const area = byID("chat-area");
        const typing = el("div", { className: "msg msg--assistant", id: "typing" });
        const bubble = el("div", { className: "msg__bubble" });
        bubble.append(spinnerEl(), el("span", { textContent: " Analyse en cours…" }));
        typing.append(el("div", { className: "msg__label", textContent: "Lex Assistant" }), bubble);
        area.appendChild(typing);
        area.scrollTop = area.scrollHeight;
        return typing;
    }
    // ─── Send message ─────────────────────────────────────────────────────────────
    async sendMessage() {
        const ta = byID("user-input");
        if (!ta)
            return;
        const text = ta.value.trim();
        if (!text || !this._activeCase)
            return;
        if (!getStoredKey()) {
            this.openSettingsModal();
            toast("Configurez votre clé API Claude.", "error");
            return;
        }
        if (!oneDrive.userName) {
            await this.connectOneDrive();
            return;
        }
        ta.value = "";
        this.autoResize(ta);
        const userMsg = {
            id: uid(),
            role: "user",
            content: text,
            timestamp: Date.now(),
            mode: this.activeMode,
        };
        this._caseMessages.push(userMsg);
        this.appendMsg(userMsg);
        if (this.activeMode === "note") {
            const note = {
                id: uid(),
                content: text,
                createdAt: Date.now(),
                updatedAt: Date.now(),
            };
            this._caseNotes.push(note);
            await this.saveNotes();
            this.renderNoteBar();
        }
        const typing = this.appendTyping();
        const sendBtn = byID("send-btn");
        if (sendBtn)
            sendBtn.disabled = true;
        try {
            const response = await callClaudeCase({
                caseName: this._activeCase.name,
                caseDomain: this._activeCase.domain,
                notes: this._caseNotes,
                docs: this._activeCase.documents,
                skills: this.skills,
                mode: this.activeMode,
                userMessage: text,
                readFile: (fileName) => this.readCaseFile(this._activeCase.folderName, fileName),
            });
            typing.remove();
            let docName;
            if (this.activeMode === "redaction") {
                const first = response
                    .split("\n")[0]
                    .replace(/^#+\s*/, "")
                    .trim();
                docName =
                    first.length > 0 && first.length < 100 ? first : "Document rédigé";
            }
            const asst = {
                id: uid(),
                role: "assistant",
                content: response,
                timestamp: Date.now(),
                mode: this.activeMode,
                generatedDocName: docName,
            };
            this._caseMessages.push(asst);
            this.appendMsg(asst);
            await this.saveConversation();
        }
        catch (err) {
            typing.remove();
            toast(err.message, "error", 6000);
            this._caseMessages.pop();
        }
        finally {
            if (sendBtn)
                sendBtn.disabled = false;
            ta.focus();
        }
    }
    // ─── Modals ───────────────────────────────────────────────────────────────────
    openCaseFormModal(existing) {
        if (!oneDrive.userName) {
            this.openSettingsModal();
            toast("Connectez OneDrive d'abord.", "error");
            return;
        }
        const isEdit = !!existing;
        const overlay = el("div", { className: "modal-overlay" });
        const dialog = el("div", { className: "modal-dialog modal-dialog--form" });
        dialog.innerHTML = `
      <h2 class="modal-title">${isEdit ? "Modifier le dossier" : "Nouveau dossier"}</h2>
      <label class="form-label">Intitulé <span class="required">*</span></label>
      <input class="form-input" id="f-name" type="text" value="${existing?.name ?? ""}"/>
      <label class="form-label">Domaine juridique</label>
      <input class="form-input" id="f-domain" type="text" placeholder="Droit commercial…" value="${existing?.domain ?? ""}"/>
      <label class="form-label">Nom du dossier OneDrive <span style="font-weight:400;color:var(--c-gray-400)">(auto si vide)</span></label>
      <input class="form-input" id="f-folder" type="text" value="${existing?.folderName ?? ""}"/>
      <label class="form-label">Statut</label>
      <select class="form-select" id="f-status">
        <option value="active" ${!existing || existing.status === "active" ? "selected" : ""}>En cours</option>
        <option value="suspended" ${existing?.status === "suspended" ? "selected" : ""}>Suspendu</option>
        <option value="closed" ${existing?.status === "closed" ? "selected" : ""}>Clôturé</option>
      </select>
      <div class="modal-btns">
        <button class="btn btn--secondary" id="f-cancel">Annuler</button>
        <button class="btn btn--primary" id="f-save">${isEdit ? "Enregistrer" : "Créer"}</button>
      </div>`;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        const nameInput = qs("#f-name", dialog);
        const folderInput = qs("#f-folder", dialog);
        nameInput.addEventListener("input", () => {
            if (!isEdit && !folderInput.value)
                folderInput.placeholder =
                    this.sanitiseFolder(nameInput.value) || "DOSSIER_NOM";
        });
        qs("#f-cancel", dialog).onclick = () => overlay.remove();
        qs("#f-save", dialog).onclick = async () => {
            const name = nameInput.value.trim();
            const domain = qs("#f-domain", dialog).value.trim() ||
                "Droit général";
            const raw = folderInput.value.trim();
            const folder = raw ? this.sanitiseFolder(raw) : this.sanitiseFolder(name);
            const status = qs("#f-status", dialog)
                .value;
            if (!name || !folder) {
                toast(!name ? "Nom obligatoire." : "Dossier invalide.", "error");
                return;
            }
            const btn = qs("#f-save", dialog);
            btn.disabled = true;
            btn.textContent = "Création…";
            const now = Date.now();
            const meta = {
                name,
                folderName: folder,
                domain,
                status,
                createdAt: existing?.createdAt ?? now,
                updatedAt: now,
                documents: existing?.documents ?? [],
            };
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
                toast(isEdit ? "Dossier modifié." : "Dossier créé.", "success");
            }
            catch (err) {
                toast("Erreur : " + err.message, "error");
                btn.disabled = false;
                btn.textContent = isEdit ? "Enregistrer" : "Créer";
            }
        };
        nameInput.focus();
    }
    openCaseContextMenu(c, x, y) {
        byID("context-menu")?.remove();
        const menu = el("div", { className: "context-menu", id: "context-menu" });
        menu.style.left = `${x}px`;
        menu.style.top = `${y}px`;
        const items = [
            { label: "Modifier", action: () => this.openCaseFormModal(c) },
            {
                label: "☁ Sync OneDrive",
                action: () => this.refreshCaseFromOneDrive(),
            },
            {
                label: "Effacer conversation",
                action: () => this.clearConversation(c.folderName),
            },
            {
                label: "Supprimer le dossier",
                action: () => this.deleteCaseIndex(c.folderName),
                danger: true,
            },
        ];
        for (const item of items) {
            const btn = el("button", {
                className: "context-menu__item" +
                    (item.danger ? " context-menu__item--danger" : ""),
                textContent: item.label,
            });
            btn.onclick = () => {
                menu.remove();
                item.action();
            };
            menu.appendChild(btn);
        }
        document.body.appendChild(menu);
        const dismiss = (e) => {
            if (!menu.contains(e.target)) {
                menu.remove();
                document.removeEventListener("click", dismiss);
            }
        };
        setTimeout(() => document.addEventListener("click", dismiss), 0);
    }
    async clearConversation(folderName) {
        const ok = await confirm("Effacer tout l'historique de conversation de ce dossier ?");
        if (!ok)
            return;
        this._caseMessages = [];
        await this.writeConversation(folderName, []);
        if (this._activeCase?.folderName === folderName)
            this.renderChat();
        toast("Conversation effacée.", "info");
    }
    async deleteCaseIndex(folderName) {
        const ok = await confirm("Supprimer ce dossier ? Les fichiers OneDrive sont conservés, seuls les fichiers Lex Assistant (_meta, _notes, _conversation) sont supprimés.");
        if (!ok)
            return;
        // Delete only Lex Assistant JSON files — preserve user documents
        for (const f of ["_meta.json", "_notes.json", "_conversation.json"]) {
            await this.deleteFilePath(`${this.casePath(folderName)}/${f}`).catch(() => { });
        }
        this._foldersMeta = this._foldersMeta.filter((c) => c.folderName !== folderName);
        this.renderCaseList();
        if (this._activeCase?.folderName === folderName) {
            this._activeCase = null;
            if (this._foldersMeta.length > 0)
                await this.selectCase(this._foldersMeta[0].folderName);
            else
                this.showEmptyState();
        }
        toast("Dossier supprimé.", "success");
    }
    openNotesModal() {
        const overlay = el("div", { className: "modal-overlay" });
        const dialog = el("div", { className: "modal-dialog modal-dialog--notes" });
        const refresh = () => {
            const list = qs("#notes-list", dialog);
            list.innerHTML = "";
            if (!this._caseNotes.length) {
                list.appendChild(el("p", {
                    className: "note-empty",
                    textContent: 'Aucune note. Ajoutez-en une ci-dessous ou utilisez le mode "Note permanente".',
                }));
                return;
            }
            for (const note of [...this._caseNotes].sort((a, b) => b.createdAt - a.createdAt)) {
                const row = el("div", { className: "note-row" });
                const del = el("button", {
                    className: "btn btn--sm btn--danger",
                    textContent: "Supprimer",
                });
                del.onclick = async () => {
                    this._caseNotes = this._caseNotes.filter((n) => n.id !== note.id);
                    await this.saveNotes();
                    this.renderNoteBar();
                    refresh();
                };
                row.append(el("p", { className: "note-content", textContent: note.content }), el("span", {
                    className: "note-meta",
                    textContent: formatDateTime(note.createdAt),
                }), del);
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
        <button class="btn btn--primary" id="n-add">Ajouter</button>
      </div>`;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        refresh();
        qs("#n-close", dialog).onclick = () => overlay.remove();
        qs("#n-add", dialog).onclick = async () => {
            const val = qs("#new-note", dialog).value.trim();
            if (!val)
                return;
            const note = {
                id: uid(),
                content: val,
                createdAt: Date.now(),
                updatedAt: Date.now(),
            };
            this._caseNotes.push(note);
            await this.saveNotes();
            this.renderNoteBar();
            refresh();
            qs("#new-note", dialog).value = "";
            toast("Note sauvegardée dans _notes.json.", "success");
        };
    }
    openSettingsModal() {
        byID("settings-overlay")?.remove();
        const overlay = el("div", {
            className: "modal-overlay",
            id: "settings-overlay",
        });
        const dialog = el("div", {
            className: "modal-dialog modal-dialog--settings",
        });
        dialog.innerHTML = `
      <h2 class="modal-title">⚙ Paramètres</h2>
      <h3 class="settings-section-title">Clé API Claude (Anthropic)</h3>
      <p class="settings-hint">Stockée dans <code>_config.json</code> sur OneDrive. Transmise uniquement à api.anthropic.com.</p>
      <input class="form-input" id="s-api" type="password" placeholder="sk-ant-api03-…" value="${getStoredKey()}" autocomplete="off"/>
      <div style="display:flex;gap:8px;margin-top:6px">
        <button class="btn btn--primary btn--sm" id="s-api-save">Enregistrer</button>
        <button class="btn btn--secondary btn--sm" id="s-api-clear">Effacer</button>
      </div>
      <hr style="margin:20px 0">
      <h3 class="settings-section-title">☁ Microsoft OneDrive (Graph API)</h3>
      <p class="settings-hint"><strong>portal.azure.com</strong> → App registrations → New registration<br>
      Type : SPA — Redirect URI : <code>${window.location.origin}</code><br>
      Permissions : <code>Files.ReadWrite</code> + <code>User.Read</code></p>
      <label class="form-label">Application (Client) ID <span class="required">*</span></label>
      <input class="form-input" id="s-client" type="text" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" value="${this.config?.clientId ?? ""}"/>
      <label class="form-label">Tenant ID</label>
      <input class="form-input" id="s-tenant" type="text" placeholder="common" value="${this.config?.tenantId ?? "common"}"/>
      <label class="form-label">Dossier racine OneDrive</label>
      <input class="form-input" id="s-root" type="text" placeholder="LexAssistant" value="${this.config?.rootFolder ?? "LexAssistant"}"/>
      <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap">
        <button class="btn btn--primary btn--sm" id="s-od-save">Enregistrer</button>
        <button class="btn btn--secondary btn--sm" id="s-od-init">Initialiser structure OneDrive</button>
        <button class="btn btn--secondary btn--sm" id="s-od-signout">Déconnecter</button>
      </div>
      ${oneDrive.userName ? `<p class="od-connected-label">✓ Connecté : ${oneDrive.userName}</p>` : ""}
      <div class="modal-btns" style="margin-top:24px">
        <button class="btn btn--secondary" id="s-close">Fermer</button>
      </div>`;
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        overlay.addEventListener("click", (e) => {
            if (e.target === overlay)
                overlay.remove();
        });
        qs("#s-close", dialog).onclick = () => overlay.remove();
        qs("#s-api-save", dialog).onclick = async () => {
            const v = qs("#s-api", dialog).value.trim();
            if (!v) {
                toast("Clé vide.", "error");
                return;
            }
            setStoredKey(v);
            if (oneDrive.userName) {
                try {
                    const existing = (await this.readAppConfig()) ?? {};
                    await this.writeAppConfig({ ...existing, apiKey: v });
                    toast("Clé API enregistrée sur OneDrive.", "success");
                }
                catch {
                    toast("Clé mémorisée mais non sauvegardée sur OneDrive (erreur).", "error");
                }
            }
            else {
                toast("Clé mémorisée pour cette session. Connectez OneDrive pour la sauvegarder définitivement.", "info");
            }
        };
        qs("#s-api-clear", dialog).onclick = async () => {
            clearStoredKey();
            qs("#s-api", dialog).value = "";
            if (oneDrive.userName) {
                try {
                    const existing = (await this.readAppConfig()) ?? {};
                    await this.writeAppConfig({ ...existing, apiKey: "" });
                }
                catch { }
            }
            toast("Clé effacée.", "info");
        };
        qs("#s-od-save", dialog).onclick = () => {
            const clientId = qs("#s-client", dialog).value.trim();
            const tenantId = qs("#s-tenant", dialog).value.trim() || "common";
            const rootFolder = qs("#s-root", dialog).value.trim() || "LexAssistant";
            if (!clientId) {
                toast("Client ID requis.", "error");
                return;
            }
            oneDrive.setConfig({ clientId, tenantId, rootFolder });
            toast("Configuration OneDrive enregistrée.", "success");
        };
        qs("#s-od-init", dialog).onclick = async () => {
            if (!oneDrive.isConfigured) {
                toast("Sauvegardez la configuration d'abord.", "error");
                return;
            }
            try {
                if (!oneDrive.userName) {
                    await oneDrive.signIn();
                    oneDrive.userName = oneDrive.getSignedInUser();
                    this.updateODStatus();
                }
                await this.initRootStructure();
                toast("Structure initialisée avec succès.", "success");
            }
            catch (err) {
                toast("Erreur : " + err.message, "error");
            }
        };
        qs("#s-od-signout", dialog).onclick = async () => {
            await oneDrive.signOut();
            oneDrive.userName = null;
            this.updateODStatus();
            toast("Déconnecté.", "info");
            overlay.remove();
        };
    }
    // ─── UI setup ─────────────────────────────────────────────────────────────────
    setupBarsBtns() {
        const btn = (id) => byID(id), onClick = this.onClick;
        [btn("btn-new-item-top"), btn("btn-new-item-sidebar")].forEach((btn) => onClick(btn, () => this.openCaseFormModal(null)));
        onClick(btn("btn-settings"), () => this.openSettingsModal());
        onClick(btn("btn-onedrive"), async () => {
            if (!oneDrive.userName)
                await this.connectOneDrive();
            else
                await this.refreshCaseFromOneDrive();
        });
        onClick(btn("btn-notes-open"), () => this.openNotesModal());
        onClick(btn("btn-od-sync"), () => this.refreshCaseFromOneDrive());
        onClick(btn("btn-upload"), () => btn("file-input")?.click());
        const tabs = qsa(".doc-filter-tab");
        tabs.forEach((tab) => onClick(tab, () => {
            setActive(tabs, tab, "active");
            this._docFilter = (tab.dataset.filter ??
                null);
            this.renderDocList();
        }));
        const btns = qsa(".mode-btn[data-mode]");
        btns.forEach((btn) => onClick(btn, () => {
            setActive(btns, btn, "active");
            this.activeMode = btn.dataset.mode;
            const hints = {
                analyse: "Posez une question, demandez une analyse du dossier…",
                redaction: "Précisez l'acte à rédiger (courrier, assignation, conclusions, contrat…)",
                modification: "Indiquez le document à modifier et les changements souhaités…",
                note: "Rédigez une correction → sauvegardée dans _notes.json…",
            };
            const ta = byID("user-input");
            if (!ta)
                return;
            ta.placeholder = hints[this.activeMode];
        }));
        const summBtn = btn("btn-case-summary");
        onClick(summBtn, async () => {
            const ta = btn("user-input");
            if (!ta)
                return;
            ta.value =
                "Fais un point complet sur ce dossier : enjeux principaux, risques identifiés, actions restantes, points d'attention prioritaires.";
            await this.sendMessage();
        });
    }
    autoResize(ta) {
        ta.style.height = "auto";
        ta.style.height = Math.min(ta.scrollHeight, 160) + "px";
    }
    setupInputArea() {
        const ta = byID("user-input");
        const sendBtn = byID("send-btn");
        if (!ta || !sendBtn)
            return;
        ta.addEventListener("input", () => this.autoResize(ta));
        ta.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                this.sendMessage();
            }
        });
        sendBtn.onclick = () => this.sendMessage();
        qsa(".quick-btn").forEach((b) => {
            b.onclick = () => {
                ta.value = b.dataset.prompt ?? "";
                this.autoResize(ta);
                ta.focus();
            };
        });
    }
    showEmptyState() {
        const area = byID("chat-area");
        if (!area)
            return;
        area.innerHTML = "";
        area.appendChild(el("div", { className: "empty-state" }, el("div", { className: "empty-icon", textContent: "⚖️" }), el("h2", { textContent: "Bienvenue dans Lex Assistant" }), el("p", {
            textContent: "Connectez OneDrive et créez votre premier dossier.",
        }), (() => {
            const b = el("button", {
                className: "btn btn--primary",
                textContent: "+ Nouveau dossier",
            });
            b.onclick = () => this.openCaseFormModal(null);
            return b;
        })()));
    }
    showNotConnected() {
        const area = byID("chat-area");
        if (!area)
            return;
        area.innerHTML = "";
        area.appendChild(el("div", { className: "empty-state" }, el("div", { className: "empty-icon", textContent: "☁" }), el("h2", { textContent: "OneDrive non connecté" }), el("p", {
            textContent: "Configurez votre App Registration Azure et connectez-vous.",
        }), (() => {
            const b = el("button", {
                className: "btn btn--primary",
                textContent: "☁ Configurer OneDrive",
            });
            b.onclick = () => this.openSettingsModal();
            return b;
        })()));
    }
    sanitiseFolder(name) {
        return name
            .replace(/[/\\:*?"<>|]/g, "_")
            .replace(/\s+/g, "_")
            .slice(0, 60);
    }
}
// ─── Library — bibliothèque scenario ─────────────────────────────────────────
export class Library extends Scenario {
    mainFolder = FOLDER_LIBRARY;
    activeDomain = "all";
    domainDocs = new Map();
    libMessages = [];
    DOMAINS = [
        { id: "commercial", label: "Commercial", icon: "🏢" },
        { id: "fiscal", label: "Fiscal", icon: "💰" },
        { id: "social", label: "Social", icon: "👥" },
        { id: "civil", label: "Civil", icon: "⚖️" },
        { id: "penal", label: "Pénal", icon: "🔒" },
        { id: "immobilier", label: "Immobilier", icon: "🏠" },
        { id: "international", label: "International", icon: "🌐" },
        { id: "autre", label: "Autre", icon: "📚" },
    ];
    domainLabel(id) {
        if (id === "all")
            return "Tous domaines";
        return this.DOMAINS.find((d) => d.id === id)?.label ?? id;
    }
    // ─── Path helpers ─────────────────────────────────────────────────────────────
    libDomainFolder(subFolder) {
        const cap = subFolder.charAt(0).toUpperCase() + subFolder.slice(1);
        return `${this._rootFolder}/${FOLDER_LIBRARY}/${cap}`;
    }
    libMetaPath(subFolder) {
        return `${this.libDomainFolder(subFolder)}/_meta.json`;
    }
    // ─── CRUD ─────────────────────────────────────────────────────────────────────
    async readLibMeta(domain) {
        return this.readJson(this.libMetaPath(domain));
    }
    async writeLibMeta(meta) {
        await this.ensureFolder(this.libDomainFolder(meta.domain));
        await this.writeJson(this.libMetaPath(meta.domain), meta);
    }
    async readLibConversation() {
        const path = this.getConvPath();
        const f = await this.readJson(path);
        return f?.messages ?? [];
    }
    async writeLibConversation(messages) {
        const path = this.getConvPath();
        await this.writeJson(path, { messages });
    }
    getConvPath() {
        return this.activeDomain === "all"
            ? `${this._rootFolder}/${FOLDER_LIBRARY}/_conversation.json`
            : `${this.libDomainFolder(this.activeDomain)}/_conversation.json`;
    }
    async listLibFiles(domain) {
        try {
            const items = await this.listFolder(this.libDomainFolder(domain));
            return items.filter((i) => i.file && !i.name.startsWith("_"));
        }
        catch {
            return [];
        }
    }
    async readLibFile(domain, fileName) {
        return this.readFilePath(`${this.libDomainFolder(domain)}/${fileName}`);
    }
    async writeLibFile(domain, fileName, data, mimeType) {
        await this.ensureFolder(this.libDomainFolder(domain));
        await this.writeFileLarge(`${this.libDomainFolder(domain)}/${fileName}`, data, mimeType);
    }
    // ─── Domain meta ─────────────────────────────────────────────────────────────
    async loadDomainMeta(domain) {
        if (this.domainDocs.has(domain))
            return this.domainDocs.get(domain);
        const meta = await this.readLibMeta(domain).catch(() => null);
        const docs = meta?.documents ?? [];
        this.domainDocs.set(domain, docs);
        return docs;
    }
    async saveDomainMeta(domain, docs) {
        this.domainDocs.set(domain, docs);
        await this.writeLibMeta({ domain, documents: docs });
    }
    // ─── Sync from OneDrive ───────────────────────────────────────────────────────
    async syncDomainFromOneDrive(domain) {
        if (!oneDrive.userName)
            await this.connectOneDrive();
        const items = await this.listLibFiles(domain);
        const existing = await this.loadDomainMeta(domain);
        let added = 0;
        for (const item of items) {
            if (!item.file || !isSupported(item.name))
                continue;
            if (existing.some((d) => d.name === item.name))
                continue;
            existing.push({
                name: item.name,
                mimeType: item.file.mimeType || "application/octet-stream",
                sizeBytes: item.size ?? 0,
                addedAt: Date.now(),
                tags: [],
            });
            added++;
        }
        // FIX: save to the correct domain being synced, not activeDomain
        if (added > 0)
            await this.saveDomainMeta(domain, existing);
        return added;
    }
    setupBarsBtns() {
        const btn = (id) => byID(id), onClick = this.onClick;
        [btn("btn-new-item-top"), btn("btn-new-item-sidebar")].forEach((btn) => onClick(btn, () => this.loadAllSubFolders()));
        onClick(btn("btn-settings"), () => this.openSettingsModal());
        onClick(btn("btn-onedrive"), async () => await this.syncCurrentDomain());
        onClick(btn("btn-od-sync"), async () => await this.syncCurrentDomain());
        onClick(btn("btn-upload"), () => btn("file-input")?.click());
        const btns = qsa(".mode-btn[data-mode]");
        const ta = btn("user-input");
        const hints = {
            analyse: "Posez une question, demandez une analyse du dossier…",
            redaction: "Précisez l'acte à rédiger (courrier, assignation, conclusions, contrat…)",
            modification: "Indiquez le document à modifier et les changements souhaités…",
            note: "Rédigez une correction → sauvegardée dans _notes.json…",
        };
        btns.forEach((btn) => onClick(btn, () => {
            setActive(btns, btn, "active");
            this.activeMode = btn.dataset.mode;
            if (!ta)
                return;
            ta.placeholder = hints[this.activeMode];
        }));
        onClick(btn("btn-case-summary"), async () => {
            const ta = btn("user-input");
            if (!ta)
                return;
            ta.value =
                "Fais un point complet sur ce dossier : enjeux principaux, risques identifiés, actions restantes, points d'attention prioritaires.";
            await this.sendMessage();
        });
    }
    sendMessage() {
        //!Missing
    }
    loadAllSubFolders() {
        //!Missing
    }
    // ─── Boot ─────────────────────────────────────────────────────────────────────
    async bootLib(container) {
        container.innerHTML = "";
        container.appendChild(this.buildUI());
        await this.loadApiKey();
        // Use `this` throughout — no second instance
        this.libMessages = await this.readLibConversation().catch(() => []);
        this.renderDomainPills();
        this.renderLibChat();
        this.renderLibDocList();
        this.setupLibInput();
        this.setupLibUpload();
        this.updateLibSkillIndicator();
    }
    // ─── UI builder ───────────────────────────────────────────────────────────────
    buildUI() {
        const wrap = el("div", { className: "lib-layout" });
        const sidebar = el("div", { className: "lib-sidebar" });
        const hdr = el("div", { className: "lib-sidebar__header" });
        hdr.appendChild(el("span", {
            className: "lib-sidebar__title",
            textContent: "Bibliothèque juridique",
        }));
        const skillBadge = el("span", {
            className: "lib-skill-badge",
            id: "lib-skill-badge",
        });
        toggle(skillBadge, false);
        hdr.appendChild(skillBadge);
        sidebar.appendChild(hdr);
        sidebar.appendChild(el("div", { className: "lib-domain-pills", id: "lib-domain-pills" }));
        sidebar.appendChild(el("div", { className: "lib-doc-list", id: "lib-doc-list" }));
        const acts = el("div", { className: "lib-action-row" });
        const fi = el("input", {
            type: "file",
            id: "lib-file-input",
            multiple: true,
        });
        fi.accept =
            ".pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt,.txt,.md";
        fi.style.display = "none";
        const upBtn = el("button", {
            className: "btn btn--ghost btn--sm",
            textContent: "⬆ Ajouter",
            id: "btn-lib-upload",
        });
        upBtn.onclick = () => fi.click();
        const synBtn = el("button", {
            className: "btn btn--ghost btn--sm",
            textContent: "⟳ Sync OneDrive",
            id: "btn-lib-sync",
        });
        synBtn.onclick = async () => await this.syncCurrentDomain();
        acts.append(fi, upBtn, synBtn);
        sidebar.appendChild(acts);
        const main = el("div", { className: "lib-main" });
        const topbar = el("div", { className: "lib-topbar" });
        const domLbl = el("span", {
            className: "lib-topbar_domain",
            id: "lib-active-domain",
            textContent: "Tous domaines",
        });
        const skillLbl = el("span", {
            className: "lib-topbar__skills",
            id: "lib-skills-top",
        });
        const clrBtn = el("button", {
            className: "btn btn--ghost btn--sm",
            textContent: "Effacer conversation",
        });
        clrBtn.onclick = () => this.clearLibConv();
        topbar.append(domLbl, skillLbl, clrBtn);
        const quickArea = el("div", {
            className: "lib-quick-prompts",
            id: "lib-quick-prompts",
        });
        const chatArea = el("div", {
            className: "chat-area",
            id: "lib-chat-area",
            role: "log",
        });
        chatArea.setAttribute("aria-live", "polite");
        const inputArea = el("div", { className: "lib-input-area" });
        const ta = el("textarea", {
            id: "lib-input",
            rows: 2,
            placeholder: "Interrogez la bibliothèque…",
        });
        const sendBtn = el("button", {
            className: "btn btn--primary",
            id: "lib-send-btn",
        });
        sendBtn.innerHTML =
            '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>';
        inputArea.append(ta, sendBtn);
        main.append(topbar, quickArea, chatArea, inputArea);
        wrap.append(sidebar, main);
        return wrap;
    }
    // ─── Domain navigation ────────────────────────────────────────────────────────
    renderDomainPills() {
        const container = byID("lib-domain-pills");
        if (!container)
            return;
        container.innerHTML = "";
        const all = el("button", {
            className: `lib-pill${this.activeDomain === "all" ? " active" : ""}`,
            textContent: "Tous",
        });
        all.onclick = () => this.switchDomain("all");
        container.appendChild(all);
        for (const d of this.DOMAINS) {
            const docs = this.domainDocs.get(d.id) ?? [];
            const pill = el("button", {
                className: `lib-pill${this.activeDomain === d.id ? " active" : ""}`,
            });
            pill.textContent = `${d.icon} ${d.label}`;
            if (docs.length)
                pill.appendChild(el("span", {
                    className: "lib-pill__count",
                    textContent: String(docs.length),
                }));
            pill.onclick = () => this.switchDomain(d.id);
            container.appendChild(pill);
        }
    }
    async switchDomain(domain) {
        this.activeDomain = domain;
        if (domain !== "all")
            await this.loadDomainMeta(domain);
        this.libMessages = await this.readLibConversation().catch(() => []);
        this.renderDomainPills();
        this.renderLibDocList();
        this.renderLibChat();
        const lbl = byID("lib-active-domain");
        if (lbl)
            lbl.textContent = this.domainLabel(domain);
        this.updateQuickPrompts();
    }
    // ─── Render: lib doc list ─────────────────────────────────────────────────────
    renderLibDocList() {
        const list = byID("lib-doc-list");
        if (!list)
            return;
        list.innerHTML = "";
        let docs = [];
        if (this.activeDomain === "all") {
            for (const [, d] of this.domainDocs)
                docs.push(...d);
        }
        else {
            docs = this.domainDocs.get(this.activeDomain) ?? [];
        }
        docs = docs.sort((a, b) => b.addedAt - a.addedAt);
        if (!docs.length) {
            list.appendChild(el("div", {
                className: "lib-doc-empty",
                textContent: "Aucun document. Ajoutez ou synchronisez.",
            }));
            return;
        }
        for (const doc of docs) {
            const item = el("div", { className: "lib-doc-item" });
            const info = el("div", { className: "doc-info" });
            info.append(el("div", { className: "doc-name", textContent: doc.name }), el("div", {
                className: "doc-meta",
                textContent: `${mimeLabel(doc.mimeType)} · ${formatSize(doc.sizeBytes)} · ${formatDate(doc.addedAt)}`,
            }));
            const del = el("button", {
                className: "doc-delete",
                textContent: "×",
                title: "Retirer de la bibliothèque",
            });
            del.onclick = async (e) => {
                e.stopPropagation();
                const ok = await confirm(`Retirer "${doc.name}" de la bibliothèque ? (Fichier OneDrive conservé.)`);
                if (!ok)
                    return;
                for (const [dom, list] of this.domainDocs) {
                    const idx = list.findIndex((d) => d.name === doc.name);
                    if (idx >= 0) {
                        list.splice(idx, 1);
                        await this.saveDomainMeta(dom, list);
                        break;
                    }
                }
                this.renderLibDocList();
                this.renderDomainPills();
                toast("Document retiré de la bibliothèque.", "info");
            };
            item.append(el("span", {
                className: "doc-icon",
                textContent: mimeIcon(doc.mimeType),
            }), info, del);
            list.appendChild(item);
        }
    }
    // ─── Render: lib chat ─────────────────────────────────────────────────────────
    renderLibChat() {
        const area = byID("lib-chat-area");
        if (!area)
            return;
        area.innerHTML = "";
        if (!this.libMessages.length) {
            area.appendChild(el("div", { className: "empty-state" }, el("div", { className: "empty-icon", textContent: "📚" }), el("h2", { textContent: "Bibliothèque juridique" }), el("p", {
                textContent: "Sélectionnez un domaine, synchronisez vos documents OneDrive, puis posez votre question.",
            })));
            return;
        }
        for (const msg of this.libMessages)
            area.appendChild(this.buildLibMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    buildLibMsgEl(msg) {
        const wrap = el("div", { className: `msg msg--${msg.role}` });
        const bubble = el("div", { className: "msg__bubble" });
        bubble.innerHTML = renderMarkdown(msg.content);
        wrap.append(el("div", {
            className: "msg__label",
            textContent: msg.role === "user" ? "Vous" : "Lex Assistant",
        }), bubble);
        if (msg.role === "assistant") {
            const acts = el("div", { className: "msg__actions" });
            const copy = el("button", {
                className: "msg-action-btn",
                textContent: "Copier",
            });
            copy.onclick = () => {
                navigator.clipboard.writeText(msg.content);
                toast("Copié.", "info", 1500);
            };
            acts.appendChild(copy);
            wrap.appendChild(acts);
        }
        return wrap;
    }
    appendLibMsg(msg) {
        const area = byID("lib-chat-area");
        if (!area)
            return;
        area.querySelector(".empty-state")?.remove();
        area.appendChild(this.buildLibMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    appendLibTyping() {
        const area = byID("lib-chat-area");
        const typing = el("div", {
            className: "msg msg--assistant",
            id: "lib-typing",
        });
        const bubble = el("div", { className: "msg__bubble" });
        bubble.append(spinnerEl(), el("span", { textContent: " Consultation de la bibliothèque…" }));
        typing.append(el("div", { className: "msg__label", textContent: "Lex Assistant" }), bubble);
        area.appendChild(typing);
        area.scrollTop = area.scrollHeight;
        return typing;
    }
    async clearLibConv() {
        const ok = await confirm("Effacer l'historique de la bibliothèque pour ce domaine ?");
        if (!ok)
            return;
        this.libMessages = [];
        await this.writeLibConversation([]);
        this.renderLibChat();
        toast("Conversation effacée.", "info");
    }
    // ─── Send lib message ─────────────────────────────────────────────────────────
    async sendLibMessage() {
        const ta = byID("lib-input");
        const sendBtn = byID("lib-send-btn");
        if (!ta || !sendBtn)
            return;
        const text = ta.value.trim();
        if (!text)
            return;
        ta.value = "";
        let docs = [];
        if (this.activeDomain === "all") {
            for (const [, d] of this.domainDocs)
                docs.push(...d);
            for (const d of this.DOMAINS) {
                if (!this.domainDocs.has(d.id)) {
                    const loaded = await this.loadDomainMeta(d.id);
                    docs.push(...loaded);
                }
            }
        }
        else {
            docs = await this.loadDomainMeta(this.activeDomain);
        }
        const userMsg = {
            id: uid(),
            role: "user",
            content: text,
            timestamp: Date.now(),
            domain: this.activeDomain,
        };
        this.libMessages.push(userMsg);
        this.appendLibMsg(userMsg);
        const typing = this.appendLibTyping();
        sendBtn.disabled = true;
        const history = this.libMessages.slice(-21, -1).map((m) => ({
            role: m.role,
            content: m.content,
        }));
        try {
            const response = await callClaudeLib({
                domain: this.activeDomain,
                docs,
                skills: this.skills,
                userMessage: text,
                history,
                readFile: (fileName) => this.readLibFile(this.activeDomain === "all" ? "all" : this.activeDomain, fileName),
            });
            typing.remove();
            const asstMsg = {
                id: uid(),
                role: "assistant",
                content: response,
                timestamp: Date.now(),
                domain: this.activeDomain,
            };
            this.libMessages.push(asstMsg);
            this.appendLibMsg(asstMsg);
            await this.writeLibConversation(this.libMessages);
        }
        catch (err) {
            typing.remove();
            toast(err.message, "error", 6000);
            this.libMessages.pop();
        }
        finally {
            sendBtn.disabled = false;
            ta.focus();
        }
    }
    // ─── Sync current domain ──────────────────────────────────────────────────────
    async syncCurrentDomain() {
        const btn = byID("btn-lib-sync");
        if (btn) {
            btn.disabled = true;
            btn.textContent = "⟳ Sync…";
        }
        try {
            let total = 0;
            if (this.activeDomain === "all") {
                for (const d of this.DOMAINS)
                    total += await this.syncDomainFromOneDrive(d.id);
            }
            else {
                total = await this.syncDomainFromOneDrive(this.activeDomain);
            }
            this.renderLibDocList();
            this.renderDomainPills();
            toast(`${total} nouveau(x) document(s) indexé(s).`, "success");
        }
        catch (err) {
            toast("Erreur sync : " + err.message, "error");
        }
        finally {
            if (btn) {
                btn.disabled = false;
                btn.textContent = "⟳ Sync OneDrive";
            }
        }
    }
    // ─── Quick prompts ────────────────────────────────────────────────────────────
    updateQuickPrompts() {
        const area = byID("lib-quick-prompts");
        if (!area)
            return;
        area.innerHTML = "";
        const prompts = {
            commercial: [
                "Jurisprudence récente sur la responsabilité du dirigeant pour insuffisance d'actif.",
                "Conditions de validité d'une clause de non-concurrence en droit commercial français.",
                "Règles applicables à la cession de fonds de commerce.",
            ],
            fiscal: [
                "Analyse la jurisprudence sur l'abus de droit fiscal (LPF art. L.64).",
                "Conditions d'application de l'acte anormal de gestion.",
                "Jurisprudence récente sur la déductibilité des charges en IS.",
            ],
            social: [
                "Conditions de validité du licenciement pour motif économique.",
                "Analyse jurisprudentielle du harcèlement moral au travail.",
                "Règles applicables au transfert du contrat de travail (L.1224-1 CT).",
            ],
            civil: [
                "Jurisprudence récente sur la responsabilité délictuelle.",
                "Conditions de la résolution pour inexécution (C.civ. art. 1224).",
                "Évolutions de la jurisprudence sur le préjudice moral.",
            ],
            penal: [
                "Éléments constitutifs de l'abus de biens sociaux.",
                "Jurisprudence sur la complicité en droit pénal des affaires.",
                "Conditions de mise en cause de la responsabilité pénale des personnes morales.",
            ],
            immobilier: [
                "Régime des baux commerciaux : droit au renouvellement et indemnité d'éviction.",
                "Conditions de l'action en garantie des vices cachés en droit immobilier.",
                "Jurisprudence sur la responsabilité du promoteur immobilier.",
            ],
            international: [
                "Conditions d'applicabilité des conventions fiscales bilatérales.",
                "Jurisprudence sur le centre des intérêts vitaux (CGI art. 4 B).",
                "Règles de conflit de lois en matière successorale (Règl. UE 650/2012).",
            ],
            all: [
                "Quels sont les documents disponibles dans la bibliothèque ?",
                "Synthèse des principales règles jurisprudentielles sur la responsabilité civile.",
                "Analyse comparative des régimes de responsabilité civile et pénale du dirigeant.",
            ],
        };
        const list = prompts[this.activeDomain] ?? prompts.all;
        for (const p of list) {
            const btn = el("button", { className: "quick-btn", textContent: p });
            btn.onclick = () => {
                const ta = byID("lib-input");
                if (ta) {
                    ta.value = p;
                    ta.focus();
                }
            };
            area.appendChild(btn);
        }
    }
    setupLibInput() {
        const ta = byID("lib-input");
        const sendBtn = byID("lib-send-btn");
        if (!ta || !sendBtn)
            return;
        ta.addEventListener("input", () => {
            ta.style.height = "auto";
            ta.style.height = Math.min(ta.scrollHeight, 160) + "px";
        });
        ta.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                this.sendLibMessage();
            }
        });
        sendBtn.onclick = () => this.sendLibMessage();
        this.updateQuickPrompts();
    }
    setupLibUpload() {
        const fi = byID("lib-file-input");
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
                    toast(`Format non supporté : ${file.name}`, "error");
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
                    toast(`"${file.name}" ajouté à ${this.domainLabel(domain)}.`, "success");
                }
                catch (err) {
                    toast(`Erreur : ${err.message}`, "error");
                }
            }
            fi.value = "";
            this.renderLibDocList();
            this.renderDomainPills();
        };
    }
    updateLibSkillIndicator() {
        const badge = byID("lib-skill-badge");
        const top = byID("lib-skills-top");
        const n = this.skills.length;
        if (badge) {
            badge.textContent = n > 0 ? `${n} skill${n > 1 ? "s" : ""}` : "";
            toggle(badge, n > 0);
        }
        if (top) {
            top.textContent =
                n > 0 ? `${n} skill${n > 1 ? "s" : ""} actif${n > 1 ? "s" : ""}` : "";
        }
    }
    // ─── Domain picker modal ──────────────────────────────────────────────────────
    pickDomainModal() {
        return new Promise((resolve) => {
            const overlay = el("div", { className: "modal-overlay" });
            const dialog = el("div", { className: "modal-dialog" });
            dialog.innerHTML =
                '<h2 class="modal-title">Domaine juridique</h2><p class="modal-subtitle">Dans quel domaine classer ce(s) document(s) ?</p>';
            const grid = el("div", { className: "domain-grid" });
            for (const d of this.DOMAINS) {
                const btn = el("button", { className: "domain-btn" });
                btn.innerHTML = `<span class="domain-btn__icon">${d.icon}</span><span>${d.label}</span>`;
                btn.onclick = () => {
                    overlay.remove();
                    resolve(d.id);
                };
                grid.appendChild(btn);
            }
            dialog.appendChild(grid);
            const cancel = el("button", {
                className: "btn btn--secondary",
                textContent: "Annuler",
            });
            cancel.style.marginTop = "16px";
            cancel.onclick = () => {
                overlay.remove();
                resolve(null);
            };
            dialog.appendChild(cancel);
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
        });
    }
}
//# sourceMappingURL=onedrive.js.map