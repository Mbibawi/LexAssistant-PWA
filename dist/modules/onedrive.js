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
import { toast, spinnerEl, el, toggle, uid, formatDate, formatDateTime, qs, qsa, setActive } from './ui.js';
import { renderMarkdown } from './markdown.js';
import { callClaudeLib, callClaudeCase, getStoredKey, setStoredKey, clearStoredKey } from './api.js';
import { isSupported, mimeLabel, mimeIcon, formatSize, makeCaseDocMeta, makeLibDocMeta, guessKind, kindLabel } from './ingest.js';
import { downloadBlob, generateDocx } from './docxgen.js';
// ─── Constants ────────────────────────────────────────────────────────────────
const LS_CONFIG = 'lex_onedrive_config';
const MSAL_CDN = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3/dist/index.iife.min.js';
const GRAPH = 'https://graph.microsoft.com/v1.0';
// ─── Configuration ────────────────────────────────────────────────────────────
class Configuration {
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
    clearConfig() { localStorage.removeItem(LS_CONFIG); }
    isConfigured() {
        const c = this.getConfig();
        return Boolean(c?.clientId && c?.rootFolder);
    }
    root() {
        return this.getConfig()?.rootFolder ?? 'LexAssistant';
    }
}
// ─── oneDrive — MSAL auth + raw Graph fetch ───────────────────────────────────
class OneDriveAuth {
    config = new Configuration();
    _scopes = ["Files.ReadWrite", "User.Read"];
    _app = null;
    _loading = null;
    _account = null;
    getMsal() {
        if (this._app)
            return Promise.resolve(this._app);
        if (this._loading)
            return this._loading;
        this._loading = new Promise((res, rej) => {
            const init = async () => {
                const cfg = this.config.getConfig();
                if (!cfg)
                    throw new Error("OneDrive non configuré");
                const app = new msal.PublicClientApplication({
                    auth: {
                        clientId: cfg.clientId,
                        authority: `https://login.microsoftonline.com/${cfg.tenantId ?? "common"}`,
                        redirectUri: window.location.origin,
                    },
                    cache: {
                        cacheLocation: "sessionStorage",
                        storeAuthStateInCookie: false,
                    },
                });
                await app.handleRedirectPromise().catch(() => null);
                this._app = app;
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
        const app = await this.getMsal();
        if (!this._account) {
            const accounts = app.getAllAccounts();
            if (accounts.length)
                this._account = accounts[0];
        }
        if (this._account) {
            try {
                return (await app.acquireTokenSilent({
                    scopes: this._scopes,
                    account: this._account,
                })).accessToken;
            }
            catch {
                /* fall through to popup */
            }
        }
        const r = await app.loginPopup({ scopes: this._scopes });
        this._account = app.getAllAccounts()[0] ?? null;
        return r.accessToken;
    }
    async signIn() {
        const app = await this.getMsal();
        await app.loginPopup({ scopes: this._scopes });
        this._account = app.getAllAccounts()[0] ?? null;
    }
    async signOut() {
        this._account = null;
        this._app = null;
        this._loading = null;
        sessionStorage.clear();
    }
    async isSignedIn() {
        try {
            const app = await this.getMsal();
            const accounts = app.getAllAccounts();
            if (accounts.length) {
                this._account = accounts[0];
                return true;
            }
            return false;
        }
        catch {
            return false;
        }
    }
    getSignedInUser() {
        return this._account?.name ?? this._account?.username ?? null;
    }
}
// ─── Folders — base file/folder/JSON operations ───────────────────────────────
class Folders extends OneDriveAuth {
    get rootFolder() { return this.config.root(); }
    token = null;
    async setToken() {
        if (!this.token)
            this.token = await this.getAccessToken();
        return this.token;
    }
    // Encode a path for Graph: "a/b/c" → "/me/drive/root:/a/b/c:"
    p(odPath) {
        const encoded = odPath.split('/').map(seg => encodeURIComponent(seg)).join('/');
        return `/me/drive/root:/${encoded}:`;
    }
    async gFetch(path, opts = {}, rawBody = false) {
        const headers = { Authorization: `Bearer ${this.token ?? await this.setToken()}` };
        if (!rawBody && opts.body && typeof opts.body === 'string') {
            headers['Content-Type'] = 'application/json';
        }
        const resp = await fetch(`${GRAPH}${path}`, { ...opts, headers });
        if (!resp.ok) {
            let msg = resp.statusText;
            try {
                const e = await resp.json();
                msg = e.error?.message ?? msg;
            }
            catch { }
            throw new Error(`Graph ${resp.status}: ${msg}`);
        }
        return resp;
    }
    async ensureFolder(folderPath) {
        try {
            await this.gFetch(this.p(folderPath));
            return;
        }
        catch { }
        const parts = folderPath.split('/');
        const name = parts.pop();
        const parentPath = parts.join('/');
        const parentEndpoint = parentPath
            ? `${this.p(parentPath)}/children`
            : `/me/drive/root/children`;
        await this.gFetch(parentEndpoint, {
            method: 'POST',
            body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }),
        });
    }
    async listFolder(folderPath) {
        const resp = await this.gFetch(`${this.p(folderPath)}/children?$select=name,size,file,folder,webUrl,lastModifiedDateTime&$top=500`);
        const data = await resp.json();
        return data.value ?? [];
    }
    async readFilePath(filePath) {
        const resp = await fetch(`${GRAPH}${this.p(filePath)}/content`, {
            headers: { Authorization: `Bearer ${this.token ?? await this.setToken()}` },
        });
        if (!resp.ok)
            throw new Error(`Read ${filePath}: ${resp.status}`);
        return resp.arrayBuffer();
    }
    async writeFilePath(filePath, data, mimeType) {
        const body = typeof data === 'string' ? new TextEncoder().encode(data) : data;
        const resp = await fetch(`${GRAPH}${this.p(filePath)}/content`, {
            method: 'PUT',
            headers: { Authorization: `Bearer ${this.token ?? await this.setToken()}`, 'Content-Type': mimeType },
            body,
        });
        if (!resp.ok) {
            let msg = resp.statusText;
            try {
                const e = await resp.json();
                msg = e.error?.message ?? msg;
            }
            catch { }
            throw new Error(`Write ${filePath}: ${msg}`);
        }
    }
    async writeFileLarge(filePath, data, mimeType) {
        if (data.byteLength <= 4 * 1024 * 1024) {
            await this.writeFilePath(filePath, data, mimeType);
            return;
        }
        const sessResp = await this.gFetch(`${this.p(filePath)}/createUploadSession`, {
            method: 'POST',
            body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
        });
        const { uploadUrl } = await sessResp.json();
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
    async deleteFilePath(filePath) {
        await this.gFetch(`${this.p(filePath)}`, { method: 'DELETE' });
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
    async writeJson(filePath, data) {
        await this.writeFilePath(filePath, JSON.stringify(data, null, 2), 'application/json');
    }
    async initRootStructure() {
        const r = this.rootFolder;
        await this.ensureFolder(r);
        await this.ensureFolder(`${r}/_Skills`);
        await this.ensureFolder(`${r}/Affaires`);
        await this.ensureFolder(`${r}/Bibliotheque`);
        for (const d of ['Commercial', 'Fiscal', 'Social', 'Civil', 'Penal', 'Immobilier', 'International']) {
            await this.ensureFolder(`${r}/Bibliotheque/${d}`);
        }
    }
}
// ─── Cases — dossiers scenario ────────────────────────────────────────────────
class Cases extends Folders {
    skills = [];
    oneDriveUser = null;
    saving = false;
    docFilter = "all";
    cases = [];
    activeCase = null;
    caseNotes = [];
    caseMessages = [];
    activeMode = "analyse";
    // ─── Path helpers ─────────────────────────────────────────────────────────────
    casePath = (f) => `${this.rootFolder}/Affaires/${f}`;
    metaPath = (f) => `${this.casePath(f)}/_meta.json`;
    notesPath = (f) => `${this.casePath(f)}/_notes.json`;
    convPath = (f) => `${this.casePath(f)}/_conversation.json`;
    // ─── CRUD ─────────────────────────────────────────────────────────────────────
    async readCaseMeta(folderName) {
        return this.readJson(this.metaPath(folderName));
    }
    async writeCaseMeta(folderName, meta) {
        await this.ensureFolder(this.casePath(folderName));
        await this.writeJson(this.metaPath(folderName), meta);
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
    async listCaseFolders() {
        try {
            const items = await this.listFolder(`${this.rootFolder}/Affaires`);
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
        if (!this.activeCase || this.saving)
            return;
        this.saving = true;
        const meta = {
            name: this.activeCase.name,
            folderName: this.activeCase.folderName,
            domain: this.activeCase.domain,
            status: this.activeCase.status,
            createdAt: this.activeCase.createdAt,
            updatedAt: Date.now(),
            documents: this.activeCase.documents,
        };
        this.activeCase.updatedAt = meta.updatedAt;
        try {
            await this.writeCaseMeta(this.activeCase.folderName, meta);
        }
        finally {
            this.saving = false;
        }
    }
    async saveNotes() {
        if (!this.activeCase)
            return;
        await this.writeNotes(this.activeCase.folderName, this.caseNotes);
    }
    async saveConversation() {
        if (!this.activeCase)
            return;
        await this.writeConversation(this.activeCase.folderName, this.caseMessages);
    }
    // ─── Skills ───────────────────────────────────────────────────────────────────
    /** Called by main.ts after fetchSkills() to load skills into both modules */
    async fetchSkills() {
        const path = `${this.rootFolder}/_Skills`;
        let items;
        try {
            items = (await this.listFolder(path));
            items = items.filter(item => !item.folder);
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
        this.syncSkills(skills);
        return skills;
    }
    syncSkills(skills) {
        this.skills = skills;
        this.updateSkillIndicator();
    }
    // ─── Load all cases ───────────────────────────────────────────────────────────
    async loadAllCases() {
        const folders = await this.listCaseFolders();
        this.cases = [];
        await Promise.all(folders.map(async (folderName) => {
            const meta = await this.readCaseMeta(folderName);
            if (!meta)
                return;
            this.cases.push({
                folderName,
                name: meta.name,
                domain: meta.domain,
                status: meta.status,
                createdAt: meta.createdAt,
                updatedAt: meta.updatedAt,
                documents: meta.documents ?? [],
            });
        }));
        this.cases.sort((a, b) => b.updatedAt - a.updatedAt);
        this.renderCaseList();
        if (this.cases.length > 0)
            await this.selectCase(this.cases[0].folderName);
        else
            this.showEmptyState();
    }
    // ─── Select case ──────────────────────────────────────────────────────────────
    async selectCase(folderName) {
        const c = this.cases.find((x) => x.folderName === folderName);
        if (!c)
            return;
        this.activeCase = c;
        this.caseNotes = await this.readNotes(folderName);
        this.caseMessages = await this.readConversation(folderName);
        qsa(".case-item").forEach((el) => el.classList.toggle("active", el.dataset.folder === folderName));
        const nameEl = document.getElementById("topbar-case-name");
        const domainEl = document.getElementById("topbar-case-domain");
        if (nameEl)
            nameEl.textContent = c.name;
        if (domainEl)
            domainEl.textContent = c.domain;
        this.renderNoteBar();
        this.renderDocList();
        this.renderChat();
        this.updateDocCount();
    }
    // ─── OneDrive connection ──────────────────────────────────────────────────────
    updateODStatus() {
        const statusEl = document.getElementById("onedrive-status");
        if (!statusEl)
            return;
        if (this.oneDriveUser) {
            statusEl.textContent = `☁ ${this.oneDriveUser}`;
            statusEl.className = "od-status od-status--connected";
        }
        else {
            statusEl.textContent = this.config.isConfigured()
                ? "☁ Non connecté"
                : "☁ Non configuré";
            statusEl.className = "od-status od-status--disconnected";
        }
    }
    async connectOneDrive() {
        try {
            await this.signIn();
            this.oneDriveUser = this.getSignedInUser();
            this.updateODStatus();
            await this.initRootStructure();
            await this.loadAllCases();
            const skills = await this.fetchSkills();
            this.syncSkills(skills);
            toast(`Connecté : ${this.oneDriveUser}`, "success");
        }
        catch (err) {
            toast("Erreur connexion OneDrive : " + err.message, "error");
        }
    }
    async refreshCaseFromOneDrive() {
        if (!this.activeCase)
            return;
        if (!this.oneDriveUser) {
            await this.connectOneDrive();
            return;
        }
        try {
            const items = await this.listCaseFiles(this.activeCase.folderName);
            let added = 0;
            for (const item of items) {
                if (!item.file || !isSupported(item.name))
                    continue;
                if (this.activeCase.documents.some((d) => d.name === item.name))
                    continue;
                this.activeCase.documents.push({
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
        const badge = document.getElementById("skills-badge");
        const n = this.skills.length;
        if (!badge)
            return;
        badge.textContent = n > 0 ? `${n} skill${n > 1 ? "s" : ""}` : "";
        toggle(badge, n > 0);
    }
    // ─── Render: case list ────────────────────────────────────────────────────────
    renderCaseList() {
        const list = document.getElementById("case-list");
        if (!list)
            return;
        list.innerHTML = "";
        const labels = {
            active: "En cours",
            closed: "Clôturé",
            suspended: "Suspendu",
        };
        for (const c of this.cases) {
            const item = el("div", {
                className: "case-item" +
                    (c.folderName === this.activeCase?.folderName ? " active" : ""),
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
        const list = document.getElementById("doc-list");
        if (!list || !this.activeCase)
            return;
        list.innerHTML = "";
        const docs = this.docFilter === "all"
            ? this.activeCase.documents
            : this.activeCase.documents.filter((d) => d.kind === this.docFilter);
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
                this.activeCase.documents = this.activeCase.documents.filter((d) => d.name !== doc.name);
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
        const countEl = document.getElementById("doc-count");
        if (countEl && this.activeCase) {
            countEl.textContent = `${this.activeCase.documents.length} pièce${this.activeCase.documents.length !== 1 ? "s" : ""}`;
        }
    }
    // ─── Render: note bar ─────────────────────────────────────────────────────────
    renderNoteBar() {
        const bar = document.getElementById("note-bar");
        if (!bar)
            return;
        bar.innerHTML = "";
        toggle(bar, this.caseNotes.length > 0);
        if (!this.caseNotes.length)
            return;
        const n = this.caseNotes.length;
        bar.append(el("span", { className: "note-badge", textContent: String(n) }), el("span", {
            className: "note-bar__text",
            textContent: `note${n > 1 ? "s" : ""} active${n > 1 ? "s" : ""} · ` +
                this.caseNotes.map((x) => x.content.slice(0, 40) + "…").join(" — "),
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
        const area = document.getElementById("chat-area");
        if (!area)
            return;
        area.innerHTML = "";
        if (!this.caseMessages.length) {
            const c = this.activeCase;
            if (!c)
                return;
            area.appendChild(el("div", { className: "msg msg--assistant" }, el("div", { className: "msg__label", textContent: "Lex Assistant" }), el("div", { className: "msg__bubble" }, el("p", {
                innerHTML: `Dossier <strong>${c.name}</strong>. ${c.documents.length} pièce(s), ${this.caseNotes.length} note(s)${this.skills.length ? `, ${this.skills.length} skill(s)` : ""}.`,
            }), el("p", { textContent: "Que souhaitez-vous faire ?" }))));
            return;
        }
        for (const msg of this.caseMessages)
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
                            caseRef: this.activeCase?.name ?? "",
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
                    if (!this.activeCase)
                        return;
                    try {
                        const blob = await generateDocx({
                            title: msg.generatedDocName ?? "Document",
                            content: msg.content,
                            caseRef: this.activeCase.name,
                        });
                        const ab = await blob.arrayBuffer();
                        const fname = (msg.generatedDocName ?? "document").replace(/[^a-z0-9_\- ]/gi, "_") + ".docx";
                        const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        await this.writeCaseFile(this.activeCase.folderName, fname, ab, mime);
                        if (!this.activeCase.documents.some((d) => d.name === fname)) {
                            this.activeCase.documents.push({
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
        const area = document.getElementById("chat-area");
        if (!area)
            return;
        area.querySelector(".empty-state")?.remove();
        area.appendChild(this.buildMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    appendTyping() {
        const area = document.getElementById("chat-area");
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
        const ta = document.getElementById("user-input");
        if (!ta)
            return;
        const text = ta.value.trim();
        if (!text || !this.activeCase)
            return;
        if (!getStoredKey()) {
            this.openSettingsModal();
            toast("Configurez votre clé API Claude.", "error");
            return;
        }
        if (!this.oneDriveUser) {
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
        this.caseMessages.push(userMsg);
        this.appendMsg(userMsg);
        if (this.activeMode === "note") {
            const note = {
                id: uid(),
                content: text,
                createdAt: Date.now(),
                updatedAt: Date.now(),
            };
            this.caseNotes.push(note);
            await this.saveNotes();
            this.renderNoteBar();
        }
        const typing = this.appendTyping();
        const sendBtn = document.getElementById("send-btn");
        if (sendBtn)
            sendBtn.disabled = true;
        try {
            const response = await callClaudeCase({
                caseName: this.activeCase.name,
                caseDomain: this.activeCase.domain,
                notes: this.caseNotes,
                docs: this.activeCase.documents,
                skills: this.skills,
                mode: this.activeMode,
                userMessage: text,
                readFile: (fileName) => this.readCaseFile(this.activeCase.folderName, fileName),
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
            this.caseMessages.push(asst);
            this.appendMsg(asst);
            await this.saveConversation();
        }
        catch (err) {
            typing.remove();
            toast(err.message, "error", 6000);
            this.caseMessages.pop();
        }
        finally {
            if (sendBtn)
                sendBtn.disabled = false;
            ta.focus();
        }
    }
    // ─── Modals ───────────────────────────────────────────────────────────────────
    openCaseFormModal(existing) {
        if (!this.oneDriveUser) {
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
                    const idx = this.cases.findIndex((x) => x.folderName === existing.folderName);
                    if (idx >= 0)
                        this.cases[idx] = c;
                    if (this.activeCase?.folderName === existing.folderName)
                        this.activeCase = c;
                }
                else {
                    this.cases.unshift(c);
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
        document.getElementById("context-menu")?.remove();
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
        this.caseMessages = [];
        await this.writeConversation(folderName, []);
        if (this.activeCase?.folderName === folderName)
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
        this.cases = this.cases.filter((c) => c.folderName !== folderName);
        this.renderCaseList();
        if (this.activeCase?.folderName === folderName) {
            this.activeCase = null;
            if (this.cases.length > 0)
                await this.selectCase(this.cases[0].folderName);
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
            if (!this.caseNotes.length) {
                list.appendChild(el("p", {
                    className: "note-empty",
                    textContent: 'Aucune note. Ajoutez-en une ci-dessous ou utilisez le mode "Note permanente".',
                }));
                return;
            }
            for (const note of [...this.caseNotes].sort((a, b) => b.createdAt - a.createdAt)) {
                const row = el("div", { className: "note-row" });
                const del = el("button", {
                    className: "btn btn--sm btn--danger",
                    textContent: "Supprimer",
                });
                del.onclick = async () => {
                    this.caseNotes = this.caseNotes.filter((n) => n.id !== note.id);
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
            this.caseNotes.push(note);
            await this.saveNotes();
            this.renderNoteBar();
            refresh();
            qs("#new-note", dialog).value = "";
            toast("Note sauvegardée dans _notes.json.", "success");
        };
    }
    openSettingsModal() {
        document.getElementById("settings-overlay")?.remove();
        const overlay = el("div", {
            className: "modal-overlay",
            id: "settings-overlay",
        });
        const dialog = el("div", {
            className: "modal-dialog modal-dialog--settings",
        });
        const cfg = this.config.getConfig();
        dialog.innerHTML = `
      <h2 class="modal-title">⚙ Paramètres</h2>
      <h3 class="settings-section-title">Clé API Claude (Anthropic)</h3>
      <p class="settings-hint">Stockée dans localStorage. Transmise uniquement à api.anthropic.com.</p>
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
      <input class="form-input" id="s-client" type="text" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" value="${cfg?.clientId ?? ""}"/>
      <label class="form-label">Tenant ID</label>
      <input class="form-input" id="s-tenant" type="text" placeholder="common" value="${cfg?.tenantId ?? "common"}"/>
      <label class="form-label">Dossier racine OneDrive</label>
      <input class="form-input" id="s-root" type="text" placeholder="LexAssistant" value="${cfg?.rootFolder ?? "LexAssistant"}"/>
      <div style="display:flex;gap:8px;margin-top:8px;flex-wrap:wrap">
        <button class="btn btn--primary btn--sm" id="s-od-save">Enregistrer</button>
        <button class="btn btn--secondary btn--sm" id="s-od-init">Initialiser structure OneDrive</button>
        <button class="btn btn--secondary btn--sm" id="s-od-signout">Déconnecter</button>
      </div>
      ${this.oneDriveUser ? `<p class="od-connected-label">✓ Connecté : ${this.oneDriveUser}</p>` : ""}
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
        qs("#s-api-save", dialog).onclick = () => {
            const v = qs("#s-api", dialog).value.trim();
            if (!v) {
                toast("Clé vide.", "error");
                return;
            }
            setStoredKey(v);
            toast("Clé API enregistrée.", "success");
        };
        qs("#s-api-clear", dialog).onclick = () => {
            clearStoredKey();
            qs("#s-api", dialog).value = "";
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
            this.config.setConfig({ clientId, tenantId, rootFolder });
            toast("Configuration OneDrive enregistrée.", "success");
        };
        qs("#s-od-init", dialog).onclick = async () => {
            if (!this.config.isConfigured()) {
                toast("Sauvegardez la configuration d'abord.", "error");
                return;
            }
            try {
                if (!this.oneDriveUser) {
                    await this.signIn();
                    this.oneDriveUser = this.getSignedInUser();
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
            await this.signOut();
            this.oneDriveUser = null;
            this.updateODStatus();
            toast("Déconnecté.", "info");
            overlay.remove();
        };
    }
    // ─── UI setup ─────────────────────────────────────────────────────────────────
    setupBarsBtns() {
        const newTopBtn = document.getElementById("btn-new-item-top");
        const settingsBtn = document.getElementById("btn-settings");
        const odTopBtn = document.getElementById("btn-onedrive");
        const newSideBtn = document.getElementById("btn-new-item-sidebar");
        const notesBtn = document.getElementById("btn-notes-open");
        const odSyncBtn = document.getElementById("btn-od-sync");
        const fileInput = document.getElementById("file-input");
        const uploadBtn = document.getElementById("btn-upload");
        [newTopBtn, newSideBtn]
            .forEach(btn => onClick(btn, () => this.openCaseFormModal(null)));
        onClick(settingsBtn, () => this.openSettingsModal());
        onClick(odTopBtn, async () => {
            if (!this.oneDriveUser)
                await this.connectOneDrive();
            else
                await this.refreshCaseFromOneDrive();
        });
        onClick(notesBtn, () => this.openNotesModal());
        onClick(odSyncBtn, () => this.refreshCaseFromOneDrive());
        onClick(uploadBtn, () => fileInput?.click());
        const tabs = qsa(".doc-filter-tab");
        tabs.forEach(tab => onClick(tab, () => {
            setActive(tabs, tab, "active");
            this.docFilter = (tab.dataset.filter ?? "all");
            this.renderDocList();
        }));
        const btns = qsa(".mode-btn[data-mode]");
        btns.forEach(btn => onClick(btn, () => {
            setActive(btns, btn, "active");
            this.activeMode = btn.dataset.mode;
            const hints = {
                analyse: "Posez une question, demandez une analyse du dossier…",
                redaction: "Précisez l'acte à rédiger (courrier, assignation, conclusions, contrat…)",
                modification: "Indiquez le document à modifier et les changements souhaités…",
                note: "Rédigez une correction → sauvegardée dans _notes.json…",
            };
            const ta = document.getElementById("user-input");
            if (!ta)
                return;
            ta.placeholder = hints[this.activeMode];
        }));
        const summBtn = document.getElementById("btn-case-summary");
        onClick(summBtn, async () => {
            const ta = document.getElementById("user-input");
            if (!ta)
                return;
            ta.value =
                "Fais un point complet sur ce dossier : enjeux principaux, risques identifiés, actions restantes, points d'attention prioritaires.";
            await this.sendMessage();
        });
        function onClick(btn, action) {
            if (!btn)
                return;
            btn.onclick = () => action();
        }
    }
    setupTopbar() {
        const newBtn = document.getElementById("btn-new-item-top");
        const setBtn = document.getElementById("btn-settings");
        const odBtn = document.getElementById("btn-onedrive");
        if (newBtn)
            newBtn.onclick = () => this.openCaseFormModal(null);
        if (setBtn)
            setBtn.onclick = () => this.openSettingsModal();
        if (odBtn)
            odBtn.onclick = async () => {
                if (!this.oneDriveUser)
                    await this.connectOneDrive();
                else
                    await this.refreshCaseFromOneDrive();
            };
    }
    setupSidebar() {
        const newSideBtn = document.getElementById("btn-new-item-sidebar");
        const notesBtn = document.getElementById("btn-notes-open");
        const odSyncBtn = document.getElementById("btn-od-sync");
        const fileInput = document.getElementById("file-input");
        const uploadBtn = document.getElementById("btn-upload");
        if (newSideBtn)
            newSideBtn.onclick = () => this.openCaseFormModal(null);
        if (notesBtn)
            notesBtn.onclick = () => this.openNotesModal();
        if (odSyncBtn)
            odSyncBtn.onclick = () => this.refreshCaseFromOneDrive();
        if (uploadBtn && fileInput)
            uploadBtn.onclick = () => fileInput.click();
        const tabs = qsa(".doc-filter-tab");
        for (const tab of tabs) {
            tab.onclick = () => {
                setActive(tabs, tab, "active");
                this.docFilter = (tab.dataset.filter ?? "all");
                this.renderDocList();
            };
        }
        if (fileInput) {
            fileInput.onchange = async () => {
                if (!this.activeCase || !fileInput.files?.length)
                    return;
                if (!this.oneDriveUser) {
                    await this.connectOneDrive();
                    return;
                }
                for (const file of Array.from(fileInput.files)) {
                    try {
                        const ab = await file.arrayBuffer();
                        const meta = makeCaseDocMeta(file);
                        await this.writeCaseFile(this.activeCase.folderName, file.name, ab, meta.mimeType);
                        if (!this.activeCase.documents.some((d) => d.name === file.name))
                            this.activeCase.documents.push(meta);
                        toast(`"${file.name}" uploadé sur OneDrive.`, "success", 2000);
                    }
                    catch (err) {
                        toast(`Erreur : ${err.message}`, "error");
                    }
                }
                fileInput.value = "";
                await this.saveMeta();
                this.renderDocList();
                this.updateDocCount();
            };
        }
    }
    setupModeBar() {
        const btns = qsa(".mode-btn[data-mode]");
        for (const btn of btns) {
            btn.onclick = () => {
                setActive(btns, btn, "active");
                this.activeMode = btn.dataset.mode;
                const hints = {
                    analyse: "Posez une question, demandez une analyse du dossier…",
                    redaction: "Précisez l'acte à rédiger (courrier, assignation, conclusions, contrat…)",
                    modification: "Indiquez le document à modifier et les changements souhaités…",
                    note: "Rédigez une correction → sauvegardée dans _notes.json…",
                };
                const ta = document.getElementById("user-input");
                if (ta)
                    ta.placeholder = hints[this.activeMode];
            };
        }
        const summBtn = document.getElementById("btn-case-summary");
        if (summBtn)
            summBtn.onclick = async () => {
                const ta = document.getElementById("user-input");
                if (ta)
                    ta.value =
                        "Fais un point complet sur ce dossier : enjeux principaux, risques identifiés, actions restantes, points d'attention prioritaires.";
                await this.sendMessage();
            };
    }
    autoResize(ta) {
        ta.style.height = "auto";
        ta.style.height = Math.min(ta.scrollHeight, 160) + "px";
    }
    setupInputArea() {
        const ta = document.getElementById("user-input");
        const sendBtn = document.getElementById("send-btn");
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
        const area = document.getElementById("chat-area");
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
        const area = document.getElementById("chat-area");
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
export class Library extends Cases {
    activeDomain = 'all';
    domainDocs = new Map();
    libMessages = [];
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
    domainLabel(id) {
        if (id === 'all')
            return 'Tous domaines';
        return this.DOMAINS.find(d => d.id === id)?.label ?? id;
    }
    // ─── Path helpers ─────────────────────────────────────────────────────────────
    libDomainFolder(domain) {
        const cap = domain.charAt(0).toUpperCase() + domain.slice(1);
        return `${this.rootFolder}/Bibliotheque/${cap}`;
    }
    libMetaPath(domain) { return `${this.libDomainFolder(domain)}/_meta.json`; }
    // ─── CRUD ─────────────────────────────────────────────────────────────────────
    async readLibMeta(domain) {
        return this.readJson(this.libMetaPath(domain));
    }
    async writeLibMeta(domain, meta) {
        await this.ensureFolder(this.libDomainFolder(domain));
        await this.writeJson(this.libMetaPath(domain), meta);
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
            ? `${this.rootFolder}/Bibliotheque/_conversation.json`
            : `${this.libDomainFolder(this.activeDomain)}/_conversation.json`;
    }
    async listLibFiles(domain) {
        try {
            const items = await this.listFolder(this.libDomainFolder(domain));
            return items.filter(i => i.file && !i.name.startsWith('_'));
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
        await this.writeLibMeta(domain, { domain, documents: docs });
    }
    // ─── Sync from OneDrive ───────────────────────────────────────────────────────
    async syncDomainFromOneDrive(domain) {
        const items = await this.listLibFiles(domain);
        const existing = await this.loadDomainMeta(domain);
        let added = 0;
        for (const item of items) {
            if (!item.file || !isSupported(item.name))
                continue;
            if (existing.some(d => d.name === item.name))
                continue;
            existing.push({ name: item.name, mimeType: item.file.mimeType || 'application/octet-stream',
                sizeBytes: item.size ?? 0, addedAt: Date.now(), tags: [] });
            added++;
        }
        // FIX: save to the correct domain being synced, not activeDomain
        if (added > 0)
            await this.saveDomainMeta(domain, existing);
        return added;
    }
    // ─── Boot ─────────────────────────────────────────────────────────────────────
    async bootLib(container) {
        container.innerHTML = '';
        container.appendChild(this.buildUI());
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
        const wrap = el('div', { className: 'lib-layout' });
        const sidebar = el('div', { className: 'lib-sidebar' });
        const hdr = el('div', { className: 'lib-sidebar__header' });
        hdr.appendChild(el('span', { className: 'lib-sidebar__title', textContent: 'Bibliothèque juridique' }));
        const skillBadge = el('span', { className: 'lib-skill-badge', id: 'lib-skill-badge' });
        toggle(skillBadge, false);
        hdr.appendChild(skillBadge);
        sidebar.appendChild(hdr);
        sidebar.appendChild(el('div', { className: 'lib-domain-pills', id: 'lib-domain-pills' }));
        sidebar.appendChild(el('div', { className: 'lib-doc-list', id: 'lib-doc-list' }));
        const acts = el('div', { className: 'lib-action-row' });
        const fi = el('input', { type: 'file', id: 'lib-file-input', multiple: true });
        fi.accept = '.pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt,.txt,.md';
        fi.style.display = 'none';
        const upBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: '⬆ Ajouter', id: 'btn-lib-upload' });
        upBtn.onclick = () => fi.click();
        const synBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: '⟳ Sync OneDrive', id: 'btn-lib-sync' });
        synBtn.onclick = () => this.syncCurrentDomain();
        acts.append(fi, upBtn, synBtn);
        sidebar.appendChild(acts);
        const main = el('div', { className: 'lib-main' });
        const topbar = el('div', { className: 'lib-topbar' });
        const domLbl = el('span', { className: 'lib-topbar__domain', id: 'lib-active-domain', textContent: 'Tous domaines' });
        const skillLbl = el('span', { className: 'lib-topbar__skills', id: 'lib-skills-top' });
        const clrBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: 'Effacer conversation' });
        clrBtn.onclick = () => this.clearLibConv();
        topbar.append(domLbl, skillLbl, clrBtn);
        const quickArea = el('div', { className: 'lib-quick-prompts', id: 'lib-quick-prompts' });
        const chatArea = el('div', { className: 'chat-area', id: 'lib-chat-area', role: 'log' });
        chatArea.setAttribute('aria-live', 'polite');
        const inputArea = el('div', { className: 'lib-input-area' });
        const ta = el('textarea', { id: 'lib-input', rows: 2, placeholder: 'Interrogez la bibliothèque…' });
        const sendBtn = el('button', { className: 'btn btn--primary', id: 'lib-send-btn' });
        sendBtn.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>';
        inputArea.append(ta, sendBtn);
        main.append(topbar, quickArea, chatArea, inputArea);
        wrap.append(sidebar, main);
        return wrap;
    }
    // ─── Domain navigation ────────────────────────────────────────────────────────
    renderDomainPills() {
        const container = document.getElementById('lib-domain-pills');
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
                pill.appendChild(el('span', { className: 'lib-pill__count', textContent: String(docs.length) }));
            pill.onclick = () => this.switchDomain(d.id);
            container.appendChild(pill);
        }
    }
    async switchDomain(domain) {
        this.activeDomain = domain;
        if (domain !== 'all')
            await this.loadDomainMeta(domain);
        this.libMessages = await this.readLibConversation().catch(() => []);
        this.renderDomainPills();
        this.renderLibDocList();
        this.renderLibChat();
        const lbl = document.getElementById('lib-active-domain');
        if (lbl)
            lbl.textContent = this.domainLabel(domain);
        this.updateQuickPrompts();
    }
    // ─── Render: lib doc list ─────────────────────────────────────────────────────
    renderLibDocList() {
        const list = document.getElementById('lib-doc-list');
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
                const ok = await confirm(`Retirer "${doc.name}" de la bibliothèque ? (Fichier OneDrive conservé.)`);
                if (!ok)
                    return;
                for (const [dom, list] of this.domainDocs) {
                    const idx = list.findIndex(d => d.name === doc.name);
                    if (idx >= 0) {
                        list.splice(idx, 1);
                        await this.saveDomainMeta(dom, list);
                        break;
                    }
                }
                this.renderLibDocList();
                this.renderDomainPills();
                toast('Document retiré de la bibliothèque.', 'info');
            };
            item.append(el('span', { className: 'doc-icon', textContent: mimeIcon(doc.mimeType) }), info, del);
            list.appendChild(item);
        }
    }
    // ─── Render: lib chat ─────────────────────────────────────────────────────────
    renderLibChat() {
        const area = document.getElementById('lib-chat-area');
        if (!area)
            return;
        area.innerHTML = '';
        if (!this.libMessages.length) {
            area.appendChild(el('div', { className: 'empty-state' }, el('div', { className: 'empty-icon', textContent: '📚' }), el('h2', { textContent: 'Bibliothèque juridique' }), el('p', { textContent: 'Sélectionnez un domaine, synchronisez vos documents OneDrive, puis posez votre question.' })));
            return;
        }
        for (const msg of this.libMessages)
            area.appendChild(this.buildLibMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    buildLibMsgEl(msg) {
        const wrap = el('div', { className: `msg msg--${msg.role}` });
        const bubble = el('div', { className: 'msg__bubble' });
        bubble.innerHTML = renderMarkdown(msg.content);
        wrap.append(el('div', { className: 'msg__label', textContent: msg.role === 'user' ? 'Vous' : 'Lex Assistant' }), bubble);
        if (msg.role === 'assistant') {
            const acts = el('div', { className: 'msg__actions' });
            const copy = el('button', { className: 'msg-action-btn', textContent: 'Copier' });
            copy.onclick = () => { navigator.clipboard.writeText(msg.content); toast('Copié.', 'info', 1500); };
            acts.appendChild(copy);
            wrap.appendChild(acts);
        }
        return wrap;
    }
    appendLibMsg(msg) {
        const area = document.getElementById('lib-chat-area');
        if (!area)
            return;
        area.querySelector('.empty-state')?.remove();
        area.appendChild(this.buildLibMsgEl(msg));
        area.scrollTop = area.scrollHeight;
    }
    appendLibTyping() {
        const area = document.getElementById('lib-chat-area');
        const typing = el('div', { className: 'msg msg--assistant', id: 'lib-typing' });
        const bubble = el('div', { className: 'msg__bubble' });
        bubble.append(spinnerEl(), el('span', { textContent: ' Consultation de la bibliothèque…' }));
        typing.append(el('div', { className: 'msg__label', textContent: 'Lex Assistant' }), bubble);
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
        toast('Conversation effacée.', 'info');
    }
    // ─── Send lib message ─────────────────────────────────────────────────────────
    async sendLibMessage() {
        const ta = document.getElementById('lib-input');
        const sendBtn = document.getElementById('lib-send-btn');
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
                if (!this.domainDocs.has(d.id)) {
                    const loaded = await this.loadDomainMeta(d.id);
                    docs.push(...loaded);
                }
            }
        }
        else {
            docs = await this.loadDomainMeta(this.activeDomain);
        }
        const userMsg = { id: uid(), role: 'user', content: text, timestamp: Date.now(), domain: this.activeDomain };
        this.libMessages.push(userMsg);
        this.appendLibMsg(userMsg);
        const typing = this.appendLibTyping();
        sendBtn.disabled = true;
        const history = this.libMessages.slice(-21, -1).map(m => ({ role: m.role, content: m.content }));
        try {
            const response = await callClaudeLib({
                domain: this.activeDomain,
                docs,
                skills: this.skills,
                userMessage: text,
                history,
                readFile: (fileName) => this.readLibFile(this.activeDomain === 'all' ? 'all' : this.activeDomain, fileName),
            });
            typing.remove();
            const asstMsg = { id: uid(), role: 'assistant', content: response, timestamp: Date.now(), domain: this.activeDomain };
            this.libMessages.push(asstMsg);
            this.appendLibMsg(asstMsg);
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
    // ─── Sync current domain ──────────────────────────────────────────────────────
    async syncCurrentDomain() {
        const btn = document.getElementById('btn-lib-sync');
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
            this.renderLibDocList();
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
    // ─── Quick prompts ────────────────────────────────────────────────────────────
    updateQuickPrompts() {
        const area = document.getElementById('lib-quick-prompts');
        if (!area)
            return;
        area.innerHTML = '';
        const prompts = {
            commercial: ["Jurisprudence récente sur la responsabilité du dirigeant pour insuffisance d'actif.",
                "Conditions de validité d'une clause de non-concurrence en droit commercial français.",
                "Règles applicables à la cession de fonds de commerce."],
            fiscal: ["Analyse la jurisprudence sur l'abus de droit fiscal (LPF art. L.64).",
                "Conditions d'application de l'acte anormal de gestion.",
                "Jurisprudence récente sur la déductibilité des charges en IS."],
            social: ["Conditions de validité du licenciement pour motif économique.",
                "Analyse jurisprudentielle du harcèlement moral au travail.",
                "Règles applicables au transfert du contrat de travail (L.1224-1 CT)."],
            civil: ["Jurisprudence récente sur la responsabilité délictuelle.",
                "Conditions de la résolution pour inexécution (C.civ. art. 1224).",
                "Évolutions de la jurisprudence sur le préjudice moral."],
            penal: ["Éléments constitutifs de l'abus de biens sociaux.",
                "Jurisprudence sur la complicité en droit pénal des affaires.",
                "Conditions de mise en cause de la responsabilité pénale des personnes morales."],
            immobilier: ["Régime des baux commerciaux : droit au renouvellement et indemnité d'éviction.",
                "Conditions de l'action en garantie des vices cachés en droit immobilier.",
                "Jurisprudence sur la responsabilité du promoteur immobilier."],
            international: ["Conditions d'applicabilité des conventions fiscales bilatérales.",
                "Jurisprudence sur le centre des intérêts vitaux (CGI art. 4 B).",
                "Règles de conflit de lois en matière successorale (Règl. UE 650/2012)."],
            all: ["Quels sont les documents disponibles dans la bibliothèque ?",
                "Synthèse des principales règles jurisprudentielles sur la responsabilité civile.",
                "Analyse comparative des régimes de responsabilité civile et pénale du dirigeant."],
        };
        const list = prompts[this.activeDomain] ?? prompts.all;
        for (const p of list) {
            const btn = el('button', { className: 'quick-btn', textContent: p });
            btn.onclick = () => { const ta = document.getElementById('lib-input'); if (ta) {
                ta.value = p;
                ta.focus();
            } };
            area.appendChild(btn);
        }
    }
    setupLibInput() {
        const ta = document.getElementById('lib-input');
        const sendBtn = document.getElementById('lib-send-btn');
        if (!ta || !sendBtn)
            return;
        ta.addEventListener('input', () => { ta.style.height = 'auto'; ta.style.height = Math.min(ta.scrollHeight, 160) + 'px'; });
        ta.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            this.sendLibMessage();
        } });
        sendBtn.onclick = () => this.sendLibMessage();
        this.updateQuickPrompts();
    }
    setupLibUpload() {
        const fi = document.getElementById('lib-file-input');
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
                    if (!existing.some(d => d.name === file.name)) {
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
            this.renderLibDocList();
            this.renderDomainPills();
        };
    }
    updateLibSkillIndicator() {
        const badge = document.getElementById('lib-skill-badge');
        const top = document.getElementById('lib-skills-top');
        const n = this.skills.length;
        if (badge) {
            badge.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''}` : '';
            toggle(badge, n > 0);
        }
        if (top) {
            top.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''} actif${n > 1 ? 's' : ''}` : '';
        }
    }
    // ─── Domain picker modal ──────────────────────────────────────────────────────
    pickDomainModal() {
        return new Promise(resolve => {
            const overlay = el('div', { className: 'modal-overlay' });
            const dialog = el('div', { className: 'modal-dialog' });
            dialog.innerHTML = '<h2 class="modal-title">Domaine juridique</h2><p class="modal-subtitle">Dans quel domaine classer ce(s) document(s) ?</p>';
            const grid = el('div', { className: 'domain-grid' });
            for (const d of this.DOMAINS) {
                const btn = el('button', { className: 'domain-btn' });
                btn.innerHTML = `<span class="domain-btn__icon">${d.icon}</span><span>${d.label}</span>`;
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