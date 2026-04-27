/**
 * main.ts — entry point
 *
 * Shows a scenario selector on first load.
 * User picks "Dossiers" or "Bibliothèque"; the chosen scenario
 * mounts its own UI into #app-content.
 *
 * A single Library instance is created at boot and reused for both
 * scenarios (Library extends Cases, so it carries all state).
 */
import { Cases, Library, OneDriveAuth } from "./modules/onedrive.js";
import { el, toggle } from './modules/ui.js';
// ─── Initiale single shared instances ───────────────────────────────────────────────────
export const oneDrive = new OneDriveAuth();
const cases = new Cases();
const library = new Library();
// ─── Active scenario ──────────────────────────────────────────────────────────
// ─── Boot ─────────────────────────────────────────────────────────────────────
async function boot() {
    buildShell();
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('/sw.js').catch(() => { });
    }
}
// ─── Shell (persistent topbar + content area) ─────────────────────────────────
function buildShell() {
    document.body.innerHTML = "";
    // Topbar Left
    const logo = el("span", { className: "topbar__logo", innerText: "⚖" });
    const topBrand = el("span", {
        className: "topbar__brand",
        innerText: "Lex Assistant",
    });
    const bread = el("div", {
        id: "topbar-breadcrumb",
        className: "topbar__breadcrumb",
    });
    const topLeft = el("div", {
        className: "topbar__Left",
    });
    [logo, topBrand, bread].forEach((e) => topLeft.appendChild(e));
    // TopBar Right
    const odStatus = el("span", {
        id: "onedrive-status",
        className: "od-status od-status--disconnected",
        innerText: "☁ OneDrive",
    });
    const skills = el("span", {
        id: "skills-badge",
        className: "skills-badge",
        innerText: "",
    });
    skills.style.display = "none";
    const home = el("button", {
        id: "btn-home",
        className: "btn btn--ghost btn--sm",
        innerText: "⌂ Accueil",
    });
    home.style.display = "none";
    const switchBtn = el("button", {
        id: "btn-switch",
        className: "btn btn--ghost btn--sm",
    });
    switchBtn.style.display = "none";
    const connection = el("button", {
        id: "btn-onedrive",
        className: "btn btn--ghost btn--sm",
        innerText: "☁ Connexion",
    });
    const settings = el("button", {
        id: "btn-settings",
        className: "btn btn--ghost btn--sm",
        innerText: "⚙",
    });
    const topRight = el("div", {
        className: "topbar__Right",
    });
    [odStatus, skills, home, switchBtn, connection, settings].forEach((e) => topRight.appendChild(e));
    // Dynamic topbar slots used by Cases/Library
    const caseInfo = el("div", {
        id: "topbar-case-info",
        className: "topbar__case-info",
    });
    caseInfo.style.display = "none";
    const caseName = el("span", {
        id: "topbar-case-name",
        className: "topbar__case",
    });
    const caseDomain = el("span", {
        id: "topbar-case-domain",
        className: "topbar__domain",
    });
    [caseName, caseDomain].forEach((el) => caseInfo.appendChild(el));
    const header = el("header", { id: "topbar" });
    [topLeft, topRight, caseInfo].forEach((el) => header.appendChild(el));
    // Content
    const content = el("div", { id: "app-content" });
    // Toast container
    const toasts = el("div", { id: "toast-container" });
    document.body.append(header, content, toasts);
    // Wire persistent topbar buttons
    home.onclick = () => showSelector(settings, content, logo.innerText, topBrand.innerText);
    home.click(); //We show the selector ui once the shell is built
}
// ─── Scenario selector ────────────────────────────────────────────────────────
function showSelector(settings, content, logo, title) {
    content.innerHTML = "";
    content.className = "selector-view";
    const selectorContainer = el("div", { className: "selector-container" }), selectorHeader = el("div", { className: "selector-header" }), selectorCards = el("div", { className: "selector-cards" });
    [selectorHeader, selectorCards].forEach((el) => selectorContainer.appendChild(el));
    content.appendChild(selectorContainer);
    const { configBtn, syncBtn } = buildSettingsShortcut(selectorContainer, cases.userName());
    const sLogo = el("div", { className: "selector-logo", textContent: logo }), sTitle = el("h1", { className: "selector-title", textContent: title }), subTitle = el("p", {
        className: "selector-subtitle",
        textContent: "Choisissez votre espace de travail",
    });
    [sLogo, sTitle, subTitle].forEach((el) => selectorHeader.appendChild(el));
    const dossiers = buildScenarioCard({
        id: "dossiers",
        icon: "📁",
        title: "Dossiers",
        description: "Gérez vos affaires, analysez les pièces, rédigez des actes, conservez vos corrections.",
        features: [
            "Pièces PDF, Word, Excel, PowerPoint",
            "Notes permanentes (_notes.json)",
            "Historique de conversation",
            "Génération et sauvegarde .docx",
        ],
        action: () => mountDossiers(configBtn, syncBtn),
    });
    const library = buildScenarioCard({
        id: "bibliotheque",
        icon: "📚",
        title: "Bibliothèque juridique",
        description: "Interrogez votre base documentaire thématique — jurisprudence, doctrine, textes.",
        features: [
            "8 domaines : Commercial, Fiscal, Social…",
            "Synchronisation OneDrive automatique",
            "Conversations multi-tours par domaine",
            "Recherche et analyse comparative",
        ],
        action: () => mountBibliotheque(configBtn, syncBtn),
    });
    [dossiers, library].forEach((el) => selectorCards.appendChild(el));
}
function buildScenarioCard(opts) {
    const card = el('div', { className: 'scenario-card', id: `card-${opts.id}` });
    const iconEl = el('div', { className: 'scenario-card__icon', textContent: opts.icon });
    const titleEl = el('h2', { className: 'scenario-card__title', textContent: opts.title });
    const descEl = el('p', { className: 'scenario-card__desc', textContent: opts.description });
    const featList = el('ul', { className: 'scenario-card__features' });
    opts.features
        .forEach(f => featList.appendChild(el('li', { textContent: f })));
    const btn = el('button', { className: 'scenario-card__btn btn btn--primary', textContent: `Ouvrir les ${opts.title}` });
    btn.onclick = opts.action;
    card.append(iconEl, titleEl, descEl, featList, btn);
    return card;
}
function buildSettingsShortcut(container, userName) {
    const row = el("div", { className: "selector-settings-row" });
    const odStatus = userName
        ? el("span", {
            className: "selector-od-status selector-od-status--connected",
            textContent: `☁ ${userName}`,
        })
        : el("span", {
            className: "selector-od-status selector-od-status--disconnected",
            textContent: "☁ OneDrive non connecté",
        });
    const configBtn = el("button", {
        className: "btn btn--secondary btn--sm",
        textContent: "⚙ Paramètres",
    });
    const syncBtn = el("button", {
        className: "btn btn--primary btn--sm",
        textContent: userName ? "☁ Synchroniser" : "☁ Connecter OneDrive",
    });
    row.append(odStatus, configBtn, syncBtn);
    container.appendChild(row);
    return { configBtn, syncBtn };
}
// ─── Mount Dossiers ───────────────────────────────────────────────────────────
async function mountDossiers(configBtn, syncBtn) {
    await updateTopbarForSelector(cases);
    configBtn.onclick = () => cases.openSettingsModal();
    syncBtn.onclick = async () => await cases.refreshCaseFromOneDrive();
    updateTopBar("Dossiers", "📚 Bibliothèque", () => mountBibliotheque(configBtn, syncBtn));
    const content = document.getElementById("app-content");
    content.innerHTML = "";
    content.className = "dossiers-view";
    content.innerHTML = buildDossiersHTML();
    // Wire all scenario-specific UI
    cases.setupBarsBtns();
    cases.setupInputArea();
    cases.updateODStatus();
    if (cases.userName()) {
        await cases.loadAllSubFolders();
        await cases.fetchSkills();
    }
    else {
        cases.showNotConnected();
    }
}
// ─── Mount Bibliothèque ───────────────────────────────────────────────────────
async function mountBibliotheque(configBtn, connectBtn) {
    await updateTopbarForSelector(library);
    configBtn.onclick = () => library.openSettingsModal();
    connectBtn.onclick = async () => await library.syncCurrentDomain();
    updateTopBar("Bibliothèque", "📁 Dossiers", () => mountDossiers(configBtn, connectBtn));
    library.setupBarsBtns();
    const content = document.getElementById("app-content");
    content.innerHTML = "";
    content.className = "bibliotheque-view";
    // Library builds its own UI into the container
    await library.bootLib(content);
    // Pass already-loaded skills
    library.updateODStatus();
    if (library.userName()) {
        await library.loadAllSubFolders();
        await library.fetchSkills();
    }
    else {
        library.showNotConnected();
    }
}
// ─── Topbar state per scenario ────────────────────────────────────────────────
async function updateTopbarForSelector(app) {
    const homeBtn = document.getElementById("btn-home");
    const switchBtn = document.getElementById("btn-switch");
    const odBtn = document.getElementById("btn-onedrive");
    const caseInfo = document.getElementById("topbar-case-info");
    const breadcrumb = document.getElementById("topbar-breadcrumb");
    toggle(homeBtn, false);
    toggle(switchBtn, false);
    toggle(caseInfo, false);
    if (breadcrumb)
        breadcrumb.textContent = "";
    // Try silent OneDrive sign-in
    if (oneDrive.isConfigured) {
        const userName = await oneDrive.isSignedIn();
        if (odBtn)
            odBtn.textContent = userName ? "☁ " + userName : "☁ Connexion";
    }
}
function updateTopBar(label, switchTo, action) {
    const homeBtn = document.getElementById("btn-home");
    const switchBtn = document.getElementById("btn-switch");
    const breadcrumb = document.getElementById("topbar-breadcrumb");
    toggle(homeBtn, true);
    toggle(switchBtn, true);
    if (switchBtn) {
        switchBtn.textContent = switchTo;
        switchBtn.onclick = () => action();
    }
    if (breadcrumb)
        breadcrumb.textContent = label;
}
// ─── OneDrive button handler ──────────────────────────────────────────────────
// ─── Dossiers HTML template ───────────────────────────────────────────────────
function buildDossiersHTML() {
    return `
    <div id="main-layout">
      <aside id="sidebar">
        <div class="sidebar__section-title">Dossiers</div>
        <div id="case-list"></div>
        <button id="btn-new-case-sidebar" class="btn btn--ghost btn--dashed">+ Nouveau dossier</button>
        <div class="sidebar__divider"></div>
        <div class="sidebar__section-title">
          Pièces <span id="doc-count" class="doc-count">0 pièce</span>
        </div>
        <div class="doc-filter-tabs">
          <button class="doc-filter-tab active" data-filter="all">Tout</button>
          <button class="doc-filter-tab" data-filter="piece">Pièces</button>
          <button class="doc-filter-tab" data-filter="jurisprudence">Juris.</button>
          <button class="doc-filter-tab" data-filter="doctrine">Doctrine</button>
          <button class="doc-filter-tab" data-filter="redige">Rédigés</button>
        </div>
        <div id="doc-list" class="doc-list"></div>
        <div class="sidebar__upload-row">
          <input type="file" id="file-input" multiple
            accept=".pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt,.txt,.md,.rtf"
            style="display:none"/>
          <button id="btn-upload"  class="btn btn--ghost btn--sm">⬆ Upload</button>
          <button id="btn-od-sync" class="btn btn--ghost btn--sm">☁ Sync</button>
          <button id="btn-notes-open" class="btn btn--ghost btn--sm">📌 Notes</button>
        </div>
      </aside>

      <div id="workspace">
        <div id="mode-bar">
          <span class="mode-bar__label">Mode :</span>
          <button class="mode-btn active" data-mode="analyse">Analyse</button>
          <button class="mode-btn" data-mode="redaction">Rédaction</button>
          <button class="mode-btn" data-mode="modification">Modification</button>
          <button class="mode-btn" data-mode="note">Note permanente</button>
          <div class="mode-bar__spacer"></div>
          <button id="btn-case-summary" class="btn btn--ghost btn--sm">Point dossier ↗</button>
        </div>
        <div id="note-bar" class="note-bar" style="display:none"></div>
        <div id="chat-area" class="chat-area" role="log" aria-live="polite"></div>
        <div id="quick-prompts">
          <button class="quick-btn" data-prompt="Analyse les risques juridiques et fiscaux du dossier et liste les points d'attention prioritaires.">Risques du dossier</button>
          <button class="quick-btn" data-prompt="Rédige une mise en demeure formelle à la partie adverse sur la base des pièces du dossier.">Mise en demeure</button>
          <button class="quick-btn" data-prompt="Fais une synthèse chronologique des faits pertinents issus des pièces du dossier.">Chronologie des faits</button>
          <button class="quick-btn" data-prompt="Analyse les données chiffrées des tableaux Excel et leurs implications juridiques et fiscales.">Analyse chiffrée</button>
          <button class="quick-btn" data-prompt="Rédige un mémorandum juridique complet sur le point de droit central avec jurisprudence applicable.">Mémorandum juridique</button>
        </div>
        <div id="input-area">
          <textarea id="user-input" rows="2"
            placeholder="Posez une question, demandez la rédaction d'un acte, ou donnez une instruction…"
            aria-label="Message"></textarea>
          <button id="send-btn" class="btn btn--primary" aria-label="Envoyer">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
          </button>
        </div>
      </div>
    </div>`;
}
document.addEventListener('DOMContentLoaded', boot);
//# sourceMappingURL=main.js.map