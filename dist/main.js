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
import { byID, el, toggle } from './modules/ui.js';
import { DocumentContext as OfficeAddin } from './modules/word-addin.js';
document.addEventListener('DOMContentLoaded', boot);
// ─── Initiale single shared instances ───────────────────────────────────────────────────
const cases = new Cases("Affaires");
const library = new Library("Bibliotheque");
export const odSingleton = new OneDriveAuth();
const officeAddin = new OfficeAddin(cases);
export const ids = {
    btnBuildKb: "btn-build-kb",
    btnCaseSync: "btn-case-sync",
    btnCaseDelete: "btn-case-delete",
    btnCaseSummary: "btn-case-summary",
    btnCaseRefresh: "btn-case-refresh",
    btnCaseUpload: "btn-case-upload",
    btnLibSync: "btn-lib-sync",
    btnLibUpload: "btn-lib-upload",
    btnNewSidebar: "btn-new-sidebar",
    btnNotesOpen: "btn-notes-open",
    btnSwitch: "btn-switch",
    btnOneDrive: "btn-onedrive",
    btnOdSync: "btn-od-sync",
    btnUpload: "btn-upload",
    chatArea: "chat-area",
    chatMessages: "chat-messages",
    caseList: "case-list",
    content: "app-content",
    contextMenu: "context-menu",
    docList: "doc-list",
    docCount: "doc-count",
    docFilter: 'doc-filter-tab',
    domainList: "domain-list",
    domainContent: "domain-content",
    fileCancel: "file-cancel",
    fileDomain: "file-domain",
    fileFolder: "file-folder",
    fileInput: "file-input",
    fileName: "file-name",
    fileSave: "file-save",
    fileStatus: "file-status",
    home: "btn-home",
    inputArea: "input-area",
    libActiveDomain: "lib-active-domain",
    libChatArea: "lib-chat-area",
    libDocList: "lib-doc-list",
    libDomainPills: "lib-domain-pills",
    libSkillsBadge: "lib-skills-badge",
    libSkillsTop: "lib-skills-top",
    logo: "topbar-logo",
    mainLayout: "main-layout",
    modeBar: "mode-bar",
    noteBar: "note-bar",
    noteContent: "note-content",
    oneDriveStatus: "ondrive-status",
    quickPrompts: "quick-prompts",
    sendBtn: "send-btn",
    settings: "btn-settings",
    settingsOverlay: "settings-overlay",
    sidebar: "sidebar",
    skillsBadge: "skills-badge",
    title: "topbar-brand",
    toast: "toast-container",
    topbar: "topbar",
    topBarBreadcrumb: "topbar-breadcrumb",
    topBarCaseInfo: "topbar-case-info",
    topBarCaseName: "topbar-case-name",
    topBarCaseDomain: "topbar-case-domain",
    typing: "typing",
    userInput: "user-input",
    workspace: "workspace",
};
// ─── Boot ─────────────────────────────────────────────────────────────────────
async function boot() {
    await officeAddin.init();
    buildShell();
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('/sw.js').catch(() => { });
    }
}
// ─── Shell (persistent topbar + content area) ─────────────────────────────────
function buildShell() {
    document.body.innerHTML = "";
    // Topbar Left
    const topLeft = el("div", {
        className: "topbar_Left",
    });
    topLeft.append(el("span", { id: ids.logo, className: "topbar_logo", innerText: "⚖" }), el("span", {
        id: ids.title,
        className: "topbar_brand",
        innerText: "Lex Assistant",
    }), el("div", {
        id: ids.topBarBreadcrumb,
        className: "topbar_breadcrumb",
    }));
    // TopBar Right
    const topRight = el("div", {
        className: "topbar_Right",
    });
    topRight.append(el("span", {
        id: ids.btnOneDrive,
        className: "od-status od-status--disconnected",
        innerText: "☁ OneDrive",
    }), el("span", {
        id: ids.skillsBadge,
        className: "skills-badge",
        innerText: "",
        style: { display: "none" },
    }), el("button", {
        id: ids.home,
        className: "btn btn--ghost btn--sm",
        innerText: "⌂ Accueil",
        style: { display: "none" }
    }), el("button", {
        id: ids.btnSwitch,
        className: "btn btn--ghost btn--sm",
        style: { display: "none" }
    }), el("button", {
        id: ids.btnOneDrive,
        className: "btn btn--ghost btn--sm",
        innerText: "☁ Connexion",
    }), el("button", {
        id: ids.settings,
        className: "btn btn--ghost btn--sm",
        innerText: "⚙",
    }));
    // Dynamic topbar slots used by Cases/Library
    const caseInfo = el("div", {
        id: ids.topBarCaseInfo,
        className: "topbar_case-info",
        style: { display: "none" }
    });
    caseInfo.append(el("span", {
        id: ids.topBarCaseName,
        className: "topbar_case",
    }), el("span", {
        id: ids.topBarCaseDomain,
        className: "topbar_domain",
    }));
    const header = el("header", { id: ids.topbar });
    header.append(topLeft, topRight, caseInfo);
    document.body.append(header, 
    // Content
    el("div", { id: ids.content }), 
    // Toast container
    el("div", { id: ids.toast }));
    // Wire persistent topbar buttons
    byID(ids.home).onclick = showSelector;
    showSelector(); //We show the selector ui once the shell is built
}
// ─── Scenario selector ────────────────────────────────────────────────────────
function showSelector() {
    const content = byID(ids.content);
    const logo = byID(ids.logo).innerText;
    const title = byID(ids.title).innerText;
    content.innerHTML = "";
    content.className = "selector-view";
    const selectorContainer = el("div", { className: "selector-container" }), selectorHeader = el("div", { className: "selector-header" }), selectorCards = el("div", { className: "selector-cards" });
    content.appendChild(selectorContainer);
    selectorContainer.append(selectorHeader, selectorCards);
    buildSettingsShortcut(selectorContainer);
    selectorHeader.append(el("div", { className: "selector-logo", textContent: logo }), el("h1", { className: "selector-title", textContent: title }), el("p", {
        className: "selector-subtitle",
        textContent: "Choisissez votre espace de travail",
    }));
    selectorCards.append(buildScenarioCard({
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
        action: () => mountDossiers(),
    }), buildScenarioCard({
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
        action: () => mountBibliotheque(),
    }));
}
function buildScenarioCard(opts) {
    const card = el('div', { className: 'scenario-card', id: `card-${opts.id}` });
    card.append(el('div', { className: 'scenario-card_icon', textContent: opts.icon }), el('h2', { className: 'scenario-card_itle', textContent: opts.title }), el('p', { className: 'scenario-card_desc', textContent: opts.description }));
    const featList = el('ul', { className: 'scenario-card_features' });
    opts.features
        .forEach(f => featList.appendChild(el('li', { textContent: f })));
    const btn = el('button', { className: 'scenario-card_btn btn btn--primary', textContent: `Ouvrir les ${opts.title}` });
    btn.onclick = opts.action;
    card.append(featList, btn);
    return card;
}
function buildSettingsShortcut(container) {
    const row = el("div", { className: "selector-settings-row" });
    const userName = odSingleton.userName;
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
    const odConnect = el("button", {
        className: "btn btn--primary btn--sm",
        textContent: userName ? "☁ Synchroniser" : "☁ Connecter OneDrive",
    });
    //configBtn.onclick = () => cases.openSettingsModal();//!add a general onedrive settings modal to be opedn
    odConnect.onclick = async () => {
        if (!odSingleton.account)
            return;
        const btnOd = byID(ids.btnOneDrive);
        btnOd.textContent = "☁ Connexion en cours...";
        await odSingleton.signIn();
        const userName = cases.userName;
        btnOd.textContent = userName ? "☁ " + userName : "☁ Connexion";
        if (userName)
            odStatus.textContent = "☁ " + userName;
    };
    //row.append(odStatus, configBtn, odConnect);
    row.append(odStatus, odConnect);
    container.appendChild(row);
}
// ─── Mount Dossiers ───────────────────────────────────────────────────────────
async function mountDossiers() {
    await updateTopBar("Dossiers", "📚 Bibliothèque", () => mountBibliotheque());
    cases.showUI();
}
// ─── Mount Bibliothèque ───────────────────────────────────────────────────────
async function mountBibliotheque() {
    await updateTopBar("Bibliothèque", "📁 Dossiers", () => mountDossiers());
    // Library builds its own UI into the container
    await library.showUI();
}
async function updateTopBar(label, switchTo, action) {
    const homeBtn = byID(ids.home);
    const switchBtn = byID(ids.btnSwitch);
    const odBtn = byID(ids.btnOneDrive);
    const breadcrumb = byID(ids.topBarBreadcrumb);
    toggle(homeBtn, true);
    toggle(switchBtn, true);
    if (switchBtn) {
        switchBtn.textContent = switchTo;
        switchBtn.onclick = () => action();
    }
    if (breadcrumb)
        breadcrumb.textContent = label;
    // Try silent OneDrive sign-in
    if (!odBtn)
        return;
    if (!odSingleton.userName)
        await odSingleton.signIn();
    const userName = odSingleton.userName;
    odBtn.textContent = userName ? "☁ " + userName : "☁ Connexion";
}
//# sourceMappingURL=main.js.map