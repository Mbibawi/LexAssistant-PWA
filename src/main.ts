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

import { Library } from './modules/onedrive.js';
import { el, toggle } from './modules/ui.js';

// ─── Single shared instance ───────────────────────────────────────────────────
const app = new Library();

// ─── Active scenario ──────────────────────────────────────────────────────────

type Scenario = 'selector' | 'dossiers' | 'bibliotheque';
let activeScenario: Scenario = 'selector';

// ─── Boot ─────────────────────────────────────────────────────────────────────

async function boot(): Promise<void> {
  buildShell();
  showSelector();

  // Try silent OneDrive sign-in
  if (app.config.isConfigured()) {
    const signed = await app.isSignedIn();
    if (signed) {
      app.oneDriveUser = app.getSignedInUser();
    }
  }

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('/sw.js').catch(() => {});
  }
}

// ─── Shell (persistent topbar + content area) ─────────────────────────────────

function buildShell(): void {
  document.body.innerHTML = '';

  // Topbar
  const topbar = el('header', { id: 'topbar' });
  topbar.innerHTML = `
    <div class="topbar__left">
      <span class="topbar__logo">⚖</span>
      <span class="topbar__brand">Lex Assistant</span>
      <div id="topbar-breadcrumb" class="topbar__breadcrumb"></div>
    </div>
    <div class="topbar__right">
      <span id="onedrive-status" class="od-status od-status--disconnected">☁ OneDrive</span>
      <span id="skills-badge" class="skills-badge" style="display:none"></span>
      <button id="btn-home"     class="btn btn--ghost btn--sm" style="display:none">⌂ Accueil</button>
      <button id="btn-switch"   class="btn btn--ghost btn--sm" style="display:none"></button>
      <button id="btn-onedrive" class="btn btn--ghost btn--sm">☁ Connexion</button>
      <button id="btn-settings" class="btn btn--ghost btn--sm">⚙</button>
    </div>`;

  // Dynamic topbar slots used by Cases/Library
  topbar.innerHTML += `
    <div id="topbar-case-info" class="topbar__case-info" style="display:none">
      <span id="topbar-case-name"   class="topbar__case"></span>
      <span id="topbar-case-domain" class="topbar__domain"></span>
    </div>`;

  // Content
  const content = el('div', { id: 'app-content' });

  // Toast container
  const toasts = el('div', { id: 'toast-container' });

  document.body.append(topbar, content, toasts);

  // Wire persistent topbar buttons
  (document.getElementById('btn-settings') as HTMLButtonElement).onclick = () => app.openSettingsModal();
  (document.getElementById('btn-onedrive') as HTMLButtonElement).onclick = () => handleOneDriveBtn();
  (document.getElementById('btn-home')     as HTMLButtonElement).onclick = () => { showSelector(); };
}

// ─── Scenario selector ────────────────────────────────────────────────────────

function showSelector(): void {
  activeScenario = 'selector';
  updateTopbarForSelector();

  const content = document.getElementById('app-content')!;
  content.innerHTML = '';
  content.className = 'selector-view';

  content.appendChild(el('div', { className: 'selector-container' },
    el('div', { className: 'selector-header' },
      el('div', { className: 'selector-logo', textContent: '⚖' }),
      el('h1',  { className: 'selector-title', textContent: 'Lex Assistant' }),
      el('p',   { className: 'selector-subtitle', textContent: 'Choisissez votre espace de travail' }),
    ),
    el('div', { className: 'selector-cards' },
      buildScenarioCard({
        id:          'dossiers',
        icon:        '📁',
        title:       'Dossiers',
        description: 'Gérez vos affaires, analysez les pièces, rédigez des actes, conservez vos corrections.',
        features:    ['Pièces PDF, Word, Excel, PowerPoint', 'Notes permanentes (_notes.json)', 'Historique de conversation', 'Génération et sauvegarde .docx'],
        action:      () => mountDossiers(),
      }),
      buildScenarioCard({
        id:          'bibliotheque',
        icon:        '📚',
        title:       'Bibliothèque juridique',
        description: 'Interrogez votre base documentaire thématique — jurisprudence, doctrine, textes.',
        features:    ['8 domaines : Commercial, Fiscal, Social…', 'Synchronisation OneDrive automatique', 'Conversations multi-tours par domaine', 'Recherche et analyse comparative'],
        action:      () => mountBibliotheque(),
      }),
    ),
    buildSettingsShortcut(),
  ));
}

function buildScenarioCard(opts: ScenarioCardOptions): HTMLElement {
  const card = el('div', { className: 'scenario-card', id: `card-${opts.id}` });
  const iconEl = el('div', { className: 'scenario-card__icon', textContent: opts.icon });
  const titleEl = el('h2', { className: 'scenario-card__title', textContent: opts.title });
  const descEl  = el('p',  { className: 'scenario-card__desc',  textContent: opts.description });

  const featList = el('ul', { className: 'scenario-card__features' });

  opts.features
    .forEach(f => featList.appendChild(el('li', { textContent: f })));

  const btn = el('button', { className: 'scenario-card__btn btn btn--primary', textContent: `Ouvrir les ${opts.title}` });
  btn.onclick = opts.action;

  card.append(iconEl, titleEl, descEl, featList, btn);
  return card;
}

function buildSettingsShortcut(): HTMLElement {
  const row = el('div', { className: 'selector-settings-row' });

  const odStatus = app.oneDriveUser
    ? el('span', { className: 'selector-od-status selector-od-status--connected', textContent: `☁ ${app.oneDriveUser}` })
    : el('span', { className: 'selector-od-status selector-od-status--disconnected', textContent: '☁ OneDrive non connecté' });

  const configBtn = el('button', { className: 'btn btn--secondary btn--sm', textContent: '⚙ Paramètres' });
  configBtn.onclick = () => app.openSettingsModal();

  const connectBtn = el('button', { className: 'btn btn--primary btn--sm', textContent: app.oneDriveUser ? '☁ Synchroniser' : '☁ Connecter OneDrive' });
  connectBtn.onclick = () => handleOneDriveBtn();

  row.append(odStatus, configBtn, connectBtn);
  return row;
}

// ─── Mount Dossiers ───────────────────────────────────────────────────────────

async function mountDossiers(): Promise<void> {
  activeScenario = "dossiers";
  updateTopBar("Dossiers", "📚 Bibliothèque", () => mountBibliotheque());
  const content = document.getElementById("app-content")!;
  content.innerHTML = "";
  content.className = "dossiers-view";
  content.innerHTML = buildDossiersHTML();

  // Wire all scenario-specific UI
  app.setupBarsBtns();
  app.setupInputArea();

  app.updateODStatus();

  if (app.oneDriveUser) {
    await app.loadAllCases();
    await app.fetchSkills();
  } else if (app.config.isConfigured()) {
    const signed = await app.isSignedIn();
    if (signed) {
      app.oneDriveUser = app.getSignedInUser();
      app.updateODStatus();
      await app.loadAllCases();
      await app.fetchSkills();
    } else {
      app.showNotConnected();
    }
  } else {
    app.showNotConnected();
  }
}

// ─── Mount Bibliothèque ───────────────────────────────────────────────────────

async function mountBibliotheque(): Promise<void> {
  activeScenario = 'bibliotheque';
  updateTopBar("Bibliothèque", "📁 Dossiers", () => mountDossiers());

  const content = document.getElementById('app-content')!;
  content.innerHTML = '';
  content.className = 'bibliotheque-view';

  // Library builds its own UI into the container
  await app.bootLib(content);

  // Pass already-loaded skills
  app.updateLibSkillIndicator();
  app.updateODStatus();
}

// ─── Topbar state per scenario ────────────────────────────────────────────────

function updateTopbarForSelector(): void {
  const homeBtn   = document.getElementById('btn-home')   as HTMLButtonElement;
  const switchBtn = document.getElementById('btn-switch') as HTMLButtonElement;
  const odBtn     = document.getElementById('btn-onedrive') as HTMLButtonElement;
  const caseInfo  = document.getElementById('topbar-case-info') as HTMLElement;
  const breadcrumb= document.getElementById('topbar-breadcrumb') as HTMLElement;

  toggle(homeBtn,   false);
  toggle(switchBtn, false);
  toggle(caseInfo,  false);
  if (breadcrumb) breadcrumb.textContent = '';
  if (odBtn) odBtn.textContent = app.oneDriveUser ? '☁ ' + app.oneDriveUser : '☁ Connexion';
}

function _accountupdateTopbarForDossiers(): void {
  const homeBtn = document.getElementById("btn-home") as HTMLButtonElement;
  const switchBtn = document.getElementById("btn-switch") as HTMLButtonElement;
  const newBtn = document.getElementById(
    "btn-new-case-top",
  ) as HTMLButtonElement | null;
  const breadcrumb = document.getElementById(
    "topbar-breadcrumb",
  ) as HTMLElement;

  toggle(homeBtn, true);
  toggle(switchBtn, true);
  if (switchBtn) {
    switchBtn.textContent = "📚 Bibliothèque";
    switchBtn.onclick = () => mountBibliotheque();
  }
  if (breadcrumb) breadcrumb.textContent = "Dossiers";
}

function updateTopBar(label: string, switchTo: string, action: Function): void {
  const homeBtn = document.getElementById("btn-home") as HTMLButtonElement;
  const switchBtn = document.getElementById("btn-switch") as HTMLButtonElement;
  const breadcrumb = document.getElementById(
    "topbar-breadcrumb",
  ) as HTMLElement;

  toggle(homeBtn, true);
  toggle(switchBtn, true);
  if (switchBtn) {
    switchBtn.textContent = switchTo;
    switchBtn.onclick = () => action();
  }
  if (breadcrumb) breadcrumb.textContent = label;
}

// ─── OneDrive button handler ──────────────────────────────────────────────────

async function handleOneDriveBtn(): Promise<void> {
  if (!app.oneDriveUser) {
    await app.connectOneDrive();
    // Refresh selector if still on it
    if (activeScenario === 'selector') showSelector();
  } else {
    if (activeScenario === 'dossiers') await app.refreshCaseFromOneDrive();
    else if (activeScenario === 'bibliotheque') await app.syncCurrentDomain();
  }
}

// ─── Dossiers HTML template ───────────────────────────────────────────────────

function buildDossiersHTML(): string {
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
