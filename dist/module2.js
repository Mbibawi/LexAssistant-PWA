/**
 * module2.ts — Legal Library (Module 2)
 *
 * State lives on OneDrive:
 *   Bibliotheque/<Domain>/_meta.json         ← LibDomainMeta (document index)
 *   Bibliotheque/<Domain>/_conversation.json ← LibConversationMessage[]
 *   Bibliotheque/<Domain>/<file>             ← actual documents
 *   Bibliotheque/_conversation.json          ← "all domains" conversation
 */
import { callClaudeLib } from './modules/api.js';
import { makeLibDocMeta, mimeIcon, mimeLabel, formatSize, isSupported, uid } from './modules/ingest.js';
import { renderMarkdown } from './modules/markdown.js';
import { el, toggle, toast, confirm, spinnerEl, formatDate } from './modules/ui.js';
import { Library } from './modules/onedrive.js';
// ─── State ────────────────────────────────────────────────────────────────────
let activeDomain = 'all';
// domain → document list (from _meta.json)
const domainDocs = new Map();
let libMessages = [];
let libSkills = [];
// ─── Domain catalogue ─────────────────────────────────────────────────────────
export const DOMAINS = [
    { id: 'commercial', label: 'Commercial', icon: '🏢' },
    { id: 'fiscal', label: 'Fiscal', icon: '💰' },
    { id: 'social', label: 'Social', icon: '👥' },
    { id: 'civil', label: 'Civil', icon: '⚖️' },
    { id: 'penal', label: 'Pénal', icon: '🔒' },
    { id: 'immobilier', label: 'Immobilier', icon: '🏠' },
    { id: 'international', label: 'International', icon: '🌐' },
    { id: 'autre', label: 'Autre', icon: '📚' },
];
function domainLabel(id) {
    if (id === 'all')
        return 'Tous domaines';
    return DOMAINS.find(d => d.id === id)?.label ?? id;
}
export function syncSkills(skills) {
    libSkills = skills;
    updateSkillIndicator();
}
// ─── Boot ─────────────────────────────────────────────────────────────────────
export async function bootLib(container) {
    container.innerHTML = '';
    const lib = new Library();
    container.appendChild(buildUI(lib));
    // Load conversation for current domain
    libMessages = await lib.readLibConversation(activeDomain).catch(() => []);
    renderDomainPills(lib);
    renderLibChat();
    renderLibDocList(lib);
    setupLibInput(lib);
    setupLibUpload(lib);
    updateSkillIndicator();
}
// ─── Persist ─────────────────────────────────────────────────────────────────
async function saveLibConversation(lib) {
    await lib.writeLibConversation(activeDomain, libMessages);
}
async function loadDomainMeta(domain, lib) {
    if (domainDocs.has(domain))
        return domainDocs.get(domain);
    const meta = await lib.readLibMeta(domain).catch(() => null);
    const docs = meta?.documents ?? [];
    domainDocs.set(domain, docs);
    return docs;
}
async function saveDomainMeta(domain, docs, lib) {
    domainDocs.set(domain, docs);
    const meta = { domain, documents: docs };
    await lib.writeLibMeta(domain, meta);
}
// ─── Sync from OneDrive ───────────────────────────────────────────────────────
async function syncDomainFromOneDrive(domain, lib) {
    const items = await lib.listLibFiles(domain);
    const existing = await loadDomainMeta(domain, lib);
    let added = 0;
    for (const item of items) {
        if (!item.file)
            continue;
        if (!isSupported(item.name))
            continue;
        if (existing.some(d => d.name === item.name))
            continue;
        existing.push({
            name: item.name,
            mimeType: item.file.mimeType || 'application/octet-stream',
            sizeBytes: item.size ?? 0,
            addedAt: Date.now(),
            tags: [],
        });
        added++;
    }
    if (added > 0)
        await saveDomainMeta(domain, existing, lib);
    return added;
}
// ─── UI builder ───────────────────────────────────────────────────────────────
function buildUI(lib) {
    const wrap = el('div', { className: 'lib-layout' });
    // ── Sidebar ──
    const sidebar = el('div', { className: 'lib-sidebar' });
    const hdr = el('div', { className: 'lib-sidebar__header' });
    hdr.appendChild(el('span', { className: 'lib-sidebar__title', textContent: 'Bibliothèque juridique' }));
    const skillBadge = el('span', { className: 'lib-skill-badge', id: 'lib-skill-badge' });
    toggle(skillBadge, false);
    hdr.appendChild(skillBadge);
    sidebar.appendChild(hdr);
    // Domain pills
    sidebar.appendChild(el('div', { className: 'lib-domain-pills', id: 'lib-domain-pills' }));
    // Doc list
    sidebar.appendChild(el('div', { className: 'lib-doc-list', id: 'lib-doc-list' }));
    // Action row
    const acts = el('div', { className: 'lib-action-row' });
    const fi = el('input', { type: 'file', id: 'lib-file-input', multiple: true });
    fi.accept = '.pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt,.txt,.md';
    fi.style.display = 'none';
    const upBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: '⬆ Ajouter', id: 'btn-lib-upload' });
    upBtn.onclick = () => fi.click();
    const syncBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: '⟳ Sync OneDrive', id: 'btn-lib-sync' });
    syncBtn.onclick = () => syncCurrentDomain(lib);
    acts.append(fi, upBtn, syncBtn);
    sidebar.appendChild(acts);
    // ── Main ──
    const main = el('div', { className: 'lib-main' });
    const topbar = el('div', { className: 'lib-topbar' });
    const domLbl = el('span', { className: 'lib-topbar__domain', id: 'lib-active-domain', textContent: 'Tous domaines' });
    const skillLbl = el('span', { className: 'lib-topbar__skills', id: 'lib-skills-top' });
    const clrBtn = el('button', { className: 'btn btn--ghost btn--sm', textContent: 'Effacer conversation' });
    clrBtn.onclick = () => clearLibConv(lib);
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
// ─── Domain pills ─────────────────────────────────────────────────────────────
function renderDomainPills(lib) {
    const container = document.getElementById('lib-domain-pills');
    if (!container)
        return;
    container.innerHTML = '';
    const all = el('button', { className: `lib-pill${activeDomain === 'all' ? ' active' : ''}`, textContent: 'Tous' });
    all.onclick = () => switchDomain('all', lib);
    container.appendChild(all);
    for (const d of DOMAINS) {
        const docs = domainDocs.get(d.id) ?? [];
        const pill = el('button', { className: `lib-pill${activeDomain === d.id ? ' active' : ''}` });
        pill.textContent = `${d.icon} ${d.label}`;
        if (docs.length) {
            const cnt = el('span', { className: 'lib-pill__count', textContent: String(docs.length) });
            pill.appendChild(cnt);
        }
        pill.onclick = () => switchDomain(d.id, lib);
        container.appendChild(pill);
    }
}
async function switchDomain(domain, lib) {
    activeDomain = domain;
    // Load docs meta for the selected domain
    if (domain !== 'all')
        await loadDomainMeta(domain, lib);
    // Load conversation for this domain
    libMessages = await lib.readLibConversation(domain).catch(() => []);
    renderDomainPills(lib);
    renderLibDocList(lib);
    renderLibChat();
    updateDomainLabel();
    updateQuickPrompts();
    const el = document.getElementById('lib-active-domain');
    if (el)
        el.textContent = domainLabel(domain);
}
function updateDomainLabel() {
    const lbl = document.getElementById('lib-active-domain');
    if (lbl)
        lbl.textContent = domainLabel(activeDomain);
}
// ─── Doc list ─────────────────────────────────────────────────────────────────
function renderLibDocList(lib) {
    const list = document.getElementById('lib-doc-list');
    if (!list)
        return;
    list.innerHTML = '';
    let docs;
    if (activeDomain === 'all') {
        docs = [];
        for (const [, d] of domainDocs)
            docs.push(...d);
    }
    else {
        docs = domainDocs.get(activeDomain) ?? [];
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
            const ok = await confirm(`Retirer "${doc.name}" de la bibliothèque ? (Le fichier OneDrive est conservé.)`);
            if (!ok)
                return;
            // Find which domain this doc belongs to
            for (const [dom, list] of domainDocs) {
                const idx = list.findIndex(d => d.name === doc.name);
                if (idx >= 0) {
                    list.splice(idx, 1);
                    await saveDomainMeta(dom, list, lib);
                    break;
                }
            }
            renderLibDocList(lib);
            renderDomainPills(lib);
            toast('Document retiré de la bibliothèque.', 'info');
        };
        item.append(el('span', { className: 'doc-icon', textContent: mimeIcon(doc.mimeType) }), info, del);
        list.appendChild(item);
    }
}
// ─── Skills indicator ─────────────────────────────────────────────────────────
function updateSkillIndicator() {
    const badge = document.getElementById('lib-skill-badge');
    const top = document.getElementById('lib-skills-top');
    const n = libSkills.length;
    if (badge) {
        badge.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''}` : '';
        toggle(badge, n > 0);
    }
    if (top) {
        top.textContent = n > 0 ? `${n} skill${n > 1 ? 's' : ''} actif${n > 1 ? 's' : ''}` : '';
    }
}
// ─── Quick prompts ────────────────────────────────────────────────────────────
function updateQuickPrompts() {
    const area = document.getElementById('lib-quick-prompts');
    if (!area)
        return;
    area.innerHTML = '';
    const prompts = {
        commercial: ["Jurisprudence récente sur la responsabilité du dirigeant pour insuffisance d'actif.",
            "Conditions de validité d'une clause de non-concurrence en droit commercial français.",
            "Quelles sont les règles applicables à la cession de fonds de commerce ?"],
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
    const list = prompts[activeDomain] ?? prompts.all;
    for (const p of list) {
        const btn = el('button', { className: 'quick-btn', textContent: p });
        btn.onclick = () => {
            const ta = document.getElementById('lib-input');
            if (ta) {
                ta.value = p;
                ta.focus();
            }
        };
        area.appendChild(btn);
    }
}
// ─── Chat ─────────────────────────────────────────────────────────────────────
function renderLibChat() {
    const area = document.getElementById('lib-chat-area');
    if (!area)
        return;
    area.innerHTML = '';
    if (!libMessages.length) {
        area.appendChild(el('div', { className: 'empty-state' }, el('div', { className: 'empty-icon', textContent: '📚' }), el('h2', { textContent: 'Bibliothèque juridique' }), el('p', { textContent: 'Sélectionnez un domaine, synchronisez vos documents OneDrive, puis posez votre question.' })));
        return;
    }
    for (const msg of libMessages)
        area.appendChild(buildLibMsgEl(msg));
    area.scrollTop = area.scrollHeight;
}
function buildLibMsgEl(msg) {
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
function appendLibMsg(msg) {
    const area = document.getElementById('lib-chat-area');
    if (!area)
        return;
    area.querySelector('.empty-state')?.remove();
    area.appendChild(buildLibMsgEl(msg));
    area.scrollTop = area.scrollHeight;
}
function appendLibTyping() {
    const area = document.getElementById('lib-chat-area');
    const typing = el('div', { className: 'msg msg--assistant', id: 'lib-typing' });
    const bubble = el('div', { className: 'msg__bubble' });
    bubble.append(spinnerEl(), el('span', { textContent: ' Consultation de la bibliothèque…' }));
    typing.append(el('div', { className: 'msg__label', textContent: 'Lex Assistant' }), bubble);
    area.appendChild(typing);
    area.scrollTop = area.scrollHeight;
    return typing;
}
async function clearLibConv(lib) {
    const ok = await confirm('Effacer l\'historique de la bibliothèque pour ce domaine ?');
    if (!ok)
        return;
    libMessages = [];
    await lib.writeLibConversation(activeDomain, []);
    renderLibChat();
    toast('Conversation effacée.', 'info');
}
// ─── Send ─────────────────────────────────────────────────────────────────────
async function sendLibMessage(lib) {
    const ta = document.getElementById('lib-input');
    const sendBtn = document.getElementById('lib-send-btn');
    if (!ta || !sendBtn)
        return;
    const text = ta.value.trim();
    if (!text)
        return;
    ta.value = '';
    // Collect docs for this domain (or all)
    let docs;
    if (activeDomain === 'all') {
        docs = [];
        for (const [, d] of domainDocs)
            docs.push(...d);
        // Also load all domains we haven't loaded yet
        for (const d of DOMAINS) {
            if (!domainDocs.has(d.id)) {
                const loaded = await loadDomainMeta(d.id, lib);
                docs.push(...loaded);
            }
        }
    }
    else {
        docs = await loadDomainMeta(activeDomain, lib);
    }
    const userMsg = {
        id: uid(), role: 'user', content: text, timestamp: Date.now(), domain: activeDomain
    };
    libMessages.push(userMsg);
    appendLibMsg(userMsg);
    const typing = appendLibTyping();
    sendBtn.disabled = true;
    // Last 20 messages as history (excluding current)
    const history = libMessages.slice(-21, -1).map(m => ({ role: m.role, content: m.content }));
    try {
        const response = await callClaudeLib({
            domain: activeDomain,
            docs,
            skills: libSkills,
            userMessage: text,
            history,
        });
        typing.remove();
        const asstMsg = {
            id: uid(), role: 'assistant', content: response, timestamp: Date.now(), domain: activeDomain
        };
        libMessages.push(asstMsg);
        appendLibMsg(asstMsg);
        await saveLibConversation(lib);
    }
    catch (err) {
        typing.remove();
        toast(err.message, 'error', 6000);
        libMessages.pop();
    }
    finally {
        sendBtn.disabled = false;
        ta.focus();
    }
}
// ─── Sync current domain ──────────────────────────────────────────────────────
async function syncCurrentDomain(lib) {
    const btn = document.getElementById('btn-lib-sync');
    if (btn) {
        btn.disabled = true;
        btn.textContent = '⟳ Sync…';
    }
    try {
        let total = 0;
        if (activeDomain === 'all') {
            for (const d of DOMAINS)
                total += await syncDomainFromOneDrive(d.id, lib);
        }
        else {
            total = await syncDomainFromOneDrive(activeDomain, lib);
        }
        renderLibDocList(lib);
        renderDomainPills(lib);
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
// ─── Input + upload setup ─────────────────────────────────────────────────────
function setupLibInput(lib) {
    const ta = document.getElementById('lib-input');
    const sendBtn = document.getElementById('lib-send-btn');
    if (!ta || !sendBtn)
        return;
    ta.addEventListener('input', () => { ta.style.height = 'auto'; ta.style.height = Math.min(ta.scrollHeight, 160) + 'px'; });
    ta.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendLibMessage(lib);
    } });
    sendBtn.onclick = () => sendLibMessage(lib);
    updateQuickPrompts();
}
function setupLibUpload(lib) {
    const fi = document.getElementById('lib-file-input');
    if (!fi)
        return;
    fi.onchange = async () => {
        if (!fi.files?.length)
            return;
        const domain = await pickDomainModal();
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
                // Upload to OneDrive library folder
                await lib.writeLibFile(domain, file.name, ab, meta.mimeType);
                // Register in domain meta
                const existing = await loadDomainMeta(domain, lib);
                if (!existing.some(d => d.name === file.name)) {
                    existing.push(meta);
                    await saveDomainMeta(domain, existing, lib);
                }
                toast(`"${file.name}" ajouté à la bibliothèque ${domainLabel(domain)}.`, 'success');
            }
            catch (err) {
                toast(`Erreur : ${err.message}`, 'error');
            }
        }
        fi.value = '';
        renderLibDocList(lib);
        renderDomainPills(lib);
    };
}
// ─── Domain picker modal ──────────────────────────────────────────────────────
function pickDomainModal() {
    return new Promise(resolve => {
        const overlay = el('div', { className: 'modal-overlay' });
        const dialog = el('div', { className: 'modal-dialog' });
        dialog.innerHTML = '<h2 class="modal-title">Domaine juridique</h2><p class="modal-subtitle">Dans quel domaine classer ce(s) document(s) ?</p>';
        const grid = el('div', { className: 'domain-grid' });
        for (const d of DOMAINS) {
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
//# sourceMappingURL=module2.js.map