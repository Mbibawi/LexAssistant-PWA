/**
 * Lightweight DOM helpers — no framework.
 * Convention: functions named el() create elements, show()/hide() toggle visibility.
 */
export function el(tag, attrs = {}, ...children) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
        if (k === 'className')
            e.className = v;
        else if (k === 'innerHTML')
            e.innerHTML = v;
        else if (k === 'textContent')
            e.textContent = v;
        else
            e[k] = v;
    }
    for (const c of children) {
        if (typeof c === 'string')
            e.appendChild(document.createTextNode(c));
        else
            e.appendChild(c);
    }
    return e;
}
export function qs(selector, root = document) {
    const found = root.querySelector(selector);
    if (!found)
        throw new Error(`Element not found: ${selector}`);
    return found;
}
export function qsa(selector, root = document) {
    return Array.from(root.querySelectorAll(selector));
}
export function show(e) { e.style.display = ''; }
export function hide(e) { e.style.display = 'none'; }
export function toggle(e, visible) { visible ? show(e) : hide(e); }
export function setActive(items, active, cls = 'active') {
    for (const item of items)
        item.classList.toggle(cls, item === active);
}
export function toast(message, type = 'info', durationMs = 3500) {
    const container = document.getElementById('toast-container') ?? (() => {
        const d = el('div', { id: 'toast-container' });
        document.body.appendChild(d);
        return d;
    })();
    const t = el('div', { className: `toast toast--${type}`, textContent: message });
    container.appendChild(t);
    requestAnimationFrame(() => t.classList.add('toast--visible'));
    setTimeout(() => {
        t.classList.remove('toast--visible');
        setTimeout(() => t.remove(), 350);
    }, durationMs);
}
export function confirm(message) {
    return new Promise(resolve => {
        const overlay = el('div', { className: 'modal-overlay' });
        const dialog = el('div', { className: 'modal-dialog' });
        const msg = el('p', { className: 'modal-msg', textContent: message });
        const btnRow = el('div', { className: 'modal-btns' });
        const btnCancel = el('button', { className: 'btn btn--secondary', textContent: 'Annuler' });
        const btnOk = el('button', { className: 'btn btn--danger', textContent: 'Confirmer' });
        btnCancel.onclick = () => { overlay.remove(); resolve(false); };
        btnOk.onclick = () => { overlay.remove(); resolve(true); };
        btnRow.append(btnCancel, btnOk);
        dialog.append(msg, btnRow);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
    });
}
export function spinnerEl() {
    return el('span', { className: 'spinner', 'aria-label': 'Chargement' });
}
export function uid() {
    return crypto.randomUUID();
}
export function formatDate(ts) {
    return new Date(ts).toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
export function formatDateTime(ts) {
    return new Date(ts).toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}
//# sourceMappingURL=ui.js.map