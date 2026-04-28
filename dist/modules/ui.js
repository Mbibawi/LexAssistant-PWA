/**
 * Lightweight DOM helpers — no framework.
 * Convention: functions named el() create elements, show()/hide() toggle visibility.
 */
import { ids } from "../main.js";
export function byID(id) {
    return document.getElementById(id);
}
/**
 * Create an element.
 * @param tag HTML tag name.
 * @param attrs Attributes and event handlers.
 * @param children Child elements or strings.
 */
export function el(tag, attrs = {}, ...children) {
    const e = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
        if (k === "className") {
            e.className = v;
        }
        else if (k === "innerHTML") {
            e.innerHTML = v;
        }
        else if (k === "textContent") {
            e.textContent = v;
        }
        else if (k === "style" && typeof v === "object") {
            Object.assign(e.style, v);
        }
        else if (k.startsWith("on") && typeof v === "function") {
            // Bonus: Support for event listeners like onClick
            const eventName = k.toLowerCase().substring(2);
            e.addEventListener(eventName, v);
        }
        else if (k in e) {
            // If the property exists on the element (like 'id', 'src', 'href')
            e[k] = v;
        }
        else {
            // For everything else, like data-attributes or aria-labels
            e.setAttribute(k, v);
        }
    });
    children.forEach(child => {
        if (typeof child === "string")
            e.appendChild(document.createTextNode(child));
        else
            e.appendChild(child);
    });
    return e;
}
/**
 * Query the DOM safely.
 * @param selector CSS selector.
 * @param root Optional root element (default: document).
 * @throws Error if element not found.
 */
export function qs(selector, root = document) {
    const found = root.querySelector(selector);
    if (!found)
        throw new Error(`Element not found: ${selector}`);
    return found;
}
/**
 * Query the DOM for multiple elements.
 * @param selector CSS selector.
 * @param root Optional root element (default: document).
 */
export function qsa(selector, root = document) {
    return Array.from(root.querySelectorAll(selector));
}
export function toggle(e, visible) {
    visible ? show(e) : hide(e);
}
/**
 * Set active class on one element in a list.
 * @param items List of elements.
 * @param active Element to activate.
 * @param cls Class name (default: "active").
 */
export function setActive(items, active, cls = "active") {
    for (const item of items)
        item.classList.toggle(cls, item === active);
}
/**
 * Show a temporary notification.
 * @param message Message to show.
 * @param type Type of notification (info, error, success).
 * @param durationMs Duration in milliseconds (default: 3500).
 */
export function toast(message, type = "info", durationMs = 3500) {
    const container = byID(ids.toast) ??
        (() => {
            const d = el("div", { id: ids.toast });
            document.body.appendChild(d);
            return d;
        })();
    const t = el("div", {
        className: `toast toast--${type}`,
        textContent: message,
    });
    container.appendChild(t);
    requestAnimationFrame(() => t.classList.add("toast--visible"));
    setTimeout(() => {
        t.classList.remove("toast--visible");
        setTimeout(() => t.remove(), 350);
    }, durationMs);
}
export function spinnerEl() {
    return el("span", {
        className: "spinner",
        "aria-label": "Chargement",
    });
}
export function uid() {
    return crypto.randomUUID();
}
export function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}
export function formatDate(ts) {
    return new Date(ts).toLocaleDateString("fr-FR", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
    });
}
export function formatDateTime(ts) {
    return new Date(ts).toLocaleString("fr-FR", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
    });
}
export function confirm(message) {
    return new Promise((resolve) => {
        const overlay = el("div", { className: "modal-overlay" });
        const dialog = el("div", { className: "modal-dialog" });
        const msg = el("p", { className: "modal-msg", textContent: message });
        const btnRow = el("div", { className: "modal-btns" });
        const btnCancel = el("button", {
            className: "btn btn--secondary",
            textContent: "Annuler",
        });
        const btnOk = el("button", {
            className: "btn btn--danger",
            textContent: "Confirmer",
        });
        btnCancel.onclick = () => {
            overlay.remove();
            resolve(false);
        };
        btnOk.onclick = () => {
            overlay.remove();
            resolve(true);
        };
        btnRow.append(btnCancel, btnOk);
        dialog.append(msg, btnRow);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
    });
}
function show(e) {
    e.style.display = "";
}
function hide(e) {
    e.style.display = "none";
}
//# sourceMappingURL=ui.js.map