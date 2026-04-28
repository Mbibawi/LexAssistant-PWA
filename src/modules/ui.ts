/**
 * Lightweight DOM helpers — no framework.
 * Convention: functions named el() create elements, show()/hide() toggle visibility.
 */
import { ids } from "../main.js";


export function byID(id: string) {
  return document.getElementById(id) as HTMLElement | null;
}

/**
 * Create an element.
 * @param tag HTML tag name.
 * @param attrs Attributes and event handlers.
 * @param children Child elements or strings.
 */
export function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, unknown> = {},
  ...children: (HTMLElement | string)[]
): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);

  Object.entries(attrs).forEach(([k, v]) => {
    if (k === "className") {
      e.className = v as string;
    } else if (k === "innerHTML") {
      e.innerHTML = v as string;
    } else if (k === "textContent") {
      e.textContent = v as string;
    } else if (k === "style" && typeof v === "object") {
      Object.assign(e.style, v);
    } else if (k.startsWith("on") && typeof v === "function") {
      // Bonus: Support for event listeners like onClick
      const eventName = k.toLowerCase().substring(2);
      e.addEventListener(eventName, v as EventListener);
    } else if (k in e) {
      // If the property exists on the element (like 'id', 'src', 'href')
      (e as any)[k] = v;
    } else {
      // For everything else, like data-attributes or aria-labels
      e.setAttribute(k, v as string);
    }
  });

  children.forEach(child => {
    if (typeof child === "string") e.appendChild(document.createTextNode(child));
    else e.appendChild(child);

  });
  return e;
}

/**
 * Query the DOM safely.
 * @param selector CSS selector.
 * @param root Optional root element (default: document).
 * @throws Error if element not found.
 */
export function qs<T extends HTMLElement>(
  selector: string,
  root: ParentNode = document,
): T {
  const found = root.querySelector<T>(selector);
  if (!found) throw new Error(`Element not found: ${selector}`);
  return found;
}
/**
 * Query the DOM for multiple elements.
 * @param selector CSS selector.
 * @param root Optional root element (default: document).
 */
export function qsa<T extends HTMLElement>(
  selector: string,
  root: ParentNode = document,
): T[] {
  return Array.from(root.querySelectorAll<T>(selector));
}
export function toggle(e: HTMLElement, visible: boolean) {
  visible ? show(e) : hide(e);
}

/**
 * Set active class on one element in a list.
 * @param items List of elements.
 * @param active Element to activate.
 * @param cls Class name (default: "active").
 */
export function setActive(
  items: HTMLElement[],
  active: HTMLElement,
  cls = "active",
) {
  for (const item of items) item.classList.toggle(cls, item === active);
}

/**
 * Show a temporary notification.
 * @param message Message to show.
 * @param type Type of notification (info, error, success).
 * @param durationMs Duration in milliseconds (default: 3500).
 */
export function toast(
  message: string,
  type: "info" | "error" | "success" = "info",
  durationMs = 3500,
) {
  const container =
    byID(ids.toast) ??
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
export function spinnerEl(): HTMLElement {
  return el("span", {
    className: "spinner",
    "aria-label": "Chargement",
  });
}

export function uid(): string {
  return crypto.randomUUID();
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

export function formatDate(ts: number): string {
  return new Date(ts).toLocaleDateString("fr-FR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
}

export function formatDateTime(ts: number): string {
  return new Date(ts).toLocaleString("fr-FR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

export function confirm(message: string): Promise<boolean> {
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

function show(e: HTMLElement) {
  e.style.display = "";
}

function hide(e: HTMLElement) {
  e.style.display = "none";
}
