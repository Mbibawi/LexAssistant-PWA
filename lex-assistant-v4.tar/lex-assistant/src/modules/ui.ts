/**
 * Lightweight DOM helpers — no framework.
 * Convention: functions named el() create elements, show()/hide() toggle visibility.
 */

export function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Partial<HTMLElementTagNameMap[K]> & { className?: string; innerHTML?: string; textContent?: string } = {},
  ...children: (HTMLElement | string)[]
): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'className') e.className = v as string;
    else if (k === 'innerHTML') e.innerHTML = v as string;
    else if (k === 'textContent') e.textContent = v as string;
    else (e as unknown as Record<string, unknown>)[k] = v;
  }
  for (const c of children) {
    if (typeof c === 'string') e.appendChild(document.createTextNode(c));
    else e.appendChild(c);
  }
  return e;
}

export function qs<T extends HTMLElement>(selector: string, root: ParentNode = document): T {
  const found = root.querySelector<T>(selector);
  if (!found) throw new Error(`Element not found: ${selector}`);
  return found;
}

export function qsa<T extends HTMLElement>(selector: string, root: ParentNode = document): T[] {
  return Array.from(root.querySelectorAll<T>(selector));
}

export function show(e: HTMLElement) { e.style.display = ''; }
export function hide(e: HTMLElement) { e.style.display = 'none'; }
export function toggle(e: HTMLElement, visible: boolean) { visible ? show(e) : hide(e); }

export function setActive(items: HTMLElement[], active: HTMLElement, cls = 'active') {
  for (const item of items) item.classList.toggle(cls, item === active);
}

export function toast(message: string, type: 'info' | 'error' | 'success' = 'info', durationMs = 3500) {
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

export function confirm(message: string): Promise<boolean> {
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

export function spinnerEl(): HTMLElement {
  return el('span', { className: 'spinner', 'aria-label': 'Chargement' } as unknown as Partial<HTMLSpanElement>);
}

export function uid(): string {
  return crypto.randomUUID();
}

export function formatDate(ts: number): string {
  return new Date(ts).toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export function formatDateTime(ts: number): string {
  return new Date(ts).toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}
