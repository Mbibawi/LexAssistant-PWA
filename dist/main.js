/**
 * main.ts — Module 1: Dossiers
 *
 * State lives entirely on OneDrive as JSON files.
 * No IndexedDB. No local persistence except:
 *   - localStorage: API key + OneDrive config (credentials only)
 *   - sessionStorage: MSAL token cache (managed by MSAL)
 *
 * Per-case OneDrive files:
 *   Affaires/<folderName>/_meta.json         ← CaseMeta
 *   Affaires/<folderName>/_notes.json        ← PermanentNote[]
 *   Affaires/<folderName>/_conversation.json ← ChatMessage[]
 *   Affaires/<folderName>/<file>             ← actual documents
 */
import { qs, toggle, } from "./modules/ui.js";
import { Cases, Library } from "./modules/onedrive.js";
// ─── App state (in-memory cache of OneDrive JSON) ─────────────────────────────
// ─── Boot ─────────────────────────────────────────────────────────────────────
async function boot() {
    const odCases = new Cases();
    const library = new Library();
    setupModuleTabs();
    setupTopbar();
    setupSidebar(odCases);
    setupModeBar(odCases);
    setupInputArea(odCases);
    // Try silent OneDrive sign-in
    if (odCases.isConfigured()) {
        const signed = await odCases.isSignedIn();
        if (signed) {
            oneDriveUser = odCases.getSignedInUser();
            updateODStatus(odCases);
            await loadAllCases(odCases);
            await odCases.loadSkills();
        }
        else {
            showNotConnected(odCases);
        }
    }
    else {
        showNotConnected(odCases);
    }
    if ("serviceWorker" in navigator) {
        navigator.serviceWorker.register("/sw.js").catch(() => { });
    }
}
// ─── Module tabs ──────────────────────────────────────────────────────────────
function setupModuleTabs() {
    const odCases = new Cases();
    const lib = new Library();
    qs("#tab-dossiers").onclick = () => switchModule(odCases);
    qs("#tab-bibliotheque").onclick = () => switchModule(lib);
}
function switchModule(mod) {
    setupTopbar(mod);
    setupSidebar(mod);
    setupModeBar(mod);
    setupInputArea(mod);
    const p1 = qs("#panel-dossiers");
    const p2 = qs("#panel-bibliotheque");
    const isCase = mod instanceof Cases;
    const isLib = mod instanceof Library;
    toggle(p1, isCase);
    toggle(p2, isLib);
    qs("#tab-dossiers").classList.toggle("active", isCase);
    qs("#tab-bibliotheque").classList.toggle("active", isLib);
    if (isLib) {
        const c = qs("#panel-bibliotheque");
        if (!c.dataset.booted) {
            mod.boot(c);
            c.dataset.booted = "1";
        }
    }
}
document.addEventListener("DOMContentLoaded", boot);
//# sourceMappingURL=main.js.map