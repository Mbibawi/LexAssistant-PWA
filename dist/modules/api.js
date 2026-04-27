import { toBase64 } from './ingest.js';
const API_URL = 'https://api.anthropic.com/v1/messages';
const MODEL = "claude-sonnet-4-6";
let _apiKey = "";
export function getStoredKey() { return _apiKey; }
export function setStoredKey(k) { _apiKey = k.trim(); }
export function clearStoredKey() { _apiKey = ''; }
// ─── MIME types Claude accepts natively ──────────────────────────────────────
const NATIVE_MIMES = new Set([
    'application/pdf',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'text/plain', 'text/html', 'text/markdown',
]);
function docPart(name, mime, base64) {
    if (NATIVE_MIMES.has(mime)) {
        return { type: 'document', source: { type: 'base64', media_type: mime, data: base64 }, title: name };
    }
    return { type: 'text', text: `[Fichier joint : ${name} — format non lu nativement]` };
}
// ─── System prompts ───────────────────────────────────────────────────────────
function modeInstruction(mode) {
    switch (mode) {
        case 'analyse':
            return "MODE ANALYSE. Expert avocat français hautement compétent et spécialisé dans les questions de droit soulevées par le dossier. Signale les risques non demandés. Ne sur - simplifie pas. N'invente jamais jurisprudence, textes ou doctrine.";
        case 'redaction':
            return "MODE RÉDACTION. Rédige un document juridique complet, sans préambule.";
        case 'modification':
            return "MODE MODIFICATION. Identifie les passages à modifier, justifie par le droit applicable, produis la version modifiée intégrale. Marque les changements avec [MODIFIÉ : …].";
        case 'note':
            return "MODE NOTE PERMANENTE. Confirme la prise en compte, résume ce qui est retenu, explique l'impact sur les prochaines analyses. Priorité absolue sur toutes tes inférences futures.";
    }
}
export function buildCaseSystem(caseName, caseDomain, notes, skills, mode) {
    const noteBlock = notes.length
        ? `\n\n## CORRECTIONS PERMANENTES (priorité absolue)\n${notes.map((n, i) => `${i + 1}. ${n.content}`).join('\n')}`
        : '';
    const skillBlock = skills.length
        ? `\n\n## INSTRUCTIONS MÉTIER (_Skills/)\n${skills.map((s) => `### ${s.name}\n${s.content}`).join("\n\n")}`
        : "";
    return `Tu es Lex Assistant, avocat collaborateur et assistant juridique personnel de Maître Mina Bibawi, avocat au Barreau de Paris (toque B0976).
  
  ## DOSSIER ACTIF
  Intitulé : ${caseName}
  Domaine : ${caseDomain}${noteBlock}${skillBlock}
  
  ## RÈGLES PERMANENTES
  - Ne sur - simplifie pas.
  - N'invente jamais jurisprudence, textes ou doctrine.
  - Ne compte jamais sur tes connaissances uniquement.
  - Ne fonde jamais ton analyse ou tes conclusions sur un texte légal, une jurisprudence ou une source doctrinale, sans en avoir vérifié l'existence et analysé le contenu exact.
  - Pour les textes légaux, vérfie systématiquement la version applicable au moment des faits ou de la situtation juridique analysée.
  - Signale l'évolution de la règle de droit ou du texte applicable postérieure à la date des faits.
  - Indique systématiquement la référence textes et jurisprudence (références exactes et complètes).
  - Inclus systématiquement un extrait du texte légale (article, alinéa et etc.), ou de la source doctrinale sur laquelle tu t'es appuyé.
- Inclus pour la jurisprudence un extrait de la motivation de la décision soutenant ton interprétation et ton analyse de sa portée.
- Intègre les éléments de fait depuis les pièces fournies.
- Inclus à la fin une liste exhaustive des pièces invoquées dans ton texte. 
- Chaque fois que tu mentionne un fait ou un élément tiré d'une pièce, Inclus une référence à la pièce invoquée. Exemple: 'en date du [date], Monsieur X a assigné la société Y en liquidation judiciaire (Pièce n°3)'.
- Emploi un style juridique très soigné et de haut niveau professionnel dans la rédaction.
- La forme et le contenu d'une qualité attendue d'un avocat hautement spécialisé et compétent dans le domaine juridique concernée.
- Respecte les styles de mise en forme indiqués par l'utilisateur.
- Signale proactivement tout risque juridique ou fiscal, même non demandé.
- Français juridique de haut niveau, niveau cabinet parisien d'affaires.
- Si un élément manque dans les pièces, le signaler explicitement.
- Ne compte jamais aveuglement sur les traductions fournie dans le dossier des pièces en langue étrangère. Analyse systématiquement la version originale de la pièce. Restitue ta propre traduction plus précise ou plus claire du contenu dans ton exposition de la portée de la pièce.
- Excel : analyse les données chiffrées et implications juridiques/fiscales.
- PowerPoint : analyse le contenu substantiel.

## ${modeInstruction(mode)}`;
}
export function buildLibSystem(domain, skills) {
    const labels = {
        commercial: 'droit commercial', fiscal: 'droit fiscal', social: 'droit social',
        civil: 'droit civil', penal: 'droit pénal', immobilier: 'droit immobilier',
        international: 'droit international', autre: 'droit général', all: 'tous domaines juridiques'
    };
    const skillBlock = skills.length
        ? `\n\n## INSTRUCTIONS MÉTIER\n${skills.map(s => `### ${s.name}\n${s.content}`).join('\n\n')}`
        : '';
    return `Tu es Lex Assistant, expert en ${labels[domain] ?? "droit français"}, au service de Maître Mina Bibawi, avocat au Barreau de Paris.
Tu as accès à une bibliothèque juridique thématique fournie avec chaque question.${skillBlock}

## RÈGLES
- Précision académique et pratique de haut niveau.
- Cite toujours la source exacte (arrêt, article, auteur, nom du document, page) issue des documents fournis.
- Structure avec des titres clairs.
- Si la question dépasse les documents, le signaler explicitement.
- Propose des analyses comparatives et chronologies jurisprudentielles.`;
}
// ─── File fetching — injected by the caller ───────────────────────────────────
// api.ts no longer imports from onedrive.ts directly; instead the caller
// provides a readFile function appropriate to the scenario.
async function buildDocParts(docs, readFile) {
    const parts = [];
    for (const doc of docs) {
        try {
            const buf = await readFile(doc.name);
            parts.push(docPart(doc.name, doc.mimeType, toBase64(buf)));
        }
        catch {
            parts.push({ type: 'text', text: `[Fichier "${doc.name}" inaccessible sur OneDrive]` });
        }
    }
    return parts;
}
export async function callClaudeCase(opts) {
    const key = getStoredKey();
    if (!key)
        throw new Error('Clé API manquante. Configurez-la dans les paramètres.');
    const system = buildCaseSystem(opts.caseName, opts.caseDomain, opts.notes, opts.skills, opts.mode);
    const docParts = await buildDocParts(opts.docs, opts.readFile);
    const content = [...docParts, { type: 'text', text: opts.userMessage }];
    return fetchClaude(key, { model: MODEL, max_tokens: 4096, system,
        messages: [{ role: 'user', content: content }] });
}
export async function callClaudeLib(opts) {
    const key = getStoredKey();
    if (!key)
        throw new Error('Clé API manquante. Configurez-la dans les paramètres.');
    const system = buildLibSystem(opts.domain, opts.skills);
    const docParts = await buildDocParts(opts.docs, opts.readFile);
    const messages = [
        ...opts.history.map(h => ({ role: h.role, content: h.content })),
        { role: 'user', content: [...docParts, { type: 'text', text: opts.userMessage }] }
    ];
    return fetchClaude(key, { model: MODEL, max_tokens: 4096, system, messages });
}
async function fetchClaude(key, body) {
    const resp = await fetch(API_URL, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'x-api-key': key,
            'anthropic-version': '2023-06-01',
            'anthropic-dangerous-direct-browser-calls': 'true'
        },
        body: JSON.stringify(body)
    });
    if (!resp.ok) {
        const e = await resp.json().catch(() => ({ error: { message: resp.statusText } }));
        throw new Error(`API Claude : ${e.error?.message ?? resp.statusText}`);
    }
    const data = await resp.json();
    return data.content.map(b => b.text).join('');
}
//# sourceMappingURL=api.js.map