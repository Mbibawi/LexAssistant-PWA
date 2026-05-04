import { oneDrive } from '../main.js';
// ─── MIME types Claude accepts natively ──────────────────────────────────────
const NATIVE_MIMES = new Set([
    'application/pdf',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'text/plain', 'text/html', 'text/markdown',
]);
// ─── Helpers ──────────────────────────────────────────────────────────────────
/** Convert any string to base64 (UTF-8 safe) */
function strToBase64(text) {
    return btoa(encodeURIComponent(text));
}
/** Decode a base64 string back to UTF-8 text */
export function base64ToStr(b64) {
    return decodeURIComponent(atob(b64));
}
/** Fingerprint a document for duplicate detection */
function docFingerprint(doc) {
    return [doc.name, doc.mimeType, doc.sizeBytes, doc.addedAt].join('|');
}
// ─── System prompts ───────────────────────────────────────────────────────────
function permanentInstructions() {
    return `## RÈGLES PERMANENTES
- Ne sur-simplifie pas.
- N'invente jamais jurisprudence, textes ou doctrine.
- Ne compte jamais sur tes propres connaissances uniquement.
- Ne fonde jamais ton analyse ou tes conclusions sur un texte légal, une jurisprudence ou une source doctrinale, sans en avoir vérifié l'existence et analysé le contenu exact.
- Pour les textes légaux, vérifie systématiquement la version applicable au moment des faits ou de la situation juridique analysée.
- Signale l'évolution de la règle de droit ou du texte applicable postérieure à la date des faits.
- Indique systématiquement la référence des textes légaux et de la jurisprudence (références exactes et complètes).
- Inclus systématiquement un extrait du texte légal (article, alinéa, etc.), ou de la source doctrinale sur laquelle tu t'es appuyé.
- Inclus pour la jurisprudence un extrait de la motivation de la décision soutenant ton interprétation et ton analyse de sa portée.
- Intègre les éléments de fait depuis les pièces fournies.
- Inclus à la fin une liste exhaustive des pièces et sources invoquées dans ton texte.
- Chaque fois que tu cites ou mentionnes un fait ou un élément tiré d'une pièce, inclus une référence à la pièce invoquée après la citation. Exemple : 'en date du [date], Monsieur X a assigné la société Y en liquidation judiciaire (Pièce n°3)'.
- Emploie un style juridique très soigné et de haut niveau professionnel dans la rédaction.
- Français juridique de haut niveau, niveau cabinet parisien d'affaires.
- La qualité, la forme et le contenu des documents générés doit être celle d'un avocat hautement spécialisé et compétent dans le domaine juridique concerné.
- Respecte les styles de mise en forme indiqués par l'utilisateur.
- Signale proactivement tout risque juridique ou fiscal, même non demandé.
- Si un élément manque dans les pièces, le signaler explicitement.
- Ne compte jamais aveuglément sur les traductions fournies dans le dossier des pièces en langue étrangère. Analyse systématiquement la version originale de la pièce. Restitue ta propre traduction plus précise ou plus claire du contenu dans ton exposition de la portée de la pièce.
- Excel : analyse les données chiffrées et implications juridiques/fiscales.
- PowerPoint : analyse le contenu substantiel.
- Structure avec des titres clairs.`;
}
function modeInstruction(mode) {
    switch (mode) {
        case 'analyse':
            return 'MODE ANALYSE. Expert avocat français hautement compétent et spécialisé dans les questions de droit soulevées par le dossier. Signale les risques non demandés. Ne sur-simplifie pas. N\'invente jamais jurisprudence, textes ou doctrine.';
        case 'redaction':
            return 'MODE RÉDACTION. Rédige un document juridique complet, sans préambule.';
        case 'modification':
            return 'MODE MODIFICATION. Identifie les passages à modifier, justifie par le droit applicable, produis la version modifiée intégrale. Marque les changements avec [MODIFIÉ : …].';
        case 'note':
            return 'MODE NOTE PERMANENTE. Confirme la prise en compte, résume ce qui est retenu, explique l\'impact sur les prochaines analyses. Priorité absolue sur toutes tes inférences futures.';
    }
}
export function buildCaseSystem(caseName, caseDomain, notes, skills, mode) {
    const noteBlock = notes.length
        ? `\n\n## CORRECTIONS PERMANENTES (priorité absolue)\n${notes.map((n, i) => `${i + 1}. ${n.content}`).join('\n')}`
        : '';
    const skillBlock = skills.length
        ? `\n\n## INSTRUCTIONS MÉTIER (_Skills/)\n${skills.map((s) => `### ${s.name}\n${s.content}`).join('\n\n')}`
        : '';
    return `Tu es Lex Assistant, avocat collaborateur et assistant juridique personnel de Maître Mina Bibawi, avocat au Barreau de Paris (toque B0976).

## DOSSIER ACTIF
Intitulé : ${caseName}
Domaine : ${caseDomain}${noteBlock}${skillBlock}

${permanentInstructions()}

## ${modeInstruction(mode)}`;
}
export function buildLibSystem(domain, skills) {
    const labels = {
        commercial: 'droit commercial', fiscal: 'droit fiscal', social: 'droit social',
        civil: 'droit civil', penal: 'droit pénal', immobilier: 'droit immobilier',
        international: 'droit international', autre: 'droit général', all: 'tous domaines juridiques',
    };
    const skillBlock = skills.length
        ? `\n\n## INSTRUCTIONS MÉTIER\n${skills.map((s) => `### ${s.name}\n${s.content}`).join('\n\n')}`
        : '';
    return `Tu es Lex Assistant, avocat collaborateur hautement spécialisé expert en ${labels[domain] ?? 'droit français'}, au service de Maître Mina Bibawi, avocat au Barreau de Paris.
Tu as accès à une bibliothèque juridique thématique fournie avec chaque question.${skillBlock}

## RÈGLES
- Précision académique et pratique de haut niveau.
- Cite toujours la source exacte (arrêt, article, auteur, nom du document, page) issue des documents fournis.
- Si la question dépasse les documents, le signaler explicitement.
- Propose des analyses comparatives et chronologies jurisprudentielles.
${permanentInstructions()}`;
}
// ─── ClaudeAPI ────────────────────────────────────────────────────────────────
export class ClaudeAPI {
    // ─── Constants ────────────────────────────────────────────────────────────────
    PATH = 'v1/messages';
    MODEL = 'claude-sonnet-4-6';
    // ─── Core fetch — routes through the Google Cloud Function proxy ──────────
    /**
     * All Claude API calls go through gFetch with the GCF proxy URL.
     * gFetch handles auth headers for Graph; for the proxy we pass rawBody=true
     * and inject the anthropic-version header ourselves since gFetch won't add it.
     */
    async callProxy(messages, api = 'claude') {
        const resp = await oneDrive.callClaudeProxy(api, this.PATH, JSON.stringify(messages), "2024-06-01");
        if (!resp.ok) {
            const e = await resp.json().catch(() => ({ error: { message: resp.statusText } }));
            throw new Error(`Claude API : ${e.error?.message ?? resp.statusText}`);
        }
        return resp.json();
    }
    claudeBody(max, messages, system) {
        const body = { model: this.MODEL, max_tokens: max, messages };
        if (system)
            body.system = system;
        return body;
    }
    extractText(data) {
        return data.content.map((b) => b.text ?? '').join('');
    }
    // ─── Doc parts builder ────────────────────────────────────────────────────
    async buildDocParts(folderName, docs, readFile) {
        const parts = [];
        for (const doc of docs) {
            try {
                const buf = await readFile(folderName, doc.name);
                parts.push(docPart(doc.name, doc.mimeType, this.toBase64(buf)));
            }
            catch {
                parts.push({
                    type: 'text',
                    text: `[Fichier "${doc.name}" inaccessible sur OneDrive]`
                });
            }
        }
        return parts;
        function docPart(name, mime, base64) {
            if (NATIVE_MIMES.has(mime)) {
                return {
                    type: 'document',
                    source: {
                        type: 'base64',
                        media_type: mime,
                        data: base64
                    },
                    title: name
                };
            }
            ;
            return {
                type: 'text',
                text: `[Fichier joint : ${name} — format non lu nativement]`
            };
        }
    }
    // ─── Duplicate detection ──────────────────────────────────────────────────
    /**
     * Compares incoming docs against those already in the knowledge base meta.
     * Returns the list of docs that are new (not already fingerprinted).
     * If duplicates are found, prompts the user to confirm resending them.
     */
    async filterNewDocs(incoming, existing) {
        const existingPrints = new Set(existing.map(docFingerprint));
        const duplicates = incoming.filter((d) => existingPrints.has(docFingerprint(d)));
        const fresh = incoming.filter((d) => !existingPrints.has(docFingerprint(d)));
        if (duplicates.length > 0) {
            const names = duplicates.map((d) => `• ${d.name}`).join('\n');
            const ok = confirm(`${duplicates.length} fichier(s) déjà inclus dans la base de connaissance existante :\n\n${names}\n\nVoulez-vous les renvoyer à Claude quand même ?`);
            return ok ? incoming : fresh;
        }
        return incoming;
    }
    // ─── Knowledge base — Cases ───────────────────────────────────────────────
    /**
     * Generates a markdown knowledge base from case documents and saves it to
     * OneDrive with a versioned filename: _kb_YYYY-MM-DD_HHmm.md
     * Returns the OneDrive path where it was saved.
     */
    async buildCaseKnowledgeBase(meta, readFile, existingKbDocs = [], appendMode = false) {
        const docsToSend = await this.filterNewDocs(meta.documents, existingKbDocs);
        if (!docsToSend.length) {
            throw new Error('Aucun nouveau document à analyser.');
        }
        const docParts = await this.buildDocParts(meta.folderName, docsToSend, readFile);
        const prompt = appendMode
            ? `Tu complètes une base de connaissance juridique existante avec de nouveaux documents.
Produis un complément en markdown structuré, couvrant uniquement les nouveaux éléments apportés par les documents fournis.
Utilise les mêmes conventions de titres et de structure que la base existante.
Ne répète pas ce qui est déjà connu. Commence directement sans préambule.`
            : `Analyse ces documents juridiques et produis une base de connaissance structurée en markdown.
Couvre : parties, dates, obligations, clauses clés, données chiffrées, risques identifiés, relations entre documents.
Structure avec des titres clairs (## et ###). Commence directement sans préambule.`;
        return await this.getMarkdown(docParts, prompt);
    }
    /**
     * Loads the most recent knowledge base file for a case folder.
     * Returns null if none exists.
     */
    async loadLatestCaseKb(folderPath, items, readFile) {
        //const items = await listFiles(folderPath).catch(() => [] as GraphDriveItem[]);
        const kbFiles = items
            .filter((i) => i.file && i.name.startsWith('_kb_') && i.name.endsWith('.md'))
            .sort((a, b) => b.name.localeCompare(a.name)); // lexicographic = chronological
        if (!kbFiles.length)
            return null;
        const latest = kbFiles[0];
        const buf = await readFile(`${folderPath}/${latest.name}`);
        return new TextDecoder().decode(buf);
    }
    // ─── Knowledge base — Library ─────────────────────────────────────────────
    async buildLibKnowledgeBase(domain, docs, readFile, existingKbDocs = [], appendMode = false) {
        const docsToSend = await this.filterNewDocs(docs, existingKbDocs);
        if (!docsToSend.length)
            throw new Error('Aucun nouveau document à analyser.');
        const docParts = await this.buildDocParts(domain, docsToSend, readFile);
        const prompt = appendMode
            ? `Tu complètes une base de connaissance juridique thématique existante (domaine : ${domain}) avec de nouveaux documents. Produis un complément en markdown structuré. Ne répète pas l'existant. Commence directement sans préambule.`
            : `Analyse ces documents juridiques (domaine : ${domain}) et produis une base de connaissance thématique structurée en markdown. Couvre : sources, règles clés, jurisprudence, doctrine, évolutions récentes. Structure avec des titres clairs. Commence directement sans préambule.`;
        return await this.getMarkdown(docParts, prompt);
    }
    async getMarkdown(docParts, prompt) {
        const data = await this.callProxy(this.claudeBody(8000, [{
                role: 'user',
                content: [...docParts, { type: 'text', text: prompt }],
            }]));
        return this.extractText(data);
    }
    // ─── Case conversation ────────────────────────────────────────────────────
    async callClaudeCase(folderName, opts) {
        const system = {
            type: 'text',
            text: buildCaseSystem(opts.caseName, opts.caseDomain, opts.notes, opts.skills, opts.mode),
        };
        // If a knowledge base is available, inject it as a cached document
        // instead of re-sending all raw files — token optimization
        let content;
        if (opts.knowledgeBase) {
            content = [
                {
                    type: 'document',
                    source: { type: 'base64', media_type: 'text/markdown', data: strToBase64(opts.knowledgeBase) },
                    title: 'Base de connaissance du dossier',
                },
                { type: 'text', text: opts.userMessage },
            ];
        }
        else {
            const docParts = await this.buildDocParts(folderName, opts.docs, opts.readFile);
            content = [...docParts, { type: 'text', text: opts.userMessage }];
        }
        const data = await this.callProxy(this.claudeBody(4096, [{ role: 'user', content }], system));
        return this.extractText(data);
    }
    // ─── Library conversation ─────────────────────────────────────────────────
    async callClaudeLib(folderName, opts) {
        const { domain, skills, knowledgeBase, history, userMessage, docs, readFile } = opts;
        const system = {
            type: 'text',
            text: buildLibSystem(domain, skills),
        };
        let firstUserContent;
        if (knowledgeBase) {
            firstUserContent = [
                {
                    type: 'document',
                    source: {
                        type: 'base64',
                        media_type: 'text/markdown',
                        data: strToBase64(knowledgeBase)
                    },
                    title: `Bibliothèque juridique — ${domain}`,
                }
            ];
        }
        else {
            firstUserContent = await this.buildDocParts(folderName, docs, readFile);
        }
        // Rebuild history: inject docs only in the first user turn
        const messages = history.length
            ? [
                {
                    role: 'user',
                    content: [...firstUserContent, history[0].content]
                },
                ...history,
                {
                    role: 'user',
                    content: { type: 'text', text: userMessage }
                },
            ]
            : [{
                    role: 'user',
                    content: [...firstUserContent, { type: 'text', text: userMessage }]
                }];
        const data = await this.callProxy(this.claudeBody(4096, messages, system));
        return this.extractText(data);
    }
    // ─── DOCX generation via Claude ──────────────────────────────────────────
    // Claude returns markdown; the caller handles local DOCX conversion.
    // This method exists so Cases/Library can request a structured redaction
    // and receive clean markdown ready for generateDocx().
    async requestRedaction(prompt, system, contextParts = []) {
        const data = await this.callProxy(this.claudeBody(6000, [
            {
                role: 'user',
                content: [...contextParts, { type: 'text', text: prompt }]
            }
        ], { type: 'text', text: system }));
        return this.extractText(data);
    }
    toBase64(buffer) {
        const bytes = new Uint8Array(buffer);
        let bin = '';
        const chunk = 8192;
        for (let i = 0; i < bytes.byteLength; i += chunk) {
            bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
        }
        return btoa(bin);
    }
    /** Build a timestamp suffix for knowledge base filenames: YYYY-MM-DD_HHmm */
    kbTimestamp() {
        const d = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
    }
}
//# sourceMappingURL=api.js.map