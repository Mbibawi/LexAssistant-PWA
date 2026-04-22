import type {
  AnthropicRequest, AnthropicResponse,
  CaseDocumentMeta, LibDocumentMeta,
  PermanentNote, WorkMode, LibDomain
} from '../types.js';
import { toBase64 } from './ingest.js';

const API_URL = 'https://api.anthropic.com/v1/messages';
const MODEL   = 'claude-sonnet-4-20250514';
const LS_KEY  = 'lex_api_key';

export function getStoredKey(): string { return localStorage.getItem(LS_KEY) ?? ''; }
export function setStoredKey(k: string): void { localStorage.setItem(LS_KEY, k.trim()); }
export function clearStoredKey(): void { localStorage.removeItem(LS_KEY); }

// ─── MIME types Claude accepts natively ──────────────────────────────────────

const NATIVE_MIMES = new Set([
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'text/plain', 'text/html', 'text/markdown',
]);

type ContentPart =
  | { type: 'text'; text: string }
  | { type: 'document'; source: { type: 'base64'; media_type: string; data: string }; title?: string };

function docPart(name: string, mime: string, base64: string): ContentPart {
  if (NATIVE_MIMES.has(mime)) {
    return { type: 'document', source: { type: 'base64', media_type: mime, data: base64 }, title: name };
  }
  return { type: 'text', text: `[Fichier joint : ${name} — format non lu nativement]` };
}

// ─── System prompts ───────────────────────────────────────────────────────────

function modeInstruction(mode: WorkMode): string {
  switch (mode) {
    case 'analyse':
      return 'MODE ANALYSE. Expert juriste français. Cite systématiquement les textes et jurisprudence (références exactes). Signale les risques non demandés. Ne sur-simplifie pas.';
    case 'redaction':
      return "MODE RÉDACTION. Rédige un document juridique complet, sans préambule. Toutes les mentions légales. Éléments de fait intégrés depuis les pièces. Termine par les signatures.";
    case 'modification':
      return "MODE MODIFICATION. Identifie les passages à modifier, justifie par le droit applicable, produis la version modifiée intégrale. Marque les changements avec [MODIFIÉ : …].";
    case 'note':
      return "MODE NOTE PERMANENTE. Confirme la prise en compte, résume ce qui est retenu, explique l'impact sur les prochaines analyses. Priorité absolue sur toutes tes inférences futures.";
  }
}

export function buildCaseSystem(
  caseName: string, caseDomain: string,
  notes: PermanentNote[], skills: Array<{ name: string; content: string }>,
  mode: WorkMode
): string {
  const noteBlock = notes.length
    ? `\n\n## CORRECTIONS PERMANENTES (priorité absolue)\n${notes.map((n, i) => `${i+1}. ${n.content}`).join('\n')}`
    : '';
  const skillBlock = skills.length
    ? `\n\n## INSTRUCTIONS MÉTIER (_Skills/)\n${skills.map(s => `### ${s.name}\n${s.content}`).join('\n\n')}`
    : '';
  return `Tu es Lex Assistant, assistant juridique personnel de Maître Mina Bibawi, avocat au Barreau de Paris (toque B0976).

## DOSSIER ACTIF
Intitulé : ${caseName}
Domaine : ${caseDomain}${noteBlock}${skillBlock}

## RÈGLES PERMANENTES
- Cite systématiquement les textes et jurisprudence (références exactes et complètes).
- Signale proactivement tout risque juridique ou fiscal, même non demandé.
- Français juridique de haut niveau, niveau cabinet parisien d'affaires.
- Si un élément manque dans les pièces, le signaler explicitement.
- Excel : analyse les données chiffrées et implications juridiques/fiscales.
- PowerPoint : analyse le contenu substantiel.

## ${modeInstruction(mode)}`;
}

export function buildLibSystem(
  domain: LibDomain | 'all', skills: Array<{ name: string; content: string }>
): string {
  const labels: Record<string, string> = {
    commercial:'droit commercial', fiscal:'droit fiscal', social:'droit social',
    civil:'droit civil', penal:'droit pénal', immobilier:'droit immobilier',
    international:'droit international', autre:'droit général', all:'tous domaines juridiques'
  };
  const skillBlock = skills.length
    ? `\n\n## INSTRUCTIONS MÉTIER\n${skills.map(s => `### ${s.name}\n${s.content}`).join('\n\n')}`
    : '';
  return `Tu es Lex Assistant, expert en ${labels[domain] ?? 'droit français'}, au service de Maître Mina Bibawi, avocat au Barreau de Paris.
Tu as accès à une bibliothèque juridique thématique fournie avec chaque question.${skillBlock}

## RÈGLES
- Précision académique et pratique de haut niveau.
- Cite toujours la source exacte (arrêt, article, auteur) issue des documents fournis.
- Structure avec des titres clairs.
- Si la question dépasse les documents, le signaler explicitement.
- Propose des analyses comparatives et chronologies jurisprudentielles.`;
}

// ─── File fetching — injected by the caller ───────────────────────────────────
// api.ts no longer imports from onedrive.ts directly; instead the caller
// provides a readFile function appropriate to the scenario.

async function buildDocParts(
  docs: (CaseDocumentMeta | LibDocumentMeta)[],
  readFile: (name: string) => Promise<ArrayBuffer>
): Promise<ContentPart[]> {
  const parts: ContentPart[] = [];
  for (const doc of docs) {
    try {
      const buf = await readFile(doc.name);
      parts.push(docPart(doc.name, doc.mimeType, toBase64(buf)));
    } catch {
      parts.push({ type: 'text', text: `[Fichier "${doc.name}" inaccessible sur OneDrive]` });
    }
  }
  return parts;
}

// ─── Case call ────────────────────────────────────────────────────────────────

export interface CaseCallOpts {
  caseName: string;
  caseDomain: string;
  notes: PermanentNote[];
  docs: CaseDocumentMeta[];
  skills: Array<{ name: string; content: string }>;
  mode: WorkMode;
  userMessage: string;
  /** Caller provides file reader scoped to the case folder */
  readFile: (fileName: string) => Promise<ArrayBuffer>;
}

export async function callClaudeCase(opts: CaseCallOpts): Promise<string> {
  const key = getStoredKey();
  if (!key) throw new Error('Clé API manquante. Configurez-la dans les paramètres.');
  const system   = buildCaseSystem(opts.caseName, opts.caseDomain, opts.notes, opts.skills, opts.mode);
  const docParts = await buildDocParts(opts.docs, opts.readFile);
  const content: ContentPart[] = [...docParts, { type: 'text', text: opts.userMessage }];
  return fetchClaude(key, { model: MODEL, max_tokens: 4096, system,
    messages: [{ role: 'user', content: content as unknown as string }] });
}

// ─── Library call ─────────────────────────────────────────────────────────────

export interface LibCallOpts {
  domain: LibDomain | 'all';
  docs: LibDocumentMeta[];
  skills: Array<{ name: string; content: string }>;
  userMessage: string;
  history: Array<{ role: 'user' | 'assistant'; content: string }>;
  /** Caller provides file reader scoped to the library domain folder */
  readFile: (fileName: string) => Promise<ArrayBuffer>;
}

export async function callClaudeLib(opts: LibCallOpts): Promise<string> {
  const key = getStoredKey();
  if (!key) throw new Error('Clé API manquante. Configurez-la dans les paramètres.');
  const system   = buildLibSystem(opts.domain, opts.skills);
  const docParts = await buildDocParts(opts.docs, opts.readFile);
  const messages: AnthropicRequest['messages'] = [
    ...opts.history.map(h => ({ role: h.role, content: h.content })),
    { role: 'user', content: [...docParts, { type: 'text', text: opts.userMessage }] as unknown as string }
  ];
  return fetchClaude(key, { model: MODEL, max_tokens: 4096, system, messages });
}

async function fetchClaude(key: string, body: AnthropicRequest): Promise<string> {
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
    const e = await resp.json().catch(() => ({ error: { message: resp.statusText } })) as { error?: { message?: string } };
    throw new Error(`API Claude : ${e.error?.message ?? resp.statusText}`);
  }
  const data = await resp.json() as AnthropicResponse;
  return data.content.map(b => b.text).join('');
}
