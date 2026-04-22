// ─── Domain types ─────────────────────────────────────────────────────────────

export type DocKind = 'piece' | 'jurisprudence' | 'doctrine' | 'redige';

export type WorkMode = 'analyse' | 'redaction' | 'modification' | 'note';

export type ChatRole = 'user' | 'assistant';

export type LibDomain =
  | 'commercial' | 'fiscal' | 'social' | 'civil'
  | 'penal' | 'immobilier' | 'international' | 'autre';

// ─── _meta.json — stored in each Affaires/<case>/ folder ─────────────────────

export interface CaseMeta {
  name: string;                // Display name
  folderName: string;          // OneDrive folder name (sanitised)
  domain: string;
  status: 'active' | 'closed' | 'suspended';
  createdAt: number;
  updatedAt: number;
  // Document registry — file names + metadata (no base64)
  documents: CaseDocumentMeta[];
}

export interface CaseDocumentMeta {
  name: string;                // exact filename on OneDrive
  kind: DocKind;
  mimeType: string;
  sizeBytes: number;
  addedAt: number;
}

// ─── _notes.json — corrections & permanent notes, stored in case folder ───────

export interface NotesFile {
  notes: PermanentNote[];
}

export interface PermanentNote {
  id: string;
  content: string;
  createdAt: number;
  updatedAt: number;
}

// ─── _conversation.json — chat history, stored in case / library folder ───────

export interface ConversationFile {
  messages: ChatMessage[];
}

export interface ChatMessage {
  id: string;
  role: ChatRole;
  content: string;
  timestamp: number;
  mode: WorkMode;
  generatedDocName?: string;
}

// ─── In-memory case object (assembled from CaseMeta + folder listing) ─────────

export interface LexCase {
  folderName: string;          // primary key — the OneDrive folder name
  name: string;
  domain: string;
  status: 'active' | 'closed' | 'suspended';
  createdAt: number;
  updatedAt: number;
  documents: CaseDocumentMeta[];
}

// ─── Module 2 — Legal Library ─────────────────────────────────────────────────

export interface LibDocumentMeta {
  name: string;
  mimeType: string;
  sizeBytes: number;
  addedAt: number;
  tags: string[];
}

// _meta.json in each Bibliotheque/<domain>/ folder
export interface LibDomainMeta {
  domain: LibDomain;
  documents: LibDocumentMeta[];
}

export interface LibConversationMessage {
  id: string;
  role: ChatRole;
  content: string;
  timestamp: number;
  domain: LibDomain | 'all';
}

// ─── OneDrive / Graph ─────────────────────────────────────────────────────────

export interface GraphDriveItem {
  name: string;
  size: number;
  file?: { mimeType: string };
  folder?: { childCount: number };
  webUrl?: string;
  lastModifiedDateTime?: string;
}

export interface OneDriveConfig {
  clientId: string;
  tenantId: string;        // 'common' for personal + org accounts
  rootFolder: string;      // e.g. 'LexAssistant'
}

// ─── API ──────────────────────────────────────────────────────────────────────

export interface AnthropicContentBlock {
  type: 'text';
  text: string;
}

export interface AnthropicRequest {
  model: string;
  max_tokens: number;
  system: string;
  messages: Array<{ role: 'user' | 'assistant'; content: string | AnthropicContentBlock[] }>;
}

export interface AnthropicResponse {
  content: AnthropicContentBlock[];
  stop_reason: string;
}
