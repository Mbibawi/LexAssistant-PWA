// ─── Domain types ─────────────────────────────────────────────────────────────

type DocKind = 'piece' | 'jurisprudence' | 'doctrine' | 'redige';

type WorkMode = 'analyse' | 'redaction' | 'modification' | 'note';

type ChatRole = 'user' | 'assistant';

type LibDomain =
  | 'commercial' | 'fiscal' | 'social' | 'civil'
  | 'penal' | 'immobilier' | 'international' | 'autre';

  // ─── MSAL types ───────────────────────────────────────────────────────────────
  
  type MsalApp = {
    loginPopup(r: { scopes: string[] }): Promise<{ accessToken: string }>;
    acquireTokenSilent(r: { scopes: string[]; account: MsalAccount }): Promise<{ accessToken: string }>;
    acquireTokenPopup(r: { scopes: string[] }): Promise<{ accessToken: string }>;
    getAllAccounts(): MsalAccount[];
    handleRedirectPromise(): Promise<null>;
  }
type MsalAccount = { homeAccountId: string; username: string; name?: string }
  
  declare const msal : {
    PublicClientApplication: new (params: object) => MsalApp;
  }
  
  type ScenarioCardOptions = {
    id:          string;
    icon:        string;
    title:       string;
    description: string;
    features:    string[];
    action:      () => void;
  }

// ─── _meta.json — stored in each Affaires/<case>/ folder ─────────────────────

type CaseMeta = {
  name: string;                // Display name
  folderName: string;          // OneDrive folder name (sanitised)
  domain: string;
  status: 'active' | 'closed' | 'suspended';
  createdAt: number;
  updatedAt: number;
  // Document registry — file names + metadata (no base64)
  documents: CaseDocumentMeta[];
}

type CaseDocumentMeta = {
  name: string;                // exact filename on OneDrive
  kind: DocKind;
  mimeType: string;
  sizeBytes: number;
  addedAt: number;
}

// ─── _notes.json — corrections & permanent notes, stored in case folder ───────

type NotesFile = {
  notes: PermanentNote[];
}

type PermanentNote = {
  id: string;
  content: string;
  createdAt: number;
  updatedAt: number;
}

// ─── _conversation.json — chat history, stored in case / library folder ───────

type ConversationFile = {
  messages: ChatMessage[];
}

type ChatMessage = {
  id: string;
  role: ChatRole;
  content: string;
  timestamp: number;
  mode: WorkMode;
  generatedDocName?: string;
}

// ─── In-memory case object (assembled from CaseMeta + folder listing) ─────────

type LexCase = {
  folderName: string;          // primary key — the OneDrive folder name
  name: string;
  domain: string;
  status: 'active' | 'closed' | 'suspended';
  createdAt: number;
  updatedAt: number;
  documents: CaseDocumentMeta[];
}

// ─── Module 2 — Legal Library ─────────────────────────────────────────────────

type LibDocumentMeta = {
  name: string;
  mimeType: string;
  sizeBytes: number;
  addedAt: number;
  tags: string[];
}

// _meta.json in each Bibliotheque/<domain>/ folder
type LibDomainMeta = {
  domain: LibDomain;
  documents: LibDocumentMeta[];
}

type LibConversationMessage = {
  id: string;
  role: ChatRole;
  content: string;
  timestamp: number;
  domain: LibDomain | 'all';
}

// ─── OneDrive / Graph ─────────────────────────────────────────────────────────

type GraphDriveItem = {
  name: string;
  size: number;
  file?: { mimeType: string };
  folder?: { childCount: number };
  webUrl?: string;
  lastModifiedDateTime?: string;
}

type OneDriveConfig = {
  clientId: string;
  tenantId: string;        // 'common' for personal + org accounts
  rootFolder: string;      // e.g. 'LexAssistant'
}

// ─── API ──────────────────────────────────────────────────────────────────────

type AnthropicContentBlock = {
  type: 'text';
  text: string;
}

type AnthropicRequest = {
  model: string;
  max_tokens: number;
  system: string;
  messages: Array<{ role: 'user' | 'assistant'; content: string | AnthropicContentBlock[] }>;
}

type AnthropicResponse = {
  content: AnthropicContentBlock[];
  stop_reason: string;
}
