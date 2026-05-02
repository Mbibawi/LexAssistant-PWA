// ─── Domain types ─────────────────────────────────────────────────────────────

type msalConfig = {
  auth: {
    clientId: string,
    authority: string,
    redirectUri: string,
  },
  cache: {
    cacheLocation: string,
    storeAuthStateInCookie: boolean
  }
}
type DocKind = 'piece' | 'jurisprudence' | 'doctrine' | 'redige';


type WorkMode = 'analyse' | 'redaction' | 'modification' | 'note';

type ChatRole = 'user' | 'assistant';

type LibDomain =
  | 'commercial' | 'fiscal' | 'social' | 'civil'
  | 'penal' | 'immobilier' | 'international' | 'autre';

  // ─── MSAL types ───────────────────────────────────────────────────────────────
  
type MsalApp = {
  initialize(): Promise<void>;
  loginPopup(r: { scopes: string[] }): Promise<{ account: MsalAccount }>;
    acquireTokenSilent(r: { scopes: string[]; account: MsalAccount }): Promise<{ accessToken: string }>;
  acquireTokenPopup(r: { scopes: string[] }): Promise<{ accessToken: string }>;
  acquireTokenRedirect(r: { scopes: string[] }): Promise<{ accessToken: string }>;
    getAllAccounts(): MsalAccount[];
  handleRedirectPromise(): Promise<{ account: MsalAccount | null; accessToken: string | null }>;
  loginRedirect(r: { scopes: string[]; prompt?: string }): Promise<void>;
  ssoSilent(r: { scopes: string[]; loginHint?: string }): Promise<{ account?: MsalAccount, accessToken?: string }>;
  setActiveAccount(account: MsalAccount): void
  getActiveAccount(): MsalAccount
  }

type MsalAccount = {
  homeAccountId: string;
  username: string;
  name?: string;
  idTokenClaims: { oid: string };
  localAccountId?: string;
}
  
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

type CaseCallOpts = {
  caseName: string;
  caseDomain: string;
  notes: PermanentNote[];
  docs: CaseDocumentMeta[];
  skills: { name: string; content: string }[];
  mode: WorkMode;
  userMessage: string;
  knowledgeBase: string | undefined;
  /** Caller provides file reader scoped to the case folder */
  readFile: (folderName, fileName: string) => Promise<ArrayBuffer>;
};

type LibCallOpts = {
  domain: LibDomain | "all";
  docs: LibDocumentMeta[];
  skills: { name: string; content: string }[];
  userMessage: string;
  history: { role: ChatRole; content: string }[];
  knowledgeBase: string | undefined;
  /** Caller provides file reader scoped to the library domain folder */
  readFile: (folderName: LibDomain | string, fileName: string) => Promise<ArrayBuffer>;
};

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
  id?: string;
  role: ChatRole;
  content: string | ContentPart[];
  timestamp?: number;
  mode?: WorkMode;
  generatedDocName?: string;
};

// ─── In-memory case object (assembled from CaseMeta + folder listing) ─────────

type FolderMeta = {
  folderName: string; // primary key — the OneDrive folder name
  name: string;
  domain: string;
  status: "active" | "closed" | "suspended";
  createdAt: number;
  updatedAt: number;
  documents: CaseDocumentMeta[];
};

//  Legal Library ─────────────────────────────────────────────────

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

type ChatBlock = {
  type: "text";
  text: string;
};

type MessageSystem = {
  type: "text";
  text: string;
  cache_control?: { type: string; ttl: string };
}[];

type ClaudeMessages = {
  model: string;
  max_tokens: number;
  system?: {
    type: "text";
    text: string;
    cache_control?: { type: string; ttl: string };
  }[];
  messages: ChatMessage[];
};

type ClaudeResponse = {
  content: ChatBlock[];
  stop_reason: string;
};

type ContentPart =
  | ChatBlock
  | {
    type: "document";
    source: { type: "base64"; media_type: string; data: string };
    title?: string;
  };

type ElAttributes = {
  id?: string;
  type?: string;
  className?: string;
  innerHTML?: string;
  innerText?: string;
  textContent?: string;
  // This allows you to pass { display: "none" } without errors
  style?: Partial<CSSStyleDeclaration>;
  [key: string]: any; // Allows other attributes like data-props
}

type ids = {
  chat: "user-input"
}
type header = {
  "Authorization": string;
  "Content-Type": string;
  [key: string]: string;
}

declare class XML {
  schema: string;
  constructor(doc: XMLDocument, lang: string);
  getTables(w: XMLDocument): Element[];
  findTableByTitle(all: Element[], title: string): Element;
  getTableRow(table: Element, index: Number): Element;
  getTableCell(table: Element, row: Number, col: Number): Element;
  editTables(xml: XML, document: XMLDocument);
  insertRowAfter(table: Element, after: any, row: any, newRow: any);
  deleteRow(table: Element, row: Number);
  appendRow(table: Element, row: any);
  createTableRow(): Element;
  getRowCells(tableRow: Element): Element[];
  createTableCell(): Element;
  getTextElement(cell: Element, index: number): Element;
  appendParagraph(cell): Element;
  setTextLanguage(cell: Element, lang: string): Element;
  setTableCellLanguage(table: Element, row: Number, col: Number, lang: string);
  getPropElement(cell: Element, index: number): Element;
  getParagraph(cell: Element): Element;
  setTextLanguage(cell: Element): Element;
  createPropElement(cell: Element): Element;
  findPropertyParagraph(paragraph: Element): Element;
  getShadowElement(tcPr: Element, n: number): Element;
  createShadowElement(): Element;
  getParagraphStyle(pPr: any, n: number): Element;
  createParagraphStyle(): Element;
  getStyle(index: number, is: boolean): string;
  getContentControls(parent: XMLDocument | Element): Element[];
  findContentControlsByTitle(ctrls: Element[], title: string): Element[];
  getContentControls(body: Element): Element[];
  editContentControlText(control: Element, value: string): void;
}

declare class JSZip {
  loadAsync(arrayBuffer: ArrayBuffer): Promise<JSZip>;
  files: JSZipFile[];
  file(name: string, serialized?: string): JSZipFile;
  generateAsync(type: { type: "blob" }): Promise<Blob>;
}

type JSZipFile = {
  name: string;
  content: string;
  async(type = 'string'): Promise<string>;
}



