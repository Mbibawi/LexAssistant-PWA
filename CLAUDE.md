# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run build    # tsc + scripts/postbuild.js (required after every change)
npm run watch    # tsc --watch for incremental compilation
npm run serve    # dev server at http://localhost:3000 (SPA fallback, COOP/COEP headers)
```

There is no linter and no test runner. TypeScript strict mode is the primary correctness check.

**Post-build script** (`scripts/postbuild.js`) rewrites relative imports in compiled `.js` files to add `.js` extensions (required for native ESM in browsers), copies `public/` → `dist/`, and bundles `docx` as a UMD vendor asset. Always run `npm run build`; `tsc` alone is not sufficient.

## Architecture

**Zero-framework vanilla TypeScript PWA.** No React, no bundler — native ES modules loaded directly in the browser. The compiled `dist/` is a static site deployable anywhere.

### Class hierarchy (all in `src/modules/onedrive.ts`)

```
Configuration   ← localStorage (OneDrive Client ID, Tenant ID, root folder)
  └── oneDrive  ← MSAL PKCE auth + Microsoft Graph API calls
        └── Folders  ← folder/file/JSON CRUD on OneDrive
              └── Cases  ← dossier (legal case) scenario logic
                    └── Library  ← extends Cases; adds bibliothèque scenario
```

`Library` is instantiated once in `src/main.ts` as the app singleton. Both scenarios share the same instance.

### Two scenarios

1. **Dossiers** — Legal case file management. Each case lives in `OneDrive/LexAssistant/Affaires/<CaseName>/` with `_meta.json`, `_notes.json`, and `_conversation.json`.
2. **Bibliothèque** — Legal knowledge base organized into thematic domains (Commercial, Fiscal, Social, Civil, Pénal, Immobilier, International).

### Key modules

| File | Purpose |
|------|---------|
| `src/main.ts` | App shell: scenario selector, topbar wiring, Library singleton |
| `src/modules/onedrive.ts` | Everything OneDrive + MSAL + Cases/Library domain logic (~2000 lines) |
| `src/modules/api.ts` | Claude API calls; system prompts for analysis, redaction, modification, notes |
| `src/modules/ingest.ts` | File upload handling (PDF, DOCX, XLSX, PPTX, TXT, MD, DOC, XLS, PPT) |
| `src/modules/docxgen.ts` | Client-side `.docx` generation via `docx` library (UMD vendor build) |
| `src/modules/markdown.ts` | Zero-dependency Markdown → HTML renderer |
| `src/modules/ui.ts` | DOM helpers: `el`, `qs`, `toast`, `confirm`, etc. |
| `src/types.d.ts` | All TypeScript types: domain models, MSAL, OneDrive Graph, Anthropic API shapes |

### Data flow

```
User action → Library/Cases method
                ├── OneDrive (Graph API) — source of truth for docs and metadata
                ├── IndexedDB — local cache for documents (base64) and chat history
                └── Claude API (api.ts) — analysis using system prompts + doc content + skills
```

### OneDrive file structure

```
OneDrive/LexAssistant/
├── _Skills/                  ← .md/.txt files injected verbatim into Claude system prompts
├── Affaires/<CaseName>/
│   ├── _meta.json            ← case metadata (name, domain, status, doc list)
│   ├── _notes.json           ← permanent corrections (highest priority in prompts)
│   ├── _conversation.json    ← chat history
│   └── <documents>
└── Bibliotheque/<Domain>/
    ├── _meta.json
    ├── _conversation.json
    └── <jurisprudence, doctrine, texts>
```

### Storage layers

- **localStorage** — API key, OneDrive config (Client ID, Tenant ID, root folder)
- **IndexedDB** — document base64 cache, chat history, notes (local)
- **OneDrive** — source of truth; synced on load and after mutations

### Claude API integration

All calls go through `src/modules/api.ts`. The model used is `claude-sonnet-4-20250514`. Skills files from `_Skills/` on OneDrive are fetched and injected into system prompts. Notes from `_notes.json` are prepended with high priority.

### External dependencies

- **`docx@^9.6.1`** — loaded as `dist/vendor/docx.umd.js` (UMD, not ESM)
- **`@azure/msal-browser@3`** — loaded from CDN (jsDelivr) in `index.html`, not npm
- **TypeScript** — dev-only; only used at build time

## Environment

No `.env` files. All configuration is entered in the in-app Settings modal:
- Anthropic API key → stored in `localStorage`
- Azure App Registration (`Client ID`, `Tenant ID`) → stored in `localStorage`
- OneDrive root folder name → stored in `localStorage`

For OneDrive to work, an Azure App Registration must exist with `Files.ReadWrite` and `User.Read` permissions, and the redirect URI must match the deployment origin (e.g., `http://localhost:3000` for local dev).
