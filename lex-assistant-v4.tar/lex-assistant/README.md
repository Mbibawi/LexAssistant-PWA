# Lex Assistant — v2

Gestionnaire de dossiers juridiques + Bibliothèque juridique thématique.
Assisté par Claude (API Anthropic). Stockage sur Microsoft OneDrive via Graph API.

**Stack : Vanilla TypeScript — ESModules natifs — aucun framework — aucun bundler.**

---

## Structure du projet

```
src/
  main.ts                 — Module 1 : dossiers, UI principale, OneDrive sync
  module2.ts              — Module 2 : bibliothèque juridique thématique
  types.ts                — Types domaine (LexCase, CaseDocument, LibDocument…)
  modules/
    api.ts                — Appels Claude API (case + library + skills injection)
    db.ts                 — IndexedDB v2 (cases, documents, notes, messages, libdocs, libmessages)
    docxgen.ts            — Génération .docx côté client (docx v9 IIFE)
    ingest.ts             — Ingestion fichiers : PDF, DOCX, XLSX, PPTX, TXT, MD
    markdown.ts           — Renderer markdown → HTML (zéro dépendance)
    onedrive.ts           — Microsoft Graph API + MSAL.js PKCE auth
    ui.ts                 — Helpers DOM (el, qs, toast, confirm…)
public/
  index.html / app.css / sw.js / manifest.json / icons/
scripts/
  postbuild.js            — Copie public/, corrige imports .js, copie docx IIFE
  serve.js                — Serveur dev (Node built-ins uniquement)
```

---

## Installation & Build

```bash
tar -xzf lex-assistant-v2.tar.gz
cd lex-assistant
npm install          # typescript + docx uniquement
npm run build        # compile + postbuild
npm run serve        # http://localhost:3000
```

---

## Structure OneDrive

L'application crée et gère automatiquement cette structure :

```
OneDrive/
└── LexAssistant/              ← rootFolder (configurable)
    ├── _Skills/               ← fichiers .md ou .txt = instructions injectées dans chaque prompt
    │   ├── style-redaction.md
    │   └── jurisprudence-fiscale.md
    ├── Affaires/
    │   ├── ISHAK_Succession/  ← un sous-dossier par dossier client
    │   │   ├── acte-notoriete.pdf
    │   │   ├── compte-succession.docx
    │   │   └── analyse-FII.xlsx
    │   └── DUPONT_Cession/
    └── Bibliotheque/
        ├── Commercial/        ← jurisprudences, articles doctrinaux, textes
        ├── Fiscal/
        ├── Social/
        ├── Civil/
        ├── Penal/
        ├── Immobilier/
        └── International/
```

---

## Configuration OneDrive (Azure App Registration)

### Étape 1 — Créer l'App Registration

1. Allez sur **https://portal.azure.com**
2. Cherchez **"App registrations"** → **New registration**
3. Remplissez :
   - **Name** : `LexAssistant` (ou tout autre nom)
   - **Supported account types** : *"Accounts in any organizational directory and personal Microsoft accounts"*
     (pour les comptes OneDrive personnels, choisir cette option)
   - **Redirect URI** : `Single-page application (SPA)` → `http://localhost:3000`
     (ajoutez aussi votre URL de production si vous déployez)
4. Cliquez **Register**
5. Notez l'**Application (client) ID** (format `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`)

### Étape 2 — Ajouter les permissions API

Dans votre App Registration → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** :
- `Files.ReadWrite` — lecture/écriture OneDrive
- `User.Read` — identification de l'utilisateur connecté

Cliquez **Grant admin consent** si vous êtes sur un tenant organisationnel.
*(Pour un compte personnel Microsoft, le consentement est donné lors de la première connexion.)*

### Étape 3 — Configurer l'authentification

Dans **Authentication** :
- Vérifiez que le Redirect URI `http://localhost:3000` est bien de type **SPA**
- Cochez **"Access tokens"** et **"ID tokens"** dans *Implicit grant*
- **Allow public client flows** : Yes

### Étape 4 — Configurer dans l'app

Ouvrez Lex Assistant → ⚙ Paramètres → section OneDrive :
- **Application (Client) ID** : votre client ID
- **Tenant ID** : `common` (compte personnel ou multi-tenant)
- **Dossier racine** : `LexAssistant` (ou autre nom)
- Cliquez **Enregistrer**
- Cliquez **Initialiser la structure OneDrive** → popup de connexion Microsoft → autorisez

---

## Skills — Instructions permanentes

Les fichiers dans `OneDrive/LexAssistant/_Skills/` sont des instructions injectées dans le `system` de **chaque appel Claude**, pour les deux modules.

Format : fichiers `.md` ou `.txt`.

Exemples de skills utiles :

**`style-redaction.md`**
```markdown
## Style de rédaction
- Utilise toujours la formule de politesse "Maître" pour les confrères
- Les courriers aux notaires commencent par "Maître,"
- Les conclusions commencent par "POUR : [nom du client]"
- Numérote toujours les paragraphes de conclusions
```

**`jurisprudence-referentielle.md`**
```markdown
## Jurisprudences de référence à citer systématiquement
- Acte anormal de gestion : CE, 27 juillet 1984, SA Renfort Service
- Centre des intérêts vitaux : CE, 12 juin 2020, n°418914
- Démembrement SCI : Cass. 3e civ., 7 nov. 2019, n°18-23.259
```

**`clients-specifiques.md`**
```markdown
## Informations client permanentes
- Dossier ISHAK : succession internationale franco-égyptienne, convention du 24 juin 1988
- M. ABDALLA : résident fiscal Égypte, article 20 de la convention franco-égyptienne
```

Les skills sont rechargés à chaque connexion OneDrive. Pour les recharger sans se déconnecter : bouton ☁ OneDrive dans la topbar.

---

## Types de fichiers supportés

| Extension | Type | Traitement |
|---|---|---|
| `.pdf` | PDF | Envoyé nativement à Claude (lecture complète) |
| `.docx` | Word | Envoyé nativement à Claude |
| `.xlsx` | Excel | Envoyé nativement à Claude (analyse des données) |
| `.pptx` | PowerPoint | Envoyé nativement à Claude (analyse du contenu) |
| `.txt` / `.md` | Texte | Extrait et envoyé comme texte |
| `.doc` / `.xls` / `.ppt` | Anciens formats Office | Envoyés en base64 |

---

## Flux OneDrive par opération

### Upload d'un document (Module 1)
1. Utilisateur upload via `⬆ Upload`
2. Fichier stocké en IndexedDB (base64)
3. Si OneDrive connecté → upload automatique dans `Affaires/<NomDossier>/`
4. Le document est marqué `source: 'onedrive'` avec son `itemId`

### Sync depuis OneDrive (`☁ Sync`)
1. Liste les fichiers dans `Affaires/<NomDossier>/`
2. Crée des stubs légers (pas de téléchargement immédiat)
3. Au moment d'envoyer à Claude : téléchargement à la volée et mise en cache IndexedDB

### Sauvegarde d'un acte rédigé
- **"Enregistrer dans le dossier"** → IndexedDB + upload automatique OneDrive
- **"☁ Sauver OneDrive"** → upload direct sans passer par IndexedDB

### Bibliothèque (Module 2)
- **`⟳ Sync OneDrive`** → indexe les fichiers de `Bibliotheque/<Domaine>/`
- Documents téléchargés à la volée lors de chaque question
- Upload local → stocké IndexedDB + synchronisé OneDrive si connecté

---

## Déploiement

Le dossier `dist/` est un site statique pur :

```bash
npm run build
# Déployez dist/ sur Vercel, Netlify, GitHub Pages, Azure Static Web Apps…
```

**Redirect URI** : ajoutez votre URL de production dans Azure App Registration → Authentication.

Pour Azure Static Web Apps, créez un fichier `dist/staticwebapp.config.json` :
```json
{
  "navigationFallback": { "rewrite": "/index.html" }
}
```

---

## Dépendances

| Package | Usage |
|---|---|
| `typescript` | Compilation (devDep) |
| `docx` | Génération .docx côté client |

Runtime (chargés dynamiquement depuis CDN) :
| Bibliothèque | Usage |
|---|---|
| `@azure/msal-browser@3` (jsdelivr) | Auth Microsoft PKCE |

Zéro framework. Zéro bundler. Deux dépendances npm.
