# ApplyPaste Development

## Runtime

ApplyPaste is a Manifest V3 Edge extension. `popup.html` and `sidepanel.html` load the shared runtime in `src/`. The Side Panel uses the same information library and browser-local state as the main workbench.

## Storage

- `chrome.storage.local` stores lightweight settings and UI state.
- IndexedDB stores the information library, attachments, versions, and application records.
- ApplyPaste uses its own extension identity and storage names. It does not migrate or read Job Paste browser data.
- `data/` and `public-template/data/` contain only the empty public bootstrap shape.

## Main Modules

- `src/app.js`: UI initialization and shared workbench behavior.
- `src/internal-db.js`: browser-local database access.
- `src/store.js`: empty profile bootstrap and state helpers.
- `src/resume-docx.js`, `src/docx.js`: Word resume import and export.
- `src/records.js`: local application records.
- `src/feishu.js`: optional Feishu CREATE/UPDATE synchronization using `localRecordId`.
- `src/workspace.js`: optional user-selected workspace import/export paths.

## Tests

Automated tests are in `tests/`; public fixtures are in `tests/fixtures/`. Run the primary checks with:

```powershell
npm run self-check
npm run test:feature-parity
npm run test:e2e
npm run test:public-build
npm run privacy-scan
```

## Build And Release

Create the public extension directory with:

```powershell
npm run build:public
```

The build target is `release/ApplyPaste-v0.1-beta/`. The build copies only runtime files, empty public data, user documentation, and icons. Recreate `release/ApplyPaste-v0.1-beta.zip` after a successful build and privacy scan.

## Release Safety

Never place personal profiles, attachments, application records, Feishu credentials, tokens, or private absolute paths in `data/`, `public-template/`, `docs/images/`, `tests/fixtures/`, or `release/`.
