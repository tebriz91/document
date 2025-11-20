## VS Code Extension Plan

### Context
- Web app boots via `index.ts` and `index.html` with an iframe. OnlyOffice API loads in `lib/onlyoffice-editor.ts` and x2t WASM loads in `lib/document-converter.ts` using `BASE_PATH` from `lib/document-utils.ts`. `BASE_PATH` currently assumes `window` and GitHub/Docker paths.
- Messaging uses `Platform.init` in `lib/events.ts` with `RENDER_OFFICE` (chunked file from host) and `CLOSE_EDITOR`. File chunks are decoded with `MessageCodec.decodeFileChunked`, then `handleDocumentOperation` opens the file. Save flow is purely browser-side (`handleSaveDocument` → `convertBinToDocumentAndDownload`).
- UI entry points live in `index.ts` (control panel, create/open document, URL params). Editor lifecycle and conversions are in `lib/converter.ts`, `lib/onlyoffice-editor.ts`, and `lib/document-utils.ts`.

### Goals / Non-Goals
- Goal: ship a VS Code extension that uses a webview + OnlyOffice/x2t to open/edit/save docx/xlsx/pptx/csv fully offline.
- Goal: keep current web app usable in browser; changes must remain backward compatible (feature flags or env switches).
- Non-goal (for now): multi-user/collab editing or remote storage providers.
- Non-goal: streaming edits back to the host on every keystroke; we will stick to save events and dirty tracking.

### Extension Shell
- Create `src/extension.ts` that registers a `CustomEditorProvider` for docx/xlsx/pptx/csv and a command `openOnlyOfficeEditor` that delegates to the provider.
- Use a `CustomDocument` implementation that caches the last saved binary, supports backup, and exposes `save`, `saveAs`, and `revert`.
- Webview creation: set `enableScripts`, `retainContextWhenHidden`, and `localResourceRoots: [extensionUri]`. Inject a JS entry (`media/webview/index.js`) and CSS; pass initial state (document URI, isNew, locale).
- Serve Vite-built assets plus static `public` payloads (`web-apps/apps/api/documents/api.js`, `wasm/x2t/x2t.js/.wasm`, `libs/sheetjs`, fonts, sdkjs/) from `media/`. Decide whether to embed fonts or prompt users to add licensed fonts.

### Asset URIs
- Compute `baseUri = webview.asWebviewUri(joinPath(extensionUri, 'media'))` and inject it as `window.__BASE_PATH` (or env variable) before app bootstrap; fall back to `getBasePath` when absent.
- Rewrite asset references in `lib/document-utils.ts`, `lib/document-converter.ts`, and `lib/onlyoffice-editor.ts` to read from `__BASE_PATH` so `wasm/x2t/x2t.js/.wasm`, OnlyOffice API script, fonts, and `web-apps` assets all resolve inside the webview.
- Keep Vite `base: './'` so the webview bundle uses relative paths; ensure copied static assets preserve folder structure under `media/`.
- Add a small pre-webview script to patch `fetch`/`URL` usage if VS Code webview restrictions require it (mainly for WASM binaries).
- Add an `index.html` placeholder for the base path plus a tiny `path-resolver` helper so both VS Code and web resolve `web-apps`/`wasm`/fonts; hide/remove the control panel when running inside VS Code to avoid duplicate UI.

### CSP
- Apply CSP on `webview.html`: `default-src 'none'; img-src ${webview.cspSource} blob: data:; script-src 'nonce-${nonce}' ${webview.cspSource}; style-src 'unsafe-inline' ${webview.cspSource}; worker-src blob:; connect-src ${webview.cspSource}; font-src ${webview.cspSource} data:; media-src ${webview.cspSource} blob:;`.
- Validate that WASM fetch uses `webview.asWebviewUri`; if VS Code disallows direct WASM execution, route via `worker-src blob:` backed by `new Worker` inside the webview.
- Ensure CSP allowances cover x2t workers and pasted images (`blob:` for worker/media) and that WASM/script URLs always go through `asWebviewUri`.

### Messaging Bridge
- Webview side: wrap `Platform.init` to listen to `window.message` and forward to existing handlers. Keep schema `{chunkIndex,data,lastModified,name,size,totalChunks,type}` for `RENDER_OFFICE` and `CLOSE_EDITOR`.
- Host side: use `MessageCodec.encodeFileChunked` (or replicate logic) to stream a file into the webview; support backpressure by waiting for `READY_FOR_CHUNK` replies if needed for large files.
- Add explicit messages: `WEBVIEW_READY`, `DOC_READY`, `SAVE_REQUESTED` (webview → host with `{format, fileName, binData}`), `SAVE_COMPLETE`, `ERROR`, and `LOG`. Gate all posts through `webview.postMessage`.
- Order messages so the webview emits `WEBVIEW_READY` before the host sends chunks and optionally uses `READY_FOR_CHUNK` pacing for large files; resend `DOC_READY` after reloads.
- Normalize `createObjectURL`/blob URLs on the webview side so existing OnlyOffice hooks keep working.

### Open/Create Flows
- Existing file: `CustomEditorProvider.resolveCustomEditor` reads file bytes, chunks into `RENDER_OFFICE` messages, and optionally sends metadata (workspace path, lastModified).
- New file: `openOnlyOfficeEditor` command prompts for format; provider creates an untitled `CustomDocument`, then sends `{type:'CREATE_NEW', extension}` so the webview calls `onCreateNew`.
- Support re-open/reload: stash last-sent file in `CustomDocument`; on reload, resend chunks and `DOC_READY` acknowledgment to prevent duplicate init.

### Save/Dirty Handling
- Webview intercepts `onSave` (currently `handleSaveDocument`), replaces `convertBinToDocumentAndDownload` with `postMessage({type:'SAVE_REQUESTED', format, fileName, data})`.
- Host writes buffer to disk; for CSV special-case keep the existing logic (`fileName` ends with `.csv` → force format CSV). Return `SAVE_COMPLETE` to clear dirty state.
- Add `onDidChangeCustomDocument` firing when a save event arrives and when edits occur (if we wire incremental updates later). Implement `backup` to persist the latest buffer in VS Code’s storage.
- Decide on conversion location: initially keep x2t in webview for parity; later optionally add host-side conversion to reduce memory.
- Have `CustomDocument` cache the current buffer, mark dirty on webview edit signals, and expose undo/redo stubs plus `revert` that reloads last saved/backup content.

### Lifecycle
- On panel dispose, post `CLOSE_EDITOR` so `window.editor.destroyEditor()` runs and chunk buffers reset.
- Handle multiple panels: keep a per-webview base path and per-document `Platform` instance; avoid global singletons that assume only one iframe.
- Respect existing queue in `onlyoffice-editor.ts` when switching files; add guards to ignore late messages after dispose.

### Build/Packaging
- Keep Vite `base: './'`; add a webview build target that outputs to `media/webview/`. Copy all `public` assets (OnlyOffice API, wasm, fonts, sdkjs) into `media/` preserving structure.
- Add npm scripts: `build:webview` (Vite), `build:ext` (`vsce package` or `@vscode/tsc -p ./`), and `package` to run both.
- Ensure `extension.ts` locates `media/webview/index.html` and injects nonce + base path. Include `pnpm` equivalents if we keep the monorepo tooling.
- Add `webpack`/`esbuild` bundle for extension host code to avoid dynamic imports in VS Code; keep it separate from Vite webview build.

### Testing Checklist
- Open/save DOCX/XLSX/PPTX/CSV; verify CSV retains CSV format.
- Paste image handling via `writeFile` still works and paths resolve with webview schemes.
- Large file chunking performance and memory: confirm backpressure prevents UI lock.
- Offline load (no external network) and CSP compliance for WASM/scripts.
- Cross-platform (Windows/macOS/Linux) paths, save dialogs, backups, and untitled untethered documents.
- Multiple editors open simultaneously do not leak state across webviews.
- Untitled flow: New -> Save As, backup, and revert all succeed; reload panel resends chunks and reinitializes cleanly.
