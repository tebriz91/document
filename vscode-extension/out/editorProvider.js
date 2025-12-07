"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.OfficeEditorProvider = void 0;
const fs = __importStar(require("node:fs"));
const path = __importStar(require("node:path"));
const vscode = __importStar(require("vscode"));
class OfficeDocument {
    constructor(uri) {
        this.uri = uri;
    }
    dispose() {
        // No resources to release yet
    }
}
class OfficeEditorProvider {
    constructor(context) {
        this.context = context;
        this.documents = new Map();
        this._onDidChangeCustomDocument = new vscode.EventEmitter();
        this.onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;
    }
    async openCustomDocument(uri) {
        const document = new OfficeDocument(uri);
        this.documents.set(uri.toString(), document);
        return document;
    }
    async resolveCustomEditor(document, webviewPanel) {
        const { webview } = webviewPanel;
        webview.options = {
            enableScripts: true,
            localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'media')],
        };
        webview.html = this.setupWebview(webview, this.context.extensionUri);
        webview.onDidReceiveMessage((message) => {
            if (message?.type === 'DOCUMENT_EDITED') {
                this._onDidChangeCustomDocument.fire({
                    document,
                    undo: () => { },
                    redo: () => { },
                });
            }
        });
        webviewPanel.onDidDispose(() => {
            try {
                webview.postMessage({ type: 'CLOSE_EDITOR' });
            }
            catch {
                // Webview might already be disposed
            }
            this.documents.delete(document.uri.toString());
        });
    }
    async saveCustomDocument(document) {
        // Save is handled by the webview for now
        return Promise.resolve();
    }
    async saveCustomDocumentAs(document, destination) {
        await vscode.workspace.fs.copy(document.uri, destination, { overwrite: true });
    }
    async revertCustomDocument(document) {
        const existingPanel = vscode.window.visibleTextEditors.find((editor) => editor.document.uri.toString() === document.uri.toString());
        if (existingPanel) {
            existingPanel.document.save();
        }
    }
    async backupCustomDocument(document, context) {
        const backupUri = vscode.Uri.joinPath(context.destination, path.basename(document.uri.fsPath));
        await vscode.workspace.fs.copy(document.uri, backupUri);
        return {
            id: backupUri.toString(),
            delete: async () => {
                try {
                    await vscode.workspace.fs.delete(backupUri);
                }
                catch {
                    // Backup cleanup best-effort
                }
            },
        };
    }
    setupWebview(webview, extensionUri) {
        const mediaUri = webview.asWebviewUri(vscode.Uri.joinPath(extensionUri, 'media'));
        const nonce = getNonce();
        const htmlPath = vscode.Uri.joinPath(extensionUri, 'media', 'index.html');
        let html;
        try {
            html = fs.readFileSync(htmlPath.fsPath, 'utf8');
        }
        catch (error) {
            return this.buildMissingHtml(error);
        }
        html = html.replace('{{BASE_URI}}', mediaUri.toString());
        const cspMeta = `<meta http-equiv="Content-Security-Policy" content="
        default-src 'none';
        img-src ${webview.cspSource} blob: data:;
        script-src 'nonce-${nonce}' ${webview.cspSource} 'unsafe-eval' 'wasm-unsafe-eval';
        style-src 'unsafe-inline' ${webview.cspSource};
        worker-src blob:;
        connect-src ${webview.cspSource};
        font-src ${webview.cspSource};
      ">`;
        const baseTag = `<base href="${mediaUri.toString()}/">`;
        const baseUriScript = `<script nonce="${nonce}">window.__VSCODE_BASE_URI__ = '${mediaUri.toString()}';</script>`;
        html = html.replace('<head>', `<head>\n    ${cspMeta}\n    ${baseTag}\n    ${baseUriScript}`);
        html = html.replace(/<script(?![^>]*nonce=)/g, `<script nonce="${nonce}"`);
        return html;
    }
    buildMissingHtml(error) {
        const details = error instanceof Error ? error.message : 'Unknown error';
        return /* html */ `
      <!doctype html>
      <html lang="en">
        <head>
          <meta charset="UTF-8">
          <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline';">
          <title>Document preview unavailable</title>
          <style>
            body { font-family: sans-serif; padding: 24px; }
            code { background: #f5f5f5; padding: 2px 6px; border-radius: 4px; }
          </style>
        </head>
        <body>
          <h2>Document preview unavailable</h2>
          <p>Unable to load <code>media/index.html</code>. Build the webview assets and try again.</p>
          <p><strong>Details:</strong> ${details}</p>
        </body>
      </html>
    `;
    }
}
exports.OfficeEditorProvider = OfficeEditorProvider;
function getNonce() {
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let nonce = '';
    for (let i = 0; i < 32; i += 1) {
        nonce += characters.charAt(Math.floor(Math.random() * characters.length));
    }
    return nonce;
}
//# sourceMappingURL=editorProvider.js.map