"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/extension.ts
var extension_exports = {};
__export(extension_exports, {
  activate: () => activate,
  deactivate: () => deactivate
});
module.exports = __toCommonJS(extension_exports);
var vscode2 = __toESM(require("vscode"));

// src/editorProvider.ts
var fs = __toESM(require("node:fs"));
var path = __toESM(require("node:path"));
var vscode = __toESM(require("vscode"));
var OfficeDocument = class {
  constructor(uri) {
    this.uri = uri;
  }
  dispose() {
  }
};
var OfficeEditorProvider = class {
  constructor(context) {
    this.context = context;
    this.documents = /* @__PURE__ */ new Map();
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
      localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, "media")]
    };
    webview.html = this.setupWebview(webview, this.context.extensionUri);
    webview.onDidReceiveMessage((message) => {
      if (message?.type === "DOCUMENT_EDITED") {
        this._onDidChangeCustomDocument.fire({
          document,
          undo: () => {
          },
          redo: () => {
          }
        });
      }
    });
    webviewPanel.onDidDispose(() => {
      try {
        webview.postMessage({ type: "CLOSE_EDITOR" });
      } catch {
      }
      this.documents.delete(document.uri.toString());
    });
  }
  async saveCustomDocument(document) {
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
        } catch {
        }
      }
    };
  }
  setupWebview(webview, extensionUri) {
    const mediaUri = webview.asWebviewUri(vscode.Uri.joinPath(extensionUri, "media"));
    const nonce = getNonce();
    const htmlPath = vscode.Uri.joinPath(extensionUri, "media", "index.html");
    let html;
    try {
      html = fs.readFileSync(htmlPath.fsPath, "utf8");
    } catch (error) {
      return this.buildMissingHtml(error);
    }
    html = html.replace("{{BASE_URI}}", mediaUri.toString());
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
    html = html.replace("<head>", `<head>
    ${cspMeta}
    ${baseTag}
    ${baseUriScript}`);
    html = html.replace(/<script(?![^>]*nonce=)/g, `<script nonce="${nonce}"`);
    return html;
  }
  buildMissingHtml(error) {
    const details = error instanceof Error ? error.message : "Unknown error";
    return (
      /* html */
      `
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
    `
    );
  }
};
function getNonce() {
  const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let nonce = "";
  for (let i = 0; i < 32; i += 1) {
    nonce += characters.charAt(Math.floor(Math.random() * characters.length));
  }
  return nonce;
}

// src/extension.ts
function activate(context) {
  const provider = new OfficeEditorProvider(context);
  const registration = vscode2.window.registerCustomEditorProvider("ranuts.officeEditor", provider, {
    webviewOptions: { retainContextWhenHidden: true },
    supportsMultipleEditorsPerDocument: false
  });
  const openCommand = vscode2.commands.registerCommand("ranuts.openOfficeEditor", async (uri) => {
    const targetUri = uri ?? await promptForOfficeFile();
    if (!targetUri) {
      return;
    }
    await vscode2.commands.executeCommand("vscode.openWith", targetUri, "ranuts.officeEditor");
  });
  context.subscriptions.push(registration);
  context.subscriptions.push(openCommand);
}
function deactivate() {
}
async function promptForOfficeFile() {
  const picked = await vscode2.window.showOpenDialog({
    canSelectFiles: true,
    canSelectFolders: false,
    canSelectMany: false,
    filters: {
      "Office docs": ["docx", "doc", "xlsx", "xls", "pptx", "ppt", "csv", "odt", "ods", "odp"],
      All: ["*"]
    }
  });
  return picked?.[0];
}
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  activate,
  deactivate
});
