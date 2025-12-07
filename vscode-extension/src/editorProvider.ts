import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';

class OfficeDocument implements vscode.CustomDocument {
  constructor(public readonly uri: vscode.Uri) {}

  dispose(): void {
    // No resources to release yet
  }
}

export class OfficeEditorProvider implements vscode.CustomEditorProvider<OfficeDocument> {
  private readonly documents = new Map<string, OfficeDocument>();
  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<OfficeDocument>>();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(uri: vscode.Uri): Promise<OfficeDocument> {
    const document = new OfficeDocument(uri);
    this.documents.set(uri.toString(), document);
    return document;
  }

  async resolveCustomEditor(document: OfficeDocument, webviewPanel: vscode.WebviewPanel): Promise<void> {
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
          undo: () => {},
          redo: () => {},
        });
      }
    });

    webviewPanel.onDidDispose(() => {
      try {
        webview.postMessage({ type: 'CLOSE_EDITOR' });
      } catch {
        // Webview might already be disposed
      }
      this.documents.delete(document.uri.toString());
    });
  }

  async saveCustomDocument(document: OfficeDocument): Promise<void> {
    // Save is handled by the webview for now
    return Promise.resolve();
  }

  async saveCustomDocumentAs(document: OfficeDocument, destination: vscode.Uri): Promise<void> {
    await vscode.workspace.fs.copy(document.uri, destination, { overwrite: true });
  }

  async revertCustomDocument(document: OfficeDocument): Promise<void> {
    const existingPanel = vscode.window.visibleTextEditors.find((editor) => editor.document.uri.toString() === document.uri.toString());
    if (existingPanel) {
      existingPanel.document.save();
    }
  }

  async backupCustomDocument(document: OfficeDocument, context: vscode.CustomDocumentBackupContext): Promise<vscode.CustomDocumentBackup> {
    const backupUri = vscode.Uri.joinPath(context.destination, path.basename(document.uri.fsPath));
    await vscode.workspace.fs.copy(document.uri, backupUri);

    return {
      id: backupUri.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(backupUri);
        } catch {
          // Backup cleanup best-effort
        }
      },
    };
  }

  private setupWebview(webview: vscode.Webview, extensionUri: vscode.Uri): string {
    const mediaUri = webview.asWebviewUri(vscode.Uri.joinPath(extensionUri, 'media'));
    const nonce = getNonce();

    const htmlPath = vscode.Uri.joinPath(extensionUri, 'media', 'index.html');
    let html: string;

    try {
      html = fs.readFileSync(htmlPath.fsPath, 'utf8');
    } catch (error) {
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

  private buildMissingHtml(error: unknown): string {
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

function getNonce(): string {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let nonce = '';
  for (let i = 0; i < 32; i += 1) {
    nonce += characters.charAt(Math.floor(Math.random() * characters.length));
  }
  return nonce;
}
