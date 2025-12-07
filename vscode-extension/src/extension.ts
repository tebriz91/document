import * as vscode from 'vscode';
import { OfficeEditorProvider } from './editorProvider';

export function activate(context: vscode.ExtensionContext) {
  const provider = new OfficeEditorProvider(context);

  const registration = vscode.window.registerCustomEditorProvider('ranuts.officeEditor', provider, {
    webviewOptions: { retainContextWhenHidden: true },
    supportsMultipleEditorsPerDocument: false,
  });

  const openCommand = vscode.commands.registerCommand('ranuts.openOfficeEditor', async (uri?: vscode.Uri) => {
    const targetUri = uri ?? (await promptForOfficeFile());
    if (!targetUri) {
      return;
    }

    await vscode.commands.executeCommand('vscode.openWith', targetUri, 'ranuts.officeEditor');
  });

  context.subscriptions.push(registration);
  context.subscriptions.push(openCommand);
}

export function deactivate() {
  // Nothing to clean up yet
}

async function promptForOfficeFile(): Promise<vscode.Uri | undefined> {
  const picked = await vscode.window.showOpenDialog({
    canSelectFiles: true,
    canSelectFolders: false,
    canSelectMany: false,
    filters: {
      'Office docs': ['docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt', 'csv', 'odt', 'ods', 'odp'],
      All: ['*'],
    },
  });

  return picked?.[0];
}
