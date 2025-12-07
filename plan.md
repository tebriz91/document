# Detailed VS Code Extension Implementation Plan

## Prerequisites

- Node.js 20+
- pnpm

## Progress

- [x] Phase 1 scaffold: Created `vscode-extension/` structure, extension manifest scripts, tsconfigs, and initial `OfficeEditorProvider` wiring.
- [x] Phase 2 (partial): Vite now outputs webview assets to `vscode-extension/media` with relative base for VS Code loading.
- [ ] Phase 2+: Remaining asset path adjustments, messaging bridge, save handling, and remaining phases pending.
## Phase 1: Project Structure Setup

### 1.1 Create Extension Directory Structure
```
vscode-extension/
├── src/
│   ├── extension.ts          # Extension entry point
│   ├── editorProvider.ts     # CustomEditorProvider implementation
│   └── webview/
│       ├── main.ts           # Webview bridge code
│       └── types.ts          # Shared types
├── media/                     # Built assets (output)
├── package.json              # Extension manifest
└── tsconfig.json
```

### 1.2 Extension Dependencies
Add to new `package.json`:
- `@types/vscode` and `@vscode/vsce` for extension development
- Copy dependencies from current web app [1](#0-0) 
- Add `esbuild` or `webpack` for bundling extension host code

## Phase 2: Asset Management & Build Configuration

### 2.1 Modify Vite Configuration
Update the current Vite config [2](#0-1)  to:
- Change `base: '/document'` to `base: './'` for relative paths
- Set `build.outDir: '../vscode-extension/media'`
- Ensure `publicDir: 'public'` copies all static assets [3](#0-2) 

### 2.2 Asset Path Resolution Strategy
The current code loads OnlyOffice API from a hardcoded path [4](#0-3)  and x2t WASM from [5](#0-4) . You need to:

**In `index.html`:**
- Remove hardcoded script tag [6](#0-5) 
- Add placeholder: `<script>window.__VSCODE_BASE_URI__ = '{{BASE_URI}}';</script>`
- Inject actual webview URI at runtime from extension host

**Create new `lib/path-resolver.ts`:**
```typescript
export function getBasePath(): string {
  // In VS Code webview context
  if (window.__VSCODE_BASE_URI__) {
    return window.__VSCODE_BASE_URI__;
  }
  // Fallback for web
  return '/document';
}
```

**Update asset loading in `lib/x2t.ts`:**
- Modify [5](#0-4)  to use `${getBasePath()}/wasm/x2t/x2t.js`
- Modify [4](#0-3)  to use `${getBasePath()}/web-apps/apps/api/documents/api.js`

## Phase 3: Messaging Bridge Implementation

### 3.1 Webview-Side Bridge (`webview/main.ts`)
The current messaging setup uses `Platform.init` with event handlers [7](#0-6) . Wrap this:

```typescript
// Detect VS Code environment
const vscode = acquireVsCodeApi();

// Intercept Platform.init to bridge messages
const originalInit = Platform.init;
Platform.init = (events) => {
  window.addEventListener('message', (event) => {
    const message = event.data;
    if (message.type && events[message.type]) {
      events[message.type](message.payload);
    }
  });
  
  // Also call original for web compatibility
  originalInit(events);
};
```

### 3.2 File Chunking Schema
The current implementation expects chunked files in this format [8](#0-7) . Your extension host must match this exactly:

**Extension host (`editorProvider.ts`):**
```typescript
import { MessageCodec } from 'ranuts/utils';

async function sendFileToWebview(panel: vscode.WebviewPanel, fileUri: vscode.Uri) {
  const fileData = await vscode.workspace.fs.readFile(fileUri);
  const file = new File([fileData], path.basename(fileUri.fsPath), {
    lastModified: Date.now(),
    type: getMimeType(fileUri.fsPath)
  });
  
  // Use MessageCodec.encodeFileChunked to create chunks
  const chunks = MessageCodec.encodeFileChunked(file);
  
  // Send each chunk to trigger RENDER_OFFICE handler
  for (const chunk of chunks) {
    panel.webview.postMessage({
      type: 'RENDER_OFFICE',
      payload: chunk
    });
  }
}
```

The webview receives these chunks and reconstructs the file [9](#0-8) .

## Phase 4: CustomEditorProvider Implementation

### 4.1 Editor Provider Registration
**In `extension.ts`:**
```typescript
export function activate(context: vscode.ExtensionContext) {
  const provider = new OfficeEditorProvider(context);
  
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(
      'ranuts.officeEditor',
      provider,
      {
        webviewOptions: { retainContextWhenHidden: true },
        supportsMultipleEditorsPerDocument: false
      }
    )
  );
}
```

**In `package.json` contribution:**
```json
"contributes": {
  "customEditors": [{
    "viewType": "ranuts.officeEditor",
    "displayName": "Office Editor",
    "selector": [
      { "filenamePattern": "*.docx" },
      { "filenamePattern": "*.xlsx" },
      { "filenamePattern": "*.pptx" },
      { "filenamePattern": "*.csv" }
    ]
  }]
}
```

### 4.2 Webview Panel Setup with CSP
```typescript
private setupWebview(webview: vscode.Webview, extensionUri: vscode.Uri): string {
  const mediaUri = webview.asWebviewUri(
    vscode.Uri.joinPath(extensionUri, 'media')
  );
  const nonce = getNonce();
  
  // Read built index.html
  const htmlPath = vscode.Uri.joinPath(extensionUri, 'media', 'index.html');
  let html = fs.readFileSync(htmlPath.fsPath, 'utf8');
  
  // Inject base URI and CSP
  html = html.replace('{{BASE_URI}}', mediaUri.toString());
  html = html.replace(
    '<head>',
    `<head>
      <meta http-equiv="Content-Security-Policy" content="
        default-src 'none';
        img-src ${webview.cspSource} blob: data:;
        script-src 'nonce-${nonce}' ${webview.cspSource};
        style-src 'unsafe-inline' ${webview.cspSource};
        worker-src blob:;
        connect-src ${webview.cspSource};
        font-src ${webview.cspSource};
      ">
    `
  );
  
  // Add nonce to all script tags
  html = html.replace(/<script/g, `<script nonce="${nonce}"`);
  
  return html;
}
```

## Phase 5: File Operations

### 5.1 Opening Existing Files
The extension must chunk files and send via `RENDER_OFFICE` event (see Phase 3.2). The webview handler [9](#0-8)  will:
1. Collect all chunks
2. Reconstruct file using `MessageCodec.decodeFileChunked`
3. Initialize x2t converter [10](#0-9) 
4. Process document [11](#0-10) 

### 5.2 Creating New Documents
The current code exposes `window.onCreateNew` [12](#0-11) . For new documents:

**Extension sends message:**
```typescript
panel.webview.postMessage({
  type: 'CREATE_NEW',
  payload: { extension: '.docx' }
});
```

**Webview handler:**
```typescript
events.CREATE_NEW = async (data: { extension: string }) => {
  await window.onCreateNew(data.extension);
};
```

### 5.3 Save Flow Modification
The current save handler downloads files [13](#0-12) . Modify to:

**In webview (`lib/x2t.ts` modification):**
```typescript
async function handleSaveDocument(event: SaveEvent) {
  const { data, option } = event.data;
  const { fileName } = getDocmentObj();
  const outputFormat = c_oAscFileType2[option.outputformat];
  
  // In VS Code, post message instead of downloading
  if (window.vscode) {
    window.vscode.postMessage({
      type: 'SAVE_DOCUMENT',
      payload: {
        bin: Array.from(data.data), // Convert Uint8Array to regular array
        fileName,
        outputFormat
      }
    });
  } else {
    // Web fallback
    await convertBinToDocumentAndDownload(data.data, fileName, outputFormat);
  }
  
  window.editor?.sendCommand({
    command: 'asc_onSaveCallback',
    data: { err_code: 0 }
  });
}
```

**Extension host receives save:**
```typescript
panel.webview.onDidReceiveMessage(async message => {
  if (message.type === 'SAVE_DOCUMENT') {
    const { bin, fileName, outputFormat } = message.payload;
    
    // Option 1: Convert in webview and receive bytes
    const uint8Array = new Uint8Array(bin);
    await vscode.workspace.fs.writeFile(document.uri, uint8Array);
    
    // Mark document as saved
    this._onDidChangeDocument.fire({
      document,
      undo: () => {},
      redo: () => {}
    });
  }
});
```

### 5.4 Dirty State Tracking
Implement `onDidChangeCustomDocument`:
```typescript
private _onDidChangeDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent>();
public readonly onDidChangeCustomDocument = this._onDidChangeDocument.event;

// Listen for any edit events from webview
panel.webview.onDidReceiveMessage(message => {
  if (message.type === 'DOCUMENT_EDITED') {
    this._onDidChangeDocument.fire({
      document,
      undo: () => {},
      redo: () => {}
    });
  }
});
```

Trigger from webview when OnlyOffice reports changes.

## Phase 6: Lifecycle Management

### 6.1 Editor Cleanup
The current code has cleanup logic [14](#0-13) . Hook into webview disposal:

```typescript
panel.onDidDispose(() => {
  // Send cleanup message
  panel.webview.postMessage({ type: 'CLOSE_EDITOR' });
  
  // Clean up resources
  this._documents.delete(document.uri.toString());
});
```

### 6.2 Prevent Race Conditions
The current implementation has queuing concerns. Add operation queue:
```typescript
private operationQueue: Promise<void> = Promise.resolve();

async openCustomDocument(uri: vscode.Uri) {
  this.operationQueue = this.operationQueue.then(async () => {
    // Open logic here
  });
  return this.operationQueue;
}
```

## Phase 7: Image Paste Handling

The current code handles pasted images via `writeFile` event [15](#0-14) . This creates blob URLs [16](#0-15) . This should work as-is in webview context, but verify blob: URLs are allowed in CSP.

## Phase 8: Build Scripts

### 8.1 Update `package.json` Scripts
```json
{
  "scripts": {
    "build:webview": "vite build",
    "build:extension": "tsc -p ./tsconfig.extension.json && esbuild src/extension.ts --bundle --outfile=out/extension.js --external:vscode --format=cjs --platform=node",
    "build": "npm run build:webview && npm run build:extension",
    "package": "vsce package",
    "watch:webview": "vite build --watch",
    "watch:extension": "tsc -p ./tsconfig.extension.json --watch"
  }
}
```

### 8.2 Copy Public Assets
Ensure Vite copies [3](#0-2)  including:
- `web-apps/` - OnlyOffice API
- `wasm/x2t/` - x2t converter
- `sdkjs/` - SDK files
- Any fonts required

## Phase 9: Testing Implementation

### 9.1 Test Cases Based on Current Code
1. **File Opening**: Test chunking logic matches [17](#0-16) 
2. **New Documents**: Verify all extensions in [18](#0-17) 
3. **Save Operations**: Test all file types in [19](#0-18) 
4. **Image Paste**: Verify writeFile handler [15](#0-14) 
5. **Large Files**: Test with files >100MB to validate chunking performance
6. **x2t Initialization**: Verify timeout handling [20](#0-19) 

### 9.2 Cross-Platform Testing
- Test file path sanitization [21](#0-20) 
- Verify WASM loading on Windows/Mac/Linux
- Test save dialogs with native file pickers

## Phase 10: Edge Cases & Error Handling

### 10.1 Handle Conversion Failures
The x2t converter has error handling [22](#0-21) . Surface these to VS Code:
```typescript
try {
  await convertDocument(file);
} catch (error) {
  vscode.window.showErrorMessage(`Conversion failed: ${error.message}`);
}
```

### 10.2 Control Panel Removal
The web app shows a control panel [23](#0-22) . In VS Code, hide it on load since file operations are handled by extension:
```typescript
if (window.vscode) {
  document.getElementById('control-panel')?.remove();
}
```

## Notes

**Key Architecture Points:**
1. The current app uses a global `fileChunks` array [24](#0-23)  that must be cleared properly to prevent memory leaks
2. The x2t converter is a singleton [25](#0-24) , so multiple editors share the same WASM instance
3. Empty document templates are stored as base64 strings [26](#0-25) 
4. OnlyOffice API must be loaded before x2t initialization [27](#0-26) 
5. File type detection relies on both MIME type and file extension [28](#0-27) 

**Critical Path Dependencies:**
- Vite build must complete before extension packaging
- `public/` assets must be in `media/` for webview access
- CSP must allow `blob:` for x2t Web Workers and image handling
- Message chunking size should match MessageCodec implementation (typically 1MB chunks)

**Potential Challenges:**
1. WASM loading in webview context - may need special headers
2. Font loading for PDF generation - ensure fonts are in `media/fonts/`
3. Worker initialization from blob URLs - verify CSP compatibility
4. File system operations on network drives or restricted paths
5. Memory management with multiple open documents sharing x2t instance

### Citations

**File:** package.json (L38-41)
```json
  "dependencies": {
    "ranui": "0.1.10-alpha.19",
    "ranuts": "0.1.0-alpha-22"
  }
```

**File:** vite.config.ts (L10-10)
```typescript
  base: '/document',
```

**File:** public/sdkjs (L1-3)
```text

```

**File:** lib/x2t.ts (L68-81)
```typescript
  private readonly DOCUMENT_TYPE_MAP: Record<string, DocumentType> = {
    docx: 'word',
    doc: 'word',
    odt: 'word',
    rtf: 'word',
    txt: 'word',
    xlsx: 'cell',
    xls: 'cell',
    ods: 'cell',
    csv: 'cell',
    pptx: 'slide',
    ppt: 'slide',
    odp: 'slide',
  };
```

**File:** lib/x2t.ts (L84-84)
```typescript
  private readonly SCRIPT_PATH = '/document/wasm/x2t/x2t.js';
```

**File:** lib/x2t.ts (L85-85)
```typescript
  private readonly INIT_TIMEOUT = 300000;
```

**File:** lib/x2t.ts (L193-216)
```typescript
  private sanitizeFileName(input: string): string {
    if (typeof input !== 'string' || !input.trim()) {
      return 'file.bin';
    }

    const parts = input.split('.');
    const ext = parts.pop() || 'bin';
    const name = parts.join('.');

    const illegalChars = /[/?<>\\:*|"]/g;
    // eslint-disable-next-line no-control-regex
    const controlChars = /[\x00-\x1f\x80-\x9f]/g;
    const reservedPattern = /^\.+$/;
    const unsafeChars = /[&'%!"{}[\]]/g;

    let sanitized = name
      .replace(illegalChars, '')
      .replace(controlChars, '')
      .replace(reservedPattern, '')
      .replace(unsafeChars, '');

    sanitized = sanitized.trim() || 'file';
    return `${sanitized.slice(0, 200)}.${ext}`; // 限制长度
  }
```

**File:** lib/x2t.ts (L286-286)
```typescript
    const fileExt = getExtensions(file?.type)[0] || fileName.split('.').pop() || '';
```

**File:** lib/x2t.ts (L319-322)
```typescript
    } catch (error) {
      throw new Error(`Document conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
```

**File:** lib/x2t.ts (L522-522)
```typescript
    script.src = './web-apps/apps/api/documents/api.js';
```

**File:** lib/x2t.ts (L534-534)
```typescript
const x2tConverter = new X2TConverter();
```

**File:** lib/x2t.ts (L630-645)
```typescript
async function handleSaveDocument(event: SaveEvent) {
  console.log('Save document event:', event);

  if (event.data && event.data.data) {
    const { data, option } = event.data;
    const { fileName } = getDocmentObj() || {};
    // 创建下载
    await convertBinToDocumentAndDownload(data.data, fileName, c_oAscFileType2[option.outputformat]);
  }

  // 告知编辑器保存完成
  window.editor?.sendCommand({
    command: 'asc_onSaveCallback',
    data: { err_code: 0 },
  });
}
```

**File:** lib/x2t.ts (L687-760)
```typescript
function handleWriteFile(event: any) {
  try {
    console.log('Write file event:', event);

    const { data: eventData } = event;
    if (!eventData) {
      console.warn('No data provided in writeFile event');
      return;
    }

    const {
      data: imageData, // Uint8Array 图片数据
      file: fileName, // 文件名，如 "display8image-174799443357-0.png"
      _target, // 目标对象，包含 frameOrigin 等信息
    } = eventData;

    // 验证数据
    if (!imageData || !(imageData instanceof Uint8Array)) {
      throw new Error('Invalid image data: expected Uint8Array');
    }

    if (!fileName || typeof fileName !== 'string') {
      throw new Error('Invalid file name');
    }

    // 从文件名中提取扩展名
    const fileExtension = fileName.split('.').pop()?.toLowerCase() || 'png';
    const mimeType = getMimeTypeFromExtension(fileExtension);

    // 创建 Blob 对象
    const blob = new Blob([imageData as unknown as BlobPart], { type: mimeType });

    // 创建对象 URL
    const objectUrl = window.URL.createObjectURL(blob);
    // 将图片 URL 添加到媒体映射中，使用原始文件名作为 key
    media[`media/${fileName}`] = objectUrl;
    window.editor?.sendCommand({
      command: 'asc_setImageUrls',
      data: {
        urls: media,
      },
    });

    window.editor?.sendCommand({
      command: 'asc_writeFileCallback',
      data: {
        // 图片 base64
        path: objectUrl,
        imgName: fileName,
      },
    });
    console.log(`Successfully processed image: ${fileName}, URL: ${media}`);
  } catch (error) {
    console.error('Error handling writeFile:', error);

    // 通知编辑器文件处理失败
    if (window.editor && typeof window.editor.sendCommand === 'function') {
      window.editor.sendCommand({
        command: 'asc_writeFileCallback',
        data: {
          success: false,
          error: error.message,
        },
      });
    }

    if (event.callback && typeof event.callback === 'function') {
      event.callback({
        success: false,
        error: error.message,
      });
    }
  }
}
```

**File:** lib/x2t.ts (L852-852)
```typescript
      const emptyBin = g_sEmpty_bin[`.${fileType}`];
```

**File:** index.html (L8-8)
```html
    <script src="/web-apps/apps/api/documents/api.js"></script>
```

**File:** index.ts (L9-17)
```typescript
interface RenderOfficeData {
  chunkIndex: number;
  data: string;
  lastModified: number;
  name: string;
  size: number;
  totalChunks: number;
  type: string;
}
```

**File:** index.ts (L28-59)
```typescript
let fileChunks: RenderOfficeData[] = [];

const events: Record<string, MessageHandler<any, unknown>> = {
  RENDER_OFFICE: async (data: RenderOfficeData) => {
    // Hide the control panel when rendering office
    const controlPanel = document.getElementById('control-panel');
    if (controlPanel) {
      controlPanel.style.display = 'none';
    }
    fileChunks.push(data);
    if (fileChunks.length >= data.totalChunks) {
      const { removeLoading } = showLoading();
      const file = await MessageCodec.decodeFileChunked(fileChunks);
      setDocmentObj({
        fileName: file.name,
        file: file,
        url: window.URL.createObjectURL(file),
      });
      await initX2T();
      const { fileName, file: fileBlob } = getDocmentObj();
      await handleDocumentOperation({ file: fileBlob, fileName, isNew: !fileBlob });
      fileChunks = [];
      removeLoading();
    }
  },
  CLOSE_EDITOR: () => {
    fileChunks = [];
    if (window.editor && typeof window.editor.destroyEditor === 'function') {
      window.editor.destroyEditor();
    }
  },
};
```

**File:** index.ts (L65-81)
```typescript
const onCreateNew = async (ext: string) => {
  const { removeLoading } = showLoading();
  setDocmentObj({
    fileName: 'New_Document' + ext,
    file: undefined,
  });
  await loadScript();
  await loadEditorApi();
  await initX2T();
  const { fileName, file: fileBlob } = getDocmentObj();
  await handleDocumentOperation({ file: fileBlob, fileName, isNew: !fileBlob });
  removeLoading();
};
// example: window.onCreateNew('.docx')
// example: window.onCreateNew('.xlsx')
// example: window.onCreateNew('.pptx')
window.onCreateNew = onCreateNew;
```

**File:** index.ts (L86-86)
```typescript
fileInput.accept = '.docx,.xlsx,.pptx,.doc,.xls,.ppt';
```

**File:** index.ts (L116-213)
```typescript
const createControlPanel = () => {
  // 创建控制面板容器
  const container = document.createElement('div');
  container.style.cssText = `
    width: 100%;
    background: linear-gradient(to right, #ffffff, #f8f9fa);
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
    border-bottom: 1px solid #eaeaea;
  `;

  const controlPanel = document.createElement('div');
  controlPanel.id = 'control-panel';
  controlPanel.style.cssText = `
    display: flex;
    flex-wrap: wrap;
    gap: 16px;
    padding: 20px;
    z-index: 1000;
    max-width: 1200px;
    margin: 0 auto;
    align-items: center;
  `;

  // 创建标题区域
  const titleSection = document.createElement('div');
  titleSection.style.cssText = `
    display: flex;
    align-items: center;
    gap: 12px;
    margin-right: auto;
  `;

  const logo = document.createElement('div');
  logo.style.cssText = `
    width: 32px;
    height: 32px;
    background: linear-gradient(135deg, #1890ff, #096dd9);
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: white;
    font-weight: bold;
    font-size: 16px;
  `;
  logo.textContent = 'W';
  titleSection.appendChild(logo);

  const title = document.createElement('div');
  title.style.cssText = `
    font-size: 18px;
    font-weight: 600;
    color: #1f1f1f;
  `;
  title.textContent = 'Web Office';
  titleSection.appendChild(title);

  controlPanel.appendChild(titleSection);

  // 创建按钮组
  const buttonGroup = document.createElement('div');
  buttonGroup.style.cssText = `
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    align-items: center;
  `;

  // Create upload button
  const uploadButton = document.createElement('r-button');
  uploadButton.textContent = 'Upload Document to view';
  uploadButton.addEventListener('click', onOpenDocument);
  buttonGroup.appendChild(uploadButton);

  // Create new document buttons
  const createDocxButton = document.createElement('r-button');
  createDocxButton.textContent = 'New Word';
  createDocxButton.addEventListener('click', () => onCreateNew('.docx'));
  buttonGroup.appendChild(createDocxButton);

  const createXlsxButton = document.createElement('r-button');
  createXlsxButton.textContent = 'New Excel';
  createXlsxButton.addEventListener('click', () => onCreateNew('.xlsx'));
  buttonGroup.appendChild(createXlsxButton);

  const createPptxButton = document.createElement('r-button');
  createPptxButton.textContent = 'New PowerPoint';
  createPptxButton.addEventListener('click', () => onCreateNew('.pptx'));
  buttonGroup.appendChild(createPptxButton);

  controlPanel.appendChild(buttonGroup);

  // 将控制面板添加到容器中
  container.appendChild(controlPanel);

  // 在 body 的最前面插入容器
  document.body.insertBefore(container, document.body.firstChild);
};
```
