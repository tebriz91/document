declare global {
  interface Window {
    __VSCODE_BASE_URI__?: string;
  }
}

/**
 * Return a base path usable for both web and VS Code webview contexts.
 * Ensures trailing slash.
 */
export function getBasePath(): string {
  const candidate =
    typeof window !== 'undefined' && window.__VSCODE_BASE_URI__
      ? window.__VSCODE_BASE_URI__
      : typeof window !== 'undefined'
        ? window.location.pathname.startsWith('/document/')
          ? '/document/'
          : window.location.pathname === '/document'
            ? '/document/'
            : '/'
        : '/';

  return candidate.endsWith('/') ? candidate : `${candidate}/`;
}
