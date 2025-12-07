declare global {
  interface Window {
    vscode?: ReturnType<typeof acquireVsCodeApi>;
    __VSCODE_BASE_URI__?: string;
  }
}

declare function acquireVsCodeApi(): any;
declare const Platform: {
  init?: (events: Record<string, (payload: unknown) => void>) => void;
};

const vscode = typeof acquireVsCodeApi === 'function' ? acquireVsCodeApi() : undefined;

if (vscode) {
  window.vscode = vscode;
}

if (typeof Platform !== 'undefined' && typeof Platform.init === 'function') {
  const originalInit = Platform.init.bind(Platform);
  Platform.init = (events) => {
    window.addEventListener('message', (event) => {
      const message = event.data;
      if (message?.type && events?.[message.type]) {
        events[message.type](message.payload);
      }
    });

    return originalInit(events);
  };
}

export {};
