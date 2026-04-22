import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';

const setupServiceWorkerUpdateHooks = () => {
  if (typeof window === 'undefined' || !('serviceWorker' in navigator)) return;

  const forceSwUpdateCheck = async () => {
    try {
      const registrations = await navigator.serviceWorker.getRegistrations();
      await Promise.all(registrations.map((registration) => registration.update().catch(() => undefined)));
    } catch {
      // noop: falha de update de SW não deve bloquear o app
    }
  };

  // Garante checagem de atualização ao carregar e ao voltar foco para a aba.
  window.addEventListener('load', () => { void forceSwUpdateCheck(); });
  window.addEventListener('focus', () => { void forceSwUpdateCheck(); });

  let reloadedAfterControllerChange = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (reloadedAfterControllerChange) return;
    reloadedAfterControllerChange = true;
    window.location.reload();
  });
};

setupServiceWorkerUpdateHooks();

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error("Could not find root element to mount to");
}

const root = ReactDOM.createRoot(rootElement);
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
