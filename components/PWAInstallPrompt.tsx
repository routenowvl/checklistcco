import React, { useState, useEffect } from 'react';
import { Download, X, Smartphone } from 'lucide-react';

interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: 'accepted' | 'dismissed' }>;
}

const PWAInstallPrompt: React.FC = () => {
  const [deferredPrompt, setDeferredPrompt] = useState<BeforeInstallPromptEvent | null>(null);
  const [showPrompt, setShowPrompt] = useState(false);
  const [isInstalled, setIsInstalled] = useState(false);

  useEffect(() => {
    // Verifica se já está instalado
    if (window.matchMedia('(display-mode: standalone)').matches) {
      setIsInstalled(true);
      return;
    }

    // Verifica se foi aberto via standalone
    if ((window.navigator as any).standalone === true) {
      setIsInstalled(true);
      return;
    }

    // Listener para o evento de instalação
    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as BeforeInstallPromptEvent);
      
      // Mostra o prompt após 2 segundos se o usuário ainda não tiver instalado
      const hasDismissed = localStorage.getItem('pwa-install-dismissed');
      const dismissedAt = hasDismissed ? parseInt(hasDismissed) : 0;
      const daysSinceDismissal = (Date.now() - dismissedAt) / (1000 * 60 * 60 * 24);
      
      // Só mostra se não foi dispensado nos últimos 7 dias
      if (daysSinceDismissal > 7 || !hasDismissed) {
        setTimeout(() => setShowPrompt(true), 2000);
      }
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);

    // Listener para quando a instalação é bem-sucedida
    window.addEventListener('appinstalled', () => {
      setIsInstalled(true);
      setShowPrompt(false);
      setDeferredPrompt(null);
    });

    return () => {
      window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    };
  }, []);

  const handleInstallClick = async () => {
    if (!deferredPrompt) return;

    try {
      await deferredPrompt.prompt();
      const { outcome } = await deferredPrompt.userChoice;
      
      if (outcome === 'accepted') {
        console.log('Usuário aceitou instalar o PWA');
      } else {
        console.log('Usuário recusou instalar o PWA');
      }
      
      setShowPrompt(false);
      setDeferredPrompt(null);
    } catch (err) {
      console.error('Erro ao instalar PWA:', err);
    }
  };

  const handleDismiss = () => {
    setShowPrompt(false);
    localStorage.setItem('pwa-install-dismissed', Date.now().toString());
  };

  if (isInstalled || !showPrompt) return null;

  return (
    <div className="fixed bottom-4 right-4 z-[100] bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-2xl shadow-2xl p-4 max-w-sm animate-in slide-in-from-bottom-4 fade-in duration-300">
      <div className="flex items-start gap-4">
        <div className="p-3 bg-blue-600 rounded-xl text-white shrink-0">
          <Smartphone size={24} />
        </div>
        <div className="flex-1">
          <h4 className="font-bold text-slate-800 dark:text-white text-sm mb-1">
            Instale o Checklist CCO
          </h4>
          <p className="text-xs text-slate-500 dark:text-slate-400 mb-3">
            Tenha acesso rápido ao sistema diretamente do seu desktop. Sem precisar abrir o navegador!
          </p>
          <div className="flex gap-2">
            <button
              onClick={handleInstallClick}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold transition-all"
            >
              <Download size={14} />
              Instalar
            </button>
            <button
              onClick={handleDismiss}
              className="p-2 text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 transition-colors"
              title="Dispensar"
            >
              <X size={16} />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default PWAInstallPrompt;
