import React, { useState, useEffect, useRef } from 'react';
import { User } from '../types';
import { LogIn, Loader2, AlertCircle, ShieldCheck, Clock, CheckCircle2 } from 'lucide-react';
import { setCurrentUser } from '../services/storageService';
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { msalInstance } from '../services/authService';

// Configuração via variáveis de ambiente
const MAX_RETRIES = Number(process.env.LIMITAR_RETRY_LOGIN) || 5;
const LOCKOUT_MINUTES = Number(process.env.LOGIN_LOCKOUT_MINUTES) || 15;
const TURNSTILE_SITE_KEY = process.env.SITE_KEY || '';

// Chaves do localStorage
const STORAGE_RETRY_KEY = 'login_retry_count';
const STORAGE_LOCKOUT_KEY = 'login_lockout_until';

// URL da API Vercel
const getApiBaseUrl = () => {
  if (import.meta.env.VITE_VERCEL_URL) {
    return `https://${import.meta.env.VITE_VERCEL_URL}`;
  }
  return window.location.origin;
};

declare global {
  interface Window {
    turnstile?: {
      render: (container: HTMLElement | string, config: TurnstileConfig) => string;
      reset: (widgetId: string) => void;
      remove: (widgetId: string) => void;
      getResponse: (widgetId: string) => string;
      onResponse: (callback: (token: string) => void, widgetId: string) => void;
      onError: (callback: () => void) => void;
    };
  }
}

interface TurnstileConfig {
  sitekey: string;
  theme?: 'light' | 'dark' | 'auto';
  callback?: (token: string) => void;
  'error-callback'?: () => void;
  'expired-callback'?: () => void;
  language?: string;
  size?: 'normal' | 'compact' | 'flexible';
}

const MicrosoftIcon = () => (
    <svg width="20" height="20" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg">
        <path fill="#f35325" d="M1 1h10v10H1z"/><path fill="#81bc06" d="M12 1h10v10H12z"/><path fill="#05a6f0" d="M1 12h10v10H1z"/><path fill="#ffba08" d="M12 12h10v10H12z"/>
    </svg>
);

function getRetryCount(): number {
  return Number(localStorage.getItem(STORAGE_RETRY_KEY)) || 0;
}

function incrementRetry(): number {
  const current = getRetryCount();
  const next = current + 1;
  localStorage.setItem(STORAGE_RETRY_KEY, String(next));
  return next;
}

function resetRetry(): void {
  localStorage.removeItem(STORAGE_RETRY_KEY);
  localStorage.removeItem(STORAGE_LOCKOUT_KEY);
}

function setLockout(): void {
  const until = Date.now() + LOCKOUT_MINUTES * 60 * 1000;
  localStorage.setItem(STORAGE_LOCKOUT_KEY, String(until));
}

function getLockoutRemaining(): number | null {
  const lockoutUntil = Number(localStorage.getItem(STORAGE_LOCKOUT_KEY));
  if (!lockoutUntil) return null;

  const remaining = lockoutUntil - Date.now();
  if (remaining <= 0) {
    localStorage.removeItem(STORAGE_LOCKOUT_KEY);
    localStorage.removeItem(STORAGE_RETRY_KEY);
    return null;
  }
  return remaining;
}

function formatCountdown(ms: number): string {
  const totalSeconds = Math.ceil(ms / 1000);
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

/**
 * Verifica se o Turnstile está ativo (chave real, não placeholder)
 */
function isTurnstileConfigured(siteKey: string): boolean {
  if (!siteKey) return false;
  // Chaves de teste do Cloudflare começam com "1x00000000000000000000AA" ou "2x00000000000000000000AA"
  // Placeholders como "su..." ou "se..." não são chaves reais
  if (siteKey.startsWith('su') || siteKey.startsWith('se') || siteKey.includes('sua_')) return false;
  return true;
}

const Login: React.FC<{ onLogin: (user: User) => void }> = ({ onLogin }) => {
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [checkingSession, setCheckingSession] = useState(true);
  const [lockoutRemaining, setLockoutRemaining] = useState<number | null>(null);
  const [turnstileToken, setTurnstileToken] = useState<string | null>(null);
  const [turnstileReady, setTurnstileReady] = useState(false);
  const [turnstileVerified, setTurnstileVerified] = useState(false);
  const [isVerifying, setIsVerifying] = useState(false);
  const [turnstileLoaded, setTurnstileLoaded] = useState(false);

  const countdownRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const turnstileWidgetId = useRef<string | null>(null);
  const turnstileContainerRef = useRef<HTMLDivElement | null>(null);

  // Countdown do lockout
  useEffect(() => {
    if (countdownRef.current) {
      clearInterval(countdownRef.current);
      countdownRef.current = null;
    }

    if (lockoutRemaining !== null && lockoutRemaining > 0) {
      countdownRef.current = setInterval(() => {
        const remaining = getLockoutRemaining();
        if (remaining === null) {
          setLockoutRemaining(null);
          if (countdownRef.current) {
            clearInterval(countdownRef.current);
            countdownRef.current = null;
          }
        } else {
          setLockoutRemaining(remaining);
        }
      }, 1000);
    }

    return () => {
      if (countdownRef.current) {
        clearInterval(countdownRef.current);
      }
    };
  }, [lockoutRemaining]);

  // Verifica se o script do Turnstile já carregou
  useEffect(() => {
    if (!isTurnstileConfigured(TURNSTILE_SITE_KEY)) {
      console.warn('[TURNSTILE] SITE_KEY não configurada — pulando verificação');
      setTurnstileLoaded(true);
      setTurnstileVerified(true);
      setTurnstileReady(true);
      return;
    }

    // Verifica se a API do Turnstile já está disponível
    const checkTurnstile = () => {
      if (window.turnstile) {
        setTurnstileLoaded(true);
      } else {
        setTimeout(checkTurnstile, 200);
      }
    };

    const timeout = setTimeout(() => {
      if (!window.turnstile) {
        console.warn('[TURNSTILE] Script não carregou após 5s — pulando verificação');
        setTurnstileLoaded(true);
        setTurnstileVerified(true);
        setTurnstileReady(true);
      }
    }, 5000);

    checkTurnstile();
    return () => clearTimeout(timeout);
  }, []);

  // Renderiza o widget do Turnstile APÓS o DOM estar montado
  useEffect(() => {
    if (!turnstileLoaded || !window.turnstile || !turnstileContainerRef.current) return;
    if (!isTurnstileConfigured(TURNSTILE_SITE_KEY)) return;
    if (turnstileWidgetId.current) return; // Já renderizado

    // Garante que o container está no DOM
    const container = turnstileContainerRef.current;
    if (!container || !container.isConnected) {
      console.warn('[TURNSTILE] Container não está no DOM, aguardando...');
      return;
    }

    try {
      const widgetId = window.turnstile.render(container, {
        sitekey: TURNSTILE_SITE_KEY,
        theme: 'dark',
        language: 'pt-BR',
        size: 'flexible',
        callback: (token: string) => {
          console.log('[TURNSTILE] Token recebido');
          setTurnstileToken(token);
          setTurnstileVerified(true);
        },
        'error-callback': () => {
          console.error('[TURNSTILE] Erro ao carregar widget');
          setError('Não foi possível carregar a verificação de segurança. Recarregue a página.');
          setTurnstileReady(true);
        },
        'expired-callback': () => {
          console.warn('[TURNSTILE] Token expirado');
          setTurnstileToken(null);
          setTurnstileVerified(false);
        },
      });

      turnstileWidgetId.current = widgetId;
      setTurnstileReady(true);
      console.log('[TURNSTILE] Widget renderizado com sucesso');
    } catch (err: any) {
      console.error('[TURNSTILE] Erro ao renderizar widget:', err.message);
      setTurnstileReady(true);
    }
  }, [turnstileLoaded]);

  // Cleanup do widget ao desmontar
  useEffect(() => {
    return () => {
      if (turnstileWidgetId.current && window.turnstile) {
        try {
          window.turnstile.remove(turnstileWidgetId.current);
        } catch { /* ignora */ }
      }
    };
  }, []);

  // Valida o token do Turnstile na API Vercel
  const verifyTurnstile = async (token: string): Promise<boolean> => {
    setIsVerifying(true);
    try {
      const baseUrl = getApiBaseUrl();
      const response = await fetch(`${baseUrl}/api/verify-turnstile`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ token }),
      });

      const data = await response.json();

      if (data.success) {
        console.log('[TURNSTILE] ✅ Verificação server-side aprovada');
        return true;
      } else {
        console.warn('[TURNSTILE] ❌ Verificação falhou:', data.errors);
        setError('⚠️ Verificação de segurança falhou. Tente novamente.');
        if (turnstileWidgetId.current && window.turnstile) {
          window.turnstile.reset(turnstileWidgetId.current);
        }
        setTurnstileToken(null);
        setTurnstileVerified(false);
        return false;
      }
    } catch (err: any) {
      console.error('[TURNSTILE] Erro ao chamar API:', err.message);
      // Dev local sem API disponível
      if (!import.meta.env.PROD) {
        console.warn('[TURNSTILE] API indisponível em dev local — pulando validação server-side');
        return true;
      }
      setError('Erro na verificação de segurança. Tente novamente.');
      return false;
    } finally {
      setIsVerifying(false);
    }
  };

  useEffect(() => {
    const initAuth = async () => {
        try {
            await msalInstance.initialize();

            const isManualLogout = localStorage.getItem('msal_manual_logout') === 'true';

            const response = await msalInstance.handleRedirectPromise();
            if (response && response.account) {
                localStorage.removeItem('msal_manual_logout');
                resetRetry();
                onLogin({
                    email: response.account.username,
                    name: response.account.name || response.account.username,
                    accessToken: response.accessToken
                });
                return;
            }

            if (!isManualLogout) {
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    try {
                        const silentRequest = {
                            scopes: ["User.Read", "Sites.ReadWrite.All"],
                            account: accounts[0]
                        };
                        const silentResponse = await msalInstance.acquireTokenSilent(silentRequest);
                        if (silentResponse) {
                            resetRetry();
                            onLogin({
                                email: silentResponse.account.username,
                                name: silentResponse.account.name || silentResponse.account.username,
                                accessToken: silentResponse.accessToken
                            });
                            return;
                        }
                    } catch (silentError) {
                        if (silentError instanceof InteractionRequiredAuthError) {
                            console.warn("Sessão expirada ou requer interação.");
                        }
                    }
                }
            } else {
                console.log("Logout manual ativo: ignorando login silencioso.");
            }
        } catch (e) {
            console.error("Erro na inicialização do MSAL:", e);
        } finally {
            setCheckingSession(false);
        }
    };

    initAuth();
  }, [onLogin]);

  const handleMicrosoftLogin = async () => {
    // Verifica lockout
    const remaining = getLockoutRemaining();
    if (remaining !== null) {
      setLockoutRemaining(remaining);
      setError(`Muitas tentativas falhas. Aguarde ${formatCountdown(remaining)} para tentar novamente.`);
      return;
    }

    // Verifica se Turnstile foi resolvido (se configurado)
    if (isTurnstileConfigured(TURNSTILE_SITE_KEY) && (!turnstileVerified || !turnstileToken)) {
      setError('⚠️ Complete a verificação de segurança antes de entrar.');
      return;
    }

    // Valida o token server-side (se configurado)
    if (isTurnstileConfigured(TURNSTILE_SITE_KEY) && turnstileToken) {
      const isValid = await verifyTurnstile(turnstileToken);
      if (!isValid) {
        return;
      }
    }

    setIsLoggingIn(true);
    setError(null);
    try {
        const loginRequest = {
            scopes: ["User.Read", "Sites.ReadWrite.All"],
            prompt: "select_account"
        };
        const response = await msalInstance.loginPopup(loginRequest);
        if (response && response.account) {
            localStorage.removeItem('msal_manual_logout');
            resetRetry();
            setCurrentUser(response.account.username);
            onLogin({
                email: response.account.username,
                name: response.account.name || response.account.username,
                accessToken: response.accessToken
            });
        }
    } catch (err: any) {
        console.error(err);

        if (err.errorCode === 'user_cancelled' || err.message?.includes('cancelada')) {
          setError('Login cancelado pelo usuário.');
          return;
        }

        const attempts = incrementRetry();
        console.warn(`[LOGIN_RETRY] Tentativa falha ${attempts}/${MAX_RETRIES}`);

        if (attempts >= MAX_RETRIES) {
          setLockout();
          const lockoutTime = getLockoutRemaining();
          setLockoutRemaining(lockoutTime);
          setError(
            `⚠️ Muitas tentativas falhas (${MAX_RETRIES}). ` +
            `Acesso bloqueado por ${LOCKOUT_MINUTES} minutos. ` +
            `Aguarde ${formatCountdown(lockoutTime!)}.`
          );
          console.warn(`[LOGIN_LOCKOUT] Bloqueio ativado por ${LOCKOUT_MINUTES} minutos.`);

          if (turnstileWidgetId.current && window.turnstile) {
            window.turnstile.reset(turnstileWidgetId.current);
          }
          setTurnstileToken(null);
          setTurnstileVerified(false);
        } else {
          const retriesLeft = MAX_RETRIES - attempts;
          setError(
            `Falha na autenticação corporativa. ` +
            `Tentativas restantes: ${retriesLeft}`
          );

          if (turnstileWidgetId.current && window.turnstile) {
            window.turnstile.reset(turnstileWidgetId.current);
          }
          setTurnstileToken(null);
          setTurnstileVerified(false);
        }
    } finally {
        setIsLoggingIn(false);
    }
  };

  const isLocked = lockoutRemaining !== null && lockoutRemaining > 0;
  const turnstileEnabled = isTurnstileConfigured(TURNSTILE_SITE_KEY);

  if (checkingSession) {
    return (
      <div className="min-h-screen bg-white flex flex-col items-center justify-center p-6">
        <Loader2 className="text-blue-600 animate-spin mb-4" size={40} />
        <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px] animate-pulse">Sincronizando Sessão...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6 transition-colors">
      <div className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-[440px] border border-slate-100 overflow-hidden animate-fade-in">
        <div className="h-2 w-full bg-blue-600"></div>
        <div className="p-10 flex flex-col items-center">
            <div className="mb-8"><img src="https://viagroup.com.br/assets/via_group-22fac685.png" alt="VIA Group" className="max-w-[180px]"/></div>
            <h1 className="text-2xl font-black text-slate-800 mb-2 uppercase tracking-tight">Checklist CCO</h1>
            <p className="text-slate-500 text-sm mb-8">Gestão de Operações em Tempo Real</p>

            {error && (
              <div className={`w-full mb-6 p-4 rounded-2xl text-xs flex items-center gap-3 border ${
                isLocked
                  ? 'bg-amber-50 text-amber-700 border-amber-200'
                  : 'bg-red-50 text-red-600 border-red-100'
              }`}>
                {isLocked ? <Clock size={20} className="shrink-0 animate-pulse" /> : <AlertCircle size={20} className="shrink-0" />}
                <span className="font-bold">{error}</span>
              </div>
            )}

            {/* Widget Turnstile */}
            {turnstileEnabled && !isLocked && (
              <div className="w-full mb-4">
                <div
                  ref={turnstileContainerRef}
                  className="flex justify-center transition-opacity duration-300"
                  style={{
                    minHeight: turnstileReady ? 'auto' : '65px',
                    opacity: (turnstileLoaded && turnstileReady) ? 1 : 0.3,
                  }}
                >
                  {(!turnstileLoaded || !turnstileReady) && (
                    <div className="flex items-center gap-2 text-slate-400 text-xs">
                      <Loader2 size={14} className="animate-spin" />
                      <span>Carregando verificação...</span>
                    </div>
                  )}
                </div>

                {/* Status da verificação */}
                {turnstileVerified && turnstileReady && (
                  <div className="flex items-center justify-center gap-2 mt-2 text-green-600 text-xs font-medium">
                    <CheckCircle2 size={14} />
                    <span>Verificação concluída</span>
                  </div>
                )}

                {isVerifying && (
                  <div className="flex items-center justify-center gap-2 mt-2 text-blue-600 text-xs font-medium">
                    <Loader2 size={14} className="animate-spin" />
                    <span>Verificando...</span>
                  </div>
                )}
              </div>
            )}

            <button
                onClick={handleMicrosoftLogin}
                disabled={isLoggingIn || isLocked || isVerifying || (turnstileEnabled && !turnstileVerified)}
                className={`w-full py-4 rounded-2xl font-bold flex items-center justify-center gap-3 transition-all shadow-lg ${
                  isLocked || isVerifying || (turnstileEnabled && !turnstileVerified)
                    ? 'bg-slate-300 text-slate-500 opacity-50 cursor-not-allowed hover:scale-100'
                    : 'bg-slate-900 text-white hover:bg-slate-800 hover:scale-[1.02] active:scale-95 disabled:opacity-70'
                } shadow-lg`}
            >
                {isLoggingIn ? (
                  <><Loader2 className="animate-spin" /><span>Conectando...</span></>
                ) : isVerifying ? (
                  <><Loader2 className="animate-spin" /><span>Verificando segurança...</span></>
                ) : isLocked ? (
                  <><Clock size={20} /><span>Bloqueado ({formatCountdown(lockoutRemaining!)})</span></>
                ) : (
                  <><MicrosoftIcon /><span>Entrar com Microsoft</span></>
                )}
            </button>

            {getRetryCount() > 0 && !isLocked && (
              <p className="mt-4 text-[10px] text-slate-400 font-medium">
                Tentativas usadas: {getRetryCount()}/{MAX_RETRIES}
              </p>
            )}

            <div className="mt-8 text-[10px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
                <ShieldCheck size={12} className="text-blue-500" /> Protegido por Cloudflare Turnstile
            </div>
        </div>
      </div>
    </div>
  );
};

export default Login;
