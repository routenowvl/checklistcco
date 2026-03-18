import { msalInstance } from './authService';
import { InteractionRequiredAuthError } from '@azure/msal-browser';

const SCOPES = ["User.Read", "Sites.ReadWrite.All"];

// Renova o token 5 minutos antes de expirar
const REFRESH_THRESHOLD_MS = 5 * 60 * 1000;

// Deduplicação: se já há um refresh em andamento, reutiliza a mesma Promise
let activeRefreshPromise: Promise<string> | null = null;

// Intervalo do loop de refresh
let refreshIntervalId: ReturnType<typeof setInterval> | null = null;

/**
 * Retorna um token válido, renovando silenciosamente se necessário.
 * É a função central — todos os componentes devem usá-la.
 */
export const getValidToken = async (): Promise<string | null> => {
  // Se já há um refresh em andamento, aguarda o mesmo resultado
  if (activeRefreshPromise) {
    try {
      return await activeRefreshPromise;
    } catch {
      return null;
    }
  }

  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      console.warn('[TOKEN] Nenhuma conta ativa no MSAL');
      return null;
    }

    const account = accounts[0];

    // acquireTokenSilent: o MSAL usa o cache interno e renova automaticamente
    // quando o token está próximo de expirar (usa o refresh_token do Azure AD)
    activeRefreshPromise = msalInstance
      .acquireTokenSilent({ scopes: SCOPES, account })
      .then(response => {
        const token = response.accessToken;
        // Mantém window.__access_token sempre atualizado para compatibilidade
        (window as any).__access_token = token;
        activeRefreshPromise = null;

        const expiresIn = response.expiresOn
          ? Math.round((response.expiresOn.getTime() - Date.now()) / 1000 / 60)
          : '?';
        console.log(`[TOKEN] ✅ Token válido — expira em ~${expiresIn} min`);
        return token;
      })
      .catch(async (err) => {
        activeRefreshPromise = null;

        if (err instanceof InteractionRequiredAuthError) {
          console.warn('[TOKEN] Refresh silencioso requer interação — tentando forceRefresh');
          // Tenta forçar refresh antes de pedir interação
          try {
            const forced = await msalInstance.acquireTokenSilent({
              scopes: SCOPES,
              account,
              forceRefresh: true,
            });
            (window as any).__access_token = forced.accessToken;
            return forced.accessToken;
          } catch {
            // Não conseguiu renovar silenciosamente — dispara evento para o App mostrar modal
            console.error('[TOKEN] ❌ Não foi possível renovar silenciosamente');
            window.dispatchEvent(new CustomEvent('token-expired'));
            throw err;
          }
        }

        console.error('[TOKEN] Erro inesperado no acquireTokenSilent:', err.message);
        throw err;
      });

    return await activeRefreshPromise;
  } catch {
    return null;
  }
};

/**
 * Versão que lança erro se não conseguir token — para uso em operações críticas.
 */
export const getValidTokenOrThrow = async (): Promise<string> => {
  const token = await getValidToken();
  if (!token) throw new Error('Sessão expirada. Por favor, renove sua sessão.');
  return token;
};

/**
 * Inicia o loop de refresh proativo.
 * Deve ser chamado UMA VEZ após o login bem-sucedido.
 * Retorna função de cleanup para parar o loop.
 */
export const startTokenRefreshLoop = (
  onTokenRefresh: (newToken: string) => void,
  onSessionExpired: () => void
): (() => void) => {
  // Para qualquer loop anterior antes de iniciar um novo
  stopTokenRefreshLoop();

  const CHECK_INTERVAL_MS = 3 * 60 * 1000; // verifica a cada 3 minutos

  const checkAndRefresh = async () => {
    try {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length === 0) return;

      const account = accounts[0];

      // Verifica se o token atual está próximo de expirar
      const cachedTokens = msalInstance.getActiveAccount();
      // Sempre tenta acquireTokenSilent — o MSAL decide se usa cache ou renova
      const response = await msalInstance.acquireTokenSilent({
        scopes: SCOPES,
        account,
      });

      if (response?.accessToken) {
        const expiresOn = response.expiresOn;
        const timeUntilExpiry = expiresOn
          ? expiresOn.getTime() - Date.now()
          : Infinity;

        // Se expira em menos de REFRESH_THRESHOLD_MS, força renovação
        if (timeUntilExpiry < REFRESH_THRESHOLD_MS) {
          console.log('[TOKEN_LOOP] Token próximo de expirar — forçando renovação');
          const forced = await msalInstance.acquireTokenSilent({
            scopes: SCOPES,
            account,
            forceRefresh: true,
          });
          (window as any).__access_token = forced.accessToken;
          onTokenRefresh(forced.accessToken);
          console.log('[TOKEN_LOOP] ✅ Token renovado proativamente');
        } else {
          // Token ainda válido — apenas sincroniza
          (window as any).__access_token = response.accessToken;
          onTokenRefresh(response.accessToken);
        }
      }
    } catch (err: any) {
      if (err instanceof InteractionRequiredAuthError) {
        console.warn('[TOKEN_LOOP] Sessão expirada — requer interação do usuário');
        onSessionExpired();
      } else {
        // Erro de rede ou temporário — não expira a sessão, tenta na próxima iteração
        console.warn('[TOKEN_LOOP] Erro temporário no refresh:', err.message);
      }
    }
  };

  // Primeira verificação após 30s do login (dá tempo do app carregar)
  const initialTimeout = setTimeout(checkAndRefresh, 30 * 1000);

  // Verificações periódicas
  refreshIntervalId = setInterval(checkAndRefresh, CHECK_INTERVAL_MS);

  console.log('[TOKEN_LOOP] 🚀 Loop de refresh iniciado (intervalo: 3 min)');

  return () => {
    clearTimeout(initialTimeout);
    stopTokenRefreshLoop();
  };
};

/**
 * Para o loop de refresh (chamado no logout).
 */
export const stopTokenRefreshLoop = () => {
  if (refreshIntervalId !== null) {
    clearInterval(refreshIntervalId);
    refreshIntervalId = null;
    console.log('[TOKEN_LOOP] 🛑 Loop de refresh parado');
  }
};

/**
 * Força renovação imediata do token (útil antes de operações críticas longas).
 */
export const forceTokenRefresh = async (): Promise<string | null> => {
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return null;

    const response = await msalInstance.acquireTokenSilent({
      scopes: SCOPES,
      account: accounts[0],
      forceRefresh: true,
    });

    (window as any).__access_token = response.accessToken;
    console.log('[TOKEN] 🔄 Force refresh concluído');
    return response.accessToken;
  } catch (err: any) {
    console.error('[TOKEN] Force refresh falhou:', err.message);
    return null;
  }
};

/**
 * Limpa estado interno (chamado no logout).
 */
export const clearTokenState = () => {
  activeRefreshPromise = null;
  stopTokenRefreshLoop();
};
