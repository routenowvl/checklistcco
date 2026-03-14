import { msalInstance } from './authService';
import { InteractionRequiredAuthError } from '@azure/msal-browser';

const SCOPES = ["User.Read", "Sites.ReadWrite.All"];

interface TokenData {
  accessToken: string;
  expiresOn: Date | null;
  account: any;
}

let tokenRefreshPromise: Promise<TokenData> | null = null;
let lastRefreshTime: number = 0;
const REFRESH_THRESHOLD_MS = 5 * 60 * 1000; // 5 minutos antes do expiry

/**
 * Obtém um token válido, fazendo refresh silencioso se necessário
 */
export const getValidToken = async (): Promise<string | null> => {
  try {
    const accounts = msalInstance.getAllAccounts();
    
    if (accounts.length === 0) {
      console.warn('[TOKEN_SERVICE] Nenhuma conta logada');
      return null;
    }

    const account = accounts[0];
    const now = Date.now();

    // Tenta obter token do cache primeiro
    const cachedToken = await getCachedToken(account);
    if (cachedToken && isTokenValid(cachedToken)) {
      console.log('[TOKEN_SERVICE] Token válido do cache');
      return cachedToken.accessToken;
    }

    // Se não tem token válido ou está perto de expirar, faz refresh
    return await refreshAccessToken(account);
  } catch (error: any) {
    console.error('[TOKEN_SERVICE] Erro ao obter token:', error.message);
    return null;
  }
};

/**
 * Obtém token do cache do MSAL
 */
const getCachedToken = async (account: any): Promise<TokenData | null> => {
  try {
    const silentRequest = {
      scopes: SCOPES,
      account: account
    };
    
    const response = await msalInstance.acquireTokenSilent(silentRequest);
    
    if (response && response.accessToken) {
      return {
        accessToken: response.accessToken,
        expiresOn: response.expiresOn,
        account: response.account
      };
    }
    
    return null;
  } catch (error) {
    return null;
  }
};

/**
 * Verifica se o token ainda é válido (não expirou e não está perto de expirar)
 */
const isTokenValid = (tokenData: TokenData): boolean => {
  if (!tokenData || !tokenData.accessToken) {
    return false;
  }

  if (!tokenData.expiresOn) {
    // Se não tem data de expiração, assume válido por 5 minutos
    const fiveMinutesFromNow = Date.now() + 5 * 60 * 1000;
    return lastRefreshTime > 0 && (fiveMinutesFromNow - lastRefreshTime) < 5 * 60 * 1000;
  }

  const now = new Date();
  const expiresOn = tokenData.expiresOn as Date;
  const timeUntilExpiry = expiresOn.getTime() - now.getTime();

  // Token válido se ainda não expirou e tem mais que REFRESH_THRESHOLD_MS de vida
  return timeUntilExpiry > REFRESH_THRESHOLD_MS;
};

/**
 * Faz refresh do access token
 */
const refreshAccessToken = async (account: any): Promise<string> => {
  // Se já tem um refresh em andamento, retorna a mesma promise
  if (tokenRefreshPromise) {
    console.log('[TOKEN_SERVICE] Reutilizando refresh em andamento');
    return (await tokenRefreshPromise).accessToken;
  }

  tokenRefreshPromise = (async () => {
    try {
      console.log('[TOKEN_SERVICE] Solicitando novo token...');
      
      const silentRequest = {
        scopes: SCOPES,
        account: account,
        forceRefresh: false // MSAL decide quando forçar refresh
      };

      const response = await msalInstance.acquireTokenSilent(silentRequest);
      
      lastRefreshTime = Date.now();
      tokenRefreshPromise = null;
      
      console.log('[TOKEN_SERVICE] Token renovado com sucesso');
      console.log(`[TOKEN_SERVICE] Expira em: ${response.expiresOn?.toLocaleString() || 'desconhecido'}`);
      
      return {
        accessToken: response.accessToken,
        expiresOn: response.expiresOn,
        account: response.account
      };
    } catch (error: any) {
      tokenRefreshPromise = null;
      
      if (error instanceof InteractionRequiredAuthError) {
        console.warn('[TOKEN_SERVICE] Refresh silencioso falhou, requer interação');
        console.warn(`[TOKEN_SERVICE] Erro: ${error.errorCode} - ${error.message}`);
        
        // Tenta login popup como fallback
        try {
          const popupResponse = await msalInstance.loginPopup({
            scopes: SCOPES,
            prompt: 'none' // Não mostra UI se não for necessário
          });
          
          lastRefreshTime = Date.now();
          return {
            accessToken: popupResponse.accessToken,
            expiresOn: popupResponse.expiresOn,
            account: popupResponse.account
          };
        } catch (popupError: any) {
          console.error('[TOKEN_SERVICE] Login popup também falhou');
          throw new Error('Sessão expirada. Por favor, faça login novamente.');
        }
      }
      
      throw error;
    }
  })();

  const result = await tokenRefreshPromise;
  return result.accessToken;
};

/**
 * Verifica periodicamente e renova o token se necessário
 * Deve ser chamado uma vez no App.tsx
 */
export const startTokenRefreshLoop = (
  onTokenRefresh: (newToken: string) => void,
  onSessionExpired: () => void
) => {
  // Verifica a cada 2 minutos
  const CHECK_INTERVAL_MS = 2 * 60 * 1000;

  const checkAndRefresh = async () => {
    try {
      const token = await getValidToken();
      
      if (token) {
        onTokenRefresh(token);
      } else {
        console.warn('[TOKEN_SERVICE] Não foi possível obter token válido');
        onSessionExpired();
      }
    } catch (error: any) {
      console.error('[TOKEN_SERVICE] Erro no refresh loop:', error.message);
      onSessionExpired();
    }
  };

  // Primeira verificação imediata
  checkAndRefresh();

  // Verificações periódicas
  const intervalId = setInterval(checkAndRefresh, CHECK_INTERVAL_MS);

  // Retorna função para limpar o intervalo
  return () => clearInterval(intervalId);
};

/**
 * Força o refresh do token (útil antes de operações críticas)
 */
export const forceTokenRefresh = async (): Promise<string | null> => {
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      return null;
    }

    const response = await msalInstance.acquireTokenSilent({
      scopes: SCOPES,
      account: accounts[0],
      forceRefresh: true
    });

    lastRefreshTime = Date.now();
    return response.accessToken;
  } catch (error: any) {
    console.error('[TOKEN_SERVICE] Force refresh falhou:', error.message);
    return null;
  }
};

/**
 * Limpa o estado do token service
 */
export const clearTokenState = () => {
  tokenRefreshPromise = null;
  lastRefreshTime = 0;
};
