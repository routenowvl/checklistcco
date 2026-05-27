
// @google/genai guidelines: Use direct process.env.API_KEY, no UI for keys, use correct model names.
// Correct models: 'gemini-3-flash-preview', 'gemini-3-pro-preview', 'gemini-2.5-flash-image', etc.

import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord, RouteDeparture, RouteOperationMapping, RouteConfig, NonCollection, ColetaPrevista, Motorista, ViewerAccessEntry } from '../types';
import { getBrazilDate, getBrazilISOString, getWeekString } from '../utils/dateUtils';

export interface DailyWarning {
  id: string;
  operacao: string; // Título
  celula: string;   // Email do responsável
  rota: string;
  descricao: string;
  dataOcorrencia: string; // ISO Date
  visualizado: boolean;
}

export interface SPNonCollection {
  id: string;
  Title: string;
  Rota: string;
  Data: string;
  Codigo: string;
  Produtor: string;
  Motivo: string;
  Observacao: string;
  Acao: string;
  DataAcao: string;
  UltimaColeta: string;
  Culpabilidade: string;
  Operacao: string;
  CausaRaiz?: string;
}

const SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH || "";
const VIEWER_ACCESS_LIST_CANDIDATES = [
  String(import.meta.env.VITE_VIEWER_ACCESS_LIST_NAME || '').trim(),
  'Usuário_filial',
  'Usuario_filial',
  'USUARIO_FILIAL'
].filter(Boolean);
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, { mapping: Record<string, string>, readOnly: Set<string>, internalNames: Set<string> }> = {};

// --- Persistência do columnMappingCache no sessionStorage ---
const COLUMN_SESSION_KEY = 'sp_columnCache_v1';
try {
  const raw = sessionStorage.getItem(COLUMN_SESSION_KEY);
  if (raw) {
    const parsed = JSON.parse(raw) as Record<string, { mapping: Record<string, string>; readOnly: string[]; internalNames: string[] }>;
    for (const [k, v] of Object.entries(parsed)) {
      columnMappingCache[k] = { mapping: v.mapping, readOnly: new Set(v.readOnly), internalNames: new Set(v.internalNames) };
    }
    if (Object.keys(columnMappingCache).length > 0) {
    }
  }
} catch { /* ignore */ }

// Cache para dados estáticos/semi-estáticos (10 minutos — otimizado para reduzir chamadas à API)
const dataCache: Record<string, { data: any, timestamp: number }> = {};
const CACHE_TTL = 10 * 60 * 1000; // 10 minutos (aumentado de 5 para reduzir consumo de API)
const MOTORISTAS_CACHE_TTL = 2 * 60 * 60 * 1000; // 2 horas
const MOTORISTAS_LOCAL_CACHE_KEY = 'sp_cache_motoristas_cco_v1';

// --- Persistência do dataCache no sessionStorage ---
const SESSION_CACHE_KEY = 'sp_dataCache_v1';
try {
  const raw = sessionStorage.getItem(SESSION_CACHE_KEY);
  if (raw) {
    const parsed = JSON.parse(raw) as Record<string, { data: any; timestamp: number }>;
    const now = Date.now();
    for (const [k, v] of Object.entries(parsed)) {
      if (v && v.data != null && now - v.timestamp < CACHE_TTL) {
        dataCache[k] = v;
      }
    }
  }
} catch { /* ignore */ }

const persistDataCacheToSession = (() => {
  let timer: ReturnType<typeof setTimeout> | null = null;
  return () => {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      try {
        const serializable: Record<string, { data: any; timestamp: number }> = {};
        // Só persiste as chaves principais que importam para o loading
        const mainKeys = ['routeConfigs_all', 'departures', 'routeOperationMappings'];
        for (const k of mainKeys) {
          if (dataCache[k]) serializable[k] = dataCache[k];
        }
        sessionStorage.setItem(SESSION_CACHE_KEY, JSON.stringify(serializable));
      } catch { /* quota exceeded, ignore */ }
    }, 500);
  };
})();

// Deduplicação de requisições archive em andamento (evita chamadas duplicadas ao mesmo range)
const inFlightArchiveRequests: Record<string, Promise<any>> = {};

// Debounce para evento token-expired (evita disparos múltiplos)
let lastTokenEventTime = 0;
const TOKEN_EVENT_DEBOUNCE_MS = 10000; // 10 segundos

/**
 * Dispara evento token-expired com debounce para evitar popups repetidos
 */
const dispatchTokenExpired = () => {
  const now = Date.now();
  if (now - lastTokenEventTime < TOKEN_EVENT_DEBOUNCE_MS) {
    console.warn('[TOKEN_EVENT] Ignorado (debounce)');
    return;
  }
  lastTokenEventTime = now;
  window.dispatchEvent(new CustomEvent('token-expired'));
};

/**
 * Delay para backoff exponencial
 */
const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Busca dados do cache se válido
 */
function getCachedData<T>(key: string): T | null {
  const cached = dataCache[key];
  if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
    return cached.data as T;
  }
  return null;
}

/**
 * Armazena dados no cache
 */
function setCachedData(key: string, data: any): void {
  dataCache[key] = { data, timestamp: Date.now() };
  persistDataCacheToSession();
}

function getMotoristasLocalCache(): Motorista[] | null {
  try {
    const raw = localStorage.getItem(MOTORISTAS_LOCAL_CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { timestamp: number; data: Motorista[] };
    if (!parsed || !Array.isArray(parsed.data)) return null;
    if (Date.now() - Number(parsed.timestamp || 0) > MOTORISTAS_CACHE_TTL) return null;
    return parsed.data;
  } catch {
    return null;
  }
}

function setMotoristasLocalCache(data: Motorista[]): void {
  try {
    localStorage.setItem(MOTORISTAS_LOCAL_CACHE_KEY, JSON.stringify({
      timestamp: Date.now(),
      data
    }));
  } catch {
    // Cache local é apenas otimização; erros aqui não devem quebrar o fluxo
  }
}

function clearMotoristasLocalCache(): void {
  try {
    localStorage.removeItem(MOTORISTAS_LOCAL_CACHE_KEY);
  } catch {
    // noop
  }
}

/**
 * Limpa cache específico
 */
export function clearCache(key?: string): void {
  if (key) {
    delete dataCache[key];
  } else {
    Object.keys(dataCache).forEach(k => delete dataCache[k]);
  }
}

/**
 * Limpa todas as chaves de cache que começam com o prefixo informado
 */
export function clearCacheByPrefix(prefix: string): void {
  Object.keys(dataCache).forEach(k => {
    if (k.startsWith(prefix)) delete dataCache[k];
  });
}

/**
 * Converte data/hora de vários formatos (DD/MM/YYYY HH:MM:SS, DD/MM/YYYY HH:MM, ISO) para ISO string.
 * Retorna null se a string for vazia.
 */
const convertToISO = (dateTimeStr: string): string | null => {
  if (!dateTimeStr || dateTimeStr.trim() === '') return null;
  const matchCompleto = dateTimeStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (matchCompleto) {
    const [, dia, mes, ano, hora, minuto, segundo] = matchCompleto;
    return new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), Number(segundo)).toISOString();
  }
  const matchSemSegundos = dateTimeStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
  if (matchSemSegundos) {
    const [, dia, mes, ano, hora, minuto] = matchSemSegundos;
    return new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), 0).toISOString();
  }
  const parsed = new Date(dateTimeStr);
  if (!isNaN(parsed.getTime())) return parsed.toISOString();
  return new Date().toISOString();
};

/**
 * Fetch com retry e backoff exponencial para lidar com throttling da Microsoft Graph
 */
async function graphFetch(
  endpoint: string, 
  token: string, 
  options: RequestInit = {},
  retryCount = 0,
  maxRetries = 4
) {
  const separator = endpoint.includes('?') ? '&' : '?';
  const url = endpoint.startsWith('https://')
    ? endpoint
    : `https://graph.microsoft.com/v1.0${endpoint}${options.method === 'GET' || !options.method ? `${separator}t=${Date.now()}` : ''}`;

  const headers: Record<string, string> = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    // Adicionado HonorNonIndexedQueriesWarningMayFailRandomly para corrigir erro de coluna não indexada (DataOperacao)
    'Prefer': 'HonorNonIndexedQueriesWarningMayFailOverLargeLists, HonorNonIndexedQueriesWarningMayFailRandomly'
  };

  try {
    const res = await fetch(url, { ...options, headers: { ...headers, ...options.headers } });

    if (!res.ok) {
      let errDetail = "";
      let errorCode = "";
      let retryAfter = 0;
      
      try {
        const err = await res.json();
        errDetail = err.error?.message || JSON.stringify(err);
        errorCode = err.error?.code || '';
      } catch(e) {
        errDetail = await res.text();
      }

      // Verifica header Retry-After
      retryAfter = parseInt(res.headers.get('Retry-After') || '0', 10);

      // Verifica se é erro de throttling (429) ou service unavailable (503)
      if ((res.status === 429 || res.status === 503) && retryCount < maxRetries) {
        // Backoff exponencial: 1s, 2s, 4s, 8s + jitter
        const delayTime = retryAfter > 0 
          ? retryAfter * 1000 
          : Math.min(1000 * Math.pow(2, retryCount) + Math.random() * 1000, 30000);
        
        console.warn(
          `[SHAREPOINT_THROTTLED] Tentativa ${retryCount + 1}/${maxRetries}. ` +
          `Retry after: ${retryAfter}s. Delay: ${delayTime}ms`
        );
        
        await delay(delayTime);
        return graphFetch(endpoint, token, options, retryCount + 1, maxRetries);
      }

      // Verifica se é erro de token expirado ou inválido
      if (res.status === 401 || errDetail.includes('expired') || errDetail.includes('invalid')) {
        console.error('[SHAREPOINT_API_FAILURE] Token expirado ou inválido. Status:', res.status);
        dispatchTokenExpired();
      }

      console.error(
        `[SHAREPOINT_API_FAILURE] URL: ${url} STATUS: ${res.status} ` +
        `ERROR: ${errDetail} CODE: ${errorCode}`
      );
      throw new Error(errDetail);
    }
    
    return res.status === 204 ? null : res.json();
  } catch (error: any) {
    // Se já atingiu o max de retries, lança o erro
    if (retryCount >= maxRetries) {
      console.error(
        `[SHAREPOINT_API_FAILURE] Máximo de retries atingido. ` +
        `Erro final: ${error.message}`
      );
      throw error;
    }
    throw error;
  }
}

async function getResolvedSiteId(token: string): Promise<string> {
  if (cachedSiteId) return cachedSiteId;
  const siteData = await graphFetch(`/sites/${SITE_PATH}`, token);
  cachedSiteId = siteData.id;
  return siteData.id;
}

async function findListByIdOrName(siteId: string, listName: string, token: string): Promise<any> {
  try { return await graphFetch(`/sites/${siteId}/lists/${listName}`, token); } 
  catch (e) {
    const data = await graphFetch(`/sites/${siteId}/lists`, token);
    const found = data.value.find((l: any) => 
      l.name?.toLowerCase() === listName.toLowerCase() || 
      l.displayName?.toLowerCase() === listName.toLowerCase()
    );
    if (found) return found;
  }
  throw new Error(`Lista '${listName}' não encontrada.`);
}

async function resolveViewerAccessList(siteId: string, token: string): Promise<any> {
  for (const candidate of VIEWER_ACCESS_LIST_CANDIDATES) {
    try {
      const list = await findListByIdOrName(siteId, candidate, token);
      return list;
    } catch {
      // Tenta o próximo candidato
    }
  }

  throw new Error(
    `Lista de visualização não encontrada. Tentativas: ${VIEWER_ACCESS_LIST_CANDIDATES.join(', ')}`
  );
}

function normalizeString(str: string): string {
  if (!str) return "";
  return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "").trim();
}

async function getListColumnMapping(siteId: string, listId: string, token: string, forceRefresh: boolean = false) {
  const cacheKey = `${siteId}_${listId}`;
  // Cache de column mapping é válido por toda a sessão — colunas do SharePoint não mudam frequentemente
  // Força refresh apenas se explicitamente solicitado (ex: admin alterou schema)
  if (columnMappingCache[cacheKey] && !forceRefresh) return columnMappingCache[cacheKey];

  const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
  const mapping: Record<string, string> = {};
  const readOnly = new Set<string>();
  const internalNames = new Set<string>();

  columns.value.forEach((col: any) => {
    const internalName = col.name;
    mapping[normalizeString(col.name)] = internalName;
    mapping[normalizeString(col.displayName)] = internalName;
    internalNames.add(internalName);
    if (col.readOnly || internalName.startsWith('_') || ['ID', 'Author', 'Created'].includes(internalName)) {
        if (internalName !== 'Title') readOnly.add(internalName);
    }
  });

  columnMappingCache[cacheKey] = { mapping, readOnly, internalNames };

  // Persiste column mapping no sessionStorage
  try {
    const serializable: Record<string, { mapping: Record<string, string>; readOnly: string[]; internalNames: string[] }> = {};
    for (const [k, v] of Object.entries(columnMappingCache)) {
      serializable[k] = { mapping: v.mapping, readOnly: [...v.readOnly], internalNames: [...v.internalNames] };
    }
    sessionStorage.setItem(COLUMN_SESSION_KEY, JSON.stringify(serializable));
  } catch { /* quota exceeded, ignore */ }

  return columnMappingCache[cacheKey];
}

function resolveFieldName(mapping: Record<string, string>, target: string): string {
  const normalized = normalizeString(target);
  if (normalized === 'titulo' || normalized === 'rota') {
      if (mapping['title']) return 'Title';
  }
  return mapping[normalized] || target;
}

function parseNumericId(value: unknown): number | null {
  if (value == null) return null;
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.trunc(value);
  }

  const raw = String(value).trim();
  if (!raw) return null;

  const match = raw.match(/-?\d+(?:[.,]\d+)?/);
  if (!match) return null;

  const parsed = Number(match[0].replace(',', '.'));
  if (!Number.isFinite(parsed)) return null;
  return Math.trunc(parsed);
}

function extractPlantFieldValue(
  fields: Record<string, any>,
  mapping: Record<string, string>
): any {
  const candidates = [
    resolveFieldName(mapping, 'Plant_id'),
    resolveFieldName(mapping, 'Plant Id'),
    resolveFieldName(mapping, 'PlantId'),
    resolveFieldName(mapping, 'plant_id'),
    resolveFieldName(mapping, 'IdPlant'),
    resolveFieldName(mapping, 'ID_PLANT')
  ];

  for (const candidate of candidates) {
    if (!candidate) continue;
    const value = fields?.[candidate];
    if (value != null && String(value).trim() !== '') {
      return value;
    }
  }

  for (const [key, value] of Object.entries(fields || {})) {
    const normalizedKey = normalizeString(key);
    if (normalizedKey.includes('plantid') || normalizedKey.includes('idplant')) {
      if (value != null && String(value).trim() !== '') {
        return value;
      }
    }
  }

  return null;
}

function collectPlantLikeFieldsForDebug(fields: Record<string, any>): Record<string, any> {
  const result: Record<string, any> = {};
  for (const [key, value] of Object.entries(fields || {})) {
    const normalizedKey = normalizeString(key);
    if (
      normalizedKey.includes('plant') ||
      normalizedKey.includes('filial') ||
      normalizedKey.includes('idplant')
    ) {
      result[key] = value;
    }
  }
  return result;
}

export const SharePointService = {
  parseEmailList(raw: string): string[] {
    return String(raw || '')
      .split(/[;,\s\n\r]+/)
      .map((email) => email.trim().toLowerCase())
      .filter(Boolean);
  },

  getViewerOperationFieldCandidates(
    mapping: Record<string, string>,
    readOnly?: Set<string>
  ): string[] {
    const unique = new Set<string>();
    const push = (value?: string) => {
      const field = String(value || '').trim();
      if (!field) return;
      unique.add(field);
    };

    push(resolveFieldName(mapping, 'OPERACAO'));
    push(resolveFieldName(mapping, 'OPERAÇÃO'));
    push(mapping['operacao']);
    push(resolveFieldName(mapping, 'FILIAL'));
    push(resolveFieldName(mapping, 'Filial'));
    push(mapping['filial']);

    Object.entries(mapping).forEach(([normalizedKey, internalName]) => {
      if (!normalizedKey) return;
      if (normalizedKey.startsWith('opera') || normalizedKey.startsWith('filial')) {
        push(internalName);
      }
    });

    const candidates = Array.from(unique);
    if (!readOnly) return candidates;

    const writable = candidates.filter((field) => !readOnly.has(field));
    return writable.length > 0 ? writable : candidates;
  },

  async getAllRouteConfigs(token: string, forceRefresh: boolean = false): Promise<RouteConfig[]> {
    try {
      if (forceRefresh) {
        clearCache('routeConfigs_all');
      }

      const cacheKey = 'routeConfigs_all';
      const cached = getCachedData<RouteConfig[]>(cacheKey);
      if (cached && !forceRefresh) {
        return cached;
      }

      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'getAll' })
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Erro ao buscar configs');

      const result: RouteConfig[] = (data.configs || []).map((row: any): RouteConfig => ({
        operacao: String(row.operacao || ''),
        email: String(row.email || '').toLowerCase().trim(),
        tolerancia: String(row.tolerancia || '00:00:00'),
        nomeExibicao: String(row.nome_exibicao || row.operacao || ''),
        plantId: row.plant_id != null ? Number(row.plant_id) : null,
        Conteudo: String(row.conteudo || ''),
        ConteudoNcoletas: String(row.conteudo_ncoletas || ''),
        ultimoEnvioSaida: String(row.ultimo_envio_saida || ''),
        Status: String(row.status || ''),
        Envio: String(row.envio || ''),
        Copia: String(row.copia || ''),
        UltimoEnvioResumoSaida: String(row.ultimo_envio_resumo_saida || ''),
        UltimoEnvioNcoletas: String(row.ultimo_envio_ncoleta || ''),
        quantidadeNcoletasRegistrada: Number(row.quantidade_ncoletas_registrada || 0),
        StatusResumoSaida: String(row.status_resumo_saida || ''),
        CodigoKmm: String(row.codigo_kmm || '')
      }));

      setCachedData(cacheKey, result);
      return result;
    } catch (e: any) {
      console.error('[PG_CONFIG] Erro ao buscar configs:', e.message);
      return [];
    }
  },

  async getTasks(token: string): Promise<SPTask[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any) => ({
          id: String(item.fields.id || item.id),
          Title: item.fields.Title || "Sem Título",
          Descricao: item.fields[resolveFieldName(mapping, 'Descricao')] || "",
          Categoria: item.fields[resolveFieldName(mapping, 'Categoria')] || "Geral",
          Horario: item.fields[resolveFieldName(mapping, 'Horario')] || "--:--",
          Ativa: item.fields[resolveFieldName(mapping, 'Ativa')] !== false,
          Ordem: Number(item.fields.Ordem) || 999
        })).sort((a: any, b: any) => a.Ordem - b.Ordem);
    } catch (e) { return []; }
  },

  async getOperations(token: string, userEmail: string): Promise<SPOperation[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const emailField = mapping['responsavel'] || 'Responsavel';
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || [])
          .map((item: any) => ({
            id: String(item.fields.id || item.id),
            Title: item.fields.Title || "OP",
            Ordem: Number(item.fields[resolveFieldName(mapping, 'Ordem')]) || 0,
            Email: (item.fields[emailField] || "").toString().trim()
          }))
          .filter((op: SPOperation) => op.Email.toLowerCase() === userEmail.toLowerCase().trim())
          .sort((a: SPOperation, b: SPOperation) => a.Ordem - b.Ordem);
    } catch (e) { return []; }
  },

  async getTeamMembers(token: string): Promise<string[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any) => item.fields.Title).filter(Boolean).sort();
    } catch (e) { return ['Logística 1', 'Logística 2', 'Supervisor']; }
  },

  async getRegisteredUsers(token: string, _userEmail?: string): Promise<string[]> { return this.getTeamMembers(token); },

  async ensureMatrix(token: string, tasks: SPTask[], ops: SPOperation[]): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);
    const today = getBrazilDate();
    const colData = resolveFieldName(mapping, 'DataReferencia');
    const filter = `fields/${colData} ge '${today}T00:00:00Z' and fields/${colData} le '${today}T23:59:59Z'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=999`, token);
    const existingKeys = new Set((existing.value || []).map((i: any) => i.fields.Title));

    for (const task of tasks) {
      if (!task.Ativa) continue;
      for (const op of ops) {
        const uniqueKey = `${today.replace(/-/g, '')}_${task.id}_${op.Title}`;
        if (!existingKeys.has(uniqueKey)) {
          const rawFields: any = { Title: uniqueKey, ChaveUnica: uniqueKey, DataReferencia: today + 'T12:00:00Z', TarefaID: task.id, OperacaoSigla: op.Title, Status: 'PR', Usuario: 'Sistema' };
          const fields: any = {};
          Object.keys(rawFields).forEach(key => {
            const int = resolveFieldName(mapping, key);
            if (internalNames.has(int) && (!readOnly.has(int) || int === 'Title')) fields[int] = rawFields[key];
          });
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) }).catch(() => null);
        }
      }
    }
  },

  async getStatusByDate(token: string, date: string): Promise<SPStatus[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const colData = resolveFieldName(mapping, 'DataReferencia');
        const filter = `fields/${colData} ge '${date}T00:00:00Z' and fields/${colData} le '${date}T23:59:59Z'`;
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=999`, token);
        return (data.value || []).map((item: any) => ({
          id: item.id, DataReferencia: item.fields[colData], TarefaID: String(item.fields[resolveFieldName(mapping, 'TarefaID')] || ""), OperacaoSigla: item.fields[resolveFieldName(mapping, 'OperacaoSigla')], Status: item.fields[resolveFieldName(mapping, 'Status')], Usuario: item.fields[resolveFieldName(mapping, 'Usuario')], Title: item.fields.Title
        }));
    } catch (e) { return []; }
  },

  async updateStatus(token: string, status: SPStatus): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const { mapping, readOnly, internalNames } = await getListColumnMapping(siteId, list.id, token);
    const escapedTitle = status.Title.replace(/'/g, "''");
    const filter = `fields/Title eq '${escapedTitle}'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    const fields: any = {};
    if (!existing.value?.length) {
        const raw = { Title: status.Title, ChaveUnica: status.Title, DataReferencia: new Date(status.DataReferencia).toISOString(), TarefaID: status.TarefaID, OperacaoSigla: status.OperacaoSigla, Status: status.Status, Usuario: status.Usuario };
        Object.keys(raw).forEach(k => { const int = resolveFieldName(mapping, k); if (internalNames.has(int)) fields[int] = (raw as any)[k]; });
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
    } else {
        const raw = { Status: status.Status, Usuario: status.Usuario };
        Object.keys(raw).forEach(k => { const int = resolveFieldName(mapping, k); if (internalNames.has(int) && !readOnly.has(int)) fields[int] = (raw as any)[k]; });
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${existing.value[0].id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);
    const celulaInternalName = mapping['celula'] || 'celula';
    const raw = { Title: record.resetBy || 'Reset', Data: new Date(record.timestamp).toISOString(), DadosJSON: JSON.stringify(record.tasks) };
    const fields: any = {};
    Object.keys(raw).forEach(k => { const int = resolveFieldName(mapping, k); if (internalNames.has(int)) fields[int] = (raw as any)[k]; });
    fields[celulaInternalName] = record.email;
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
  },

  async getHistory(token: string, userEmail: string, fetchAll?: boolean): Promise<HistoryRecord[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);
      const celulaField = mapping['celula'] || 'celula';
      
      
      // Busca todos os itens com paginação (SharePoint retorna max 100 por página)
      let allItems: any[] = [];
      let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;
      
      while (nextUrl) {
        const data = await graphFetch(nextUrl, token);
        allItems = allItems.concat(data.value || []);
        nextUrl = data['@odata.nextLink'] || null;
      }
      
      
      const result = allItems.map((item: any) => ({ 
        id: item.id, 
        timestamp: item.fields.Data, 
        resetBy: item.fields.Title, 
        email: (item.fields[celulaField] || "").toString().trim(), 
        tasks: JSON.parse(item.fields.DadosJSON || '[]') 
      })).filter((record: HistoryRecord) => fetchAll ? true : record.email?.toLowerCase() === userEmail.toLowerCase().trim())
        .sort((a: any, b: any) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
      
      return result;
    } catch (e) { return []; }
  },

  async getRouteConfigs(token: string, userEmail: string, forceRefresh: boolean = false): Promise<RouteConfig[]> {
    try {
        if (forceRefresh) {
            clearCache('routeConfigs');
        }

        const configs = await this.getAllRouteConfigs(token, forceRefresh);
        const result = configs.filter(c => c.email === userEmail.toLowerCase().trim());

        return result;
    } catch (e: any) {
      console.error('[SHAREPOINT] Erro ao buscar CONFIG_OPERACAO_SAIDA_DE_ROTAS:', e.message);
      // Retorna array vazio se a lista não existir
      return [];
    }
  },

  async getViewerAccessEntries(token: string, forceRefresh: boolean = false): Promise<ViewerAccessEntry[]> {
    try {
      if (forceRefresh) {
        clearCache('viewer_access_entries');
      }

      const cacheKey = 'viewer_access_entries';
      const cached = getCachedData<ViewerAccessEntry[]>(cacheKey);
      if (cached && !forceRefresh) {
        return cached;
      }

      const siteId = await getResolvedSiteId(token);
      const list = await resolveViewerAccessList(siteId, token);
      const { mapping, readOnly } = await getListColumnMapping(siteId, list.id, token);
      const operationFieldCandidates = this.getViewerOperationFieldCandidates(mapping, readOnly);

      let allItems: any[] = [];
      let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=200`;
      while (nextUrl) {
        const data = await graphFetch(nextUrl, token);
        allItems = allItems.concat(data.value || []);
        nextUrl = data['@odata.nextLink'] || null;
      }

      const rows = allItems
        .map((item: any): ViewerAccessEntry => ({
          id: String(item.id || ''),
          email: String(item.fields?.Title || '').trim().toLowerCase(),
          operacao: String(
            operationFieldCandidates
              .map((field) => item.fields?.[field])
              .find((value) => String(value || '').trim().length > 0) || ''
          ).trim()
        }))
        .filter((row) => row.id && row.email && row.operacao);

      setCachedData(cacheKey, rows);
      return rows;
    } catch (e: any) {
      console.error('[SHAREPOINT] Erro ao buscar acessos de visualização:', e.message);
      return [];
    }
  },

  async replaceViewerAccessForEmail(
    token: string,
    viewerEmail: string,
    selectedOperations: string[],
    scopeOperations: string[] = []
  ): Promise<void> {
    const normalizedEmail = String(viewerEmail || '').trim().toLowerCase();
    if (!normalizedEmail) return;

    const siteId = await getResolvedSiteId(token);
    const list = await resolveViewerAccessList(siteId, token);
    const { mapping, readOnly } = await getListColumnMapping(siteId, list.id, token);
    const operationFieldCandidates = this.getViewerOperationFieldCandidates(mapping, readOnly);

    const normalizeOp = (value: string): string =>
      String(value || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim()
        .toUpperCase();

    const opByNormalized = new Map<string, string>();
    selectedOperations
      .map((op) => String(op || '').trim())
      .filter(Boolean)
      .forEach((op) => {
        const key = normalizeOp(op);
        if (!opByNormalized.has(key)) opByNormalized.set(key, op);
      });

    const scopeSet = new Set(
      scopeOperations
        .map((op) => normalizeOp(op))
        .filter(Boolean)
    );

    const allEntries = await this.getViewerAccessEntries(token, true);
    const scopedExisting = allEntries.filter((entry) => {
      if (entry.email !== normalizedEmail) return false;
      if (scopeSet.size === 0) return true;
      return scopeSet.has(normalizeOp(entry.operacao));
    });

    const byOperation = new Map<string, ViewerAccessEntry[]>();
    scopedExisting.forEach((entry) => {
      const key = normalizeOp(entry.operacao);
      if (!byOperation.has(key)) byOperation.set(key, []);
      byOperation.get(key)!.push(entry);
    });

    const deleteIds: string[] = [];
    byOperation.forEach((entries, opKey) => {
      if (!opByNormalized.has(opKey)) {
        entries.forEach((entry) => deleteIds.push(entry.id));
        return;
      }

      // Mantém apenas um registro por (email, operação)
      entries.slice(1).forEach((entry) => deleteIds.push(entry.id));
    });

    for (const itemId of deleteIds) {
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}`, token, {
        method: 'DELETE'
      });
    }

    for (const [opKey, opValue] of opByNormalized.entries()) {
      const existing = byOperation.get(opKey);
      if (existing && existing.length > 0) continue;

      let created = false;
      let lastError: any = null;

      for (const operacaoField of operationFieldCandidates) {
        try {
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
            method: 'POST',
            body: JSON.stringify({
              fields: {
                Title: normalizedEmail,
                [operacaoField]: opValue
              }
            })
          });
          created = true;
          break;
        } catch (err: any) {
          lastError = err;
          const msg = String(err?.message || '').toLowerCase();
          if (
            msg.includes('not recognized') ||
            msg.includes('does not exist') ||
            msg.includes('invalid request')
          ) {
            continue;
          }
          throw err;
        }
      }

      if (!created) {
        throw new Error(
          `Não foi possível mapear a coluna de operação na lista de visualização. Último erro: ${String(lastError?.message || lastError || 'desconhecido')}`
        );
      }
    }

    clearCache('viewer_access_entries');
  },

  async getRouteOperationMappings(token: string): Promise<RouteOperationMapping[]> {
    try {
        const cacheKey = 'routeOperationMappings';
        const cached = getCachedData<RouteOperationMapping[]>(cacheKey);
        if (cached) return cached;

        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Rotas_Operacao_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        
        const result = (data.value || []).map((item: any) => ({ 
          id: item.id, 
          Title: item.fields.Title, 
          OPERACAO: item.fields[resolveFieldName(mapping, 'OPERACAO')] 
        }));
        
        setCachedData(cacheKey, result);
        return result;
    } catch (e) { return []; }
  },

  async getMotoristas(token: string, forceRefresh: boolean = false): Promise<Motorista[]> {
    try {
      const memoryCacheKey = 'motoristas_cco';

      if (!forceRefresh) {
        const memoryCached = getCachedData<Motorista[]>(memoryCacheKey);
        if (memoryCached) return memoryCached;

        const localCached = getMotoristasLocalCache();
        if (localCached) {
          setCachedData(memoryCacheKey, localCached);
          return localCached;
        }
      }

      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'motoristas_cco', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);

      const nomeField = resolveFieldName(mapping, 'nome');
      const operacaoField = resolveFieldName(mapping, 'operacao');
      const contatoField = resolveFieldName(mapping, 'Contato');

      let allItems: any[] = [];
      let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=200`;

      while (nextUrl) {
        const data = await graphFetch(nextUrl, token);
        allItems = allItems.concat(data.value || []);
        nextUrl = data['@odata.nextLink'] || null;
      }

      const result: Motorista[] = allItems
        .map((item: any) => ({
          id: String(item.id || ''),
          codigo: String(item.fields?.ID ?? item.id ?? ''),
          motorista: String(item.fields?.[nomeField] || '').trim(),
          operacao: String(item.fields?.[operacaoField] || '').trim(),
          contato: String(item.fields?.[contatoField] || '').replace(/\D/g, '')
        }))
        .filter((item: Motorista) => item.id !== '')
        .sort((a, b) => Number(a.codigo) - Number(b.codigo));

      dataCache[memoryCacheKey] = { data: result, timestamp: Date.now() };
      setMotoristasLocalCache(result);
      return result;
    } catch (e) {
      return [];
    }
  },

  async getRouteConfigsByAccess(token: string, userEmail: string, forceRefresh: boolean = false): Promise<{ configs: RouteConfig[]; canEdit: boolean; isAllViewer: boolean }> {
    try {
      const normalizedEmail = String(userEmail || '').trim().toLowerCase();
      const allConfigs = await this.getAllRouteConfigs(token, forceRefresh);

      const editableConfigs = allConfigs.filter((config) => config.email === normalizedEmail);
      if (editableConfigs.length > 0) {
        return { configs: editableConfigs, canEdit: true, isAllViewer: false };
      }

      // Modo visualização: acesso concedido pela lista de cadastro (Title=email, OPERACAO)
      const viewerEntries = await this.getViewerAccessEntries(token, forceRefresh);
      const userViewerOps = viewerEntries
        .filter((entry) => entry.email === normalizedEmail)
        .map((entry) => String(entry.operacao || '').trim().toUpperCase())
        .filter(Boolean);

      // Se o usuário tem "ALL" na OPERACAO, retorna todas as configs como read-only
      if (userViewerOps.includes('ALL')) {
        return { configs: allConfigs, canEdit: false, isAllViewer: true };
      }

      const allowedOps = new Set(userViewerOps);
      const readableConfigs = allConfigs.filter((config) =>
        allowedOps.has(String(config.operacao || '').trim().toUpperCase())
      );

      return { configs: readableConfigs, canEdit: false, isAllViewer: false };
    } catch (e: any) {
      console.error('[SHAREPOINT] Erro ao buscar configs por acesso:', e.message);
      return { configs: [], canEdit: false, isAllViewer: false };
    }
  },

  async updateMotorista(token: string, id: string, fieldsToUpdate: { operacao?: string; contato?: string }): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'motoristas_cco', token);
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

    const payload: Record<string, string> = {};

    if (fieldsToUpdate.operacao !== undefined) {
      const opField = resolveFieldName(mapping, 'operacao');
      if (internalNames.has(opField) && !readOnly.has(opField)) {
        payload[opField] = fieldsToUpdate.operacao;
      }
    }

    if (fieldsToUpdate.contato !== undefined) {
      const contatoField = resolveFieldName(mapping, 'Contato');
      if (internalNames.has(contatoField) && !readOnly.has(contatoField)) {
        payload[contatoField] = String(fieldsToUpdate.contato).replace(/\D/g, '');
      }
    }

    if (Object.keys(payload).length === 0) return;

    await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${id}/fields`, token, {
      method: 'PATCH',
      body: JSON.stringify(payload)
    });

    delete dataCache['motoristas_cco'];
    clearMotoristasLocalCache();
  },

  async addRouteOperationMapping(token: string, routeName: string, operation: string): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Rotas_Operacao_Checklist', token);
    const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);
    const fields: any = { Title: routeName, [resolveFieldName(mapping, 'OPERACAO')]: operation };
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
    
    // Invalida cache após adicionar
    clearCache('routeOperationMappings');
  },

  /**
   * Atualiza o campo UltimoEnvioSaida na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateUltimoEnvioSaida(token: string, operacao: string, dataHoraEnvio: string): Promise<void> {
    try {
      const dataISO = convertToISO(dataHoraEnvio);
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateField', operacao, field: 'ultimo_envio_saida', value: dataISO })
      });
      const data = await res.json();
      if (!data.success) console.warn('[PG_CONFIG] Falha ao atualizar UltimoEnvioSaida:', data.error);
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar UltimoEnvioSaida:', error.message);
    }
  },

  /**
   * Atualiza o campo UltimoEnvioNcoletas na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateUltimoEnvioNaoColetas(token: string, operacao: string, dataHoraEnvio: string): Promise<void> {
    try {
      const dataISO = convertToISO(dataHoraEnvio);
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateField', operacao, field: 'ultimo_envio_ncoleta', value: dataISO })
      });
      const data = await res.json();
      if (!data.success) console.warn('[PG_CONFIG] Falha ao atualizar UltimoEnvioNcoletas:', data.error);
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar UltimoEnvioNcoletas:', error.message);
    }
  },

  /**
   * Atualiza o campo quantidade_ncoletas_registrada no operacao_config
   */
  async updateQuantidadeNcoletasRegistrada(token: string, operacao: string, quantidade: number): Promise<void> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateField', operacao, field: 'quantidade_ncoletas_registrada', value: quantidade })
      });
      const data = await res.json();
      if (!data.success) console.warn('[PG_CONFIG] Falha ao atualizar quantidade_ncoletas_registrada:', data.error);
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar quantidade_ncoletas_registrada:', error.message);
    }
  },

  /**
   * Atualiza o campo Status na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateStatusOperacao(token: string, operacao: string, status: string): Promise<void> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateField', operacao, field: 'status', value: status })
      });
      const data = await res.json();
      if (!data.success) console.warn('[PG_CONFIG] Falha ao atualizar Status:', data.error);
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar Status:', error.message);
    }
  },

  /**
   * Atualiza os campos Envio e Copia na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateRouteConfigEmails(token: string, operacao: string, envio: string, copia: string): Promise<void> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateFields', operacao, fields: { envio, copia } })
      });
      const data = await res.json();
      if (!data.success) {
        console.warn('[PG_CONFIG] Falha ao atualizar emails:', data.error);
        throw new Error(data.error || 'Erro ao atualizar emails');
      }
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar emails:', error.message);
      throw error;
    }
  },

  /**
   * Atualiza a coluna Conteudo da configuração da operação apenas quando houver mudança.
   * Retorna true quando houve PATCH, false quando manteve sem alteração.
   */
  async updateRouteConfigConteudoIfChanged(token: string, operacao: string, conteudo: string): Promise<boolean> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateConteudoIfChanged', operacao, conteudo })
      });
      const data = await res.json();
      if (data.success && data.changed) {
        clearCache('routeConfigs_all');
        clearCache('routeConfigs');
      }
      return data.changed || false;
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar Conteudo:', error.message);
      return false;
    }
  },

  /**
   * Atualiza a coluna ConteudoNcoletas da configuração da operação apenas quando houver mudança.
   * Retorna true quando houve PATCH, false quando manteve sem alteração.
   */
  async updateRouteConfigConteudoNcoletasIfChanged(token: string, operacao: string, conteudoNcoletas: string): Promise<boolean> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateConteudoNcoletasIfChanged', operacao, conteudoNcoletas })
      });
      const data = await res.json();
      if (data.success && data.changed) {
        clearCache('routeConfigs_all');
        clearCache('routeConfigs');
      }
      return data.changed || false;
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar ConteudoNcoletas:', error.message);
      return false;
    }
  },

  /**
   * Atualiza o campo UltimoEnvioResumoSaida na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateUltimoEnvioResumoSaida(token: string, operacao: string, dataHoraEnvio: string): Promise<void> {
    try {
      const dataISO = convertToISO(dataHoraEnvio);
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateField', operacao, field: 'ultimo_envio_resumo_saida', value: dataISO })
      });
      const data = await res.json();
      if (!data.success) console.warn('[PG_CONFIG] Falha ao atualizar UltimoEnvioResumoSaida:', data.error);
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar UltimoEnvioResumoSaida:', error.message);
    }
  },

  /**
   * Atualiza o campo StatusResumoSaida na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateStatusResumoSaida(token: string, operacao: string, status: string): Promise<void> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'updateField', operacao, field: 'status_resumo_saida', value: status })
      });
      const data = await res.json();
      if (!data.success) console.warn('[PG_CONFIG] Falha ao atualizar StatusResumoSaida:', data.error);
    } catch (error: any) {
      console.error('[PG_CONFIG] ❌ Erro ao atualizar StatusResumoSaida:', error.message);
    }
  },

  async getDepartures(token: string, forceRefresh: boolean = false): Promise<RouteDeparture[]> {
    try {
      const cacheKey = 'departures';
      if (forceRefresh) {
        clearCache(cacheKey);
      } else {
        const cached = getCachedData<RouteDeparture[]>(cacheKey);
        if (cached) return cached;
      }

      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "departures", action: 'getAll' })
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Erro ao buscar departures');

      const formatDateFromPg = (v: any): string => {
        if (!v) return '';
        const s = String(v).trim();
        const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
      };
      const formatTimeFromPg = (v: any): string => {
        if (!v) return '';
        // Se o pg driver retornou um objeto Date, extrai no fuso de São Paulo
        if (v instanceof Date) {
          const br = new Date(v.toLocaleString('en-US', { timeZone: 'America/Sao_Paulo' }));
          const dd = String(br.getDate()).padStart(2, '0');
          const mm = String(br.getMonth() + 1).padStart(2, '0');
          const yyyy = br.getFullYear();
          const hh = String(br.getHours()).padStart(2, '0');
          const mi = String(br.getMinutes()).padStart(2, '0');
          const ss = String(br.getSeconds()).padStart(2, '0');
          return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
        }
        const s = String(v).trim();
        // ISO com T: "2026-05-24T23:48:20.000Z" ou "2026-05-24T23:48:20"
        const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
        if (isoMatch) return `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]} ${isoMatch[4]}:${isoMatch[5]}:${isoMatch[6]}`;
        // TIMESTAMP completo "YYYY-MM-DD HH:MM:SS" → "DD/MM/YYYY HH:MM:SS"
        const tsMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}:\d{2}:\d{2})/);
        if (tsMatch) return `${tsMatch[3]}/${tsMatch[2]}/${tsMatch[1]} ${tsMatch[4]}`;
        // Apenas hora "HH:MM:SS" ou "HH:MM"
        const m = s.match(/^(\d{2}):(\d{2}):?(\d{2})?/);
        return m ? `${m[1]}:${m[2]}:${m[3] || '00'}` : s;
      };

      const result: RouteDeparture[] = (data.departures || []).map((row: any): RouteDeparture => ({
        id: String(row.id),
        semana: '',
        rota: String(row.rota || ''),
        data: formatDateFromPg(row.data_operacao),
        inicio: formatTimeFromPg(row.hora_prevista),
        motorista: String(row.motorista || ''),
        codPessoa: '',
        contato: String(row.celular_motorista || '').replace(/\D/g, ''),
        placa: String(row.placa_veiculo || ''),
        saida: formatTimeFromPg(row.hora_saida),
        motivo: String(row.motivo_atraso || ''),
        observacao: String(row.observacao || ''),
        statusGeral: String(row.status_saida || ''),
        aviso: 'NÃO',
        operacao: String(row.operacao || ''),
        statusOp: String(row.status_rota || 'Previsto'),
        tempo: '',
        createdAt: String(row.criado_em || new Date().toISOString()),
        checklistMotorista: String(row.checklist_motorista || ''),
        retornoMotorista: String(row.retorno_motorista || ''),
        causaRaiz: String(row.causa_raiz || ''),
        tempoResposta: String(row.tempo_resposta || ''),
        logTempoResposta: String(row.log_tempo_resposta || '')
      }));

      setCachedData(cacheKey, result);
      return result;
    } catch (e: any) {
      console.error('[PG_DEPARTURES] Erro ao buscar departures:', e.message);
      return [];
    }
  },

  async getArchivedDepartures(token: string, operation: string | null, startDate: string, endDate: string, signal?: AbortSignal): Promise<RouteDeparture[]> {
    const cacheKey = `archived_departures_${startDate}_${endDate}_${operation || 'all'}`;

    // 1. Cache: retorna imediatamente se já buscou esse range recentemente
    const cached = getCachedData<RouteDeparture[]>(cacheKey);
    if (cached) {
      return cached;
    }

    // 2. Deduplicação: se já existe uma requisição em andamento para o mesmo range, reutiliza
    if (inFlightArchiveRequests[cacheKey]) {
      return inFlightArchiveRequests[cacheKey];
    }

    const executeQuery = async (): Promise<RouteDeparture[]> => {
      try {
        const siteId = await getResolvedSiteId(token);
        // History List ID provided by User: {856bf9d5-6081-4360-bcad-e771cbabfda8}
        const historyListId = "856bf9d5-6081-4360-bcad-e771cbabfda8";
        const { mapping } = await getListColumnMapping(siteId, historyListId, token);

        const colData = resolveFieldName(mapping, 'DataOperacao');
        const colOp = resolveFieldName(mapping, 'Operacao');

        let filter = `fields/${colData} ge '${startDate}T00:00:00Z' and fields/${colData} le '${endDate}T23:59:59Z'`;
        if (operation) {
            filter += ` and fields/${colOp} eq '${operation}'`;
        }


        // Busca todos os itens com paginação (SharePoint retorna max 100 por página)
        let allItems: any[] = [];
        let nextUrl: string | null = `/sites/${siteId}/lists/${historyListId}/items?expand=fields&$filter=${filter}&$top=100`;

        while (nextUrl) {
          // Verifica se a requisição foi cancelada antes de cada página
          if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
          const data = await graphFetch(nextUrl, token, signal ? { signal } : {});
          allItems = allItems.concat(data.value || []);
          nextUrl = data['@odata.nextLink'] || null;
        }

        const results = allItems.map((item: any) => {
          const f = item.fields;
          const dataStr = f[colData] ? f[colData].split('T')[0] : "";
          const semanaFromSharePoint = f[resolveFieldName(mapping, 'Semana')] || "";

          return {
            id: String(item.id),
            semana: semanaFromSharePoint || getWeekString(dataStr), // Calcula se não vier do SharePoint
            rota: f.Title || "",
            data: dataStr,
            inicio: f[resolveFieldName(mapping, 'HorarioInicio')] || "",
            motorista: f[resolveFieldName(mapping, 'Motorista')] || "",
            codPessoa: String(f[resolveFieldName(mapping, 'CodPessoa')] || ""),
            contato: String(f[resolveFieldName(mapping, 'Contato')] || "").replace(/\D/g, ''),
            placa: f[resolveFieldName(mapping, 'Placa')] || "",
            saida: f[resolveFieldName(mapping, 'HorarioSaida')] || "",
            motivo: f[resolveFieldName(mapping, 'MotivoAtraso')] || "",
            observacao: f[resolveFieldName(mapping, 'Observacao')] || "",
            statusGeral: f[resolveFieldName(mapping, 'StatusGeral')] || "",
            aviso: f[resolveFieldName(mapping, 'Aviso')] || "NÃO",
            operacao: f[colOp] || "",
            statusOp: f[resolveFieldName(mapping, 'StatusOp')] || "Pendente",
            tempo: f[resolveFieldName(mapping, 'TempGab')] || f[resolveFieldName(mapping, 'TempoGap')] || "",
            createdAt: f.Created || new Date().toISOString(),
            checklistMotorista: f[resolveFieldName(mapping, 'ChecklistMotorista')] || "",
            retornoMotorista: f[resolveFieldName(mapping, 'RetornoMotorista')] || f.RetornoMotorista || "",
            causaRaiz: f[resolveFieldName(mapping, 'CausaRaiz')] || "",
            tempoResposta: f[resolveFieldName(mapping, 'TempoResposta')] || "",
            logTempoResposta: f[resolveFieldName(mapping, 'LogTempoResposta')] || ""
          };
        });

        setCachedData(cacheKey, results);
        return results;
      } catch (e: any) {
        if (e.name === 'AbortError') {
          return [];
        }
        console.error("[ARCHIVE_FETCH_ERROR] Error fetching archived data:", e.message);
        throw e;
      } finally {
        delete inFlightArchiveRequests[cacheKey];
      }
    };

    const promise = executeQuery();
    inFlightArchiveRequests[cacheKey] = promise;
    return promise;
  },

  async updateDeparture(token: string, departure: RouteDeparture): Promise<string> {
    const res = await fetch('/api/checklist', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ domain: "departures", action: 'upsert', departure })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error || 'Erro ao salvar departure');

    clearCache('departures');
    return String(data.id);
  },

  async updateArchivedDeparture(token: string, departure: RouteDeparture): Promise<string> {
    const siteId = await getResolvedSiteId(token);
    const historyListId = "856bf9d5-6081-4360-bcad-e771cbabfda8";
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, historyListId, token);

    // Calcula a semana com base na data, usando a mesma lógica do Excel
    const semana = departure.semana || getWeekString(departure.data);

    const raw: any = {
        Title: departure.rota,
        Semana: semana,
        DataOperacao: departure.data ? new Date(departure.data + 'T12:00:00Z').toISOString() : null,
        HorarioInicio: departure.inicio,
        Motorista: departure.motorista,
        CodPessoa: departure.codPessoa || '',
        Contato: departure.contato || '',
        Placa: departure.placa,
        HorarioSaida: departure.saida,
        MotivoAtraso: departure.motivo,
        Observacao: departure.observacao,
        StatusGeral: departure.statusGeral,
        Aviso: departure.aviso,
        Operacao: departure.operacao,
        StatusOp: departure.statusOp,
        TempGab: departure.tempo,
        ChecklistMotorista: departure.checklistMotorista || '',
        RetornoMotorista: departure.retornoMotorista || '',
        CausaRaiz: departure.causaRaiz || '',
        TempoResposta: departure.tempoResposta || '',
        LogTempoResposta: departure.logTempoResposta || ''
    };

    const fields: any = {};
    Object.keys(raw).forEach(k => {
        const int = resolveFieldName(mapping, k);
        if (int === 'Title' || (internalNames.has(int) && !readOnly.has(int))) {
            fields[int] = raw[k];
        }
    });

    const isUpdate = departure.id && departure.id !== "" && departure.id !== "0" && !isNaN(Number(departure.id));
    let result: string;

    if (isUpdate) {
      await graphFetch(`/sites/${siteId}/lists/${historyListId}/items/${departure.id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
      result = departure.id;
    } else {
      const res = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
      result = String(res.id);
    }

    // Invalida cache de histórico após atualização
    clearCacheByPrefix('archived_departures_');
    return result;
  },

  async deleteDeparture(token: string, id: string): Promise<void> {
    const res = await fetch('/api/checklist', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ domain: "departures", action: 'delete', id })
    });
    const data = await res.json();
    if (!data.success) throw new Error(data.error || 'Erro ao deletar departure');
    clearCache('departures');
  },

  async moveDeparturesToHistory(token: string, items: RouteDeparture[]): Promise<{ success: number, failed: number, lastError?: string }> {
    const siteId = await getResolvedSiteId(token);
    const historyListId = "856bf9d5-6081-4360-bcad-e771cbabfda8";
    const { mapping: histMapping, internalNames: histInternals } = await getListColumnMapping(siteId, historyListId, token);

    let successCount = 0;
    let failedCount = 0;
    let lastErrorMessage = "";

    for (const item of items) {
        try {
            const semana = item.semana || getWeekString(item.data);
            const raw: any = {
                Title: item.rota, Semana: semana,
                DataOperacao: item.data ? (() => {
                    const d = String(item.data);
                    const dm = d.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                    const iso = dm ? `${dm[3]}-${dm[2]}-${dm[1]}` : d;
                    return new Date(iso + 'T12:00:00Z').toISOString();
                })() : null,
                HorarioInicio: item.inicio, Motorista: item.motorista,
                CodPessoa: item.codPessoa || '', Contato: item.contato || '',
                Placa: item.placa, HorarioSaida: item.saida,
                MotivoAtraso: item.motivo, Observacao: item.observacao,
                StatusGeral: item.statusGeral, Aviso: item.aviso,
                Operacao: item.operacao, StatusOp: item.statusOp,
                TempGab: item.tempo,
                ChecklistMotorista: item.checklistMotorista || '',
                RetornoMotorista: item.retornoMotorista || '',
                CausaRaiz: item.causaRaiz || '',
                TempoResposta: item.tempoResposta || ''
            };
            const histFields: any = {};
            Object.keys(raw).forEach(k => { const int = resolveFieldName(histMapping, k); if (histInternals.has(int)) histFields[int] = raw[k]; });
            const postRes = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, { method: 'POST', body: JSON.stringify({ fields: histFields }) });
            if (postRes && postRes.id) {
                // Delete from PG instead of SharePoint
                await fetch('/api/checklist', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
                  body: JSON.stringify({ domain: "departures", action: 'delete', id: item.id })
                });
                successCount++;
            } else { failedCount++; lastErrorMessage = "Failed to confirm archived ID."; }
        } catch (err: any) { failedCount++; lastErrorMessage = err.message; console.error(`[ARCHIVE_ERROR] Falha ao arquivar rota ${item.rota}:`, err.message, err?.detail || ''); }
    }
    clearCache('departures');
    return { success: successCount, failed: failedCount, lastError: lastErrorMessage };
  },

  async addDailyWarning(token: string, warning: Omit<DailyWarning, 'id' | 'visualizado'>): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'avisos_diarios_checklist', token);
    const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);
    
    const raw: any = {
        Title: warning.operacao || 'SEM OPERACAO',
        celula: warning.celula,
        rota: warning.rota,
        descricao: warning.descricao,
        data_referencia: new Date(warning.dataOcorrencia + 'T12:00:00Z').toISOString(),
        visualizado: "false" 
    };

    const fields: any = {};
    Object.keys(raw).forEach(k => {
        const int = resolveFieldName(mapping, k);
        if (internalNames.has(int)) {
            fields[int] = raw[k];
        } else if (internalNames.has(k)) {
            fields[k] = raw[k];
        }
    });

    if (!fields['Title']) fields['Title'] = raw.Title;

    try {
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { 
            method: 'POST', 
            body: JSON.stringify({ fields }) 
        });
    } catch (error: any) {
        console.error('[DEBUG ERROR] Critical failure saving warning:', error.message || error);
        throw error;
    }
  },

  async getDailyWarnings(token: string, userEmail: string): Promise<DailyWarning[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'avisos_diarios_checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        
        const celulaCol = resolveFieldName(mapping, 'celula');
        const visualizadoCol = resolveFieldName(mapping, 'visualizado');
        const rotaCol = resolveFieldName(mapping, 'rota');
        const descCol = resolveFieldName(mapping, 'descricao');
        const dataCol = resolveFieldName(mapping, 'data_referencia');

        const filter = `fields/${celulaCol} eq '${userEmail.trim()}' and fields/${visualizadoCol} eq 'false'`;
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        
        return (data.value || []).map((item: any) => {
            const f = item.fields;
            return {
                id: String(item.id),
                operacao: f.Title || "",
                celula: f[celulaCol] || "",
                rota: f[rotaCol] || "",
                descricao: f[descCol] || "",
                dataOcorrencia: f[dataCol] || "",
                visualizado: f[visualizadoCol] === 'true'
            };
        });
    } catch (e) {
        console.error("Erro ao carregar avisos:", e);
        return [];
    }
  },

  async markWarningAsViewed(token: string, id: string): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'avisos_diarios_checklist', token);
    const { mapping } = await getListColumnMapping(siteId, list.id, token);
    const visualizadoCol = resolveFieldName(mapping, 'visualizado');
    
    const fields: any = { [visualizadoCol]: "true" }; 
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
  },

  async getAllListsMetadata(token: string): Promise<any[]> {
    const siteId = await getResolvedSiteId(token);
    const results: any[] = [];
    
    const listsToQuery = [
      'Tarefas_Checklist',
      'Operacoes_Checklist',
      'Status_Checklist',
      'Historico_checklist_web',
      'Dados_Saida_de_rotas',
      'Rotas_Operacao_Checklist',
      'CONFIG_OPERACAO_SAIDA_DE_ROTAS',
      'Usuarios_cco',
      'avisos_diarios_checklist',
      '856bf9d5-6081-4360-bcad-e771cbabfda8'
    ];

    for (const listName of listsToQuery) {
      try {
        const list = await findListByIdOrName(siteId, listName, token);
        const columns = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, token);
        results.push({
          list: {
            id: list.id,
            displayName: list.displayName,
            webUrl: list.webUrl
          },
          columns: columns.value || [],
          error: false
        });
      } catch (err: any) {
        results.push({
          list: { displayName: listName },
          columns: [],
          error: true,
          errorMessage: err.message
        });
      }
    }
    return results;
  },

  /**
   * Limpa o cache de dados (útil após operações de escrita)
   */
  clearCache,

  /**
   * Verifica se há uma trava ativa para envio de resumo de uma operação
   * @returns null se não houver trava, ou objeto com info da trava
   */
  async checkSendLock(token: string, operacao: string): Promise<{ locked: boolean; user?: string; timestamp?: string; expired?: boolean } | null> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'getLockStatus', operacao })
      });
      const data = await res.json();
      if (!data.success || !data.lock) return { locked: false };

      const lockStatus = String(data.lock.lock_envio || '');
      const lockUser = String(data.lock.lock_user || '');
      const lockTimestamp = String(data.lock.lock_timestamp || '');

      if (!lockStatus || (lockStatus.toLowerCase() !== 'true' && lockStatus !== '1')) {
        return { locked: false };
      }

      // Verifica se a trava expirou (timeout de 2 minutos)
      if (lockTimestamp) {
        const lockDate = new Date(lockTimestamp);
        if (!isNaN(lockDate.getTime())) {
          const diffMs = Date.now() - lockDate.getTime();
          const timeoutMs = 2 * 60 * 1000;
          if (diffMs > timeoutMs) {
            return { locked: false, user: lockUser, timestamp: lockTimestamp, expired: true };
          }
        }
      }

      return { locked: true, user: lockUser, timestamp: lockTimestamp, expired: false };
    } catch (e: any) {
      console.error('[LOCK_CHECK] Erro ao verificar trava:', e.message);
      return { locked: false };
    }
  },

  /**
   * Adquire trava para envio de resumo
   * @returns true se conseguiu adquirir, false se outra pessoa já tem a trava
   */
  async acquireSendLock(token: string, operacao: string, userEmail: string): Promise<{ success: boolean; message?: string }> {
    try {
      // First check if already locked by someone else
      const lockInfo = await this.checkSendLock(token, operacao);
      if (lockInfo?.locked && lockInfo.user?.toLowerCase() !== userEmail.toLowerCase()) {
        return {
          success: false,
          message: `Outro usuário (${lockInfo.user}) está enviando os dados. Aguarde alguns segundos e tente novamente.`
        };
      }

      const timestamp = new Date().toISOString();
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'acquireLock', operacao, userEmail, timestamp })
      });
      const data = await res.json();
      if (!data.success) return { success: false, message: data.error || 'Erro ao adquirir trava' };

      return { success: true };
    } catch (e: any) {
      console.error('[LOCK_ACQUIRE] Erro ao adquirir trava:', e.message);
      return { success: false, message: `Erro ao adquirir trava: ${e.message}` };
    }
  },

  /**
   * Libera trava de envio
   */
  async releaseSendLock(token: string, operacao: string): Promise<void> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "config", action: 'releaseLock', operacao })
      });
      const data = await res.json();
      if (!data.success) console.warn('[LOCK_RELEASE] Falha ao liberar trava:', data.error);
    } catch (e: any) {
      console.error('[LOCK_RELEASE] Erro ao liberar trava:', e.message);
    }
  },

  /**
   * Busca não coletas da lista do SharePoint
   * Lista: Dados_Nao_Coletas (ID: 83e8cfb9-1982-47ae-b515-3fec112da457)
   */
  async getNonCollections(token: string, userEmail: string): Promise<NonCollection[]> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "non-collections", action: 'getAll' })
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Erro ao buscar non-collections');

      // Converte data ISO/YYYY-MM-DD vinda do PostgreSQL para DD/MM/YYYY
      const pgDateToBR = (v: any): string => {
        if (!v) return '';
        const s = String(v).trim();
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
        // ISO com T: "2026-05-28T15:00:00Z" ou "1970-01-01T00:00:00.000Z"
        const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
        return s;
      };

      return (data.nonCollections || []).map((row: any): NonCollection => ({
        id: String(row.id),
        semana: String(row.semana || ''),
        rota: String(row.rota || ''),
        data: pgDateToBR(row.data_operacao || row.data),
        codigo: String(row.codigo || ''),
        produtor: String(row.produtor || ''),
        motivo: String(row.motivo || ''),
        observacao: String(row.observacao || ''),
        acao: String(row.acao || ''),
        dataAcao: pgDateToBR(row.data_acao),
        ultimaColeta: pgDateToBR(row.ultima_coleta),
        Culpabilidade: String(row.culpabilidade || ''),
        operacao: String(row.operacao || ''),
        causaRaiz: String(row.causa_raiz || '')
      }));
    } catch (e: any) {
      console.error('[PG_NC] Erro ao buscar non-collections:', e.message);
      return [];
    }
  },

  /**
   * Salva não coleta na lista do SharePoint
   * Lista: Dados_Nao_Coletas (ID: 83e8cfb9-1982-47ae-b515-3fec112da457)
   * Nomes internos conforme schema XML:
   * - Semana → Title
   * - Rota → Rota
   * - Data → Data (DateTime)
   * - Código → C_x00f3_digo
   * - Produtor → Produtor
   * - Motivo → Motivo
   * - Observação → Observa_x00e7__x00e3_o
   * - Ação → A_x00e7__x00e3_o
   * - Data Ação → DataA_x00e7__x00e3_o
   * - Última Coleta → _x00da_ltimaColeta
   * - Culpabilidade → Culpabilidade
   * - Operação → Opera_x00e7__x00e3_o
   */
  async saveNonCollection(token: string, nonCollection: NonCollection): Promise<string> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "non-collections", action: 'insert', nonCollection })
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Erro ao salvar non-collection');
      return String(data.id);
    } catch (e: any) {
      console.error('[PG_NC] Erro ao salvar não coleta:', e.message);
      throw e;
    }
  },

  /**
   * Atualiza não coleta existente na lista do SharePoint
   * Lista: Dados_Nao_Coletas (ID: 83e8cfb9-1982-47ae-b515-3fec112da457)
   */
  async updateNonCollection(token: string, nonCollection: NonCollection): Promise<void> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "non-collections", action: 'update', nonCollection })
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Erro ao atualizar non-collection');
    } catch (e: any) {
      console.error('[PG_NC] Erro ao atualizar não coleta:', e.message);
      throw e;
    }
  },

  /**
   * Atualiza não coleta arquivada na lista de histórico
   * Lista: nao_coletas_web_hist (ID: 1702fe62-6a47-4fd1-b935-0e3258073bb6)
   */
  async updateArchivedNonCollection(token: string, nonCollection: NonCollection): Promise<void> {
    try {
      const siteId = await getResolvedSiteId(token);
      const historyListId = '1702fe62-6a47-4fd1-b935-0e3258073bb6';
      const { mapping } = await getListColumnMapping(siteId, historyListId, token);
      const historyObservacaoField = resolveFieldName(mapping, 'Observação');

      // Constrói payload removendo campos vazios (SharePoint rejeita DateTime com "")
      const payload: any = {};

      // Regra do histórico: Title SEMPRE representa a Rota
      if (nonCollection.rota && nonCollection.rota.trim() !== '') {
        payload.Title = nonCollection.rota;
      }
      if (nonCollection.data) {
        const parsedData = parseDateForSharePoint(nonCollection.data);
        if (parsedData) payload.Data = parsedData;
      }
      if (nonCollection.codigo) payload.C_x00f3_digo = nonCollection.codigo;
      if (nonCollection.produtor) payload.Produtor = nonCollection.produtor;
      if (nonCollection.motivo) payload.Motivo = nonCollection.motivo;
      if (nonCollection.observacao) payload[historyObservacaoField] = nonCollection.observacao;
      if (nonCollection.acao) payload.A_x00e7__x00e3_o = nonCollection.acao;
      // Campos DateTime: só envia se parse resultou em valor válido
      { const v = parseDateForSharePoint(nonCollection.dataAcao); if (v) payload.DataA_x00e7__x00e3_o = v; }
      { const v = parseDateForSharePoint(nonCollection.ultimaColeta); if (v) payload._x00da_ltimaColeta = v; }
      if (nonCollection.Culpabilidade) payload.Culpabilidade = nonCollection.Culpabilidade;
      if (nonCollection.operacao) payload.Opera_x00e7__x00e3_o = nonCollection.operacao;
      if (nonCollection.causaRaiz) payload.CausaRaiz = nonCollection.causaRaiz;

      await graphFetch(`/sites/${siteId}/lists/${historyListId}/items/${nonCollection.id}`, token, {
        method: 'PATCH',
        body: JSON.stringify({ fields: payload })
      });

      // Invalida cache das consultas de histórico de não coletas
      clearCacheByPrefix('archived_noncollections_');

    } catch (e: any) {
      console.error('[NonCollectionsHistory] Erro ao atualizar não coleta de histórico:', e.message);
      throw e;
    }
  },

  /**
   * Exclui não coleta existente da lista do SharePoint
   * Lista: Dados_Nao_Coletas (ID: 83e8cfb9-1982-47ae-b515-3fec112da457)
   */
  async deleteNonCollection(token: string, id: string): Promise<void> {
    try {
      const res = await fetch('/api/checklist', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ domain: "non-collections", action: 'delete', id })
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.error || 'Erro ao deletar non-collection');
    } catch (e: any) {
      console.error('[PG_NC] Erro ao excluir não coleta:', e.message);
      throw e;
    }
  },

  /**
   * Busca não coletas arquivadas no histórico.
   * Lista: nao_coletas_web_hist (ID: 1702fe62-6a47-4fd1-b935-0e3258073bb6)
   */
  async getArchivedNonCollections(token: string, userEmail: string, startDate: string, endDate: string, signal?: AbortSignal): Promise<NonCollection[]> {
    const cacheKey = `archived_noncollections_v2_${startDate}_${endDate}`;

    // 1. Cache: retorna imediatamente se já buscou esse range recentemente
    const cached = getCachedData<NonCollection[]>(cacheKey);
    if (cached) {
      return cached;
    }

    // 2. Deduplicação: se já existe uma requisição em andamento para o mesmo range, reutiliza
    if (inFlightArchiveRequests[cacheKey]) {
      return inFlightArchiveRequests[cacheKey];
    }

    const executeQuery = async (): Promise<NonCollection[]> => {
      try {
        const siteId = await getResolvedSiteId(token);
        const historyListId = '1702fe62-6a47-4fd1-b935-0e3258073bb6';
        const { mapping } = await getListColumnMapping(siteId, historyListId, token);

        const colData = resolveFieldName(mapping, 'Data');
        const colOp = resolveFieldName(mapping, 'Operação');

        let filter = `fields/${colData} ge '${startDate}T00:00:00Z' and fields/${colData} le '${endDate}T23:59:59Z'`;


        // Busca todos os itens com paginação
        let allItems: any[] = [];
        let nextUrl: string | null = `/sites/${siteId}/lists/${historyListId}/items?expand=fields&$filter=${filter}&$top=100`;

        while (nextUrl) {
          // Verifica se a requisição foi cancelada antes de cada página
          if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
          const data = await graphFetch(nextUrl, token, signal ? { signal } : {});
          allItems = allItems.concat(data.value || []);
          nextUrl = data['@odata.nextLink'] || null;
        }

        const results = allItems.map((item: any) => {
          const f = item.fields;
          const dataStr = f[colData] ? f[colData].split('T')[0] : "";
          const ultimaColetaStr = f[resolveFieldName(mapping, 'ÚltimaColeta')] ? f[resolveFieldName(mapping, 'ÚltimaColeta')].split('T')[0] : "";
          const dataAcaoStr = f[resolveFieldName(mapping, 'DataAção')] ? f[resolveFieldName(mapping, 'DataAção')].split('T')[0] : "";
          const observacaoValue =
            f[resolveFieldName(mapping, 'Observação')] ||
            f[resolveFieldName(mapping, 'Observacao')] ||
            f[resolveFieldName(mapping, 'Observa_x00e7__x00e3_o')] ||
            f.Observação ||
            f.Observacao ||
            f.Observa_x00e7__x00e3_o ||
            "";

          // Histórico: Title armazena a rota.
          // "Semana" é apenas informativa na UI e é calculada pela data.
          const semanaCalc = dataStr ? getWeekString(dataStr) : "";

          return {
            id: String(item.id),
            semana: semanaCalc,
            rota: f.Title || f[resolveFieldName(mapping, 'Rota')] || "",
            data: dataStr,
            codigo: f[resolveFieldName(mapping, 'Código')] || "",
            produtor: f[resolveFieldName(mapping, 'Produtor')] || "",
            motivo: f[resolveFieldName(mapping, 'Motivo')] || "",
            observacao: observacaoValue,
            acao: f[resolveFieldName(mapping, 'Ação')] || "",
            dataAcao: dataAcaoStr,
            ultimaColeta: ultimaColetaStr,
            Culpabilidade: f[resolveFieldName(mapping, 'Culpabilidade')] || "",
            operacao: f[colOp] || "",
            causaRaiz: f[resolveFieldName(mapping, 'CausaRaiz')] || ""
          };
        });

        setCachedData(cacheKey, results);
        return results;
      } catch (e: any) {
        if (e.name === 'AbortError') {
          return [];
        }
        console.error("[NC_ARCHIVE_FETCH_ERROR] Error fetching archived non-collections:", e.message);
        throw e;
      } finally {
        delete inFlightArchiveRequests[cacheKey];
      }
    };

    const promise = executeQuery();
    inFlightArchiveRequests[cacheKey] = promise;
    return promise;
  },

  /**
   * Busca coletas previstas da lista Coletas_previstas_cco.
   * Filtra por data e retorna operação + quantidade.
   */
  async getColetasPrevistas(
    token: string,
    date: string,
    userEmail: string,
    userOperations: string[] = []
  ): Promise<ColetaPrevista[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Coletas_previstas_cco', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);

      // Data em formato ISO para filtro
      const startISO = `${date}T00:00:00Z`;
      const endISO = `${date}T23:59:59Z`;
      const colData = resolveFieldName(mapping, 'Data');

      const fetchAllItems = async (initialUrl: string): Promise<any[]> => {
        let allItems: any[] = [];
        let nextUrl: string | null = initialUrl;
        let page = 0;

        while (nextUrl) {
          const data = await graphFetch(nextUrl, token);
          allItems = allItems.concat(data.value || []);
          nextUrl = data['@odata.nextLink'] || null;
          page++;
          if (page > 500) {
            console.warn('[COLETAS_PREVISTAS] Limite de segurança de paginação atingido (500 páginas).');
            break;
          }
        }

        return allItems;
      };

      const normalizeDateField = (value: any): string => {
        const raw = String(value || '').trim();
        if (!raw) return '';
        if (/^\d{4}-\d{2}-\d{2}T/.test(raw)) return raw.slice(0, 10);
        if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
          const [d, m, y] = raw.split('/');
          return `${y}-${m}-${d}`;
        }
        return '';
      };

      const normalizeOperation = (value: any): string =>
        String(value || '')
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim()
          .toUpperCase()
          .replace(/\s+/g, ' ');

      const rangeFilter = `fields/${colData} ge '${startISO}' and fields/${colData} le '${endISO}'`;

      // 1) Tentativa padrão por intervalo de data/hora
      let allItems = await fetchAllItems(
        `/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${rangeFilter}&$top=100`
      );

      // 2) Fallback para colunas Date-only (sem hora)
      if (allItems.length === 0) {
        const eqDateFilter = `fields/${colData} eq '${date}'`;
        console.warn('[COLETAS_PREVISTAS] Busca por range retornou 0. Tentando filtro date-only (eq).');
        allItems = await fetchAllItems(
          `/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${eqDateFilter}&$top=100`
        );
      }

      // 3) Fallback final: busca ampla com paginação e filtra a data no cliente
      if (allItems.length === 0) {
        console.warn('[COLETAS_PREVISTAS] Filtro no servidor retornou 0. Aplicando fallback com filtro de data no cliente.');
        const broadItems = await fetchAllItems(
          `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`
        );
        allItems = broadItems.filter((item: any) => normalizeDateField(item?.fields?.[colData]) === date);
      }

      if (allItems.length > 0) {
      }

      // Busca configurações do usuário para filtrar pelas operações dele
      const operationSource =
        userOperations.length > 0
          ? userOperations
          : (await this.getRouteConfigs(token, userEmail, true)).map(c => c.operacao);

      const myOps = new Set(operationSource.map(normalizeOperation).filter(Boolean));


      const result = (allItems || [])
        .map((item: any): ColetaPrevista => {
          const f = item.fields;
          const dataRaw = f[colData];
          const normalizedDate = normalizeDateField(dataRaw);
          const dataISO = normalizedDate ? `${normalizedDate}T12:00:00Z` : '';
          const operacaoTitle = String(f.Title || '').trim();

          return {
            id: String(item.id),
            Title: operacaoTitle,
            QntColeta: Number(f[resolveFieldName(mapping, 'QntColeta')] || 0),
            Data: dataISO
          };
        });


      const filtered = result.filter(c => myOps.size === 0 || myOps.has(normalizeOperation(c.Title)));

      if (myOps.size > 0) {
        const resultOps = new Set(result.map(r => normalizeOperation(r.Title)).filter(Boolean));
        const missingOps = Array.from(myOps).filter(op => !resultOps.has(op));
        if (missingOps.length > 0) {
          console.warn('[COLETAS_PREVISTAS] Operações do usuário sem correspondência na lista de previstas:', missingOps);
        }
      }


      return filtered;
    } catch (e: any) {
      console.error('[COLETAS_PREVISTAS] Erro ao buscar:', e.message);
      return [];
    }
  },

  /**
   * Move não coletas para a lista de histórico permanente.
   * Lista origem: Dados_Nao_Coletas (ID: 83e8cfb9-1982-47ae-b515-3fec112da457)
   * Lista destino: nao_coletas_web_hist (ID: 1702fe62-6a47-4fd1-b935-0e3258073bb6)
   */
  async moveNonCollectionsToHistory(token: string, items: NonCollection[]): Promise<{ success: number, failed: number, lastError?: string }> {
    const siteId = await getResolvedSiteId(token);
    const historyListId = '1702fe62-6a47-4fd1-b935-0e3258073bb6';
    const { mapping: histMapping, internalNames: histInternals } = await getListColumnMapping(siteId, historyListId, token);

    const safeToISO = (dateStr: string | undefined): string | null => {
      if (!dateStr || dateStr.trim() === '') return null;
      if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr + 'T12:00:00Z';
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) { const [d, m, y] = dateStr.split('/'); return `${y}-${m}-${d}T12:00:00Z`; }
      const parsed = new Date(dateStr); if (isNaN(parsed.getTime())) return null; return parsed.toISOString();
    };

    let successCount = 0;
    let failedCount = 0;
    let lastErrorMessage = "";

    for (const item of items) {
      try {
        const semana = item.semana || getWeekString(item.data);
        const fieldMap: Record<string, any> = {
          Semana: semana, Rota: item.rota, Data: safeToISO(item.data),
          'Código': item.codigo, Produtor: item.produtor, Motivo: item.motivo,
          'Observação': item.observacao, Observacao: item.observacao, 'Observa_x00e7__x00e3_o': item.observacao,
          Ação: item.acao, 'DataAção': safeToISO(item.dataAcao), 'ÚltimaColeta': safeToISO(item.ultimaColeta),
          Culpabilidade: item.Culpabilidade, 'Operação': item.operacao, CausaRaiz: item.causaRaiz || ''
        };
        const readOnlyFields = new Set(['LinkTitle','LinkTitleNoMenu','ID','ContentType','Modified','Created','Author','Editor','_UIVersionString','Attachments','Edit','DocIcon','ItemChildCount','FolderChildCount','_ComplianceFlags','_ComplianceTag','_ComplianceTagWrittenTime','_ComplianceTagUserId','_IsRecord','AppAuthor','AppEditor','Title']);
        const histFields: any = { Title: item.rota };
        Object.entries(fieldMap).forEach(([displayName, value]) => {
          const intName = resolveFieldName(histMapping, displayName);
          if (intName && histInternals.has(intName) && !readOnlyFields.has(intName)) histFields[intName] = value;
        });

        const postRes = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, { method: 'POST', body: JSON.stringify({ fields: histFields }) });
        if (postRes && postRes.id) {
          // Delete from PG instead of SharePoint
          await fetch('/api/checklist', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
            body: JSON.stringify({ domain: "non-collections", action: 'delete', id: item.id })
          });
          successCount++;
        } else { failedCount++; lastErrorMessage = "Failed to confirm archived NC ID."; }
      } catch (err: any) { failedCount++; lastErrorMessage = err.message; }
    }
    return { success: successCount, failed: failedCount, lastError: lastErrorMessage };
  }
};

/**
 * Converte data do SharePoint (ISO) para formato BR (DD/MM/YYYY)
 */
function formatDateFromSharePoint(isoDate: string): string {
  if (!isoDate) return '';
  try {
    const date = new Date(isoDate);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  } catch {
    return isoDate;
  }
}

/**
 * Converte data do formato BR (DD/MM/YYYY) para ISO (para SharePoint)
 */
function parseDateForSharePoint(brDate: string): string {
  if (!brDate || brDate.trim() === '' || brDate === '-') return '';
  try {
    const [day, month, year] = brDate.split('/');
    // Usa hora 12:00 para evitar que o fuso horário mova a data para o dia anterior
    // ao converter para ISO (ex: 00:00 UTC-3 vira 03:00 UTC, mas 23:00 do dia anterior em alguns fusos)
    const date = new Date(Number(year), Number(month) - 1, Number(day), 12, 0, 0);
    if (isNaN(date.getTime())) return '';
    return date.toISOString();
  } catch {
    return '';
  }
}
