
// @google/genai guidelines: Use direct process.env.API_KEY, no UI for keys, use correct model names.
// Correct models: 'gemini-3-flash-preview', 'gemini-3-pro-preview', 'gemini-2.5-flash-image', etc.

import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord, RouteDeparture, RouteOperationMapping, RouteConfig, NonCollection, ColetaPrevista } from '../types';
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
}

const SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH || "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, { mapping: Record<string, string>, readOnly: Set<string>, internalNames: Set<string> }> = {};

// Cache para dados estáticos/semi-estáticos (10 minutos — otimizado para reduzir chamadas à API)
const dataCache: Record<string, { data: any, timestamp: number }> = {};
const CACHE_TTL = 10 * 60 * 1000; // 10 minutos (aumentado de 5 para reduzir consumo de API)

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
  console.log('[TOKEN_EVENT] Disparando evento token-expired');
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
    console.log(`[CACHE_HIT] ${key}`);
    return cached.data as T;
  }
  return null;
}

/**
 * Armazena dados no cache
 */
function setCachedData(key: string, data: any): void {
  dataCache[key] = { data, timestamp: Date.now() };
  console.log(`[CACHE_SET] ${key}`);
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
  console.log(`[COLUMN_CACHE] Cache criado para ${cacheKey} (${columns.value.length} colunas)`);
  return columnMappingCache[cacheKey];
}

function resolveFieldName(mapping: Record<string, string>, target: string): string {
  const normalized = normalizeString(target);
  if (normalized === 'titulo' || normalized === 'rota') {
      if (mapping['title']) return 'Title';
  }
  return mapping[normalized] || target;
}

export const SharePointService = {
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
    const filter = `fields/Title eq '${status.Title}'`;
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

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);
      const celulaField = mapping['celula'] || 'celula';
      
      console.log('[HISTORY_QUERY] Buscando histórico com paginação...');
      
      // Busca todos os itens com paginação (SharePoint retorna max 100 por página)
      let allItems: any[] = [];
      let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;
      
      while (nextUrl) {
        const data = await graphFetch(nextUrl, token);
        allItems = allItems.concat(data.value || []);
        nextUrl = data['@odata.nextLink'] || null;
        console.log(`[HISTORY_QUERY] Página carregada. Total acumulado: ${allItems.length}`);
      }
      
      console.log(`[HISTORY_QUERY] Total de registros brutos: ${allItems.length}`);
      
      const result = allItems.map((item: any) => ({ 
        id: item.id, 
        timestamp: item.fields.Data, 
        resetBy: item.fields.Title, 
        email: (item.fields[celulaField] || "").toString().trim(), 
        tasks: JSON.parse(item.fields.DadosJSON || '[]') 
      })).filter((record: HistoryRecord) => record.email?.toLowerCase() === userEmail.toLowerCase().trim())
        .sort((a: any, b: any) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
      
      console.log(`[HISTORY_QUERY] ✅ ${result.length} registros filtrados por usuário ${userEmail}`);
      return result;
    } catch (e) { return []; }
  },

  async getRouteConfigs(token: string, userEmail: string, forceRefresh: boolean = false): Promise<RouteConfig[]> {
    try {
        // Se forceRefresh for true, limpa o cache antes de buscar
        if (forceRefresh) {
            clearCache('routeConfigs');
        }
        
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        
        // Adiciona timestamp para evitar cache do browser
        const timestamp = Date.now();
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&t=${timestamp}`, token);

        console.log('[DEBUG_SHAREPOINT] CONFIG_OPERACAO_SAIDA_DE_ROTAS raw data:', data);

        const result = (data.value || []).map((item: any): RouteConfig => {
          const f = item.fields;
          const config = {
            operacao: String(f[resolveFieldName(mapping, 'OPERACAO')] || ""),
            email: String(f[resolveFieldName(mapping, 'EMAIL')] || "").toString().toLowerCase().trim(),
            tolerancia: String(f[resolveFieldName(mapping, 'TOLERANCIA')] || "00:00:00"),
            nomeExibicao: String(f[resolveFieldName(mapping, 'NomeExibicao')] || String(f[resolveFieldName(mapping, 'OPERACAO')] || "")),
            ultimoEnvioSaida: String(f[resolveFieldName(mapping, 'UltimoEnvioSaida')] || ""),
            Status: String(f[resolveFieldName(mapping, 'Status')] || ""), // Status retornado pelo webhook
            Envio: String(f[resolveFieldName(mapping, 'Envio')] || ""), // Emails para envio principal
            Copia: String(f[resolveFieldName(mapping, 'Copia')] || ""), // Emails para cópia
            UltimoEnvioResumoSaida: String(f[resolveFieldName(mapping, 'UltimoEnvioResumoSaida')] || ""), // Último envio de resumo
            UltimoEnvioNcoletas: String(f[resolveFieldName(mapping, 'UltimoEnvioNcoletas')] || ""), // Último envio de não coletas
            StatusResumoSaida: String(f[resolveFieldName(mapping, 'StatusResumoSaida')] || ""), // Status do resumo
            CodigoKmm: String(f[resolveFieldName(mapping, 'CodigoKmm')] || "") // Código KMM da operação
          };
          console.log('[DEBUG_SHAREPOINT] Config item:', {
            operacao: config.operacao,
            ultimoEnvioSaida_raw: f[resolveFieldName(mapping, 'UltimoEnvioSaida')],
            ultimoEnvioSaida: config.ultimoEnvioSaida,
            Status: config.Status,
            Envio: config.Envio,
            Copia: config.Copia,
            CodigoKmm: config.CodigoKmm,
            UltimoEnvioResumoSaida: config.UltimoEnvioResumoSaida,
            StatusResumoSaida: config.StatusResumoSaida
          });
          return config;
        }).filter(c => c.email === userEmail.toLowerCase().trim());

        console.log('[DEBUG_SHAREPOINT] Configs filtradas por email:', result);
        return result;
    } catch (e: any) {
      console.error('[SHAREPOINT] Erro ao buscar CONFIG_OPERACAO_SAIDA_DE_ROTAS:', e.message);
      // Retorna array vazio se a lista não existir
      return [];
    }
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
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

      // Busca o item da operação específica
      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (existing.value && existing.value.length > 0) {
        const itemId = existing.value[0].id;
        const ultimoEnvioField = resolveFieldName(mapping, 'UltimoEnvioSaida');

        // Se string vazia, limpa o campo
        let dataISO: string | null = dataHoraEnvio;

        if (!dataHoraEnvio || dataHoraEnvio.trim() === '') {
          dataISO = null;
          console.log(`[DATA_CONVERSAO] 🧹 Limpando campo UltimoEnvioSaida para ${operacao}`);
        } else {
          console.log(`[DATA_CONVERSAO] Recebido: "${dataHoraEnvio}" (tipo: ${typeof dataHoraEnvio})`);

          // Tenta formato completo: DD/MM/YYYY HH:MM:SS
          const matchCompleto = dataHoraEnvio.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
          // Tenta formato sem segundos: DD/MM/YYYY HH:MM
          const matchSemSegundos = dataHoraEnvio.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);

          if (matchCompleto) {
            const [, dia, mes, ano, hora, minuto, segundo] = matchCompleto;
            const localDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), Number(segundo));
            dataISO = localDate.toISOString();
            console.log(`[DATA_CONVERSAO] ✅ Formato DD/MM/YYYY HH:MM:SS: "${dataHoraEnvio}" → "${dataISO}"`);
          } else if (matchSemSegundos) {
            const [, dia, mes, ano, hora, minuto] = matchSemSegundos;
            const localDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), 0);
            dataISO = localDate.toISOString();
            console.log(`[DATA_CONVERSAO] ✅ Formato DD/MM/YYYY HH:MM: "${dataHoraEnvio}" → "${dataISO}"`);
          } else {
            const parsed = new Date(dataHoraEnvio);
            if (!isNaN(parsed.getTime())) {
              dataISO = parsed.toISOString();
              console.log(`[DATA_CONVERSAO] ⚠️ Parseado como ISO genérico: "${dataHoraEnvio}" → "${dataISO}"`);
            } else {
              console.warn(`[DATA_CONVERSAO] ❌ Formato não reconhecido: "${dataHoraEnvio}", usando data atual`);
              dataISO = new Date().toISOString();
            }
          }
        }

        const fields: any = {
          [ultimoEnvioField]: dataISO
        };
        
        console.log(`[SHAREPOINT] Enviando PATCH para item ${itemId} campo ${ultimoEnvioField} = ${dataISO}`);
        
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, { 
          method: 'PATCH', 
          body: JSON.stringify(fields) 
        });
        
        console.log(`[SHAREPOINT] ✅ UltimoEnvioSaida atualizado para ${operacao}: ${dataISO}`);
      } else {
        console.warn(`[SHAREPOINT] Operação "${operacao}" não encontrada na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS`);
      }
    } catch (error: any) {
      console.error('[SHAREPOINT] ❌ Erro ao atualizar UltimoEnvioSaida:', error.message);
    }
  },

  /**
   * Atualiza o campo UltimoEnvioNcoletas na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateUltimoEnvioNaoColetas(token: string, operacao: string, dataHoraEnvio: string): Promise<void> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

      // Busca o item da operação específica
      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (existing.value && existing.value.length > 0) {
        const itemId = existing.value[0].id;
        const ultimoEnvioField = resolveFieldName(mapping, 'UltimoEnvioNcoletas');

        // Converte data/hora para ISO
        let dataISO: string | null = dataHoraEnvio;

        if (!dataHoraEnvio || dataHoraEnvio.trim() === '') {
          dataISO = null;
          console.log(`[DATA_CONVERSAO_NC] 🧹 Limpando campo UltimoEnvioNcoletas para ${operacao}`);
        } else {
          console.log(`[DATA_CONVERSAO_NC] Recebido: "${dataHoraEnvio}"`);

          // Tenta formato completo: DD/MM/YYYY HH:MM:SS
          const matchCompleto = dataHoraEnvio.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
          // Tenta formato sem segundos: DD/MM/YYYY HH:MM
          const matchSemSegundos = dataHoraEnvio.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);

          if (matchCompleto) {
            const [, dia, mes, ano, hora, minuto, segundo] = matchCompleto;
            const localDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), Number(segundo));
            dataISO = localDate.toISOString();
            console.log(`[DATA_CONVERSAO_NC] ✅ Formato DD/MM/YYYY HH:MM:SS: "${dataHoraEnvio}" → "${dataISO}"`);
          } else if (matchSemSegundos) {
            const [, dia, mes, ano, hora, minuto] = matchSemSegundos;
            const localDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), 0);
            dataISO = localDate.toISOString();
            console.log(`[DATA_CONVERSAO_NC] ✅ Formato DD/MM/YYYY HH:MM: "${dataHoraEnvio}" → "${dataISO}"`);
          } else {
            const parsed = new Date(dataHoraEnvio);
            if (!isNaN(parsed.getTime())) {
              dataISO = parsed.toISOString();
              console.log(`[DATA_CONVERSAO_NC] ⚠️ Parseado como ISO genérico: "${dataHoraEnvio}" → "${dataISO}"`);
            } else {
              console.warn(`[DATA_CONVERSAO_NC] ❌ Formato não reconhecido: "${dataHoraEnvio}", usando data atual`);
              dataISO = new Date().toISOString();
            }
          }
        }

        const fields: any = {
          [ultimoEnvioField]: dataISO
        };

        console.log(`[SHAREPOINT_NC] Enviando PATCH para item ${itemId} campo ${ultimoEnvioField} = ${dataISO}`);

        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
          method: 'PATCH',
          body: JSON.stringify(fields)
        });

        console.log(`[SHAREPOINT_NC] ✅ UltimoEnvioNcoletas atualizado para ${operacao}: ${dataISO}`);
      } else {
        console.warn(`[SHAREPOINT_NC] Operação "${operacao}" não encontrada na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS`);
      }
    } catch (error: any) {
      console.error('[SHAREPOINT_NC] ❌ Erro ao atualizar UltimoEnvioNcoletas:', error.message);
    }
  },

  /**
   * Atualiza o campo Status na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateStatusOperacao(token: string, operacao: string, status: string): Promise<void> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

      // Busca o item da operação específica
      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (existing.value && existing.value.length > 0) {
        const itemId = existing.value[0].id;
        const statusField = resolveFieldName(mapping, 'Status');

        const fields: any = {
          [statusField]: status
        };

        console.log(`[SHAREPOINT] Enviando PATCH para item ${itemId} campo ${statusField} = ${status}`);

        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
          method: 'PATCH',
          body: JSON.stringify(fields)
        });

        console.log(`[SHAREPOINT] ✅ Status atualizado para ${operacao}: ${status}`);
      } else {
        console.warn(`[SHAREPOINT] Operação "${operacao}" não encontrada na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS`);
      }
    } catch (error: any) {
      console.error('[SHAREPOINT] ❌ Erro ao atualizar Status:', error.message);
    }
  },

  /**
   * Atualiza os campos Envio e Copia na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateRouteConfigEmails(token: string, operacao: string, envio: string, copia: string): Promise<void> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

      // Busca o item da operação específica
      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (existing.value && existing.value.length > 0) {
        const itemId = existing.value[0].id;
        const envioField = resolveFieldName(mapping, 'Envio');
        const copiaField = resolveFieldName(mapping, 'Copia');

        const fields: any = {
          [envioField]: envio,
          [copiaField]: copia
        };

        console.log(`[SHAREPOINT] Enviando PATCH para item ${itemId}:`);
        console.log(`  - ${envioField} = ${envio}`);
        console.log(`  - ${copiaField} = ${copia}`);

        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
          method: 'PATCH',
          body: JSON.stringify(fields)
        });

        console.log(`[SHAREPOINT] ✅ Emails atualizados para ${operacao}:`);
        console.log(`  - Envio: ${envio}`);
        console.log(`  - Copia: ${copia}`);
      } else {
        console.warn(`[SHAREPOINT] Operação "${operacao}" não encontrada na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS`);
        throw new Error(`Operação "${operacao}" não encontrada no SharePoint`);
      }
    } catch (error: any) {
      console.error('[SHAREPOINT] ❌ Erro ao atualizar emails:', error.message);
      throw error;
    }
  },

  /**
   * Atualiza o campo UltimoEnvioResumoSaida na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateUltimoEnvioResumoSaida(token: string, operacao: string, dataHoraEnvio: string): Promise<void> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

      // Busca o item da operação específica
      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (existing.value && existing.value.length > 0) {
        const itemId = existing.value[0].id;
        const ultimoEnvioResumoField = resolveFieldName(mapping, 'UltimoEnvioResumoSaida');

        // Se string vazia, limpa o campo
        let dataISO: string | null = dataHoraEnvio;

        if (!dataHoraEnvio || dataHoraEnvio.trim() === '') {
          dataISO = null;
          console.log(`[DATA_CONVERSAO] 🧹 Limpando campo UltimoEnvioResumoSaida para ${operacao}`);
        } else {
          console.log(`[DATA_CONVERSAO] Recebido: "${dataHoraEnvio}" (tipo: ${typeof dataHoraEnvio})`);

          // Tenta formato completo: DD/MM/YYYY HH:MM:SS
          const matchCompleto = dataHoraEnvio.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
          // Tenta formato sem segundos: DD/MM/YYYY HH:MM
          const matchSemSegundos = dataHoraEnvio.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);

          if (matchCompleto) {
            const [, dia, mes, ano, hora, minuto, segundo] = matchCompleto;
            const localDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), Number(segundo));
            dataISO = localDate.toISOString();
            console.log(`[DATA_CONVERSAO] ✅ Formato DD/MM/YYYY HH:MM:SS: "${dataHoraEnvio}" → "${dataISO}"`);
          } else if (matchSemSegundos) {
            const [, dia, mes, ano, hora, minuto] = matchSemSegundos;
            const localDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(hora), Number(minuto), 0);
            dataISO = localDate.toISOString();
            console.log(`[DATA_CONVERSAO] ✅ Formato DD/MM/YYYY HH:MM: "${dataHoraEnvio}" → "${dataISO}"`);
          } else {
            const parsed = new Date(dataHoraEnvio);
            if (!isNaN(parsed.getTime())) {
              dataISO = parsed.toISOString();
              console.log(`[DATA_CONVERSAO] ⚠️ Parseado como ISO genérico: "${dataHoraEnvio}" → "${dataISO}"`);
            } else {
              console.warn(`[DATA_CONVERSAO] ❌ Formato não reconhecido: "${dataHoraEnvio}", usando data atual`);
              dataISO = new Date().toISOString();
            }
          }
        }

        const fields: any = {
          [ultimoEnvioResumoField]: dataISO
        };

        console.log(`[SHAREPOINT] Enviando PATCH para item ${itemId} campo ${ultimoEnvioResumoField} = ${dataISO}`);

        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
          method: 'PATCH',
          body: JSON.stringify(fields)
        });

        console.log(`[SHAREPOINT] ✅ UltimoEnvioResumoSaida atualizado para ${operacao}: ${dataISO}`);
      } else {
        console.warn(`[SHAREPOINT] Operação "${operacao}" não encontrada na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS`);
      }
    } catch (error: any) {
      console.error('[SHAREPOINT] ❌ Erro ao atualizar UltimoEnvioResumoSaida:', error.message);
    }
  },

  /**
   * Atualiza o campo StatusResumoSaida na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
   */
  async updateStatusResumoSaida(token: string, operacao: string, status: string): Promise<void> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

      // Busca o item da operação específica
      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (existing.value && existing.value.length > 0) {
        const itemId = existing.value[0].id;
        const statusResumoField = resolveFieldName(mapping, 'StatusResumoSaida');

        const fields: any = {
          [statusResumoField]: status
        };

        console.log(`[SHAREPOINT] Enviando PATCH para item ${itemId} campo ${statusResumoField} = ${status}`);

        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
          method: 'PATCH',
          body: JSON.stringify(fields)
        });

        console.log(`[SHAREPOINT] ✅ StatusResumoSaida atualizado para ${operacao}: ${status}`);
      } else {
        console.warn(`[SHAREPOINT] Operação "${operacao}" não encontrada na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS`);
      }
    } catch (error: any) {
      console.error('[SHAREPOINT] ❌ Erro ao atualizar StatusResumoSaida:', error.message);
    }
  },

  async getDepartures(token: string, forceRefresh: boolean = false): Promise<RouteDeparture[]> {
    try {
      const cacheKey = 'departures';

      // Se forceRefresh for true, limpa o cache antes de buscar
      if (forceRefresh) {
        clearCache(cacheKey);
      } else {
        const cached = getCachedData<RouteDeparture[]>(cacheKey);
        if (cached) return cached;
      }

      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);

      // Adiciona timestamp para evitar cache do browser
      const timestamp = Date.now();

      console.log('[GET_DEPARTURES] Buscando todas as rotas com paginação...');

      // Busca todos os itens com paginação (SharePoint retorna max ~200 por página)
      let allItems: any[] = [];
      let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100&t=${timestamp}`;

      while (nextUrl) {
        const data = await graphFetch(nextUrl, token);
        allItems = allItems.concat(data.value || []);
        nextUrl = data['@odata.nextLink'] || null;
        console.log(`[GET_DEPARTURES] Página carregada. Total acumulado: ${allItems.length}`);
      }

      console.log(`[GET_DEPARTURES] Total de rotas carregadas: ${allItems.length}`);

      const result = allItems.map((item: any) => {
        const f = item.fields;
        const dataStr = f[resolveFieldName(mapping, 'DataOperacao')] ? f[resolveFieldName(mapping, 'DataOperacao')].split('T')[0] : "";
        const semanaFromSharePoint = f[resolveFieldName(mapping, 'Semana')] || "";

        return {
          id: String(item.id),
          semana: semanaFromSharePoint || getWeekString(dataStr), // Calcula se não vier do SharePoint
          rota: f.Title || "",
          data: dataStr,
          inicio: f[resolveFieldName(mapping, 'HorarioInicio')] || "",
          motorista: f[resolveFieldName(mapping, 'Motorista')] || "",
          placa: f[resolveFieldName(mapping, 'Placa')] || "",
          saida: f[resolveFieldName(mapping, 'HorarioSaida')] || "",
          motivo: f[resolveFieldName(mapping, 'MotivoAtraso')] || "",
          observacao: f[resolveFieldName(mapping, 'Observacao')] || "",
          statusGeral: f[resolveFieldName(mapping, 'StatusGeral')] || "",
          aviso: f[resolveFieldName(mapping, 'Aviso')] || "NÃO",
          operacao: f[resolveFieldName(mapping, 'Operacao')] || "",
          statusOp: f[resolveFieldName(mapping, 'StatusOp')] || "Previsto",
          tempo: f[resolveFieldName(mapping, 'TempGab')] || f[resolveFieldName(mapping, 'TempoGap')] || "",
          createdAt: f.Created || new Date().toISOString(),
          checklistMotorista: f[resolveFieldName(mapping, 'ChecklistMotorista')] || "",
          causaRaiz: f[resolveFieldName(mapping, 'CausaRaiz')] || ""
        };
      });

      setCachedData(cacheKey, result);
      return result;
    } catch (e) { return []; }
  },

  async getArchivedDepartures(token: string, operation: string | null, startDate: string, endDate: string, signal?: AbortSignal): Promise<RouteDeparture[]> {
    const cacheKey = `archived_departures_${startDate}_${endDate}_${operation || 'all'}`;

    // 1. Cache: retorna imediatamente se já buscou esse range recentemente
    const cached = getCachedData<RouteDeparture[]>(cacheKey);
    if (cached) {
      console.log(`[ARCHIVE_QUERY] Cache hit para ${cacheKey}`);
      return cached;
    }

    // 2. Deduplicação: se já existe uma requisição em andamento para o mesmo range, reutiliza
    if (inFlightArchiveRequests[cacheKey]) {
      console.log(`[ARCHIVE_QUERY] Reutilizando requisição em andamento para ${cacheKey}`);
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

        console.log(`[ARCHIVE_QUERY] URL: /sites/${siteId}/lists/${historyListId}/items Filter: ${filter}`);

        // Busca todos os itens com paginação (SharePoint retorna max 100 por página)
        let allItems: any[] = [];
        let nextUrl: string | null = `/sites/${siteId}/lists/${historyListId}/items?expand=fields&$filter=${filter}&$top=100`;

        while (nextUrl) {
          // Verifica se a requisição foi cancelada antes de cada página
          if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
          const data = await graphFetch(nextUrl, token, signal ? { signal } : {});
          allItems = allItems.concat(data.value || []);
          nextUrl = data['@odata.nextLink'] || null;
          console.log(`[ARCHIVE_QUERY] Página carregada. Total acumulado: ${allItems.length}`);
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
            causaRaiz: f[resolveFieldName(mapping, 'CausaRaiz')] || ""
          };
        });

        console.log(`[ARCHIVE_QUERY] Search success. Found ${results.length} records.`);
        setCachedData(cacheKey, results);
        return results;
      } catch (e: any) {
        if (e.name === 'AbortError') {
          console.log('[ARCHIVE_QUERY] Requisição cancelada pelo usuário.');
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
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);

    // Calcula a semana com base na data, usando a mesma lógica do Excel
    // Se departure.semana já existir, usa; caso contrário, calcula automaticamente
    const semana = departure.semana || getWeekString(departure.data);

    const raw: any = {
        Title: departure.rota,
        Semana: semana,
        DataOperacao: departure.data ? new Date(departure.data + 'T12:00:00Z').toISOString() : null,
        HorarioInicio: departure.inicio,
        Motorista: departure.motorista,
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
        CausaRaiz: departure.causaRaiz || ''
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
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${departure.id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
      result = departure.id;
    } else {
      const res = await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
      result = String(res.id);
    }

    // Invalida cache após atualização
    clearCache('departures');
    return result;
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
        CausaRaiz: departure.causaRaiz || ''
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
      console.log(`[HISTORY_UPDATE] Atualizando item ${departure.id} na lista de histórico`);
      await graphFetch(`/sites/${siteId}/lists/${historyListId}/items/${departure.id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
      result = departure.id;
    } else {
      console.log(`[HISTORY_UPDATE] Criando novo item na lista de histórico`);
      const res = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
      result = String(res.id);
    }

    // Invalida cache de histórico após atualização
    clearCacheByPrefix('archived_departures_');
    return result;
  },

  async deleteDeparture(token: string, id: string): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${id}`, token, { method: 'DELETE' });
    
    // Invalida cache após deletar
    clearCache('departures');
  },

  async moveDeparturesToHistory(token: string, items: RouteDeparture[]): Promise<{ success: number, failed: number, lastError?: string }> {
    console.log(`[ARCHIVE_START] Starting migration of ${items.length} items to permanent history.`);
    const siteId = await getResolvedSiteId(token);
    const sourceList = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    const historyListId = "856bf9d5-6081-4360-bcad-e771cbabfda8";
    const { mapping: histMapping, internalNames: histInternals } = await getListColumnMapping(siteId, historyListId, token);
    
    let successCount = 0;
    let failedCount = 0;
    let lastErrorMessage = "";

    for (const item of items) {
        try {
            // Calcula a semana com base na data, usando a mesma lógica do Excel
            // Se item.semana já existir, usa; caso contrário, calcula automaticamente
            const semana = item.semana || getWeekString(item.data);
            
            const raw: any = { 
                Title: item.rota, 
                Semana: semana, 
                DataOperacao: item.data ? new Date(item.data + 'T12:00:00Z').toISOString() : null, 
                HorarioInicio: item.inicio, 
                Motorista: item.motorista, 
                Placa: item.placa, 
                HorarioSaida: item.saida, 
                MotivoAtraso: item.motivo, 
                Observacao: item.observacao, 
                StatusGeral: item.statusGeral, 
                Aviso: item.aviso, 
                Operacao: item.operacao, 
                StatusOp: item.statusOp, 
                TempGab: item.tempo, 
                ChecklistMotorista: item.checklistMotorista || '',
                CausaRaiz: item.causaRaiz || '' 
            };
            const histFields: any = {};
            Object.keys(raw).forEach(k => { const int = resolveFieldName(histMapping, k); if (histInternals.has(int)) histFields[int] = raw[k]; });
            const postRes = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, { method: 'POST', body: JSON.stringify({ fields: histFields }) });
            if (postRes && postRes.id) {
                await graphFetch(`/sites/${siteId}/lists/${sourceList.id}/items/${item.id}`, token, { method: 'DELETE' });
                successCount++;
            } else { failedCount++; lastErrorMessage = "Failed to confirm archived ID."; }
        } catch (err: any) { failedCount++; lastErrorMessage = err.message; }
    }
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
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);

      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (!existing.value || existing.value.length === 0) {
        return { locked: false };
      }

      const item = existing.value[0];
      const f = item.fields;
      
      const lockField = resolveFieldName(mapping, 'LockEnvio');
      const lockUserField = resolveFieldName(mapping, 'LockUser');
      const lockTimestampField = resolveFieldName(mapping, 'LockTimestamp');

      const lockStatus = f[lockField] || '';
      const lockUser = f[lockUserField] || '';
      const lockTimestamp = f[lockTimestampField] || '';

      // Se não tem trava, retorna false
      if (!lockStatus || lockStatus.toLowerCase() !== 'true') {
        return { locked: false };
      }

      // Verifica se a trava expirou (timeout de 2 minutos)
      if (lockTimestamp) {
        let lockDate: Date | null = null;
        if (lockTimestamp.includes('T')) {
          lockDate = new Date(lockTimestamp);
        } else if (lockTimestamp.includes('/')) {
          const [data, hora] = lockTimestamp.split(' ');
          const [dia, mes, ano] = data.split('/');
          const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
          lockDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
        }

        if (lockDate && !isNaN(lockDate.getTime())) {
          const now = new Date();
          const diffMs = now.getTime() - lockDate.getTime();
          const timeoutMs = 2 * 60 * 1000; // 2 minutos de timeout

          if (diffMs > timeoutMs) {
            console.log(`[LOCK_CHECK] Trava expirada para ${operacao} (usuário: ${lockUser}, tempo: ${Math.floor(diffMs / 1000)}s)`);
            return { locked: false, user: lockUser, timestamp: lockTimestamp, expired: true };
          }
        }
      }

      console.log(`[LOCK_CHECK] Trava ativa para ${operacao} por ${lockUser} em ${lockTimestamp}`);
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
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);

      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (!existing.value || existing.value.length === 0) {
        return { success: false, message: 'Operação não encontrada' };
      }

      const itemId = existing.value[0].id;
      const lockField = resolveFieldName(mapping, 'LockEnvio');
      const lockUserField = resolveFieldName(mapping, 'LockUser');
      const lockTimestampField = resolveFieldName(mapping, 'LockTimestamp');

      // Verifica estado atual da trava
      const f = existing.value[0].fields;
      const currentLock = f[lockField] || '';
      const currentLockUser = f[lockUserField] || '';
      const currentLockTimestamp = f[lockTimestampField] || '';

      // Verifica se a trava está ativa e não é do usuário atual
      if (currentLock && currentLock.toLowerCase() === 'true' && currentLockUser.toLowerCase() !== userEmail.toLowerCase()) {
        // Verifica se não expirou
        if (currentLockTimestamp) {
          let lockDate: Date | null = null;
          if (currentLockTimestamp.includes('T')) {
            lockDate = new Date(currentLockTimestamp);
          } else if (currentLockTimestamp.includes('/')) {
            const [data, hora] = currentLockTimestamp.split(' ');
            const [dia, mes, ano] = data.split('/');
            const [h, m, s] = hora ? hora.split(':') : ['00', '00', '00'];
            lockDate = new Date(Number(ano), Number(mes) - 1, Number(dia), Number(h), Number(m), Number(s));
          }

          if (lockDate && !isNaN(lockDate.getTime())) {
            const now = new Date();
            const diffMs = now.getTime() - lockDate.getTime();
            const timeoutMs = 2 * 60 * 1000; // 2 minutos

            if (diffMs <= timeoutMs) {
              return { 
                success: false, 
                message: `Outro usuário (${currentLockUser}) está enviando os dados. Aguarde alguns segundos e tente novamente.`
              };
            }
          }
        }
      }

      // Adquire a trava
      const now = new Date();
      const timestamp = now.toISOString();
      
      const fields: any = {
        [lockField]: 'true',
        [lockUserField]: userEmail,
        [lockTimestampField]: timestamp
      };

      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
        method: 'PATCH',
        body: JSON.stringify(fields)
      });

      console.log(`[LOCK_ACQUIRE] Trava adquirida por ${userEmail} para ${operacao} em ${timestamp}`);
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
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
      const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);

      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const filter = `fields/${operacaoField} eq '${operacao}'`;
      const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=1`, token);

      if (!existing.value || existing.value.length === 0) {
        console.warn(`[LOCK_RELEASE] Operação ${operacao} não encontrada`);
        return;
      }

      const itemId = existing.value[0].id;
      const lockField = resolveFieldName(mapping, 'LockEnvio');
      const lockUserField = resolveFieldName(mapping, 'LockUser');
      const lockTimestampField = resolveFieldName(mapping, 'LockTimestamp');

      const fields: any = {
        [lockField]: '',
        [lockUserField]: '',
        [lockTimestampField]: ''
      };

      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
        method: 'PATCH',
        body: JSON.stringify(fields)
      });

      console.log(`[LOCK_RELEASE] Trava liberada para ${operacao}`);
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
      const siteId = await getResolvedSiteId(token);
      const listId = '83e8cfb9-1982-47ae-b515-3fec112da457';

      // Busca todos os itens com paginação (SharePoint retorna max 100 por página)
      let allItems: any[] = [];
      let nextUrl: string | null = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=100`;

      while (nextUrl) {
        const data = await graphFetch(nextUrl, token);
        allItems = allItems.concat(data.value || []);
        nextUrl = data['@odata.nextLink'] || null;
        console.log(`[NonCollections] Página carregada. Total acumulado: ${allItems.length}`);
      }

      console.log(`[NonCollections] Total bruto: ${allItems.length} itens`);

      return allItems.map((item: any) => {
        const f = item.fields || {};
        return {
          id: item.id.toString(),
          semana: f.Title || '',
          rota: f.Rota || '',
          data: f.Data ? formatDateFromSharePoint(f.Data) : '',
          codigo: f.C_x00f3_digo || '',
          produtor: f.Produtor || '',
          motivo: f.Motivo || '',
          observacao: f.Observa_x00e7__x00e3_o || '',
          acao: f.A_x00e7__x00e3_o || '',
          dataAcao: f.DataA_x00e7__x00e3_o ? formatDateFromSharePoint(f.DataA_x00e7__x00e3_o) : '',
          ultimaColeta: f._x00da_ltimaColeta ? formatDateFromSharePoint(f._x00da_ltimaColeta) : '',
          Culpabilidade: f.Culpabilidade || '',
          operacao: f.Opera_x00e7__x00e3_o || ''
        };
      });
    } catch (e: any) {
      console.error('[NonCollections] Erro ao buscar não coletas:', e.message);
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
      const siteId = await getResolvedSiteId(token);
      const listId = '83e8cfb9-1982-47ae-b515-3fec112da457';

      // Constrói payload removendo campos vazios (SharePoint rejeita DateTime com "")
      const payload: any = {};

      if (nonCollection.semana) payload.Title = nonCollection.semana;
      if (nonCollection.rota) payload.Rota = nonCollection.rota;
      if (nonCollection.data) {
        const parsedData = parseDateForSharePoint(nonCollection.data);
        if (parsedData) payload.Data = parsedData;
      }
      if (nonCollection.codigo) payload.C_x00f3_digo = nonCollection.codigo;
      if (nonCollection.produtor) payload.Produtor = nonCollection.produtor;
      if (nonCollection.motivo) payload.Motivo = nonCollection.motivo;
      if (nonCollection.observacao) payload.Observa_x00e7__x00e3_o = nonCollection.observacao;
      if (nonCollection.acao) payload.A_x00e7__x00e3_o = nonCollection.acao;
      // Campos DateTime: só envia se parse resultou em valor válido
      { const v = parseDateForSharePoint(nonCollection.dataAcao); if (v) payload.DataA_x00e7__x00e3_o = v; }
      { const v = parseDateForSharePoint(nonCollection.ultimaColeta); if (v) payload._x00da_ltimaColeta = v; }
      if (nonCollection.Culpabilidade) payload.Culpabilidade = nonCollection.Culpabilidade;
      if (nonCollection.operacao) payload.Opera_x00e7__x00e3_o = nonCollection.operacao;

      console.log('[NonCollections] Salvando payload:', JSON.stringify(payload));

      const response = await graphFetch(`/sites/${siteId}/lists/${listId}/items`, token, {
        method: 'POST',
        body: JSON.stringify({ fields: payload })
      });

      const spId = response.id?.toString();
      console.log('[NonCollections] ✅ Não coleta salva com sucesso, ID:', spId);
      return spId;
    } catch (e: any) {
      console.error('[NonCollections] Erro ao salvar não coleta:', e.message);
      throw e;
    }
  },

  /**
   * Atualiza não coleta existente na lista do SharePoint
   * Lista: Dados_Nao_Coletas (ID: 83e8cfb9-1982-47ae-b515-3fec112da457)
   */
  async updateNonCollection(token: string, nonCollection: NonCollection): Promise<void> {
    try {
      const siteId = await getResolvedSiteId(token);
      const listId = '83e8cfb9-1982-47ae-b515-3fec112da457';

      // Constrói payload removendo campos vazios (SharePoint rejeita DateTime com "")
      const payload: any = {};

      if (nonCollection.semana) payload.Title = nonCollection.semana;
      if (nonCollection.rota) payload.Rota = nonCollection.rota;
      if (nonCollection.data) {
        const parsedData = parseDateForSharePoint(nonCollection.data);
        if (parsedData) payload.Data = parsedData;
      }
      if (nonCollection.codigo) payload.C_x00f3_digo = nonCollection.codigo;
      if (nonCollection.produtor) payload.Produtor = nonCollection.produtor;
      if (nonCollection.motivo) payload.Motivo = nonCollection.motivo;
      if (nonCollection.observacao) payload.Observa_x00e7__x00e3_o = nonCollection.observacao;
      if (nonCollection.acao) payload.A_x00e7__x00e3_o = nonCollection.acao;
      // Campos DateTime: só envia se parse resultou em valor válido
      { const v = parseDateForSharePoint(nonCollection.dataAcao); if (v) payload.DataA_x00e7__x00e3_o = v; }
      { const v = parseDateForSharePoint(nonCollection.ultimaColeta); if (v) payload._x00da_ltimaColeta = v; }
      if (nonCollection.Culpabilidade) payload.Culpabilidade = nonCollection.Culpabilidade;
      if (nonCollection.operacao) payload.Opera_x00e7__x00e3_o = nonCollection.operacao;

      console.log('[NonCollections] Atualizando payload:', JSON.stringify(payload));
      console.log('[NonCollections] ID do item:', nonCollection.id);

      await graphFetch(`/sites/${siteId}/lists/${listId}/items/${nonCollection.id}`, token, {
        method: 'PATCH',
        body: JSON.stringify({ fields: payload })
      });

      console.log('[NonCollections] ✅ Não coleta atualizada com sucesso:', nonCollection.rota, '-', nonCollection.codigo);
    } catch (e: any) {
      console.error('[NonCollections] Erro ao atualizar não coleta:', e.message);
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
      if (nonCollection.observacao) payload.Observa_x00e7__x00e3_o = nonCollection.observacao;
      if (nonCollection.acao) payload.A_x00e7__x00e3_o = nonCollection.acao;
      // Campos DateTime: só envia se parse resultou em valor válido
      { const v = parseDateForSharePoint(nonCollection.dataAcao); if (v) payload.DataA_x00e7__x00e3_o = v; }
      { const v = parseDateForSharePoint(nonCollection.ultimaColeta); if (v) payload._x00da_ltimaColeta = v; }
      if (nonCollection.Culpabilidade) payload.Culpabilidade = nonCollection.Culpabilidade;
      if (nonCollection.operacao) payload.Opera_x00e7__x00e3_o = nonCollection.operacao;

      await graphFetch(`/sites/${siteId}/lists/${historyListId}/items/${nonCollection.id}`, token, {
        method: 'PATCH',
        body: JSON.stringify({ fields: payload })
      });

      // Invalida cache das consultas de histórico de não coletas
      clearCacheByPrefix('archived_noncollections_');

      console.log('[NonCollectionsHistory] ✅ Não coleta de histórico atualizada com sucesso:', nonCollection.rota, '-', nonCollection.codigo);
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
      const siteId = await getResolvedSiteId(token);
      const listId = '83e8cfb9-1982-47ae-b515-3fec112da457';

      await graphFetch(`/sites/${siteId}/lists/${listId}/items/${id}`, token, {
        method: 'DELETE'
      });

      console.log('[NonCollections] ✅ Não coleta excluída com sucesso, ID:', id);
    } catch (e: any) {
      console.error('[NonCollections] Erro ao excluir não coleta:', e.message);
      throw e;
    }
  },

  /**
   * Busca não coletas arquivadas no histórico.
   * Lista: nao_coletas_web_hist (ID: 1702fe62-6a47-4fd1-b935-0e3258073bb6)
   */
  async getArchivedNonCollections(token: string, userEmail: string, startDate: string, endDate: string, signal?: AbortSignal): Promise<NonCollection[]> {
    const cacheKey = `archived_noncollections_${startDate}_${endDate}`;

    // 1. Cache: retorna imediatamente se já buscou esse range recentemente
    const cached = getCachedData<NonCollection[]>(cacheKey);
    if (cached) {
      console.log(`[NC_ARCHIVE_QUERY] Cache hit para ${cacheKey}`);
      return cached;
    }

    // 2. Deduplicação: se já existe uma requisição em andamento para o mesmo range, reutiliza
    if (inFlightArchiveRequests[cacheKey]) {
      console.log(`[NC_ARCHIVE_QUERY] Reutilizando requisição em andamento para ${cacheKey}`);
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

        console.log(`[NC_ARCHIVE_QUERY] URL: /sites/${siteId}/lists/${historyListId}/items Filter: ${filter}`);

        // Busca todos os itens com paginação
        let allItems: any[] = [];
        let nextUrl: string | null = `/sites/${siteId}/lists/${historyListId}/items?expand=fields&$filter=${filter}&$top=100`;

        while (nextUrl) {
          // Verifica se a requisição foi cancelada antes de cada página
          if (signal?.aborted) throw new DOMException('Aborted', 'AbortError');
          const data = await graphFetch(nextUrl, token, signal ? { signal } : {});
          allItems = allItems.concat(data.value || []);
          nextUrl = data['@odata.nextLink'] || null;
          console.log(`[NC_ARCHIVE_QUERY] Página carregada. Total acumulado: ${allItems.length}`);
        }

        const results = allItems.map((item: any) => {
          const f = item.fields;
          const dataStr = f[colData] ? f[colData].split('T')[0] : "";
          const ultimaColetaStr = f[resolveFieldName(mapping, 'ÚltimaColeta')] ? f[resolveFieldName(mapping, 'ÚltimaColeta')].split('T')[0] : "";
          const dataAcaoStr = f[resolveFieldName(mapping, 'DataAção')] ? f[resolveFieldName(mapping, 'DataAção')].split('T')[0] : "";

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
            observacao: f[resolveFieldName(mapping, 'Observação')] || "",
            acao: f[resolveFieldName(mapping, 'Ação')] || "",
            dataAcao: dataAcaoStr,
            ultimaColeta: ultimaColetaStr,
            Culpabilidade: f[resolveFieldName(mapping, 'Culpabilidade')] || "",
            operacao: f[colOp] || ""
          };
        });

        console.log(`[NC_ARCHIVE_QUERY] Search success. Found ${results.length} records.`);
        setCachedData(cacheKey, results);
        return results;
      } catch (e: any) {
        if (e.name === 'AbortError') {
          console.log('[NC_ARCHIVE_QUERY] Requisição cancelada pelo usuário.');
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
          console.log(`[COLETAS_PREVISTAS] Página ${page} carregada. Total acumulado: ${allItems.length}`);
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
      console.log(`[COLETAS_PREVISTAS] URL base: /sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${rangeFilter}`);

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
        console.log(`[COLETAS_PREVISTAS] Fallback cliente encontrou ${allItems.length} itens para ${date}.`);
      }

      console.log(`[COLETAS_PREVISTAS] Raw data count: ${allItems.length}`);
      if (allItems.length > 0) {
        console.log(`[COLETAS_PREVISTAS] Primeiro item raw:`, JSON.stringify(allItems[0], null, 2));
      }

      // Busca configurações do usuário para filtrar pelas operações dele
      const operationSource =
        userOperations.length > 0
          ? userOperations
          : (await this.getRouteConfigs(token, userEmail, true)).map(c => c.operacao);

      const myOps = new Set(operationSource.map(normalizeOperation).filter(Boolean));

      console.log(`[COLETAS_PREVISTAS] Operações do usuário (${userEmail}):`, Array.from(myOps));

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

      console.log(`[COLETAS_PREVISTAS] Antes do filtro:`, result.map(c => `${c.Title}=${c.QntColeta}`));

      const filtered = result.filter(c => myOps.size === 0 || myOps.has(normalizeOperation(c.Title)));

      if (myOps.size > 0) {
        const resultOps = new Set(result.map(r => normalizeOperation(r.Title)).filter(Boolean));
        const missingOps = Array.from(myOps).filter(op => !resultOps.has(op));
        if (missingOps.length > 0) {
          console.warn('[COLETAS_PREVISTAS] Operações do usuário sem correspondência na lista de previstas:', missingOps);
        }
      }

      console.log(`[COLETAS_PREVISTAS] Depois do filtro:`, filtered.map(c => `${c.Title}=${c.QntColeta}`));
      console.log(`[COLETAS_PREVISTAS] Total: ${filtered.length}, Soma QntColeta: ${filtered.reduce((sum, c) => sum + c.QntColeta, 0)}`);

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
    console.log(`[NC_ARCHIVE_START] Starting migration of ${items.length} items to permanent history.`);
    const siteId = await getResolvedSiteId(token);
    const sourceListId = '83e8cfb9-1982-47ae-b515-3fec112da457';
    const historyListId = '1702fe62-6a47-4fd1-b935-0e3258073bb6';
    const { mapping: histMapping, internalNames: histInternals } = await getListColumnMapping(siteId, historyListId, token);

    console.log('[NC_ARCHIVE] histMapping:', histMapping);
    console.log('[NC_ARCHIVE] histInternals:', Array.from(histInternals));

    // Função segura para converter data DD/MM/YYYY ou YYYY-MM-DD para ISO
    const safeToISO = (dateStr: string | undefined): string | null => {
      if (!dateStr || dateStr.trim() === '') return null;
      // Já está em formato YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
        return dateStr + 'T12:00:00Z';
      }
      // Formato DD/MM/YYYY
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
        const [d, m, y] = dateStr.split('/');
        return `${y}-${m}-${d}T12:00:00Z`;
      }
      // Tenta parse genérico
      const parsed = new Date(dateStr);
      if (isNaN(parsed.getTime())) return null;
      return parsed.toISOString();
    };

    let successCount = 0;
    let failedCount = 0;
    let lastErrorMessage = "";

    for (const item of items) {
      try {
        const semana = item.semana || getWeekString(item.data);

        // Tenta resolver cada campo usando o mapping
        const fieldMap: Record<string, any> = {
          Semana: semana,
          Rota: item.rota,
          Data: safeToISO(item.data),
          'Código': item.codigo,
          Produtor: item.produtor,
          Motivo: item.motivo,
          'Observação': item.observacao,
          Ação: item.acao,
          'DataAção': safeToISO(item.dataAcao),
          'ÚltimaColeta': safeToISO(item.ultimaColeta),
          Culpabilidade: item.Culpabilidade,
          'Operação': item.operacao
        };

        // Campos read-only que NÃO podem ser escritos via Graph API
        const readOnlyFields = new Set([
          'LinkTitle', 'LinkTitleNoMenu', 'ID', 'ContentType', 'Modified', 'Created',
          'Author', 'Editor', '_UIVersionString', 'Attachments', 'Edit', 'DocIcon',
          'ItemChildCount', 'FolderChildCount', '_ComplianceFlags', '_ComplianceTag',
          '_ComplianceTagWrittenTime', '_ComplianceTagUserId', '_IsRecord', 'AppAuthor',
          'AppEditor', 'Title'
        ]);

        const histFields: any = { Title: item.rota };
        Object.entries(fieldMap).forEach(([displayName, value]) => {
          const intName = resolveFieldName(histMapping, displayName);
          console.log(`[NC_ARCHIVE] resolveFieldName("${displayName}") -> "${intName}" | readOnly: ${readOnlyFields.has(intName)}`);
          if (intName && histInternals.has(intName) && !readOnlyFields.has(intName)) {
            histFields[intName] = value;
          }
        });

        console.log('[NC_ARCHIVE] histFields final:', histFields);

        const postRes = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, { method: 'POST', body: JSON.stringify({ fields: histFields }) });
        console.log('[NC_ARCHIVE] POST result:', postRes);
        if (postRes && postRes.id) {
          await graphFetch(`/sites/${siteId}/lists/${sourceListId}/items/${item.id}`, token, { method: 'DELETE' });
          successCount++;
          console.log(`[NC_ARCHIVE] ✅ Item ${item.id} arquivado com sucesso`);
        } else {
          failedCount++;
          lastErrorMessage = "Failed to confirm archived NC ID.";
          console.error('[NC_ARCHIVE] ❌ Falha ao confirmar ID arquivado:', postRes);
        }
      } catch (err: any) {
        failedCount++;
        lastErrorMessage = err.message;
        console.error(`[NC_ARCHIVE] ❌ Erro ao arquivar item ${item.id}:`, err.message);
      }
    }
    console.log(`[NC_ARCHIVE] Final: success=${successCount}, failed=${failedCount}`);
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
