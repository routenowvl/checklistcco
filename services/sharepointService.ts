
// @google/genai guidelines: Use direct process.env.API_KEY, no UI for keys, use correct model names.
// Correct models: 'gemini-3-flash-preview', 'gemini-3-pro-preview', 'gemini-2.5-flash-image', etc.

import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord, RouteDeparture, RouteOperationMapping, RouteConfig } from '../types';

export interface DailyWarning {
  id: string;
  operacao: string; // Título
  celula: string;   // Email do responsável
  rota: string;
  descricao: string;
  dataOcorrencia: string; // ISO Date
  visualizado: boolean;
}

const SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH || "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, { mapping: Record<string, string>, readOnly: Set<string>, internalNames: Set<string> }> = {};

// Cache para dados estáticos/semi-estáticos (5 minutos)
const dataCache: Record<string, { data: any, timestamp: number }> = {};
const CACHE_TTL = 5 * 60 * 1000; // 5 minutos

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
        window.dispatchEvent(new CustomEvent('token-expired'));
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
    const today = new Date().toISOString().split('T')[0];
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
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
      return (data.value || []).map((item: any) => ({ id: item.id, timestamp: item.fields.Data, resetBy: item.fields.Title, email: (item.fields[celulaField] || "").toString().trim(), tasks: JSON.parse(item.fields.DadosJSON || '[]') })).filter((record: HistoryRecord) => record.email?.toLowerCase() === userEmail.toLowerCase().trim()).sort((a: any, b: any) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
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
            StatusResumoSaida: String(f[resolveFieldName(mapping, 'StatusResumoSaida')] || "") // Status do resumo
          };
          console.log('[DEBUG_SHAREPOINT] Config item:', {
            operacao: config.operacao,
            ultimoEnvioSaida_raw: f[resolveFieldName(mapping, 'UltimoEnvioSaida')],
            ultimoEnvioSaida: config.ultimoEnvioSaida,
            Status: config.Status,
            Envio: config.Envio,
            Copia: config.Copia,
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
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&t=${timestamp}`, token);

      const result = (data.value || []).map((item: any) => {
        const f = item.fields;
        return {
          id: String(item.id),
          semana: f[resolveFieldName(mapping, 'Semana')] || "",
          rota: f.Title || "",
          data: f[resolveFieldName(mapping, 'DataOperacao')] ? f[resolveFieldName(mapping, 'DataOperacao')].split('T')[0] : "",
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
          checklistMotorista: f[resolveFieldName(mapping, 'ChecklistMotorista')] || ""
        };
      });

      setCachedData(cacheKey, result);
      return result;
    } catch (e) { return []; }
  },

  async getArchivedDepartures(token: string, operation: string | null, startDate: string, endDate: string): Promise<RouteDeparture[]> {
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
      const data = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items?expand=fields&$filter=${filter}&$top=999`, token);
      
      const results = (data.value || []).map((item: any) => {
        const f = item.fields;
        return {
          id: String(item.id),
          semana: f[resolveFieldName(mapping, 'Semana')] || "",
          rota: f.Title || "",
          data: f[colData] ? f[colData].split('T')[0] : "",
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
          checklistMotorista: f[resolveFieldName(mapping, 'ChecklistMotorista')] || ""
        };
      });

      console.log(`[ARCHIVE_QUERY] Search success. Found ${results.length} records.`);
      return results;
    } catch (e: any) {
        console.error("[ARCHIVE_FETCH_ERROR] Error fetching archived data:", e.message);
        throw e;
    }
  },

  async updateDeparture(token: string, departure: RouteDeparture): Promise<string> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token, true);
    const raw: any = { Title: departure.rota, Semana: departure.semana, DataOperacao: departure.data ? new Date(departure.data + 'T12:00:00Z').toISOString() : null, HorarioInicio: departure.inicio, Motorista: departure.motorista, Placa: departure.placa, HorarioSaida: departure.saida, MotivoAtraso: departure.motivo, Observacao: departure.observacao, StatusGeral: departure.statusGeral, Aviso: departure.aviso, Operacao: departure.operacao, StatusOp: departure.statusOp, TempGab: departure.tempo, ChecklistMotorista: departure.checklistMotorista || '' };

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
            const raw: any = { Title: item.rota, Semana: item.semana, DataOperacao: item.data ? new Date(item.data + 'T12:00:00Z').toISOString() : null, HorarioInicio: item.inicio, Motorista: item.motorista, Placa: item.placa, HorarioSaida: item.saida, MotivoAtraso: item.motivo, Observacao: item.observacao, StatusGeral: item.statusGeral, Aviso: item.aviso, Operacao: item.operacao, StatusOp: item.statusOp, TempGab: item.tempo, ChecklistMotorista: item.checklistMotorista || '' };
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
  clearCache
};
