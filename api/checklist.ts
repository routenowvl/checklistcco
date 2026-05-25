// Endpoint consolidado Checklist (config, departures, non-collections)
import type { VercelRequest, VercelResponse } from '@vercel/node';
import {
  getAllConfigs,
  getConfigByOperacao,
  updateConfigField,
  updateConfigFields,
  updateConteudoIfChanged,
  updateConteudoNcoletasIfChanged,
  getLockStatus,
  acquireLock,
  releaseLock,
  getDepartures,
  upsertDeparture,
  deleteDeparture,
  getNonCollections,
  insertNonCollection,
  updateNonCollection,
  deleteNonCollection,
  insertConfig
} from './lib-checklistDb.js';
import { getGraphAppToken } from './lib-graphAppAuth.js';

const validateToken = async (authHeader: string | undefined): Promise<boolean> => {
  if (!authHeader) return false;
  const token = authHeader.startsWith('Bearer ')
    ? authHeader.slice(7).trim()
    : authHeader.trim();
  if (!token) return false;
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${token}` }
    });
    return res.ok;
  } catch {
    return false;
  }
};

const handleConfig = async (action: string, body: any, res: VercelResponse) => {
  switch (action) {
    case 'getAll': {
      const configs = await getAllConfigs();
      return res.status(200).json({ success: true, configs });
    }
    case 'getByOperacao': {
      const { operacao } = body;
      if (!operacao) return res.status(400).json({ success: false, error: 'operacao é obrigatório' });
      const config = await getConfigByOperacao(String(operacao));
      return res.status(200).json({ success: true, config });
    }
    case 'updateField': {
      const { operacao, field, value } = body;
      if (!operacao || !field) return res.status(400).json({ success: false, error: 'operacao e field são obrigatórios' });
      await updateConfigField(String(operacao), String(field), value);
      return res.status(200).json({ success: true });
    }
    case 'updateFields': {
      const { operacao, fields } = body;
      if (!operacao || !fields || typeof fields !== 'object') return res.status(400).json({ success: false, error: 'operacao e fields são obrigatórios' });
      await updateConfigFields(String(operacao), fields);
      return res.status(200).json({ success: true });
    }
    case 'updateConteudoIfChanged': {
      const { operacao, conteudo } = body;
      if (!operacao || conteudo === undefined) return res.status(400).json({ success: false, error: 'operacao e conteudo são obrigatórios' });
      const changed = await updateConteudoIfChanged(String(operacao), String(conteudo));
      return res.status(200).json({ success: true, changed });
    }
    case 'updateConteudoNcoletasIfChanged': {
      const { operacao, conteudoNcoletas } = body;
      if (!operacao || conteudoNcoletas === undefined) return res.status(400).json({ success: false, error: 'operacao e conteudoNcoletas são obrigatórios' });
      const changed = await updateConteudoNcoletasIfChanged(String(operacao), String(conteudoNcoletas));
      return res.status(200).json({ success: true, changed });
    }
    case 'getLockStatus': {
      const { operacao } = body;
      if (!operacao) return res.status(400).json({ success: false, error: 'operacao é obrigatório' });
      const lockStatus = await getLockStatus(String(operacao));
      return res.status(200).json({ success: true, lock: lockStatus });
    }
    case 'acquireLock': {
      const { operacao, userEmail, timestamp } = body;
      if (!operacao || !userEmail || !timestamp) return res.status(400).json({ success: false, error: 'operacao, userEmail e timestamp são obrigatórios' });
      await acquireLock(String(operacao), String(userEmail), String(timestamp));
      return res.status(200).json({ success: true });
    }
    case 'releaseLock': {
      const { operacao } = body;
      if (!operacao) return res.status(400).json({ success: false, error: 'operacao é obrigatório' });
      await releaseLock(String(operacao));
      return res.status(200).json({ success: true });
    }
    default:
      return res.status(400).json({ success: false, error: `Ação desconhecida: ${action}` });
  }
};

const handleDepartures = async (action: string, body: any, res: VercelResponse) => {
  switch (action) {
    case 'getAll': {
      const departures = await getDepartures();
      return res.status(200).json({ success: true, departures });
    }
    case 'upsert': {
      const { departure } = body;
      if (!departure || typeof departure !== 'object') return res.status(400).json({ success: false, error: 'departure é obrigatório' });
      const id = await upsertDeparture(departure);
      return res.status(200).json({ success: true, id });
    }
    case 'delete': {
      const { id } = body;
      if (!id) return res.status(400).json({ success: false, error: 'id é obrigatório' });
      await deleteDeparture(Number(id));
      return res.status(200).json({ success: true });
    }
    default:
      return res.status(400).json({ success: false, error: `Ação desconhecida: ${action}` });
  }
};

const handleNonCollections = async (action: string, body: any, res: VercelResponse) => {
  switch (action) {
    case 'getAll': {
      const nonCollections = await getNonCollections();
      return res.status(200).json({ success: true, nonCollections });
    }
    case 'insert': {
      const { nonCollection } = body;
      if (!nonCollection || typeof nonCollection !== 'object') return res.status(400).json({ success: false, error: 'nonCollection é obrigatório' });
      const id = await insertNonCollection(nonCollection);
      return res.status(200).json({ success: true, id });
    }
    case 'update': {
      const { nonCollection } = body;
      if (!nonCollection || typeof nonCollection !== 'object') return res.status(400).json({ success: false, error: 'nonCollection é obrigatório' });
      await updateNonCollection(nonCollection);
      return res.status(200).json({ success: true });
    }
    case 'delete': {
      const { id } = body;
      if (!id) return res.status(400).json({ success: false, error: 'id é obrigatório' });
      await deleteNonCollection(Number(id));
      return res.status(200).json({ success: true });
    }
    default:
      return res.status(400).json({ success: false, error: `Ação desconhecida: ${action}` });
  }
};

// ---------------------------------------------------------------------------
// Migration helpers (TEMPORÁRIO — remover após migração)
// ---------------------------------------------------------------------------

const SITE_PATH = process.env.VITE_SHAREPOINT_SITE_PATH || '';

const graphFetch = async (endpoint: string, token: string): Promise<any> => {
  const url = endpoint.startsWith('https://')
    ? endpoint
    : `https://graph.microsoft.com/v1.0${endpoint}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text.slice(0, 400)}`);
  }
  return res.status === 204 ? null : res.json();
};

const normalizeStr = (str: string): string =>
  str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '').trim();

const resolveField = (mapping: Record<string, string>, target: string): string =>
  mapping[normalizeStr(target)] || target;

const formatISOtoBR = (iso: any): string => {
  if (!iso) return '';
  const s = String(iso).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
};

/** Converte "DD/MM/YYYY HH:MM:SS" → "YYYY-MM-DDTHH:MM:SS" ou retorna null */
const brDatetimeToISO = (v: any): string | null => {
  if (!v) return null;
  const s = String(v).trim();
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}:\d{2}:\d{2})$/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}T${m[4]}`;
  // Já é ISO?
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s;
  return null;
};

const parseUltimoEnvioNcoletas = (raw: any): { datetime: string | null; quantidade: number } => {
  if (!raw) return { datetime: null, quantidade: 0 };
  const s = String(raw).trim();
  const match = s.match(/^(.+?\d{2}:\d{2}:\d{2})\s+(\d+)$/);
  if (match) return { datetime: match[1].trim(), quantidade: parseInt(match[2], 10) || 0 };
  return { datetime: s, quantidade: 0 };
};

const parseNumericId = (value: unknown): number | null => {
  if (value == null) return null;
  if (typeof value === 'number' && Number.isFinite(value)) return Math.trunc(value);
  const raw = String(value).trim();
  if (!raw) return null;
  const match = raw.match(/-?\d+(?:[.,]\d+)?/);
  if (!match) return null;
  const parsed = Number(match[0].replace(',', '.'));
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
};

const extractPlantId = (fields: Record<string, any>, mapping: Record<string, string>): any => {
  for (const c of ['Plant_id', 'Plant Id', 'PlantId', 'plant_id', 'IdPlant']) {
    const resolved = resolveField(mapping, c);
    if (fields?.[resolved] != null && String(fields[resolved]).trim() !== '') return fields[resolved];
  }
  for (const [key, value] of Object.entries(fields || {})) {
    const nk = normalizeStr(key);
    if ((nk.includes('plantid') || nk.includes('idplant')) && value != null && String(value).trim() !== '') return value;
  }
  return null;
};

const getColumnMapping = async (siteId: string, listId: string, token: string): Promise<Record<string, string>> => {
  const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
  const mapping: Record<string, string> = {};
  for (const col of columns.value || []) {
    mapping[normalizeStr(col.name)] = col.name;
    mapping[normalizeStr(col.displayName)] = col.name;
  }
  return mapping;
};

const formatTime = (v: any): string => {
  if (!v) return '';
  const s = String(v).trim();
  if (s === '-') return '';
  // BR datetime "DD/MM/YYYY HH:MM:SS" → extrai hora
  const brMatch = s.match(/(\d{2}:\d{2}):\d{2}$/);
  if (brMatch) return brMatch[1] + ':00';
  const dtMatch = s.match(/T(\d{2}:\d{2})/);
  if (dtMatch) return dtMatch[1] + ':00';
  const tMatch = s.match(/^(\d{2}:\d{2})/);
  return tMatch ? tMatch[1] + ':00' : '';
};

const handleMigrateConfig = async (_action: string, _body: any, res: VercelResponse) => {
  const appToken = await getGraphAppToken();
  const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
  const siteId = siteData.id;

  let list: any;
  try {
    list = await graphFetch(`/sites/${siteId}/lists/CONFIG_OPERACAO_SAIDA_DE_ROTAS`, appToken);
  } catch {
    const listsData = await graphFetch(`/sites/${siteId}/lists`, appToken);
    list = (listsData.value || []).find(
      (l: any) => l.name?.toLowerCase() === 'config_operacao_saida_de_rotas' || l.displayName?.toLowerCase() === 'config_operacao_saida_de_rotas'
    );
    if (!list) throw new Error('Lista CONFIG_OPERACAO_SAIDA_DE_ROTAS não encontrada');
  }

  const mapping = await getColumnMapping(siteId, list.id, appToken);

  let allItems: any[] = [];
  let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;
  while (nextUrl) {
    const data = await graphFetch(nextUrl, appToken);
    allItems = allItems.concat(data.value || []);
    nextUrl = data['@odata.nextLink'] || null;
  }

  let upserted = 0, skipped = 0;
  const errors: string[] = [];

  for (const item of allItems) {
    const f = item.fields || {};
    const operacaoField = resolveField(mapping, 'OPERACAO');
    const operacao = String(f[operacaoField] || f.Title || '').trim();
    if (!operacao) { skipped++; continue; }

    try {
      const plantRaw = extractPlantId(f, mapping);
      const plantId = parseNumericId(plantRaw);
      const ncoleta = parseUltimoEnvioNcoletas(f[resolveField(mapping, 'UltimoEnvioNcoleta')]);

      const row: Record<string, unknown> = {
        operacao,
        email: String(f[resolveField(mapping, 'EMAIL')] || '').toLowerCase().trim(),
        tolerancia: String(f[resolveField(mapping, 'TOLERANCIA')] || '00:00:00'),
        nome_exibicao: String(f[resolveField(mapping, 'NomeExibicao')] || operacao),
        plant_id: plantId,
        ultimo_envio_saida: brDatetimeToISO(f[resolveField(mapping, 'UltimoEnvioSaida')]),
        status: String(f[resolveField(mapping, 'Status')] || ''),
        envio: String(f[resolveField(mapping, 'Envio')] || ''),
        copia: String(f[resolveField(mapping, 'Copia')] || ''),
        ultimo_envio_resumo_saida: brDatetimeToISO(f[resolveField(mapping, 'UltimoEnvioResumoSaida')]),
        status_resumo_saida: String(f[resolveField(mapping, 'StatusResumoSaida')] || ''),
        ultimo_envio_ncoleta: brDatetimeToISO(ncoleta.datetime),
        quantidade_ncoletas_registrada: ncoleta.quantidade,
        conteudo: String(f[resolveField(mapping, 'Conteudo')] || ''),
        conteudo_ncoletas: String(f[resolveField(mapping, 'ConteudoNcoletas')] || ''),
        lock_envio: f[resolveField(mapping, 'LockEnvio')] || null,
        lock_user: String(f[resolveField(mapping, 'LockUser')] || ''),
        lock_timestamp: brDatetimeToISO(f[resolveField(mapping, 'LockTimestamp')]),
      };
      await insertConfig(row);
      upserted++;
    } catch (err: any) {
      errors.push(`${operacao}: ${err.message}`);
    }
  }

  return res.status(200).json({ success: true, total: allItems.length, upserted, skipped, errors: errors.length > 0 ? errors.slice(0, 20) : undefined });
};

const handleMigrateDepartures = async (_action: string, _body: any, res: VercelResponse) => {
  const appToken = await getGraphAppToken();
  const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
  const siteId = siteData.id;

  let list: any;
  try {
    list = await graphFetch(`/sites/${siteId}/lists/Dados_Saida_de_rotas`, appToken);
  } catch {
    const listsData = await graphFetch(`/sites/${siteId}/lists`, appToken);
    list = (listsData.value || []).find(
      (l: any) => l.name?.toLowerCase() === 'dados_saida_de_rotas' || l.displayName?.toLowerCase() === 'dados_saida_de_rotas'
    );
    if (!list) throw new Error('Lista Dados_Saida_de_rotas não encontrada');
  }

  const mapping = await getColumnMapping(siteId, list.id, appToken);

  let allItems: any[] = [];
  let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;
  while (nextUrl) {
    const data = await graphFetch(nextUrl, appToken);
    allItems = allItems.concat(data.value || []);
    nextUrl = data['@odata.nextLink'] || null;
  }

  let upserted = 0;
  const errors: string[] = [];

  for (const item of allItems) {
    const f = item.fields || {};
    try {
      const dataBR = formatISOtoBR(f[resolveField(mapping, 'DataOperacao')]);
      if (!dataBR) { errors.push(`Item ${item.id}: sem data`); continue; }

      const row: Record<string, unknown> = {
        operacao: String(f[resolveField(mapping, 'Operacao')] || '').trim(),
        rota: String(f.Title || '').trim(),
        motorista: String(f[resolveField(mapping, 'Motorista')] || '').trim(),
        placa: String(f[resolveField(mapping, 'Placa')] || '').trim(),
        contato: String(f[resolveField(mapping, 'Contato')] || '').replace(/\D/g, ''),
        inicio: formatTime(f[resolveField(mapping, 'HorarioInicio')]),
        saida: formatTime(f[resolveField(mapping, 'HorarioSaida')]),
        statusGeral: String(f[resolveField(mapping, 'StatusGeral')] || '').trim(),
        motivo: String(f[resolveField(mapping, 'MotivoAtraso')] || '').trim(),
        observacao: String(f[resolveField(mapping, 'Observacao')] || '').trim(),
        data: dataBR,
        statusOp: String(f[resolveField(mapping, 'StatusOp')] || 'Previsto').trim(),
        checklistMotorista: String(f[resolveField(mapping, 'ChecklistMotorista')] || '').trim(),
        retornoMotorista: String(f[resolveField(mapping, 'RetornoMotorista')] || '').trim(),
        causaRaiz: String(f[resolveField(mapping, 'CausaRaiz')] || '').trim(),
        tempoResposta: String(f[resolveField(mapping, 'TempoResposta')] || '').trim(),
        logTempoResposta: String(f[resolveField(mapping, 'LogTempoResposta')] || '').trim(),
      };
      await upsertDeparture(row);
      upserted++;
    } catch (err: any) {
      errors.push(`Item ${item.id}: ${err.message}`);
    }
  }

  return res.status(200).json({ success: true, total: allItems.length, upserted, errors: errors.length > 0 ? errors.slice(0, 20) : undefined });
};

const handleMigrateNonCollections = async (_action: string, _body: any, res: VercelResponse) => {
  const LIST_ID = '83e8cfb9-1982-47ae-b515-3fec112da457';
  const appToken = await getGraphAppToken();
  const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
  const siteId = siteData.id;

  const mapping = await getColumnMapping(siteId, LIST_ID, appToken);

  let allItems: any[] = [];
  let nextUrl: string | null = `/sites/${siteId}/lists/${LIST_ID}/items?expand=fields&$top=100`;
  while (nextUrl) {
    const data = await graphFetch(nextUrl, appToken);
    allItems = allItems.concat(data.value || []);
    nextUrl = data['@odata.nextLink'] || null;
  }

  let upserted = 0;
  const errors: string[] = [];

  for (const item of allItems) {
    const f = item.fields || {};
    try {
      const dataRaw = f[resolveField(mapping, 'Data')];
      const dataBR = formatISOtoBR(dataRaw);
      if (!dataBR) { errors.push(`Item ${item.id}: sem data`); continue; }

      const row: Record<string, unknown> = {
        operacao: String(f[resolveField(mapping, 'Operacao')] || f['Opera_x00e7__x00e3_o'] || '').trim(),
        rota: String(f.Title || '').trim(),
        data: dataBR,
        semana: String(f[resolveField(mapping, 'Semana')] || f.Title || '').trim(),
        codigo: String(f[resolveField(mapping, 'Codigo')] || f['C_x00f3_digo'] || '').trim(),
        produtor: String(f[resolveField(mapping, 'Produtor')] || '').trim(),
        motivo: String(f[resolveField(mapping, 'Motivo')] || '').trim(),
        observacao: String(f[resolveField(mapping, 'Observacao')] || f['Observa_x00e7__x00e3_o'] || '').trim(),
        acao: String(f[resolveField(mapping, 'Acao')] || f['A_x00e7__x00e3_o'] || '').trim(),
        dataAcao: formatISOtoBR(f[resolveField(mapping, 'DataAcao')] || f['DataA_x00e7__x00e3_o']),
        ultimaColeta: formatISOtoBR(f[resolveField(mapping, 'UltimaColeta')] || f['_x00da_ltimaColeta']),
        Culpabilidade: String(f[resolveField(mapping, 'Culpabilidade')] || '').trim(),
        causaRaiz: String(f[resolveField(mapping, 'CausaRaiz')] || '').trim(),
      };
      await insertNonCollection(row);
      upserted++;
    } catch (err: any) {
      errors.push(`Item ${item.id}: ${err.message}`);
    }
  }

  return res.status(200).json({ success: true, total: allItems.length, upserted, errors: errors.length > 0 ? errors.slice(0, 20) : undefined });
};

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const { domain, action } = req.body;
    if (!domain) return res.status(400).json({ success: false, error: 'domain é obrigatório' });

    // Migration domains (TEMPORÁRIO — sem auth)
    if (domain === 'migrate-config') return await handleMigrateConfig(action, req.body, res);
    if (domain === 'migrate-departures') return await handleMigrateDepartures(action, req.body, res);
    if (domain === 'migrate-non-collections') return await handleMigrateNonCollections(action, req.body, res);

    // Domains normais exigem auth
    const isAuthenticated = await validateToken(req.headers.authorization);
    if (!isAuthenticated) {
      return res.status(401).json({ success: false, error: 'Unauthorized' });
    }

    switch (domain) {
      case 'config':
        return await handleConfig(action, req.body, res);
      case 'departures':
        return await handleDepartures(action, req.body, res);
      case 'non-collections':
        return await handleNonCollections(action, req.body, res);
      default:
        return res.status(400).json({ success: false, error: `Domínio desconhecido: ${domain}` });
    }
  } catch (error: any) {
    console.error('[CHECKLIST] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: 'Erro ao processar operação' });
  }
}
