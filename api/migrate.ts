import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getGraphAppToken } from '../utils/graphAppAuth';
import { insertConfig, upsertDeparture, insertNonCollection } from '../utils/checklistDb';

/**
 * Endpoint temporário de migração: lê listas do SharePoint e copia para o PostgreSQL.
 *
 * Uso: POST /api/migrate
 * Body: { "list": "config"|"departures"|"non-collections" }
 *
 * DELETAR este arquivo após a migração.
 */

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

const normalizeString = (str: string): string =>
  str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '').trim();

const resolveFieldName = (mapping: Record<string, string>, target: string): string =>
  mapping[normalizeString(target)] || target;

const formatISOtoBR = (iso: any): string => {
  if (!iso) return '';
  const s = String(iso).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : s;
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

const brDatetimeToISO = (s: string): string | null => {
  const m = s.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  return m ? `${m[3]}-${m[2]}-${m[1]}T${m[4]}:${m[5]}:${m[6]}` : null;
};

const parseUltimoEnvioNcoletas = (raw: any): { datetime: string | null; quantidade: number } => {
  if (!raw) return { datetime: null, quantidade: 0 };
  const s = String(raw).trim();
  const match = s.match(/^(.+?\d{2}:\d{2}:\d{2})\s+(\d+)$/);
  if (match) {
    const iso = brDatetimeToISO(match[1].trim());
    return { datetime: iso || match[1].trim(), quantidade: parseInt(match[2], 10) || 0 };
  }
  const iso2 = brDatetimeToISO(s);
  return iso2 ? { datetime: iso2, quantidade: 0 } : { datetime: s, quantidade: 0 };
};

const formatTime = (v: any): string => {
  if (!v) return '';
  const s = String(v).trim();
  if (!s || s === '-') return '';
  const brDt = s.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
  if (brDt) return `${brDt[4]}:${brDt[5]}:00`;
  const dtMatch = s.match(/T(\d{2}:\d{2})/);
  if (dtMatch) return dtMatch[1] + ':00';
  const tMatch = s.match(/^(\d{2}:\d{2})/);
  return tMatch ? tMatch[1] + ':00' : '';
};

const extractPlantFieldValue = (fields: Record<string, any>, mapping: Record<string, string>): any => {
  const candidates = ['Plant_id', 'Plant Id', 'PlantId', 'plant_id', 'IdPlant', 'ID_PLANT'].map(c => resolveFieldName(mapping, c));
  for (const candidate of candidates) {
    if (!candidate) continue;
    const value = fields?.[candidate];
    if (value != null && String(value).trim() !== '') return value;
  }
  for (const [key, value] of Object.entries(fields || {})) {
    const normalizedKey = normalizeString(key);
    if (normalizedKey.includes('plantid') || normalizedKey.includes('idplant')) {
      if (value != null && String(value).trim() !== '') return value;
    }
  }
  return null;
};

const fetchAllItems = async (siteId: string, listId: string, token: string): Promise<any[]> => {
  let allItems: any[] = [];
  let nextUrl: string | null = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=100`;
  while (nextUrl) {
    const data = await graphFetch(nextUrl, token);
    allItems = allItems.concat(data.value || []);
    nextUrl = data['@odata.nextLink'] || null;
  }
  return allItems;
};

const getColumnMapping = async (siteId: string, listId: string, token: string): Promise<Record<string, string>> => {
  const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
  const mapping: Record<string, string> = {};
  for (const col of columns.value || []) {
    mapping[normalizeString(col.name)] = col.name;
    mapping[normalizeString(col.displayName)] = col.name;
  }
  return mapping;
};

const findList = async (siteId: string, idOrName: string, token: string): Promise<any> => {
  try {
    return await graphFetch(`/sites/${siteId}/lists/${idOrName}`, token);
  } catch {
    const listsData = await graphFetch(`/sites/${siteId}/lists`, token);
    const found = (listsData.value || []).find(
      (l: any) => l.name?.toLowerCase() === idOrName.toLowerCase() || l.displayName?.toLowerCase() === idOrName.toLowerCase()
    );
    if (!found) throw new Error(`Lista ${idOrName} não encontrada`);
    return found;
  }
};

// ─── Config migration ────────────────────────────────────────────────────

const migrateConfig = async (siteId: string, token: string): Promise<{ total: number; inserted: number; alreadyExists: number; skipped: number; errors: string[] }> => {
  const list = await findList(siteId, 'CONFIG_OPERACAO_SAIDA_DE_ROTAS', token);
  const mapping = await getColumnMapping(siteId, list.id, token);
  const allItems = await fetchAllItems(siteId, list.id, token);

  let inserted = 0, skipped = 0, alreadyExists = 0;
  const errors: string[] = [];

  for (const item of allItems) {
    const f = item.fields || {};
    const operacaoField = resolveFieldName(mapping, 'OPERACAO');
    const operacao = String(f[operacaoField] || f.Title || '').trim();
    if (!operacao) { errors.push(`Item ${item.id}: sem operacao`); skipped++; continue; }

    try {
      const plantRaw = extractPlantFieldValue(f, mapping);
      const plantId = parseNumericId(plantRaw);
      const nc = parseUltimoEnvioNcoletas(f[resolveFieldName(mapping, 'UltimoEnvioNcoleta')]);

      const row: Record<string, unknown> = {
        operacao,
        email: String(f[resolveFieldName(mapping, 'EMAIL')] || '').toLowerCase().trim(),
        tolerancia: String(f[resolveFieldName(mapping, 'TOLERANCIA')] || '00:00:00'),
        nome_exibicao: String(f[resolveFieldName(mapping, 'NomeExibicao')] || operacao),
        plant_id: plantId,
        ultimo_envio_saida: f[resolveFieldName(mapping, 'UltimoEnvioSaida')] || null,
        status: String(f[resolveFieldName(mapping, 'Status')] || ''),
        envio: String(f[resolveFieldName(mapping, 'Envio')] || ''),
        copia: String(f[resolveFieldName(mapping, 'Copia')] || ''),
        ultimo_envio_resumo_saida: f[resolveFieldName(mapping, 'UltimoEnvioResumoSaida')] || null,
        status_resumo_saida: String(f[resolveFieldName(mapping, 'StatusResumoSaida')] || ''),
        ultimo_envio_ncoleta: nc.datetime,
        quantidade_ncoletas_registrada: nc.quantidade,
        conteudo: String(f[resolveFieldName(mapping, 'Conteudo')] || ''),
        conteudo_ncoletas: String(f[resolveFieldName(mapping, 'ConteudoNcoletas')] || ''),
        lock_envio: f[resolveFieldName(mapping, 'LockEnvio')] || null,
        lock_user: String(f[resolveFieldName(mapping, 'LockUser')] || ''),
        lock_timestamp: f[resolveFieldName(mapping, 'LockTimestamp')] || null,
      };

      const id = await insertConfig(row);
      if (id > 0) inserted++; else alreadyExists++;
    } catch (err: any) {
      errors.push(`${operacao}: ${err.message}`);
    }
  }

  return { total: allItems.length, inserted, alreadyExists, skipped, errors };
};

// ─── Departures migration ────────────────────────────────────────────────

const migrateDepartures = async (siteId: string, token: string): Promise<{ total: number; inserted: number; errors: string[] }> => {
  const list = await findList(siteId, 'Dados_Saida_de_rotas', token);
  const mapping = await getColumnMapping(siteId, list.id, token);
  const allItems = await fetchAllItems(siteId, list.id, token);

  let inserted = 0;
  const errors: string[] = [];

  for (const item of allItems) {
    const f = item.fields || {};
    try {
      const colData = resolveFieldName(mapping, 'DataOperacao');
      const colOp = resolveFieldName(mapping, 'Operacao');
      const dataBR = formatISOtoBR(f[colData]);
      if (!dataBR) { errors.push(`Item ${item.id}: sem data`); continue; }

      const row: Record<string, unknown> = {
        operacao: String(f[colOp] || '').trim(),
        rota: String(f.Title || '').trim(),
        motorista: String(f[resolveFieldName(mapping, 'Motorista')] || '').trim(),
        placa: String(f[resolveFieldName(mapping, 'Placa')] || '').trim(),
        contato: String(f[resolveFieldName(mapping, 'Contato')] || '').replace(/\D/g, ''),
        inicio: formatTime(f[resolveFieldName(mapping, 'HorarioInicio')]),
        saida: formatTime(f[resolveFieldName(mapping, 'HorarioSaida')]),
        statusGeral: String(f[resolveFieldName(mapping, 'StatusGeral')] || '').trim(),
        motivo: String(f[resolveFieldName(mapping, 'MotivoAtraso')] || '').trim(),
        observacao: String(f[resolveFieldName(mapping, 'Observacao')] || '').trim(),
        data: dataBR,
        statusOp: String(f[resolveFieldName(mapping, 'StatusOp')] || 'Previsto').trim(),
        checklistMotorista: String(f[resolveFieldName(mapping, 'ChecklistMotorista')] || '').trim(),
        retornoMotorista: String(f[resolveFieldName(mapping, 'RetornoMotorista')] || '').trim(),
        causaRaiz: String(f[resolveFieldName(mapping, 'CausaRaiz')] || '').trim(),
        tempoResposta: String(f[resolveFieldName(mapping, 'TempoResposta')] || '').trim(),
        logTempoResposta: String(f[resolveFieldName(mapping, 'LogTempoResposta')] || '').trim(),
      };

      await upsertDeparture(row);
      inserted++;
    } catch (err: any) {
      errors.push(`Item ${item.id}: ${err.message}`);
    }
  }

  return { total: allItems.length, inserted, errors };
};

// ─── Non-collections migration ───────────────────────────────────────────

const migrateNonCollections = async (siteId: string, token: string): Promise<{ total: number; inserted: number; errors: string[] }> => {
  const LIST_ID = '83e8cfb9-1982-47ae-b515-3fec112da457';
  const mapping = await getColumnMapping(siteId, LIST_ID, token);
  const allItems = await fetchAllItems(siteId, LIST_ID, token);

  let inserted = 0;
  const errors: string[] = [];

  for (const item of allItems) {
    const f = item.fields || {};
    try {
      const colOp = resolveFieldName(mapping, 'Operacao');
      const colObs = resolveFieldName(mapping, 'Observacao');
      const dataRaw = f[resolveFieldName(mapping, 'Data')];
      const dataBR = formatISOtoBR(dataRaw);
      if (!dataBR) { errors.push(`Item ${item.id}: sem data`); continue; }

      const row: Record<string, unknown> = {
        operacao: String(f[colOp] || f['Opera_x00e7__x00e3_o'] || '').trim(),
        rota: String(f.Title || '').trim(),
        data: dataBR,
        semana: String(f[resolveFieldName(mapping, 'Semana')] || f.Title || '').trim(),
        codigo: String(f[resolveFieldName(mapping, 'Codigo')] || f['C_x00f3_digo'] || '').trim(),
        produtor: String(f[resolveFieldName(mapping, 'Produtor')] || '').trim(),
        motivo: String(f[resolveFieldName(mapping, 'Motivo')] || '').trim(),
        observacao: String(f[colObs] || f['Observa_x00e7__x00e3_o'] || '').trim(),
        acao: String(f[resolveFieldName(mapping, 'Acao')] || f['A_x00e7__x00e3_o'] || '').trim(),
        dataAcao: formatISOtoBR(f[resolveFieldName(mapping, 'DataAcao')] || f['DataA_x00e7__x00e3_o']),
        ultimaColeta: formatISOtoBR(f[resolveFieldName(mapping, 'UltimaColeta')] || f['_x00da_ltimaColeta']),
        Culpabilidade: String(f[resolveFieldName(mapping, 'Culpabilidade')] || '').trim(),
        causaRaiz: String(f[resolveFieldName(mapping, 'CausaRaiz')] || '').trim(),
      };

      await insertNonCollection(row);
      inserted++;
    } catch (err: any) {
      errors.push(`Item ${item.id}: ${err.message}`);
    }
  }

  return { total: allItems.length, inserted, errors };
};

// ─── Handler ─────────────────────────────────────────────────────────────

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  try {
    const { list } = req.body || {};
    if (!list) return res.status(400).json({ success: false, error: 'list é obrigatório (config, departures, non-collections)' });

    const appToken = await getGraphAppToken();
    const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
    const siteId = siteData.id;

    let result: any;
    switch (list) {
      case 'config':
        result = await migrateConfig(siteId, appToken);
        break;
      case 'departures':
        result = await migrateDepartures(siteId, appToken);
        break;
      case 'non-collections':
        result = await migrateNonCollections(siteId, appToken);
        break;
      default:
        return res.status(400).json({ success: false, error: `Lista desconhecida: ${list}` });
    }

    return res.status(200).json({ success: true, ...result, errors: result.errors?.length > 0 ? result.errors.slice(0, 20) : undefined });
  } catch (error: any) {
    console.error('[MIGRATE] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: error?.message || 'Erro na migração' });
  }
}
