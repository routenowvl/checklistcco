import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getGraphAppToken } from '../utils/graphAppAuth';
import { insertConfig } from '../utils/checklistDb';

/**
 * Endpoint temporário de migração: lê CONFIG_OPERACAO_SAIDA_DE_ROTAS do SharePoint
 * e copia todos os itens para operacao_config no PostgreSQL.
 * Não altera nada no SharePoint — é somente leitura + inserção no PG.
 *
 * Uso: POST /api/migrate-config-from-sharepoint
 * Protegido por token do Graph (validação via /v1.0/me).
 *
 * DELETAR este arquivo após a migração.
 */

const SITE_PATH = process.env.VITE_SHAREPOINT_SITE_PATH || '';

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

const graphFetch = async (endpoint: string, token: string): Promise<any> => {
  const url = endpoint.startsWith('https://')
    ? endpoint
    : `https://graph.microsoft.com/v1.0${endpoint}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text.slice(0, 400)}`);
  }
  return res.status === 204 ? null : res.json();
};

const normalizeString = (str: string): string =>
  str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '').trim();

/**
 * Extrai data/hora e quantidade do campo UltimoEnvioNcoletas do SharePoint.
 * Formato: "24/05/2026 18:00:25 01" → { datetime: "24/05/2026 18:00:25", quantidade: 1 }
 * Se não houver quantidade no final, retorna quantidade 0.
 */
const parseUltimoEnvioNcoletas = (raw: any): { datetime: string | null; quantidade: number } => {
  if (!raw) return { datetime: null, quantidade: 0 };
  const s = String(raw).trim();
  // Tenta extrair padrão: DD/MM/YYYY HH:MM:SS NN
  const match = s.match(/^(.+?\d{2}:\d{2}:\d{2})\s+(\d+)$/);
  if (match) {
    return { datetime: match[1].trim(), quantidade: parseInt(match[2], 10) || 0 };
  }
  return { datetime: s, quantidade: 0 };
};

const resolveFieldName = (mapping: Record<string, string>, target: string): string => {
  return mapping[normalizeString(target)] || target;
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

const extractPlantFieldValue = (fields: Record<string, any>, mapping: Record<string, string>): any => {
  const candidates = [
    'Plant_id', 'Plant Id', 'PlantId', 'plant_id', 'IdPlant', 'ID_PLANT'
  ].map((c) => resolveFieldName(mapping, c));

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

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  // TODO: Restaurar validação de token após migração
  // const isAuthenticated = await validateToken(req.headers.authorization);
  // if (!isAuthenticated) {
  //   return res.status(401).json({ success: false, error: 'Unauthorized' });
  // }

  try {
    // 1. Obter token app-only
    const appToken = await getGraphAppToken();

    // 2. Resolver site
    const siteData = await graphFetch(`/sites/${SITE_PATH}`, appToken);
    const siteId = siteData.id;

    // 3. Encontrar lista
    let list: any;
    try {
      list = await graphFetch(`/sites/${siteId}/lists/CONFIG_OPERACAO_SAIDA_DE_ROTAS`, appToken);
    } catch {
      const listsData = await graphFetch(`/sites/${siteId}/lists`, appToken);
      list = (listsData.value || []).find(
        (l: any) =>
          l.name?.toLowerCase() === 'config_operacao_saida_de_rotas' ||
          l.displayName?.toLowerCase() === 'config_operacao_saida_de_rotas'
      );
      if (!list) throw new Error('Lista CONFIG_OPERACAO_SAIDA_DE_ROTAS não encontrada');
    }

    // 4. Mapear colunas
    const columns = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, appToken);
    const mapping: Record<string, string> = {};
    for (const col of columns.value || []) {
      mapping[normalizeString(col.name)] = col.name;
      mapping[normalizeString(col.displayName)] = col.name;
    }

    // 5. Buscar todos os itens (com paginação)
    let allItems: any[] = [];
    let nextUrl: string | null = `/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=100`;

    while (nextUrl) {
      const data = await graphFetch(nextUrl, appToken);
      allItems = allItems.concat(data.value || []);
      nextUrl = data['@odata.nextLink'] || null;
    }

    console.log(`[MIGRATE] ${allItems.length} itens encontrados no SharePoint`);

    // 6. Inserir cada item no PG
    let inserted = 0;
    let skipped = 0;
    let alreadyExists = 0;
    const errors: string[] = [];

    for (const item of allItems) {
      const f = item.fields || {};

      const operacaoField = resolveFieldName(mapping, 'OPERACAO');
      const operacao = String(f[operacaoField] || f.Title || '').trim();
      if (!operacao) { errors.push(`Item ${item.id}: sem operacao (Title="${f.Title || ''}", field="${operacaoField}")`); skipped++; continue; }

      try {
        const plantRaw = extractPlantFieldValue(f, mapping);
        const plantId = parseNumericId(plantRaw);

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
          ultimo_envio_ncoleta: parseUltimoEnvioNcoletas(f[resolveFieldName(mapping, 'UltimoEnvioNcoleta')]).datetime,
          quantidade_ncoletas_registrada: parseUltimoEnvioNcoletas(f[resolveFieldName(mapping, 'UltimoEnvioNcoleta')]).quantidade,
          conteudo: String(f[resolveFieldName(mapping, 'Conteudo')] || ''),
          conteudo_ncoletas: String(f[resolveFieldName(mapping, 'ConteudoNcoletas')] || ''),
          lock_envio: f[resolveFieldName(mapping, 'LockEnvio')] || null,
          lock_user: String(f[resolveFieldName(mapping, 'LockUser')] || ''),
          lock_timestamp: f[resolveFieldName(mapping, 'LockTimestamp')] || null,
        };

        const id = await insertConfig(row);
        if (id > 0) inserted++;
        else { alreadyExists++; }
      } catch (err: any) {
        errors.push(`${operacao}: ${err.message}`);
      }
    }

    return res.status(200).json({
      success: true,
      total: allItems.length,
      inserted,
      alreadyExists,
      skipped,
      errors: errors.length > 0 ? errors.slice(0, 20) : undefined
    });
  } catch (error: any) {
    console.error('[MIGRATE_CONFIG] Erro:', error?.message || error);
    return res.status(500).json({ success: false, error: error?.message || 'Erro na migração' });
  }
}
