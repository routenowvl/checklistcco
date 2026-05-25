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
  deleteNonCollection
} from './lib-checklistDb.js';

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

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Method not allowed' });
  }

  const isAuthenticated = await validateToken(req.headers.authorization);
  if (!isAuthenticated) {
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }

  try {
    const { domain, action } = req.body;
    if (!domain) return res.status(400).json({ success: false, error: 'domain é obrigatório (config, departures, non-collections)' });

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
