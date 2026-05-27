import pg from 'pg';

const { Pool } = pg;

let pool: pg.Pool | null = null;

const getPool = (): pg.Pool => {
  if (pool) return pool;

  const dbUrl = String(process.env.CHECKLIST_DB_URL || '').trim();
  if (!dbUrl) {
    throw new Error('CHECKLIST_DB_URL não configurada');
  }

  const ssl = String(process.env.CHECKLIST_DB_SSL || 'true').trim().toLowerCase();
  const schema = String(process.env.CHECKLIST_DB_SCHEMA || 'public').trim();

  pool = new Pool({
    connectionString: dbUrl,
    ssl: ssl === 'true' || ssl === '1' ? { rejectUnauthorized: false } : undefined,
    max: 4,
    idleTimeoutMillis: 15_000,
    connectionTimeoutMillis: 15_000,
    keepAlive: true,
    keepAliveInitialDelayMillis: 10_000,
    allowExitOnIdle: false
  });

  pool.on('connect', (client) => {
    client.query(`SET search_path TO ${schema}`);
  });

  pool.on('error', (err) => {
    console.error('[CHECKLIST_DB] Pool error (idle connection):', err.message);
  });

  pool.on('remove', (client) => {
    console.log('[CHECKLIST_DB] Connection removed from pool');
  });

  return pool;
};

// ---------------------------------------------------------------------------
// operacao_config
// ---------------------------------------------------------------------------

export type ConfigRow = {
  operacao: string;
  email: string;
  tolerancia: string;
  nome_exibicao: string;
  plant_id: number | null;
  conteudo: string;
  conteudo_ncoletas: string;
  ultimo_envio_saida: string | null;
  status: string;
  envio: string;
  copia: string;
  ultimo_envio_resumo_saida: string | null;
  ultimo_envio_ncoleta: string | null;
  quantidade_ncoletas_registrada: number;
  status_resumo_saida: string;
  codigo_kmm: string;
  lock_envio: string;
  lock_user: string;
  lock_timestamp: string;
};

export const getAllConfigs = async (): Promise<ConfigRow[]> => {
  const client = getPool();
  const result = await client.query('SELECT * FROM operacao_config ORDER BY operacao');
  return result.rows as ConfigRow[];
};

export const getConfigByOperacao = async (operacao: string): Promise<ConfigRow | null> => {
  const client = getPool();
  const result = await client.query('SELECT * FROM operacao_config WHERE operacao = $1', [operacao]);
  return (result.rows[0] as ConfigRow) || null;
};

export const updateConfigField = async (operacao: string, field: string, value: unknown): Promise<void> => {
  const client = getPool();
  const safeField = field.replace(/[^a-z_]/g, '');
  await client.query(
    `UPDATE operacao_config SET ${safeField} = $1, atualizado_em = NOW() WHERE operacao = $2`,
    [value, operacao]
  );
};

export const updateConfigFields = async (operacao: string, fields: Record<string, unknown>): Promise<void> => {
  const client = getPool();
  const keys = Object.keys(fields).map(k => k.replace(/[^a-z_]/g, ''));
  const values = Object.values(fields);
  const setClauses = keys.map((k, i) => `${k} = $${i + 1}`).join(', ');
  await client.query(
    `UPDATE operacao_config SET ${setClauses}, atualizado_em = NOW() WHERE operacao = $${values.length + 1}`,
    [...values, operacao]
  );
};

export const updateConteudoIfChanged = async (operacao: string, conteudo: string): Promise<boolean> => {
  const client = getPool();
  const result = await client.query(
    'SELECT conteudo FROM operacao_config WHERE operacao = $1',
    [operacao]
  );
  if (result.rows.length === 0) return false;
  const atual = String(result.rows[0].conteudo || '');
  if (atual === conteudo) return false;
  await client.query(
    'UPDATE operacao_config SET conteudo = $1, atualizado_em = NOW() WHERE operacao = $2',
    [conteudo, operacao]
  );
  return true;
};

export const updateConteudoNcoletasIfChanged = async (operacao: string, conteudoNcoletas: string): Promise<boolean> => {
  const client = getPool();
  const result = await client.query(
    'SELECT conteudo_ncoletas FROM operacao_config WHERE operacao = $1',
    [operacao]
  );
  if (result.rows.length === 0) return false;
  const atual = String(result.rows[0].conteudo_ncoletas || '');
  if (atual === conteudoNcoletas) return false;
  await client.query(
    'UPDATE operacao_config SET conteudo_ncoletas = $1, atualizado_em = NOW() WHERE operacao = $2',
    [conteudoNcoletas, operacao]
  );
  return true;
};

export const getLockStatus = async (operacao: string): Promise<{ lock_envio: string; lock_user: string; lock_timestamp: string } | null> => {
  const client = getPool();
  const result = await client.query(
    'SELECT lock_envio, lock_user, lock_timestamp FROM operacao_config WHERE operacao = $1',
    [operacao]
  );
  return result.rows[0] || null;
};

export const acquireLock = async (operacao: string, userEmail: string, timestamp: string): Promise<void> => {
  const client = getPool();
  await client.query(
    'UPDATE operacao_config SET lock_envio = $1, lock_user = $2, lock_timestamp = $3, atualizado_em = NOW() WHERE operacao = $4',
    [true, userEmail, timestamp, operacao]
  );
};

export const releaseLock = async (operacao: string): Promise<void> => {
  const client = getPool();
  await client.query(
    'UPDATE operacao_config SET lock_envio = $1, lock_user = $2, lock_timestamp = $3, atualizado_em = NOW() WHERE operacao = $4',
    [false, '', null, operacao]
  );
};

export const insertConfig = async (row: Record<string, unknown>): Promise<number> => {
  const client = getPool();
  const result = await client.query(
    `INSERT INTO operacao_config (operacao, email, tolerancia, nome_exibicao, plant_id,
       ultimo_envio_saida, status, envio, copia, ultimo_envio_resumo_saida,
       status_resumo_saida, ultimo_envio_ncoleta, quantidade_ncoletas_registrada,
       conteudo, conteudo_ncoletas,
       lock_envio, lock_user, lock_timestamp)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18)
     ON CONFLICT (operacao) DO UPDATE SET
       email = EXCLUDED.email,
       tolerancia = EXCLUDED.tolerancia,
       nome_exibicao = EXCLUDED.nome_exibicao,
       plant_id = EXCLUDED.plant_id,
       ultimo_envio_saida = EXCLUDED.ultimo_envio_saida,
       status = EXCLUDED.status,
       envio = EXCLUDED.envio,
       copia = EXCLUDED.copia,
       ultimo_envio_resumo_saida = EXCLUDED.ultimo_envio_resumo_saida,
       status_resumo_saida = EXCLUDED.status_resumo_saida,
       ultimo_envio_ncoleta = EXCLUDED.ultimo_envio_ncoleta,
       quantidade_ncoletas_registrada = EXCLUDED.quantidade_ncoletas_registrada,
       conteudo = EXCLUDED.conteudo,
       conteudo_ncoletas = EXCLUDED.conteudo_ncoletas,
       lock_envio = EXCLUDED.lock_envio,
       lock_user = EXCLUDED.lock_user,
       lock_timestamp = EXCLUDED.lock_timestamp
     RETURNING id`,
    [
      String(row.operacao || ''),
      String(row.email || ''),
      String(row.tolerancia || '00:00:00'),
      String(row.nome_exibicao || ''),
      row.plant_id != null ? Number(row.plant_id) : null,
      row.ultimo_envio_saida || null,
      String(row.status || ''),
      String(row.envio || ''),
      String(row.copia || ''),
      row.ultimo_envio_resumo_saida || null,
      String(row.status_resumo_saida || ''),
      row.ultimo_envio_ncoleta || null,
      Number(row.quantidade_ncoletas_registrada) || 0,
      String(row.conteudo || ''),
      String(row.conteudo_ncoletas || ''),
      row.lock_envio || null,
      String(row.lock_user || ''),
      row.lock_timestamp || null
    ]
  );
  return result.rows[0]?.id || 0;
};

// ---------------------------------------------------------------------------
// departures
// ---------------------------------------------------------------------------

export type DepartureRow = {
  id: number;
  operacao: string;
  rota: string;
  motorista: string;
  placa_veiculo: string;
  celular_motorista: string;
  hora_prevista: string;
  hora_saida: string;
  status_saida: string;
  motivo_atraso: string;
  observacao: string;
  data_operacao: string;
  celula: string;
  status_rota: string;
  tipo_veiculo: string;
  km: string;
  conferente: string;
  tipo_servico: string;
  regional: string;
  base: string;
  turno: string;
  rota_origem: string;
  total_pacotes: number | null;
  checklist_motorista: string;
  retorno_motorista: string;
  causa_raiz: string;
  tempo_resposta: string;
  log_tempo_resposta: string;
  criado_em: string;
  atualizado_em: string;
};

export const getDepartures = async (): Promise<DepartureRow[]> => {
  const client = getPool();
  const result = await client.query('SELECT * FROM departures ORDER BY id');
  return result.rows as DepartureRow[];
};

export const upsertDeparture = async (d: Record<string, unknown>): Promise<number> => {
  const client = getPool();
  const id = d.id ? Number(d.id) : null;

  // Mapeamento do schema do frontend (camelCase) para colunas do banco
  const operacao = String(d.operacao || '');
  const rota = String(d.rota || '');
  const motorista = String(d.motorista || '');
  const placa_veiculo = String(d.placa || d.placa_veiculo || '');
  const celular_motorista = String(d.contato || d.celular_motorista || '');
  const hora_prevista_raw = String(d.inicio || d.hora_prevista || '').trim();
  const hora_saida_raw = String(d.saida || d.hora_saida || '').trim();
  // Converte "DD/MM/AAAA HH:MM:SS" → "YYYY-MM-DD HH:MM:SS" para TIMESTAMP do PostgreSQL
  const toTimestamp = (raw: string): string | null => {
    if (!raw) return null;
    const trimmed = raw.replace(/\s+/g, ' ').trim();
    // DD/MM/AAAA HH:MM:SS → YYYY-MM-DD HH:MM:SS
    const m = trimmed.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}:\d{2}:\d{2})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]} ${m[4]}`;
    // DD/MM/AAAA → YYYY-MM-DD 00:00:00
    const dm = trimmed.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (dm) return `${dm[3]}-${dm[2]}-${dm[1]} 00:00:00`;
    // Já ISO (YYYY-MM-DD HH:MM:SS)
    if (/^\d{4}-\d{2}-\d{2}(\s+\d{2}:\d{2}:\d{2})?$/.test(trimmed)) return trimmed;
    // Valor simples (HH:MM:SS) retorna como está
    return trimmed || null;
  };
  const hora_prevista = toTimestamp(hora_prevista_raw);
  const hora_saida = toTimestamp(hora_saida_raw);
  const status_saida = String(d.statusGeral || d.status_saida || '');
  const motivo_atraso = String(d.motivo || d.motivo_atraso || '');
  const observacao = String(d.observacao || '');
  const data_operacao_raw = String(d.data || d.data_operacao || '').trim();
  // Converte DD/MM/AAAA ou DD/MM/AAAA HH:MM:SS -> formato ISO para o PostgreSQL
  let data_operacao: string | null = null;
  if (data_operacao_raw) {
    const dm = data_operacao_raw.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (dm) {
      data_operacao = `${dm[3]}-${dm[2]}-${dm[1]}`;
    } else if (/^\d{4}-\d{2}-\d{2}/.test(data_operacao_raw)) {
      data_operacao = data_operacao_raw.slice(0, 10);
    } else {
      data_operacao = data_operacao_raw;
    }
  }
  const status_rota = String(d.statusOp || d.status_rota || 'Previsto');
  const checklist_motorista = String(d.checklistMotorista || d.checklist_motorista || '');
  const retorno_motorista = String(d.retornoMotorista || d.retorno_motorista || '');
  const causa_raiz = String(d.causaRaiz || d.causa_raiz || '');
  const tempo_resposta = String(d.tempoResposta || d.tempo_resposta || '');
  const log_tempo_resposta = String(d.logTempoResposta || d.log_tempo_resposta || '');

  if (id && Number.isFinite(id) && id > 0) {
    const updResult = await client.query(
      `UPDATE departures SET operacao=$1, rota=$2, motorista=$3, placa_veiculo=$4,
       celular_motorista=$5, hora_prevista=$6, hora_saida=$7, status_saida=$8,
       motivo_atraso=$9, observacao=$10, data_operacao=$11, status_rota=$12,
       checklist_motorista=$13, retorno_motorista=$14, causa_raiz=$15,
       tempo_resposta=$16, log_tempo_resposta=$17, atualizado_em=NOW()
       WHERE id = $18`,
      [operacao, rota, motorista, placa_veiculo, celular_motorista, hora_prevista,
       hora_saida, status_saida, motivo_atraso, observacao, data_operacao, status_rota,
       checklist_motorista, retorno_motorista, causa_raiz, tempo_resposta,
       log_tempo_resposta, id]
    );
    if ((updResult.rowCount ?? 0) > 0) {
      console.log(`[upsertDeparture] UPDATE OK id=${id}, rowCount=${updResult.rowCount}, hora_saida=${hora_saida}`);
      return id;
    }
    // ID não existe no PostgreSQL (provavelmente ID do SharePoint) — faz INSERT
    console.log(`[upsertDeparture] UPDATE matched 0 rows for id=${id}, falling back to INSERT`);
  }

  const result = await client.query(
    `INSERT INTO departures (operacao, rota, motorista, placa_veiculo, celular_motorista,
     hora_prevista, hora_saida, status_saida, motivo_atraso, observacao, data_operacao,
     status_rota, checklist_motorista, retorno_motorista, causa_raiz, tempo_resposta,
     log_tempo_resposta)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17)
     RETURNING id`,
    [operacao, rota, motorista, placa_veiculo, celular_motorista, hora_prevista,
     hora_saida, status_saida, motivo_atraso, observacao, data_operacao, status_rota,
     checklist_motorista, retorno_motorista, causa_raiz, tempo_resposta,
     log_tempo_resposta]
  );
  return result.rows[0].id;
};

export const deleteDeparture = async (id: number): Promise<void> => {
  const client = getPool();
  await client.query('DELETE FROM departures WHERE id = $1', [id]);
};

// ---------------------------------------------------------------------------
// non_collections
// ---------------------------------------------------------------------------

export type NonCollectionRow = {
  id: number;
  operacao: string;
  data_operacao: string;
  rota: string;
  status: string;
  quantidade: number | null;
  observacao: string;
  celula: string;
  regional: string;
  base: string;
  turno: string;
  tipo_ocorrencia: string;
  responsavel: string;
  previsto: string;
  semana: string;
  data: string;
  codigo: string;
  produtor: string;
  motivo: string;
  acao: string;
  data_acao: string;
  ultima_coleta: string;
  culpabilidade: string;
  causa_raiz: string;
  criado_em: string;
  atualizado_em: string;
};

export const getNonCollections = async (): Promise<NonCollectionRow[]> => {
  const client = getPool();
  const result = await client.query('SELECT * FROM non_collections ORDER BY id');
  return result.rows as NonCollectionRow[];
};

export const insertNonCollection = async (nc: Record<string, unknown>): Promise<number> => {
  const client = getPool();
  // Converte data DD/MM/AAAA -> AAAA-MM-DD para data_operacao (DATE NOT NULL)
  const dataRaw = String(nc.data || '').trim();
  const dm = dataRaw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  const dataOperacao = dm ? `${dm[3]}-${dm[2]}-${dm[1]}` : (dataRaw || '1970-01-01');
  const result = await client.query(
    `INSERT INTO non_collections (operacao, data_operacao, rota, observacao, semana, data, codigo,
     produtor, motivo, acao, data_acao, ultima_coleta, culpabilidade, causa_raiz)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)
     RETURNING id`,
    [
      String(nc.operacao || ''),
      dataOperacao,
      String(nc.rota || ''),
      String(nc.observacao || ''),
      String(nc.semana || ''),
      dataRaw,
      String(nc.codigo || ''),
      String(nc.produtor || ''),
      String(nc.motivo || ''),
      String(nc.acao || ''),
      String(nc.dataAcao || nc.data_acao || ''),
      String(nc.ultimaColeta || nc.ultima_coleta || ''),
      String(nc.Culpabilidade || nc.culpabilidade || ''),
      String(nc.causaRaiz || nc.causa_raiz || '')
    ]
  );
  return result.rows[0].id;
};

export const updateNonCollection = async (nc: Record<string, unknown>): Promise<void> => {
  const client = getPool();
  const id = Number(nc.id);
  // Converte data DD/MM/AAAA -> AAAA-MM-DD para data_operacao (DATE NOT NULL)
  const dataRaw = String(nc.data || '').trim();
  const dm = dataRaw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  const dataOperacao = dm ? `${dm[3]}-${dm[2]}-${dm[1]}` : (dataRaw || '1970-01-01');
  await client.query(
    `UPDATE non_collections SET operacao=$1, data_operacao=$2, rota=$3, observacao=$4, semana=$5,
     data=$6, codigo=$7, produtor=$8, motivo=$9, acao=$10, data_acao=$11,
     ultima_coleta=$12, culpabilidade=$13, causa_raiz=$14, atualizado_em=NOW()
     WHERE id = $15`,
    [
      String(nc.operacao || ''),
      dataOperacao,
      String(nc.rota || ''),
      String(nc.observacao || ''),
      String(nc.semana || ''),
      dataRaw,
      String(nc.codigo || ''),
      String(nc.produtor || ''),
      String(nc.motivo || ''),
      String(nc.acao || ''),
      String(nc.dataAcao || nc.data_acao || ''),
      String(nc.ultimaColeta || nc.ultima_coleta || ''),
      String(nc.Culpabilidade || nc.culpabilidade || ''),
      String(nc.causaRaiz || nc.causa_raiz || ''),
      id
    ]
  );
};

export const deleteNonCollection = async (id: number): Promise<void> => {
  const client = getPool();
  await client.query('DELETE FROM non_collections WHERE id = $1', [id]);
};

/**
 * Corrige a coluna rota do PostgreSQL cruzando operacao + codigo com os dados do SharePoint.
 * spItems = array de { operacao, codigo, rota } vindos do SharePoint.
 * Só atualiza se encontrar match de operacao+codigo E se o SharePoint tiver rota preenchida.
 */
export const fixNonCollectionsRoutes = async (
  spItems: { operacao: string; codigo: string; rota: string }[]
): Promise<{ updated: number; skipped: number; details: string[] }> => {
  const client = getPool();
  let updated = 0;
  let skipped = 0;
  const details: string[] = [];

  for (const item of spItems) {
    const rota = item.rota.trim();
    if (!rota) { skipped++; continue; }

    const operacao = item.operacao.trim();
    const codigo = item.codigo.trim();
    if (!operacao || !codigo) { skipped++; continue; }

    const result = await client.query(
      `UPDATE non_collections SET rota = $1, atualizado_em = NOW()
       WHERE operacao = $2 AND codigo = $3 AND (rota IS NULL OR rota = '' OR rota != $1)
       RETURNING id`,
      [rota, operacao, codigo]
    );

    if (result.rowCount && result.rowCount > 0) {
      updated += result.rowCount;
      details.push(`operacao="${operacao}" codigo="${codigo}" -> rota="${rota}" (${result.rowCount} rows)`);
    } else {
      skipped++;
    }
  }

  return { updated, skipped, details };
};

// ---------------------------------------------------------------------------
// Pool cleanup
// ---------------------------------------------------------------------------

export const closeChecklistPool = async (): Promise<void> => {
  if (pool) {
    await pool.end();
    pool = null;
  }
};
