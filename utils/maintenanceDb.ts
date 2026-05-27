import pg from 'pg';

const { Pool } = pg;

type PoolHolder = typeof globalThis & {
  __maintenanceDbPool?: pg.Pool;
};

const normalizePlate = (value: string): string =>
  String(value || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();

const normalizeDate = (value: string): string => {
  const raw = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return '';
  return parsed.toISOString().slice(0, 10);
};

const sanitizeIdentifier = (value: string, fallback: string): string => {
  const clean = String(value || fallback).trim();
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(clean)) return clean;
  return fallback;
};

const getBooleanEnv = (value: string | undefined, defaultValue: boolean): boolean => {
  if (value == null) return defaultValue;
  const normalized = value.trim().toLowerCase();
  if (normalized === 'true' || normalized === '1' || normalized === 'yes') return true;
  if (normalized === 'false' || normalized === '0' || normalized === 'no') return false;
  return defaultValue;
};

const getPool = (): pg.Pool => {
  const holder = globalThis as PoolHolder;
  if (holder.__maintenanceDbPool) return holder.__maintenanceDbPool;

  const connectionString = process.env.MAINT_DB_URL || process.env.DATABASE_URL;
  if (!connectionString) {
    throw new Error('MAINT_DB_URL/DATABASE_URL não configurada');
  }

  const sslEnabled = getBooleanEnv(process.env.MAINT_DB_SSL, true);

  holder.__maintenanceDbPool = new Pool({
    connectionString,
    ssl: sslEnabled ? { rejectUnauthorized: false } : undefined,
    max: 3,
    idleTimeoutMillis: 15000,
    connectionTimeoutMillis: 10000
  });

  return holder.__maintenanceDbPool;
};

export type MaintenanceLookupItem = {
  placa?: string;
  data?: string;
};

export type MaintenanceEventRow = {
  placa: string;
  tipo: string;
  area: string;
  status: string;
  titulo: string;
  categoria: string;
  data_planejada: string;
};

export const queryMaintenanceEvents = async (rawItems: MaintenanceLookupItem[]): Promise<MaintenanceEventRow[]> => {
  const normalizedItems = rawItems
    .map((item) => ({
      placa: normalizePlate(item?.placa || ''),
      data: normalizeDate(item?.data || '')
    }))
    .filter((item) => item.placa && item.data);

  if (normalizedItems.length === 0) return [];

  const uniquePlates = Array.from(new Set(normalizedItems.map((item) => item.placa)));
  const uniqueDates = Array.from(new Set(normalizedItems.map((item) => item.data)));

  if (uniquePlates.length === 0 || uniqueDates.length === 0) return [];

  const schema = sanitizeIdentifier(process.env.MAINT_DB_SCHEMA || '', 'public');
  const table = sanitizeIdentifier(process.env.MAINT_DB_TABLE || '', 'manutencoes');

  const query = `
    SELECT
      placa,
      tipo,
      area,
      status,
      titulo,
      categoria,
      data_planejada::date AS data_planejada
    FROM "${schema}"."${table}"
    WHERE data_planejada::date = ANY($1::date[])
      AND UPPER(REGEXP_REPLACE(COALESCE(placa, ''), '[^A-Za-z0-9]', '', 'g')) = ANY($2::text[])
    ORDER BY data_planejada ASC
  `;

  try {
    const pool = getPool();
    const result = await pool.query(query, [uniqueDates, uniquePlates]);
    return (result.rows || []).map((row: any) => ({
      placa: normalizePlate(String(row.placa || '')),
      tipo: String(row.tipo || ''),
      area: String(row.area || ''),
      status: String(row.status || ''),
      titulo: String(row.titulo || ''),
      categoria: String(row.categoria || ''),
      data_planejada: normalizeDate(String(row.data_planejada || ''))
    }));
  } catch (error: any) {
    console.warn('[MAINT_DB] Consulta falhou (banco pode estar inacessível):', error?.message || error);
    return [];
  }
};
