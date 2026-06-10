import pg from 'pg';

const { Pool } = pg;

let pool: pg.Pool | null = null;

const getPool = (): pg.Pool => {
  if (pool) return pool;

  const dbUrl = String(process.env.RWE_DB_URL || '').trim();
  if (!dbUrl) {
    throw new Error('RWE_DB_URL não configurada');
  }

  const ssl = String(process.env.RWE_DB_SSL || 'true').trim().toLowerCase();
  const schema = String(process.env.RWE_DB_SCHEMA || 'public').trim();

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
    console.error('[RWE_DB] Pool error (idle connection):', err.message);
  });

  return pool;
};

export type RouteWebEventRow = {
  route_id: number;
  event_id: number | null;
  plant_id: number;
  filial: string;
  operacao: string;
  rota_codigo: string;
  motorista: string;
  placa: string;
  type_name: string;
  reference: string;
  reference_code: string;
  status: string;
  executed: boolean | null;
  expected_arrival: string | null;
  actual_arrival: string | null;
  expected_departure: string | null;
  actual_departure: string | null;
  motivo: string;
  status_type: string;
  occurrence_id: number | null;
  occurrence_type_id: number | null;
  occurrence_type_description: string;
  occurrence_inserted_by: string;
  occurrence_inserted_at: string | null;
  event_created_at: string | null;
  event_updated_at: string | null;
  data_referencia: string;
};

const toNullIfEmpty = (value: unknown): string | null => {
  const str = String(value ?? '').trim();
  return str || null;
};

const toBooleanOrNull = (value: unknown): boolean | null => {
  if (value == null) return null;
  if (typeof value === 'boolean') return value;
  const raw = String(value).trim().toLowerCase();
  if (raw === 'true' || raw === '1') return true;
  if (raw === 'false' || raw === '0') return false;
  return null;
};

export const upsertRouteWebEvents = async (rows: RouteWebEventRow[]): Promise<number> => {
  if (rows.length === 0) return 0;

  const client = getPool();
  let inserted = 0;

  const sql = `
    INSERT INTO route_web_events (
      route_id, event_id, plant_id, filial, operacao, rota_codigo, motorista, placa,
      type_name, reference, reference_code, status, executed,
      expected_arrival, actual_arrival, expected_departure, actual_departure,
      motivo, status_type,
      occurrence_id, occurrence_type_id, occurrence_type_description,
      occurrence_inserted_by, occurrence_inserted_at,
      event_created_at, event_updated_at, data_referencia
    ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27)
    ON CONFLICT (route_id, event_id, occurrence_id, data_referencia)
    DO UPDATE SET
      filial = EXCLUDED.filial,
      operacao = EXCLUDED.operacao,
      rota_codigo = EXCLUDED.rota_codigo,
      motorista = EXCLUDED.motorista,
      placa = EXCLUDED.placa,
      type_name = EXCLUDED.type_name,
      reference = EXCLUDED.reference,
      reference_code = EXCLUDED.reference_code,
      status = EXCLUDED.status,
      executed = EXCLUDED.executed,
      expected_arrival = EXCLUDED.expected_arrival,
      actual_arrival = EXCLUDED.actual_arrival,
      expected_departure = EXCLUDED.expected_departure,
      actual_departure = EXCLUDED.actual_departure,
      motivo = EXCLUDED.motivo,
      status_type = EXCLUDED.status_type,
      occurrence_type_id = EXCLUDED.occurrence_type_id,
      occurrence_type_description = EXCLUDED.occurrence_type_description,
      occurrence_inserted_by = EXCLUDED.occurrence_inserted_by,
      occurrence_inserted_at = EXCLUDED.occurrence_inserted_at,
      event_created_at = EXCLUDED.event_created_at,
      event_updated_at = EXCLUDED.event_updated_at,
      fetched_at = NOW()
  `;

  for (const row of rows) {
    const values = [
      row.route_id,
      row.event_id,
      row.plant_id,
      row.filial,
      row.operacao,
      row.rota_codigo,
      row.motorista,
      row.placa,
      row.type_name,
      row.reference,
      row.reference_code,
      row.status,
      row.executed,
      toNullIfEmpty(row.expected_arrival),
      toNullIfEmpty(row.actual_arrival),
      toNullIfEmpty(row.expected_departure),
      toNullIfEmpty(row.actual_departure),
      row.motivo,
      row.status_type,
      row.occurrence_id,
      row.occurrence_type_id,
      row.occurrence_type_description,
      row.occurrence_inserted_by,
      toNullIfEmpty(row.occurrence_inserted_at),
      toNullIfEmpty(row.event_created_at),
      toNullIfEmpty(row.event_updated_at),
      row.data_referencia
    ];

    await client.query(sql, values);
    inserted += 1;
  }

  return inserted;
};

export const deleteRouteWebEventsByDate = async (dataReferencia: string): Promise<number> => {
  const client = getPool();
  const result = await client.query(
    'DELETE FROM route_web_events WHERE data_referencia = $1',
    [dataReferencia]
  );
  return result.rowCount ?? 0;
};

export type RouteWebEventDbRow = {
  id: number;
  route_id: number;
  event_id: number | null;
  plant_id: number;
  filial: string;
  operacao: string;
  rota_codigo: string;
  motorista: string;
  placa: string;
  type_name: string;
  reference: string;
  reference_code: string;
  status: string;
  executed: boolean | null;
  expected_arrival: string | null;
  actual_arrival: string | null;
  expected_departure: string | null;
  actual_departure: string | null;
  motivo: string;
  status_type: string;
  is_already_launched: boolean;
  occurrence_id: number | null;
  occurrence_type_id: number | null;
  occurrence_type_description: string;
  occurrence_inserted_by: string;
  occurrence_inserted_at: string | null;
  event_created_at: string | null;
  event_updated_at: string | null;
  fetched_at: string;
  data_referencia: string;
};

export const getRouteWebEventsByDateAndPlants = async (
  dataReferencia: string,
  plantIds: number[]
): Promise<RouteWebEventDbRow[]> => {
  const client = getPool();

  const dateFilter = `data_referencia = $1::date`;

  if (plantIds.length === 0) {
    const result = await client.query(
      `SELECT * FROM route_web_events WHERE ${dateFilter} ORDER BY route_id, event_id, occurrence_id`,
      [dataReferencia]
    );
    return result.rows as RouteWebEventDbRow[];
  }

  const intPlantIds = plantIds.map((id) => Number(id)).filter(Number.isFinite);
  const placeholders = intPlantIds.map((_, i) => `$${i + 2}::int`).join(',');
  const result = await client.query(
    `SELECT * FROM route_web_events WHERE ${dateFilter} AND plant_id = ANY(ARRAY[${placeholders}]) ORDER BY route_id, event_id, occurrence_id`,
    [dataReferencia, ...intPlantIds]
  );
  return result.rows as RouteWebEventDbRow[];
};

export const markEventsAsLaunched = async (
  dataReferencia: string,
  eventIds: number[]
): Promise<number> => {
  if (eventIds.length === 0) return 0;
  const client = getPool();
  const intEventIds = eventIds.map((id) => Number(id)).filter(Number.isFinite);
  const placeholders = intEventIds.map((_, i) => `$${i + 2}::int`).join(',');
  const result = await client.query(
    `UPDATE route_web_events SET is_already_launched = true WHERE data_referencia = $1 AND event_id = ANY(ARRAY[${placeholders}])`,
    [dataReferencia, ...intEventIds]
  );
  return result.rowCount ?? 0;
};

export const closeRwePool = async (): Promise<void> => {
  if (pool) {
    await pool.end();
    pool = null;
  }
};

// ---------------------------------------------------------------------------
// Route Web Routes (dados da rota para integração de saídas)
// ---------------------------------------------------------------------------

export type RouteWebRouteRow = {
  route_id: number;
  route_plan_id: string | null;
  schedule_order_id: number | null;
  plant_id: number;
  datalake_plant_id: number | null;
  filial: string;
  operacao: string;
  roadmap_code: string;
  status: string;
  specific_status: string;
  general_status: string;
  placa: string;
  motorista: string;
  last_driver_id: number | null;
  last_vehicle_id: number | null;
  smartquestion_actual_start_time: string | null;
  smartquestion_actual_end_time: string | null;
  smartquestion_collected_liters: number | null;
  smartquestion_unloading_plate: string;
  start_time: string | null;
  actual_start_time: string | null;
  expected_end_time: string | null;
  actual_end_time: string | null;
  expected_liters: number | null;
  collected_liters: number | null;
  unloaded_liters: number | null;
  volume: number | null;
  expected_km: number | null;
  actual_km: number | null;
  last_landmark: string;
  data_referencia: string;
};

const toNumericOrNull = (value: unknown): number | null => {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : null;
};

export const upsertRouteWebRoutes = async (rows: RouteWebRouteRow[]): Promise<number> => {
  if (rows.length === 0) return 0;
  const client = getPool();

  const sql = `
    INSERT INTO route_web_routes (
      route_id, route_plan_id, schedule_order_id, plant_id, datalake_plant_id,
      filial, operacao, roadmap_code, status, specific_status, general_status,
      placa, motorista, last_driver_id, last_vehicle_id,
      smartquestion_actual_start_time, smartquestion_actual_end_time,
      smartquestion_collected_liters, smartquestion_unloading_plate,
      start_time, actual_start_time, expected_end_time, actual_end_time,
      expected_liters, collected_liters, unloaded_liters, volume,
      expected_km, actual_km, last_landmark, data_referencia
    ) VALUES (
      $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,
      $11,$12,$13,$14,$15,$16,$17,$18,$19,$20,
      $21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31
    )
    ON CONFLICT (route_id, data_referencia)
    DO UPDATE SET
      route_plan_id = EXCLUDED.route_plan_id,
      schedule_order_id = EXCLUDED.schedule_order_id,
      plant_id = EXCLUDED.plant_id,
      datalake_plant_id = EXCLUDED.datalake_plant_id,
      filial = EXCLUDED.filial,
      operacao = EXCLUDED.operacao,
      roadmap_code = EXCLUDED.roadmap_code,
      status = EXCLUDED.status,
      specific_status = EXCLUDED.specific_status,
      general_status = EXCLUDED.general_status,
      placa = EXCLUDED.placa,
      motorista = EXCLUDED.motorista,
      last_driver_id = EXCLUDED.last_driver_id,
      last_vehicle_id = EXCLUDED.last_vehicle_id,
      smartquestion_actual_start_time = EXCLUDED.smartquestion_actual_start_time,
      smartquestion_actual_end_time = EXCLUDED.smartquestion_actual_end_time,
      smartquestion_collected_liters = EXCLUDED.smartquestion_collected_liters,
      smartquestion_unloading_plate = EXCLUDED.smartquestion_unloading_plate,
      start_time = EXCLUDED.start_time,
      actual_start_time = EXCLUDED.actual_start_time,
      expected_end_time = EXCLUDED.expected_end_time,
      actual_end_time = EXCLUDED.actual_end_time,
      expected_liters = EXCLUDED.expected_liters,
      collected_liters = EXCLUDED.collected_liters,
      unloaded_liters = EXCLUDED.unloaded_liters,
      volume = EXCLUDED.volume,
      expected_km = EXCLUDED.expected_km,
      actual_km = EXCLUDED.actual_km,
      last_landmark = EXCLUDED.last_landmark,
      fetched_at = NOW()
  `;

  let inserted = 0;
  for (const row of rows) {
    const values = [
      row.route_id,
      toNullIfEmpty(row.route_plan_id),
      row.schedule_order_id,
      row.plant_id,
      row.datalake_plant_id,
      row.filial,
      row.operacao,
      row.roadmap_code,
      row.status,
      row.specific_status,
      row.general_status,
      row.placa,
      row.motorista,
      row.last_driver_id,
      row.last_vehicle_id,
      toNullIfEmpty(row.smartquestion_actual_start_time),
      toNullIfEmpty(row.smartquestion_actual_end_time),
      toNumericOrNull(row.smartquestion_collected_liters),
      row.smartquestion_unloading_plate,
      toNullIfEmpty(row.start_time),
      toNullIfEmpty(row.actual_start_time),
      toNullIfEmpty(row.expected_end_time),
      toNullIfEmpty(row.actual_end_time),
      toNumericOrNull(row.expected_liters),
      toNumericOrNull(row.collected_liters),
      toNumericOrNull(row.unloaded_liters),
      toNumericOrNull(row.volume),
      toNumericOrNull(row.expected_km),
      toNumericOrNull(row.actual_km),
      row.last_landmark,
      row.data_referencia
    ];
    await client.query(sql, values);
    inserted += 1;
  }
  return inserted;
};

export type RouteWebRouteDbRow = {
  id: number;
  route_id: number;
  route_plan_id: string | null;
  schedule_order_id: number | null;
  plant_id: number;
  datalake_plant_id: number | null;
  filial: string;
  operacao: string;
  roadmap_code: string;
  status: string;
  specific_status: string;
  general_status: string;
  placa: string;
  motorista: string;
  last_driver_id: number | null;
  last_vehicle_id: number | null;
  smartquestion_actual_start_time: string | null;
  smartquestion_actual_end_time: string | null;
  smartquestion_collected_liters: number | null;
  smartquestion_unloading_plate: string;
  start_time: string | null;
  actual_start_time: string | null;
  expected_end_time: string | null;
  actual_end_time: string | null;
  expected_liters: number | null;
  collected_liters: number | null;
  unloaded_liters: number | null;
  volume: number | null;
  expected_km: number | null;
  actual_km: number | null;
  last_landmark: string;
  data_referencia: string;
  fetched_at: string;
};

export const getRouteWebRoutesByDateAndPlants = async (
  dataReferencia: string,
  plantIds: number[]
): Promise<RouteWebRouteDbRow[]> => {
  const client = getPool();

  if (plantIds.length === 0) {
    const result = await client.query(
      `SELECT * FROM route_web_routes WHERE data_referencia = $1::date ORDER BY roadmap_code`,
      [dataReferencia]
    );
    return result.rows as RouteWebRouteDbRow[];
  }

  const intPlantIds = plantIds.map((id) => Number(id)).filter(Number.isFinite);
  const placeholders = intPlantIds.map((_, i) => `$${i + 2}::int`).join(',');
  const result = await client.query(
    `SELECT * FROM route_web_routes WHERE data_referencia = $1::date AND plant_id = ANY(ARRAY[${placeholders}]) ORDER BY roadmap_code`,
    [dataReferencia, ...intPlantIds]
  );
  return result.rows as RouteWebRouteDbRow[];
};
