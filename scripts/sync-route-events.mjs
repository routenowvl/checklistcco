#!/usr/bin/env node
/**
 * Standalone Route Web Events sync script.
 * Runs on EC2 via cron — no Vercel timeout constraints.
 *
 * Usage:
 *   node sync-route-events.mjs
 *
 * Reads env vars from .env file in CWD or parent dirs, or from environment.
 */

import pg from 'pg';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ---------------------------------------------------------------------------
// Env reader
// ---------------------------------------------------------------------------

const readEnvFromDotEnvFiles = (name) => {
  const candidates = [path.join(__dirname, '..', '.env.local'), path.join(__dirname, '..', '.env')];
  for (const file of candidates) {
    if (!fs.existsSync(file)) continue;
    const content = fs.readFileSync(file, 'utf-8');
    for (const rawLine of content.split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith('#')) continue;
      const eqIndex = line.indexOf('=');
      if (eqIndex <= 0) continue;
      const key = line.slice(0, eqIndex).trim().replace(/^export\s+/i, '');
      if (key !== name) continue;
      return line.slice(eqIndex + 1).trim().replace(/^['"]|['"]$/g, '').trim();
    }
  }
  return '';
};

const readEnv = (name) => {
  const fromProcess = String(process.env[name] || '').trim();
  const fromDotEnv = readEnvFromDotEnvFiles(name);
  return fromDotEnv || fromProcess;
};

const readRequiredEnv = (name) => {
  const value = readEnv(name);
  if (!value) {
    console.error(`[FATAL] ${name} não configurada`);
    process.exit(1);
  }
  return value;
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const toOptionalInt = (value) => {
  if (value == null) return null;
  if (typeof value === 'number' && Number.isFinite(value)) return Math.trunc(value);
  const raw = String(value).trim();
  if (!raw) return null;
  const match = raw.match(/-?\d+(?:[.,]\d+)?/);
  if (!match) return null;
  const parsed = Number(match[0].replace(',', '.'));
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
};

const normalizeText = (value) =>
  String(value ?? '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

const getOccurrenceDescription = (occ) =>
  String(occ?.occurrence_type?.description || occ?.occurrence_type_description || occ?.description || occ?.type_name || occ?.name || occ?.title || '').trim();

const NON_COLLECTION_TECHNICAL_OCCURRENCE_IDS = new Set([2]);
const NON_COLLECTION_EXCLUDED_REASON_PATTERNS = [
  'troca de caminhao', 'troca de caminhão', 'evento extra', 'evento_extra',
  'tanque comunitario', 'tanque_comunitario', 'alteracao de horario',
  'troca de reboque', 'falta de sinal do rastreador'
];

const isTechnicalOccurrence = (occ) => {
  const occurrenceTypeId = toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id);
  if (occurrenceTypeId != null && NON_COLLECTION_TECHNICAL_OCCURRENCE_IDS.has(occurrenceTypeId)) return true;
  const desc = normalizeText(getOccurrenceDescription(occ));
  if (!desc) return false;
  return desc.includes('atualizacao de posicao pelo rastreador') || desc.includes('evento fora de ordem');
};

const isNonCollectionOccurrence = (occ) => {
  if (!occ || typeof occ !== 'object') return false;
  const desc = normalizeText(getOccurrenceDescription(occ));
  if (NON_COLLECTION_EXCLUDED_REASON_PATTERNS.some((p) => desc.includes(normalizeText(p)))) return false;
  if (normalizeText(occ?.inserted_by) === 'scrapersmartquestion') return true;
  if (isTechnicalOccurrence(occ)) return false;
  const occurrenceTypeId = toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id);
  return occurrenceTypeId != null || Boolean(desc);
};

const getRouteCode = (route) =>
  String(route?.roadmap_code || route?.route_code || route?.code || route?.route || route?.route_plan_id || route?.id || '-');

const getDriverName = (route) =>
  String(route?.driver_name || route?.last_driver_name || route?.driver?.name || route?.last_driver?.name || '-');

const getPlate = (event, route) => {
  const raw = String(route?.unloading_plate || event?.trailer_plate || event?.plate || route?.trailer_plate || route?.plate || route?.vehicle_plate || route?.last_vehicle?.plate || '-').trim();
  return raw || '-';
};

const getScraperOccurrenceReason = (event) => {
  if (!Array.isArray(event?.occurrences)) return '';
  for (const occ of event.occurrences) {
    if (normalizeText(occ?.inserted_by) !== 'scrapersmartquestion') continue;
    const reason = getOccurrenceDescription(occ);
    if (reason) return reason;
  }
  return '';
};

const pickArray = (payload, ...keys) => {
  if (Array.isArray(payload)) return payload;
  for (const key of keys) {
    const val = payload?.[key];
    if (Array.isArray(val)) return val;
  }
  return [];
};

const pickRoutesArray = (payload) => pickArray(payload, 'data', 'routes', 'items', 'results');
const pickEventsArray = (payload) => pickArray(payload, 'data', 'events', 'items', 'results');

const getCurrentDayDate = () => {
  const now = new Date();
  // Usa timezone de Brasília para consistência com o frontend (getBrazilDate)
  const parts = now.toLocaleDateString('sv-SE', { timeZone: 'America/Sao_Paulo' }).split('-');
  return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
};

const getPreviousDayDate = () => {
  const today = getCurrentDayDate();
  const d = new Date(today + 'T12:00:00Z');
  d.setDate(d.getDate() - 1);
  const parts = d.toISOString().split('T')[0].split('-');
  return `${parts[0]}-${parts[1]}-${parts[2]}`;
};

const getD2Date = () => {
  const today = getCurrentDayDate();
  const d = new Date(today + 'T12:00:00Z');
  d.setDate(d.getDate() - 2);
  const parts = d.toISOString().split('T')[0].split('-');
  return `${parts[0]}-${parts[1]}-${parts[2]}`;
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const normalizeString = (str) =>
  str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '').trim();

// ---------------------------------------------------------------------------
// Route Web Token
// ---------------------------------------------------------------------------

const requestRouteWebToken = async () => {
  const baseUrl = readRequiredEnv('ROUTE_WEB_URL').replace(/\/+$/, '');
  const url = `${baseUrl}/api/oauth/token`;
  const clientId = readRequiredEnv('ROUTE_WEB_CLIENT_ID');
  const clientSecret = readRequiredEnv('ROUTE_WEB_CLIENT_SECRET');
  const username = readRequiredEnv('ROUTE_WEB_USERNAME');
  const password = readRequiredEnv('ROUTE_WEB_PASSWORD');
  const scope = readRequiredEnv('ROUTE_WEB_SCOPE');

  // Try JSON first, then form
  const attempts = [
    { headers: { 'Content-Type': 'application/json', Accept: 'application/json' }, body: JSON.stringify({ client_id: clientId, client_secret: clientSecret, username, password, scope, grant_type: 'password' }) },
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' }, body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, username, password, scope, grant_type: 'password' }).toString() }
  ];

  for (const attempt of attempts) {
    const response = await fetch(url, { method: 'POST', headers: attempt.headers, body: attempt.body });
    const data = await response.json().catch(() => null);
    if (!response.ok) continue;

    const tokenFields = ['access_token', 'token', 'TOKEN', 'bearer', 'Bearer', 'accessToken'];
    for (const field of tokenFields) {
      const val = data?.[field] || data?.data?.[field] || data?.result?.[field];
      if (typeof val === 'string' && val.trim()) return val.trim();
    }
  }

  throw new Error('Falha ao obter token Route Web');
};

// ---------------------------------------------------------------------------
// Azure AD Graph Token (Client Credentials)
// ---------------------------------------------------------------------------

let cachedGraphToken = null;

const getGraphAppToken = async () => {
  if (cachedGraphToken && Date.now() < cachedGraphToken.expiresAt) return cachedGraphToken.token;

  const clientId = readRequiredEnv('VITE_AZURE_CLIENT_ID');
  const clientSecret = readRequiredEnv('VITE_AZURE_CLIENT_SECRET');
  const tenantId = readRequiredEnv('VITE_AZURE_TENANT_ID');

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({ client_id: clientId, client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }).toString();

  const response = await fetch(tokenUrl, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph token error ${response.status}: ${text.slice(0, 400)}`);
  }

  const data = await response.json();
  const token = String(data.access_token || '').trim();
  if (!token) throw new Error('Token vazio na resposta do Azure AD');

  const expiresIn = Number(data.expires_in || 3600);
  cachedGraphToken = { token, expiresAt: Date.now() + expiresIn * 1000 - 60_000 };
  console.log(`[GRAPH] Token obtido, expira em ${expiresIn}s`);
  return token;
};

// ---------------------------------------------------------------------------
// SharePoint Plant Configs
// ---------------------------------------------------------------------------

const graphFetch = async (endpoint, token) => {
  const url = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text.slice(0, 400)}`);
  }
  return res.status === 204 ? null : res.json();
};

const resolveFieldName = (mapping, target) => {
  const normalized = normalizeString(target);
  if (normalized === 'titulo' || normalized === 'rota') {
    if (mapping['title']) return 'Title';
  }
  return mapping[normalized] || target;
};

const extractPlantFieldValue = (fields, mapping) => {
  const candidates = ['Plant_id', 'Plant Id', 'PlantId', 'plant_id', 'IdPlant', 'ID_PLANT'].map((c) => resolveFieldName(mapping, c));
  for (const candidate of candidates) {
    if (!candidate) continue;
    const value = fields?.[candidate];
    if (value != null && String(value).trim() !== '') return value;
  }
  for (const [key, value] of Object.entries(fields || {})) {
    const nk = normalizeString(key);
    if (nk.includes('plantid') || nk.includes('idplant')) {
      if (value != null && String(value).trim() !== '') return value;
    }
  }
  return null;
};

const getPlantConfigsFromSharePoint = async () => {
  const sitePath = readRequiredEnv('VITE_SHAREPOINT_SITE_PATH');
  const token = await getGraphAppToken();

  const siteData = await graphFetch(`/sites/${sitePath}`, token);
  const siteId = siteData.id;

  let list;
  try {
    list = await graphFetch(`/sites/${siteId}/lists/CONFIG_OPERACAO_SAIDA_DE_ROTAS`, token);
  } catch {
    const listsData = await graphFetch(`/sites/${siteId}/lists`, token);
    list = (listsData.value || []).find((l) => l.name?.toLowerCase() === 'config_operacao_saida_de_rotas' || l.displayName?.toLowerCase() === 'config_operacao_saida_de_rotas');
    if (!list) throw new Error('Lista CONFIG_OPERACAO_SAIDA_DE_ROTAS não encontrada');
  }

  const columns = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, token);
  const mapping = {};
  for (const col of columns.value || []) {
    mapping[normalizeString(col.name)] = col.name;
    mapping[normalizeString(col.displayName)] = col.name;
  }

  const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);

  const configs = [];
  for (const item of data.value || []) {
    const f = item.fields || {};
    const operacao = String(f[resolveFieldName(mapping, 'OPERACAO')] || '').trim();
    if (!operacao) continue;
    const plantRaw = extractPlantFieldValue(f, mapping);
    const plantId = toOptionalInt(plantRaw);
    if (plantId == null) continue;
    const filial = String(f[resolveFieldName(mapping, 'NomeExibicao')] || operacao).trim();
    configs.push({ plantId, operacao, filial });
  }

  console.log(`[SHAREPOINT] ${configs.length} plant configs:`, configs.map((c) => `${c.filial}(plantId=${c.plantId})`));
  return configs;
};

// ---------------------------------------------------------------------------
// Database
// ---------------------------------------------------------------------------

let pool = null;

const getPool = () => {
  if (pool) return pool;
  const dbUrl = readRequiredEnv('RWE_DB_URL');
  const ssl = String(readEnv('RWE_DB_SSL') || 'true').trim().toLowerCase();
  const schema = String(readEnv('RWE_DB_SCHEMA') || 'public').trim();

  pool = new pg.Pool({
    connectionString: dbUrl,
    ssl: ssl === 'true' || ssl === '1' ? { rejectUnauthorized: false } : undefined,
    max: 4,
    idleTimeoutMillis: 15_000,
    connectionTimeoutMillis: 8_000
  });

  pool.on('connect', (client) => {
    client.query(`SET search_path TO ${schema}`);
  });

  return pool;
};

const toNullIfEmpty = (value) => {
  const str = String(value ?? '').trim();
  return str || null;
};

const upsertRouteWebEvents = async (rows) => {
  if (rows.length === 0) return 0;
  const client = getPool();

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
      filial = EXCLUDED.filial, operacao = EXCLUDED.operacao, rota_codigo = EXCLUDED.rota_codigo,
      motorista = EXCLUDED.motorista, placa = EXCLUDED.placa, type_name = EXCLUDED.type_name,
      reference = EXCLUDED.reference, reference_code = EXCLUDED.reference_code, status = EXCLUDED.status,
      executed = EXCLUDED.executed, expected_arrival = EXCLUDED.expected_arrival,
      actual_arrival = EXCLUDED.actual_arrival, expected_departure = EXCLUDED.expected_departure,
      actual_departure = EXCLUDED.actual_departure, motivo = EXCLUDED.motivo, status_type = EXCLUDED.status_type,
      occurrence_type_id = EXCLUDED.occurrence_type_id, occurrence_type_description = EXCLUDED.occurrence_type_description,
      occurrence_inserted_by = EXCLUDED.occurrence_inserted_by, occurrence_inserted_at = EXCLUDED.occurrence_inserted_at,
      event_created_at = EXCLUDED.event_created_at, event_updated_at = EXCLUDED.event_updated_at,
      fetched_at = NOW()
  `;

  let inserted = 0;
  for (const row of rows) {
    const values = [
      row.route_id, row.event_id, row.plant_id, row.filial, row.operacao, row.rota_codigo,
      row.motorista, row.placa, row.type_name, row.reference, row.reference_code, row.status,
      row.executed, toNullIfEmpty(row.expected_arrival), toNullIfEmpty(row.actual_arrival),
      toNullIfEmpty(row.expected_departure), toNullIfEmpty(row.actual_departure), row.motivo,
      row.status_type, row.occurrence_id, row.occurrence_type_id, row.occurrence_type_description,
      row.occurrence_inserted_by, toNullIfEmpty(row.occurrence_inserted_at),
      toNullIfEmpty(row.event_created_at), toNullIfEmpty(row.event_updated_at), row.data_referencia
    ];
    await client.query(sql, values);
    inserted += 1;
  }
  return inserted;
};

// ---------------------------------------------------------------------------
// Route Web Routes DB
// ---------------------------------------------------------------------------

const upsertRouteWebRoutes = async (rows) => {
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
      row.route_id, toNullIfEmpty(row.route_plan_id), row.schedule_order_id,
      row.plant_id, row.datalake_plant_id,
      row.filial, row.operacao, row.roadmap_code, row.status, row.specific_status,
      row.general_status, row.placa, row.motorista, row.last_driver_id, row.last_vehicle_id,
      toNullIfEmpty(row.smartquestion_actual_start_time), toNullIfEmpty(row.smartquestion_actual_end_time),
      toNumericOrNull(row.smartquestion_collected_liters), row.smartquestion_unloading_plate,
      toNullIfEmpty(row.start_time), toNullIfEmpty(row.actual_start_time),
      toNullIfEmpty(row.expected_end_time), toNullIfEmpty(row.actual_end_time),
      toNumericOrNull(row.expected_liters), toNumericOrNull(row.collected_liters),
      toNumericOrNull(row.unloaded_liters), toNumericOrNull(row.volume),
      toNumericOrNull(row.expected_km), toNumericOrNull(row.actual_km),
      row.last_landmark, row.data_referencia
    ];
    await client.query(sql, values);
    inserted += 1;
  }
  return inserted;
};

const toNumericOrNull = (value) => {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : null;
};

// ---------------------------------------------------------------------------
// Main sync
// ---------------------------------------------------------------------------

const syncForDate = async (dateRef, bearerToken, plantConfigs, routesUrlBase) => {
  console.log(`[SYNC] Sincronizando data: ${dateRef}`);

  const allRows = [];
  const allRouteRows = [];

  let totalRoutes = 0;
  let totalEvents = 0;
  const errors = [];

  // Fetch routes per plant
  for (const config of plantConfigs) {
    try {
      const query = new URLSearchParams({
        plant_id: String(config.plantId),
        per_page: '60',
        strict_date: '1',
        initial_expected_start_date: `${dateRef}T00:00:00Z`,
        final_expected_start_date: `${dateRef}T23:59:59Z`
      });

      const routesUrl = `${routesUrlBase}?${query.toString()}`;

      const response = await fetch(routesUrl, {
        method: 'GET',
        headers: {
          Authorization: `Bearer ${bearerToken}`,
          'Content-Type': 'application/json',
          'X-Requested-With': 'XMLHttpRequest',
          'x-requested_with': 'XLMHttpRequest'
        }
      });

      if (!response.ok) {
        const text = await response.text().catch(() => '');
        errors.push(`Plant ${config.plantId}: ${response.status} ${text.slice(0, 200)}`);
        continue;
      }

      const payload = await response.json().catch(() => null);
      const routes = pickRoutesArray(payload);
      totalRoutes += routes.length;
      console.log(`[SYNC] Plant ${config.plantId} (${config.filial}): ${routes.length} rotas`);

      // Extrair dados das rotas para a tabela route_web_routes
      for (const route of routes) {
        const routeId = toOptionalInt(route?.id);
        if (routeId == null) continue;

        const filialName = String(route?.plant?.display_name || route?.plant?.name || route?.plant_name || config.filial).trim();
        const operacao = String(config.operacao || '').trim();
        const placa = String(route?.unloading_plate || route?.smartquestion_unloading_plate || route?.last_vehicle?.registration_number || '').trim();
        const motorista = getDriverName(route);

        allRouteRows.push({
          route_id: routeId,
          route_plan_id: route?.route_plan_id || null,
          schedule_order_id: toOptionalInt(route?.schedule_order_id),
          plant_id: config.plantId,
          datalake_plant_id: toOptionalInt(route?.datalake_plant_id),
          filial: filialName,
          operacao,
          roadmap_code: getRouteCode(route),
          status: String(route?.status || '').trim(),
          specific_status: String(route?.specific_status || '').trim(),
          general_status: String(route?.general_status || '').trim(),
          placa,
          motorista,
          last_driver_id: toOptionalInt(route?.last_driver_id || route?.last_driver?.id),
          last_vehicle_id: toOptionalInt(route?.last_vehicle_id || route?.last_vehicle?.id),
          smartquestion_actual_start_time: route?.smartquestion_actual_start_time || route?.start_date || null,
          smartquestion_actual_end_time: route?.smartquestion_actual_end_time || route?.end_date || null,
          smartquestion_collected_liters: toNumericOrNull(route?.smartquestion_collected_liters),
          smartquestion_unloading_plate: String(route?.smartquestion_unloading_plate || '').trim(),
          start_time: route?.start_time || null,
          actual_start_time: route?.actual_start_time || null,
          expected_end_time: route?.expected_end_time || null,
          actual_end_time: route?.actual_end_time || null,
          expected_liters: toNumericOrNull(route?.expected_liters),
          collected_liters: toNumericOrNull(route?.collected_liters),
          unloaded_liters: toNumericOrNull(route?.unloaded_liters),
          volume: toNumericOrNull(route?.volume),
          expected_km: toNumericOrNull(route?.expected_km),
          actual_km: toNumericOrNull(route?.actual_km),
          last_landmark: String(route?.last_landmark || '').trim(),
          data_referencia: dateRef
        });
      }

      // Fetch events per route (batch of 4)
      const batchSize = 4;
      for (let i = 0; i < routes.length; i += batchSize) {
        const batch = routes.slice(i, i + batchSize);

        await Promise.all(batch.map(async (route) => {
          const routeId = toOptionalInt(route?.id);
          if (routeId == null) return;

          try {
            const eventsUrl = `${routesUrlBase}/${routeId}/events?with_occurrences=true`;
            const eventsResponse = await fetch(eventsUrl, {
              method: 'GET',
              headers: {
                Authorization: `Bearer ${bearerToken}`,
                'Content-Type': 'application/json',
                'X-Requested-With': 'XMLHttpRequest',
                'x-requested_with': 'XLMHttpRequest'
              }
            });

            if (!eventsResponse.ok) return;

            const eventsPayload = await eventsResponse.json().catch(() => null);
            const events = pickEventsArray(eventsPayload);
            const rotaCodigo = getRouteCode(route);
            const motorista = getDriverName(route);
            const filialName = String(route?.plant?.display_name || route?.plant?.name || route?.plant_name || config.filial).trim();
            const operacao = String(config.operacao || '').trim();

            for (const event of events) {
              const typeNameNormalized = normalizeText(event?.type_name);
              if (typeNameNormalized !== 'coleta') continue;

              const occurrences = Array.isArray(event?.occurrences) ? event.occurrences : [];
              const nonCollectionOccs = occurrences.filter(isNonCollectionOccurrence);
              const placa = getPlate(event, route);
              const eventId = toOptionalInt(event?.id);
              const isScheduled = !event?.executed && normalizeText(event?.status).startsWith('scheduled');

              // Não coletas (ocorrências)
              if (nonCollectionOccs.length > 0) {
                totalEvents += 1;
                const scraperReason = getScraperOccurrenceReason(event);
                const fallbackReason = getOccurrenceDescription(nonCollectionOccs[0]);
                const motivo = scraperReason || fallbackReason || String(event?.motivo || event?.reason || '').trim();

                for (const occ of nonCollectionOccs) {
                  allRows.push({
                    route_id: routeId,
                    event_id: eventId,
                    plant_id: config.plantId,
                    filial: filialName,
                    operacao,
                    rota_codigo: rotaCodigo,
                    motorista,
                    placa,
                    type_name: String(event?.type_name || ''),
                    reference: String(event?.reference || ''),
                    reference_code: String(event?.reference_code || ''),
                    status: String(event?.status || ''),
                    executed: event?.executed == null ? null : Boolean(event.executed),
                    expected_arrival: event?.expected_arrival || null,
                    actual_arrival: event?.actual_arrival || null,
                    expected_departure: event?.expected_departure || null,
                    actual_departure: event?.actual_departure || null,
                    motivo,
                    status_type: 'nao-coleta',
                    occurrence_id: toOptionalInt(occ?.id),
                    occurrence_type_id: toOptionalInt(occ?.occurrence_type_id ?? occ?.occurrence_type?.id),
                    occurrence_type_description: getOccurrenceDescription(occ),
                    occurrence_inserted_by: String(occ?.inserted_by || ''),
                    occurrence_inserted_at: occ?.inserted_at || null,
                    event_created_at: event?.created_at || null,
                    event_updated_at: event?.updated_at || null,
                    data_referencia: dateRef
                  });
                }
              }

              // Coletas previstas (scheduled, not executed, sem ocorrências relevantes)
              if (isScheduled) {
                totalEvents += 1;
                allRows.push({
                  route_id: routeId,
                  event_id: eventId,
                  plant_id: config.plantId,
                  filial: filialName,
                  operacao,
                  rota_codigo: rotaCodigo,
                  motorista,
                  placa,
                  type_name: String(event?.type_name || ''),
                  reference: String(event?.reference || ''),
                  reference_code: String(event?.reference_code || ''),
                  status: String(event?.status || ''),
                  executed: false,
                  expected_arrival: event?.expected_arrival || null,
                  actual_arrival: null,
                  expected_departure: event?.expected_departure || null,
                  actual_departure: null,
                  motivo: 'Coleta prevista',
                  status_type: 'coleta-prevista',
                  occurrence_id: -1,
                  occurrence_type_id: null,
                  occurrence_type_description: '',
                  occurrence_inserted_by: '',
                  occurrence_inserted_at: null,
                  event_created_at: event?.created_at || null,
                  event_updated_at: event?.updated_at || null,
                  data_referencia: dateRef
                });
              }

            }
          } catch (err) {
            errors.push(`Route ${routeId}: ${err?.message || 'erro'}`);
          }
        }));

        if (i + batchSize < routes.length) await sleep(100);
      }
    } catch (err) {
      errors.push(`Plant ${config.plantId}: ${err?.message || 'erro'}`);
    }
  }

  return { allRows, allRouteRows, totalRoutes, totalEvents, errors };
};

const syncAll = async () => {
  const startedAt = Date.now();

  // Sincroniza sempre 3 dias: d-2, d-1 e dia atual
  const dates = [getD2Date(), getPreviousDayDate(), getCurrentDayDate()];
  console.log(`[SYNC] Iniciando sincronização para ${dates.join(', ')}`);

  // 1. Token
  console.log('[SYNC] Obtendo token Route Web...');
  const bearerToken = await requestRouteWebToken();
  console.log('[SYNC] Token obtido');

  // 2. SharePoint configs
  console.log('[SYNC] Lendo configs do SharePoint...');
  const plantConfigs = await getPlantConfigsFromSharePoint();
  if (plantConfigs.length === 0) {
    console.error('[SYNC] Nenhuma plant config encontrada');
    return;
  }

  // 3. Routes URL
  const routesUrlBase = readRequiredEnv('ROUTE_WEB_ROUTES_URL').replace(/\/+$/, '');

  let grandTotalRoutes = 0;
  let grandTotalEvents = 0;
  let grandTotalUpserted = 0;
  let grandTotalRoutesUpserted = 0;
  const allErrors = [];

  for (const dateRef of dates) {
    const result = await syncForDate(dateRef, bearerToken, plantConfigs, routesUrlBase);
    grandTotalRoutes += result.totalRoutes;
    grandTotalEvents += result.totalEvents;

    // Persist events
    if (result.allRows.length > 0) {
      console.log(`[SYNC] Persistindo ${result.allRows.length} eventos (${dateRef}) no banco...`);
      try {
        grandTotalUpserted += await upsertRouteWebEvents(result.allRows);
      } catch (err) {
        allErrors.push(`DB events ${dateRef}: ${err?.message || 'erro ao persistir'}`);
      }
    }

    // Persist routes
    if (result.allRouteRows.length > 0) {
      console.log(`[SYNC] Persistindo ${result.allRouteRows.length} rotas (${dateRef}) no banco...`);
      try {
        grandTotalRoutesUpserted += await upsertRouteWebRoutes(result.allRouteRows);
      } catch (err) {
        allErrors.push(`DB routes ${dateRef}: ${err?.message || 'erro ao persistir rotas'}`);
      }
    }

    allErrors.push(...result.errors);
  }

  // Close pool
  if (pool) {
    await pool.end().catch(() => {});
    pool = null;
  }

  const durationMs = Date.now() - startedAt;
  console.log(`[SYNC] Concluído em ${(durationMs / 1000).toFixed(1)}s: ${grandTotalRoutes} rotas, ${grandTotalEvents} eventos, ${grandTotalUpserted} events upserted, ${grandTotalRoutesUpserted} routes upserted, ${allErrors.length} erros`);
  if (allErrors.length > 0) {
    console.error('[SYNC] Erros:', allErrors);
  }
};

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------

syncAll().catch((err) => {
  console.error('[FATAL]', err?.message || err);
  process.exit(1);
});
