import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { AlertCircle, Check, CheckSquare, Filter, Loader2, Plus, Search, Settings2, Square } from 'lucide-react';
import { NonCollection, RouteConfig, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import { getWeekString } from '../utils/dateUtils';

type RouteQueryVariables = {
  perPage: number;
  strictDate: number;
  initialExpectedStartDate: string;
  finalExpectedStartDate: string;
};

type RouteWebLabResult = {
  success: boolean;
  tokenPreview?: string;
  totalPlantIds: number;
  queriedPlantIds: number[];
  totalRouteIds: number;
  queriedRouteIds: number[];
  count: number;
  totalEvents: number;
  routes: any[];
  query: RouteQueryVariables;
  failedPlantIds: Array<{ plantId: number; error: string; status?: number }>;
  failedRouteIds: Array<{ routeId: number; error: string; status?: number }>;
};

type CollectionFilterTab = 'nao-coletas' | 'coletas-previstas';

type CollectionRow = {
  key: string;
  routeId: number | null;
  eventRowId: number | null;
  plantId: number;
  filial: string;
  data: string;
  produtor: string;
  codigoProdutor: string;
  rota: string;
  motorista: string;
  placa: string;
  horarioPrevisto: string;
  horarioRealizado: string;
  motivo: string;
  statusLabel: 'Realizada' | 'Pendente' | 'Não Coleta' | 'Prevista';
  statusType: 'realizada' | 'pendente' | 'nao-coleta' | 'coleta-prevista';
  rawStatus: string;
  typeName: string;
  operacao: string;
  isAlreadyLaunched: boolean;
  rawEvent: any;
};

type FilterColumn = 'rota' | 'codigoProdutor' | 'produtor' | 'motivo' | 'motorista' | 'placa' | 'horario' | 'operacao' | 'status';
type SelectedFilters = Record<FilterColumn, string[]>;
type ColFilters = Record<FilterColumn, string>;

const FILTER_COLUMNS: FilterColumn[] = [
  'rota', 'codigoProdutor', 'produtor', 'motivo', 'motorista', 'placa', 'horario', 'operacao', 'status'
];

const createEmptySelectedFilters = (): SelectedFilters => ({
  rota: [], codigoProdutor: [], produtor: [], motivo: [], motorista: [], placa: [], horario: [], operacao: [], status: []
});

type RowsCachePayload = {
  version: 1;
  userEmail: string;
  dayRef: string;
  savedAt: number;
  rows: Array<Omit<CollectionRow, 'rawEvent'> & { rawEvent?: any }>;
};

const ROUTES_QUERY_DEFAULTS = {
  perPage: 60,
  strictDate: 1
};
const ROWS_CACHE_VERSION = 1 as const;
const ROWS_CACHE_MAX_AGE_MS = 12 * 60 * 60 * 1000;
const LAUNCH_STATUS_FILTER_OPTIONS = ['todas', 'nao-lancadas'] as const;
type LaunchStatusFilter = (typeof LAUNCH_STATUS_FILTER_OPTIONS)[number];
type RowsPerPageOption = number | 'all';

const TABLE_HEADER_COLUMNS: Array<{ key: FilterColumn; label: string }> = [
  { key: 'rota', label: 'Rota' },
  { key: 'codigoProdutor', label: 'Código Produtor' },
  { key: 'produtor', label: 'Produtor' },
  { key: 'motivo', label: 'Motivo' },
  { key: 'motorista', label: 'Motorista' },
  { key: 'placa', label: 'Placa' },
  { key: 'horario', label: 'Horário' },
  { key: 'operacao', label: 'Operação' },
  { key: 'status', label: 'Status' }
];

const toNumericId = (value: unknown): number | null => {
  if (value == null) return null;
  if (typeof value === 'number' && Number.isFinite(value)) return Math.trunc(value);
  const raw = String(value).trim();
  if (!raw) return null;
  const match = raw.match(/-?\d+(?:[.,]\d+)?/);
  if (!match) return null;
  const parsed = Number(match[0].replace(',', '.'));
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
};

const normalizeText = (value: unknown): string =>
  String(value ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

const normalizeProducerCode = (value: unknown): string => String(value ?? '').trim().toUpperCase();

const buildLaunchCodeKey = (dayRefIso: string, producerCode: unknown): string => {
  const code = normalizeProducerCode(producerCode);
  if (!dayRefIso || !code) return '';
  return `${dayRefIso}|${code}`;
};

const buildLaunchKeyFromNonCollection = (row: NonCollection): string => {
  const dayRef = toIsoDay(row?.data || '');
  if (!dayRef) return '';
  return buildLaunchCodeKey(dayRef, row?.codigo);
};

const getCurrentDayDate = (): string => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const isValidDateInput = (value: string): boolean => /^\d{4}-\d{2}-\d{2}$/.test(String(value || '').trim());

const buildRouteQueryVariables = (dayRef?: string): RouteQueryVariables => {
  const currentDay = getCurrentDayDate();
  const safeDay = isValidDateInput(String(dayRef || '')) ? String(dayRef).trim() : currentDay;
  return {
    perPage: ROUTES_QUERY_DEFAULTS.perPage,
    strictDate: ROUTES_QUERY_DEFAULTS.strictDate,
    initialExpectedStartDate: `${safeDay}T00:00:00Z`,
    finalExpectedStartDate: `${safeDay}T23:59:59Z`
  };
};

const toIsoDay = (value: unknown): string | null => {
  const raw = String(value ?? '').trim();
  if (!raw) return null;

  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const brMatch = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (brMatch) {
    return `${brMatch[3]}-${brMatch[2]}-${brMatch[1]}`;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return null;
  const year = String(parsed.getFullYear());
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const day = String(parsed.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const getMostFrequentReferenceDay = (rows: NonCollection[]): string | null => {
  const dayCounter = new Map<string, number>();

  rows.forEach((row) => {
    const isoDay = toIsoDay(row?.data);
    if (!isoDay) return;
    dayCounter.set(isoDay, (dayCounter.get(isoDay) || 0) + 1);
  });

  if (dayCounter.size === 0) return null;

  const sorted = Array.from(dayCounter.entries()).sort((a, b) => {
    const byCount = b[1] - a[1];
    if (byCount !== 0) return byCount;
    return b[0].localeCompare(a[0]);
  });

  return sorted[0][0];
};

const getRowsCacheKey = (userEmail: string, dayRef: string): string =>
  `route-web-nao-coletas-cache::${normalizeText(userEmail)}::${dayRef}`;

const normalizeRowsFromCache = (rows: any[]): CollectionRow[] => {
  if (!Array.isArray(rows)) return [];
  return rows
    .map((row, index): CollectionRow | null => {
      if (!row || typeof row !== 'object') return null;
      const routeId = toNumericId(row.routeId);
      const eventRowId = toNumericId(row.eventRowId);
      const key = String(row.key || `${routeId || 'sem-rota'}-${eventRowId || index}`);
      return {
        key,
        routeId,
        eventRowId,
        plantId: toNumericId(row.plantId) || 0,
        filial: String(row.filial || ''),
        data: String(row.data || '--'),
        produtor: String(row.produtor || 'Sem produtor'),
        codigoProdutor: String(row.codigoProdutor || '-'),
        rota: String(row.rota || '-'),
        motorista: String(row.motorista || '-'),
        placa: String(row.placa || '-'),
        horarioPrevisto: String(row.horarioPrevisto || '--'),
        horarioRealizado: String(row.horarioRealizado || '--'),
        motivo: String(row.motivo || ''),
        statusLabel: 'Não Coleta',
        statusType: 'nao-coleta',
        rawStatus: String(row.rawStatus || ''),
        typeName: String(row.typeName || 'Coleta'),
        operacao: String(row.operacao || row.filial || ''),
        isAlreadyLaunched: Boolean(row.isAlreadyLaunched),
        rawEvent: row.rawEvent
      };
    })
    .filter((row): row is CollectionRow => row != null);
};

const readRowsCache = (userEmail: string, dayRef: string): CollectionRow[] => {
  try {
    const key = getRowsCacheKey(userEmail, dayRef);
    const raw = localStorage.getItem(key);
    if (!raw) return [];

    const parsed = JSON.parse(raw) as RowsCachePayload;
    if (!parsed || parsed.version !== ROWS_CACHE_VERSION) return [];
    if (normalizeText(parsed.userEmail) !== normalizeText(userEmail)) return [];
    if (String(parsed.dayRef || '') !== dayRef) return [];
    if (!parsed.savedAt || Date.now() - parsed.savedAt > ROWS_CACHE_MAX_AGE_MS) return [];

    return normalizeRowsFromCache(parsed.rows || []);
  } catch {
    return [];
  }
};

const writeRowsCache = (userEmail: string, dayRef: string, rows: CollectionRow[]): void => {
  try {
    const key = getRowsCacheKey(userEmail, dayRef);
    const payload: RowsCachePayload = {
      version: ROWS_CACHE_VERSION,
      userEmail: normalizeText(userEmail),
      dayRef,
      savedAt: Date.now(),
      rows: rows.map((row) => ({
        ...row,
        rawEvent: undefined
      }))
    };
    localStorage.setItem(key, JSON.stringify(payload));
  } catch {
    // cache é apenas otimização local
  }
};

const getMinutesWaiting = (expectedArrival: unknown): number | null => {
  const raw = String(expectedArrival ?? '').trim();
  if (!raw) return null;
  const expected = new Date(raw);
  if (Number.isNaN(expected.getTime())) return null;
  const now = new Date();
  const diffMs = now.getTime() - expected.getTime();
  return Math.max(0, Math.floor(diffMs / 60000));
};

const formatWaitingTime = (minutes: number | null): string => {
  if (minutes == null) return '';
  if (minutes < 60) return `${minutes}min`;
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return mins > 0 ? `${hours}h${mins}min` : `${hours}h`;
};

const getWaitingColorClass = (minutes: number | null): string => {
  if (minutes == null) return 'text-amber-600 dark:text-amber-400/70';
  if (minutes >= 120) return 'text-red-600 dark:text-red-400 font-bold';
  if (minutes >= 60) return 'text-orange-600 dark:text-orange-400';
  return 'text-amber-600 dark:text-amber-400/70';
};

const formatHour = (value: unknown): string => {
  const raw = String(value ?? '').trim();
  if (!raw) return '--';

  const directMatch = raw.match(/(\d{2}):(\d{2})(?::\d{2})?/);
  if (directMatch) {
    return `${directMatch[1]}:${directMatch[2]}`;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return '--';

  return parsed.toLocaleTimeString('pt-BR', {
    hour: '2-digit',
    minute: '2-digit'
  });
};

const formatDateBR = (value: unknown): string => {
  const raw = String(value ?? '').trim();
  if (!raw) return '--';

  const directDate = raw.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (directDate) {
    return `${directDate[3]}/${directDate[2]}/${directDate[1]}`;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return '--';

  const day = String(parsed.getDate()).padStart(2, '0');
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const year = String(parsed.getFullYear());
  return `${day}/${month}/${year}`;
};

const getStatusBadgeClass = (type: CollectionRow['statusType']): string => {
  if (type === 'realizada') return 'bg-emerald-500/15 text-emerald-300 border border-emerald-500/30';
  if (type === 'nao-coleta') return 'bg-rose-500/15 text-rose-300 border border-rose-500/30';
  if (type === 'coleta-prevista') return 'bg-amber-500/15 text-amber-300 border border-amber-500/30';
  return 'bg-amber-500/15 text-amber-300 border border-amber-500/30';
};

const getStatusPriority = (statusType: CollectionRow['statusType']): number => {
  if (statusType === 'nao-coleta') return 0;
  if (statusType === 'pendente') return 1;
  if (statusType === 'coleta-prevista') return 3;
  return 2;
};

const getRowStatusFilterValue = (row: CollectionRow): string =>
  row.isAlreadyLaunched ? 'Não coleta lançada' : 'Não lançada';

const getRowColumnValue = (row: CollectionRow, column: FilterColumn): string => {
  if (column === 'rota') return String(row.rota || '-');
  if (column === 'codigoProdutor') return String(row.codigoProdutor || '-');
  if (column === 'produtor') return String(row.produtor || '-');
  if (column === 'motivo') return String(row.motivo || '-');
  if (column === 'motorista') return String(row.motorista || '-');
  if (column === 'placa') return String(row.placa || '-');
  if (column === 'horario') return `Previsto ${row.horarioPrevisto} | Realizado ${row.horarioRealizado}`;
  if (column === 'operacao') return String(row.operacao || '-');
  return getRowStatusFilterValue(row);
};

const fetchJsonWithTimeout = async (
  url: string,
  init: RequestInit,
  timeoutMs: number
): Promise<{ response: Response; data: any }> => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      ...init,
      signal: controller.signal
    });

    const data = await response.json().catch(() => ({}));
    return { response, data };
  } finally {
    clearTimeout(timer);
  }
};

const RouteWebLabView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [isLoading, setIsLoading] = useState(false);
  const [isHydratingEvents, setIsHydratingEvents] = useState(false);
  const [eventProgress, setEventProgress] = useState<{ done: number; total: number }>({ done: 0, total: 0 });
  const [fetchError, setFetchError] = useState<string | null>(null);
  const [referenceDate, setReferenceDate] = useState<string>(getCurrentDayDate());
  const [result, setResult] = useState<RouteWebLabResult | null>(null);
  const [collectionRowsState, setCollectionRowsState] = useState<CollectionRow[]>([]);
  const [activeTab, setActiveTab] = useState<CollectionFilterTab>('nao-coletas');
  const [searchText, setSearchText] = useState('');
  const [filialFilter, setFilialFilter] = useState('todas');
  const [launchStatusFilter, setLaunchStatusFilter] = useState<LaunchStatusFilter>('todas');
  const [rowsPerPage, setRowsPerPage] = useState<RowsPerPageOption>(60);
  const [selectedFilters, setSelectedFilters] = useState<SelectedFilters>(createEmptySelectedFilters);
  const [colFilters, setColFilters] = useState<ColFilters>({} as ColFilters);
  const [filterPos, setFilterPos] = useState<{ top: number; left: number } | null>(null);
  const [activeFilterCol, setActiveFilterCol] = useState<FilterColumn | null>(null);
  const [addingRowKeys, setAddingRowKeys] = useState<Set<string>>(new Set());
  const [currentPage, setCurrentPage] = useState<number>(1);

  const fetchRunIdRef = useRef(0);
  const rowCacheMapRef = useRef<Map<string, { signature: string; row: CollectionRow }>>(new Map());
  const routeRowKeysRef = useRef<Map<number, Set<string>>>(new Map());
  const launchedCodeKeysRef = useRef<Set<string>>(new Set());
  const persistRowsTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const latestDayRefRef = useRef<string>(getCurrentDayDate());
  const filterDropdownRef = useRef<HTMLDivElement | null>(null);
  const filterAnchorRef = useRef<HTMLElement | null>(null);

  const getRowSignature = useCallback((row: CollectionRow): string => {
    return [
      row.routeId ?? '',
      row.eventRowId ?? '',
      row.data,
      row.codigoProdutor,
      row.produtor,
      row.motivo,
      row.motorista,
      row.placa,
      row.horarioPrevisto,
      row.horarioRealizado,
      row.statusType,
      row.isAlreadyLaunched ? '1' : '0'
    ].join('|');
  }, []);

  const scheduleRowsCachePersist = useCallback((dayRef: string) => {
    latestDayRefRef.current = dayRef;
    if (persistRowsTimerRef.current) {
      clearTimeout(persistRowsTimerRef.current);
    }

    persistRowsTimerRef.current = setTimeout(() => {
      const rows = Array.from(rowCacheMapRef.current.values()).map((entry) => entry.row);
      writeRowsCache(currentUser.email, latestDayRefRef.current, rows);
      persistRowsTimerRef.current = null;
    }, 500);
  }, [currentUser.email]);

  const resolveReferenceDateFromNonCollections = useCallback(
    async (
      token: string,
      configs: RouteConfig[]
    ): Promise<{ dayRef: string; totalRows: number; filteredRowsCount: number; launchedCodeKeys: Set<string> }> => {
      const allRows = await SharePointService.getNonCollections(token, currentUser.email);
      const operationSet = new Set(
        (configs || [])
          .map((cfg) => normalizeText(cfg.operacao))
          .filter(Boolean)
      );

      const filteredRows = (allRows || []).filter((row) => {
        if (operationSet.size === 0) return true;
        return operationSet.has(normalizeText(row.operacao));
      });

      const dayRef =
        getMostFrequentReferenceDay(filteredRows) ||
        getMostFrequentReferenceDay(allRows) ||
        getCurrentDayDate();

      const launchedCodeKeys = new Set<string>();
      filteredRows.forEach((row) => {
        const rowDay = toIsoDay(row?.data || '');
        if (rowDay !== dayRef) return;
        const launchKey = buildLaunchKeyFromNonCollection(row);
        if (!launchKey) return;
        launchedCodeKeys.add(launchKey);
      });

      return {
        dayRef,
        totalRows: allRows.length,
        filteredRowsCount: filteredRows.length,
        launchedCodeKeys
      };
    },
    [currentUser.email]
  );

  useEffect(() => {
    return () => {
      if (persistRowsTimerRef.current) {
        clearTimeout(persistRowsTimerRef.current);
        persistRowsTimerRef.current = null;
      }
    };
  }, []);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 'l') {
        event.preventDefault();
        setSelectedFilters(createEmptySelectedFilters());
        setColFilters({} as ColFilters);
        setActiveFilterCol(null);
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, []);

  useEffect(() => {
    if (!activeFilterCol) return;
    const handlePointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (filterDropdownRef.current?.contains(target)) return;
      setActiveFilterCol(null);
      setFilterPos(null);
    };
    const handleEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') { setActiveFilterCol(null); setFilterPos(null); }
    };
    document.addEventListener('mousedown', handlePointerDown);
    document.addEventListener('keydown', handleEscape);
    return () => {
      document.removeEventListener('mousedown', handlePointerDown);
      document.removeEventListener('keydown', handleEscape);
    };
  }, [activeFilterCol]);

  const markRowAsLaunched = useCallback(
    (rowKey: string, dayRefIso: string, producerCode: string) => {
      const launchKey = buildLaunchCodeKey(dayRefIso, producerCode);
      if (launchKey) {
        launchedCodeKeysRef.current.add(launchKey);
      }

      setCollectionRowsState((prev) => {
        let changed = false;
        const nextRows = prev.map((row) => {
          const sameProducerCode = normalizeProducerCode(row.codigoProdutor) === normalizeProducerCode(producerCode);
          const sameDay = toIsoDay(row.data || '') === dayRefIso;
          if (row.key !== rowKey && !(sameProducerCode && sameDay)) return row;
          if (row.isAlreadyLaunched) return row;
          changed = true;
          const updated = { ...row, isAlreadyLaunched: true };
          const signature = getRowSignature(updated);
          rowCacheMapRef.current.set(updated.key, { signature, row: updated });
          return updated;
        });
        if (changed) {
          scheduleRowsCachePersist(dayRefIso);
        }
        return nextRows;
      });
    },
    [getRowSignature, scheduleRowsCachePersist]
  );

  const handleLaunchNonCollection = useCallback(
    async (row: CollectionRow) => {
      if (!row || row.isAlreadyLaunched) return;
      if (addingRowKeys.has(row.key)) return;

      const dayRefIso = toIsoDay(row.data || '') || latestDayRefRef.current || getCurrentDayDate();
      const launchKey = buildLaunchCodeKey(dayRefIso, row.codigoProdutor);
      if (launchKey && launchedCodeKeysRef.current.has(launchKey)) {
        markRowAsLaunched(row.key, dayRefIso, row.codigoProdutor);
        return;
      }

      setAddingRowKeys((prev) => {
        const next = new Set(prev);
        next.add(row.key);
        return next;
      });

      try {
        const token = (await getValidToken()) || currentUser.accessToken;
        if (!token) {
          throw new Error('Token de sessão indisponível para lançar não coleta.');
        }

        const payload: NonCollection = {
          id: '',
          semana: getWeekString(row.data || ''),
          rota: row.rota || '',
          data: row.data || formatDateBR(dayRefIso),
          codigo: row.codigoProdutor || '',
          produtor: row.produtor || '',
          motivo: row.motivo || '',
          observacao: '',
          acao: '',
          dataAcao: '',
          ultimaColeta: '',
          Culpabilidade: 'Não se aplica',
          operacao: row.operacao || row.filial || ''
        };

        await SharePointService.saveNonCollection(token, payload);
        markRowAsLaunched(row.key, dayRefIso, row.codigoProdutor);
      } catch (error: any) {
        console.error('[ROUTE_WEB_LAUNCH_DEBUG] Erro ao lançar não coleta:', {
          rowKey: row.key,
          codigoProdutor: row.codigoProdutor,
          error: error?.message || error
        });
        setFetchError(error?.message || 'Falha ao lançar não coleta na lista.');
      } finally {
        setAddingRowKeys((prev) => {
          const next = new Set(prev);
          next.delete(row.key);
          return next;
        });
      }
    },
    [addingRowKeys, currentUser.accessToken, markRowAsLaunched]
  );

  const handleFetchRoutes = useCallback(async () => {
    const runId = fetchRunIdRef.current + 1;
    fetchRunIdRef.current = runId;

    setIsLoading(true);
    setIsHydratingEvents(false);
    setEventProgress({ done: 0, total: 0 });
    setFetchError(null);
    setResult(null);
    setCollectionRowsState([]);
    rowCacheMapRef.current.clear();
    routeRowKeysRef.current.clear();
    launchedCodeKeysRef.current = new Set();

    try {
      const graphToken = (await getValidToken()) || currentUser.accessToken;
      if (!graphToken) {
        throw new Error('Token de sessão indisponível para identificar as operações do login.');
      }

      const accessResult = await SharePointService.getRouteConfigsByAccess(graphToken, currentUser.email, false);
      const allowedPlantIds = Array.from(
        new Set(
          (accessResult.configs || [])
            .map((cfg) => toNumericId(cfg.plantId))
            .filter((id): id is number => id != null)
        )
      ).sort((a, b) => a - b);

      const plantDisplayMap = new Map<number, string>();
      (accessResult.configs || []).forEach((cfg) => {
        const plantId = toNumericId(cfg.plantId);
        if (plantId == null || plantDisplayMap.has(plantId)) return;
        plantDisplayMap.set(plantId, String(cfg.nomeExibicao || cfg.operacao || `Plant ${plantId}`).trim());
      });

      if (allowedPlantIds.length === 0) {
        setReferenceDate(getCurrentDayDate());
        setResult({
          success: true,
          totalPlantIds: 0,
          queriedPlantIds: [],
          totalRouteIds: 0,
          queriedRouteIds: [],
          count: 0,
          totalEvents: 0,
          routes: [],
          query: buildRouteQueryVariables(),
          failedPlantIds: [],
          failedRouteIds: []
        });
        return;
      }

      const referenceDateMeta = await resolveReferenceDateFromNonCollections(graphToken, accessResult.configs || []);
      if (runId !== fetchRunIdRef.current) return;

      setReferenceDate(referenceDateMeta.dayRef);
      launchedCodeKeysRef.current = referenceDateMeta.launchedCodeKeys;
      latestDayRefRef.current = referenceDateMeta.dayRef;
      const queryVariables = buildRouteQueryVariables(referenceDateMeta.dayRef);

      // Query events from database
      const { response: eventsResponse, data: eventsData } = await fetchJsonWithTimeout(
        '/api/route-web-db',
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            entity: 'events',
            dataReferencia: referenceDateMeta.dayRef,
            plantIds: allowedPlantIds
          })
        },
        15000
      );

      if (runId !== fetchRunIdRef.current) return;

      if (!eventsResponse.ok || !eventsData?.success) {
        throw new Error(eventsData?.error || 'Falha ao consultar eventos do banco de dados');
      }

      let dbEvents: any[] = eventsData.events || [];

      // Deduplicar: um único registro por event_id (manter o com occurrence_inserted_at mais recente)
      const seenEvents = new Map<string, any>();
      for (const row of dbEvents) {
        const key = String(row.event_id ?? '');
        if (!key) continue;
        const existing = seenEvents.get(key);
        if (!existing) {
          seenEvents.set(key, row);
          continue;
        }
        const existingTime = String(existing.occurrence_inserted_at || existing.event_updated_at || '');
        const rowTime = String(row.occurrence_inserted_at || row.event_updated_at || '');
        if (rowTime > existingTime) {
          seenEvents.set(key, row);
        }
      }
      if (seenEvents.size > 0) {
        dbEvents = Array.from(seenEvents.values());
      }

      // Build CollectionRow[] from DB rows
      const rows: CollectionRow[] = dbEvents.map((row: any, index: number): CollectionRow => {
        const routeId = toNumericId(row.route_id);
        const eventRowId = toNumericId(row.event_id);
        const plantId = toNumericId(row.plant_id) || 0;
        const filial = String(row.filial || plantDisplayMap.get(plantId) || `Plant ${plantId}`).trim();
        const isLaunched = Boolean(row.is_already_launched);
        const isColetaPrevista = String(row.status_type || '') === 'coleta-prevista';

        return {
          key: `${routeId || 'sem-rota'}-${eventRowId || index}`,
          routeId,
          eventRowId,
          plantId,
          filial,
          data: formatDateBR(row.expected_arrival || row.expected_departure || row.event_created_at),
          produtor: String(row.reference || 'Sem produtor'),
          codigoProdutor: String(row.reference_code || '-'),
          rota: String(row.rota_codigo || '-'),
          motorista: String(row.motorista || '-'),
          placa: String(row.placa || '-'),
          horarioPrevisto: formatHour(row.expected_arrival || row.expected_departure || row.event_created_at),
          horarioRealizado: isColetaPrevista ? '--' : formatHour(row.actual_arrival || row.actual_departure || row.event_updated_at),
          motivo: String(row.motivo || '').trim(),
          statusLabel: isColetaPrevista ? ('Prevista' as const) : ('Não Coleta' as const),
          statusType: isColetaPrevista ? ('coleta-prevista' as const) : ('nao-coleta' as const),
          rawStatus: String(row.status || ''),
          typeName: String(row.type_name || ''),
          operacao: String(row.operacao || filial).trim(),
          isAlreadyLaunched:
            isLaunched ||
            launchedCodeKeysRef.current.has(
              buildLaunchCodeKey(referenceDateMeta.dayRef, row.reference_code)
            ),
          rawEvent: row
        };
      });

      setCollectionRowsState(rows);

      // Populate caches
      rows.forEach((row) => {
        rowCacheMapRef.current.set(row.key, { signature: getRowSignature(row), row });
        if (row.routeId != null) {
          const prev = routeRowKeysRef.current.get(row.routeId) || new Set<string>();
          prev.add(row.key);
          routeRowKeysRef.current.set(row.routeId, prev);
        }
      });

      scheduleRowsCachePersist(referenceDateMeta.dayRef);

      const uniqueRouteIds = Array.from(new Set(rows.map((r) => r.routeId).filter((id): id is number => id != null)));

      setResult({
        success: true,
        totalPlantIds: allowedPlantIds.length,
        queriedPlantIds: allowedPlantIds,
        totalRouteIds: uniqueRouteIds.length,
        queriedRouteIds: uniqueRouteIds,
        count: rows.length,
        totalEvents: dbEvents.length,
        routes: [],
        query: queryVariables,
        failedPlantIds: [],
        failedRouteIds: []
      });

    } catch (error: any) {
      if (runId !== fetchRunIdRef.current) return;
      setIsHydratingEvents(false);
      setEventProgress({ done: 0, total: 0 });
      setFetchError(error?.message || 'Falha ao carregar eventos do banco de dados');
    } finally {
      if (runId === fetchRunIdRef.current) {
        setIsLoading(false);
      }
    }
  }, [
    currentUser.accessToken,
    currentUser.email,
    getRowSignature,
    resolveReferenceDateFromNonCollections,
    scheduleRowsCachePersist
  ]);

  useEffect(() => {
    void handleFetchRoutes();
    return () => {
      fetchRunIdRef.current += 1;
    };
  }, [handleFetchRoutes]);

  const collectionRows = useMemo<CollectionRow[]>(() => {
    return collectionRowsState;
  }, [collectionRowsState]);

  const filialOptions = useMemo(() => {
    const options = Array.from(new Set(collectionRows.map((row) => row.filial))).sort((a, b) => a.localeCompare(b));
    return ['todas', ...options];
  }, [collectionRows]);

  const rowsForFilters = useMemo(() => {
    let rows = collectionRows;

    if (activeTab === 'nao-coletas') {
      rows = rows.filter((row) => row.statusType === 'nao-coleta');
    } else if (activeTab === 'coletas-previstas') {
      rows = rows.filter((row) => row.statusType === 'coleta-prevista');
    }

    if (filialFilter !== 'todas') {
      rows = rows.filter((row) => row.filial === filialFilter);
    }

    if (launchStatusFilter === 'nao-lancadas') {
      rows = rows.filter((row) => !row.isAlreadyLaunched);
    }

    const normalizedSearch = normalizeText(searchText);
    if (normalizedSearch) {
      rows = rows.filter((row) => {
        const combined = normalizeText(`${row.data} ${row.placa} ${row.produtor} ${row.codigoProdutor} ${row.filial} ${row.rota} ${row.motorista} ${row.motivo}`);
        return combined.includes(normalizedSearch);
      });
    }

    return rows;
  }, [collectionRows, filialFilter, launchStatusFilter, searchText, activeTab]);

  const filterOptionsByColumn = useMemo(() => {
    const next: Record<FilterColumn, string[]> = { rota: [], codigoProdutor: [], produtor: [], motivo: [], motorista: [], placa: [], horario: [], operacao: [], status: [] };
    FILTER_COLUMNS.forEach((column) => {
      const set = new Set<string>();
      rowsForFilters.forEach((row) => {
        const value = getRowColumnValue(row, column);
        if (!value) return;
        set.add(value);
      });
      next[column] = Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    });
    return next;
  }, [rowsForFilters]);

  const filteredRows = useMemo(() => {
    let rows = rowsForFilters;

    rows = rows.filter((row) => {
      for (const column of FILTER_COLUMNS) {
        const selectedValues = selectedFilters[column] || [];
        if (selectedValues.length === 0) continue;
        const cellValue = getRowColumnValue(row, column);
        if (!selectedValues.includes(cellValue)) return false;
      }
      return true;
    });

    return [...rows].sort((a, b) => {
      const byPriority = getStatusPriority(a.statusType) - getStatusPriority(b.statusType);
      if (byPriority !== 0) return byPriority;
      // Coletas previstas: mais tempo esperando primeiro (descendente)
      const aWait = getMinutesWaiting(a.rawEvent?.expected_arrival) ?? 0;
      const bWait = getMinutesWaiting(b.rawEvent?.expected_arrival) ?? 0;
      if (aWait !== bWait) return bWait - aWait;
      const byFilial = String(a.filial || '').localeCompare(String(b.filial || ''), 'pt-BR');
      if (byFilial !== 0) return byFilial;
      const byRota = String(a.rota || '').localeCompare(String(b.rota || ''), 'pt-BR');
      if (byRota !== 0) return byRota;
      return (a.eventRowId || 0) - (b.eventRowId || 0);
    });
  }, [selectedFilters, rowsForFilters]);

  const tabCounts = useMemo(() => {
    const naoColetas = collectionRows.filter((row) => row.statusType === 'nao-coleta').length;
    const coletasPrevistas = collectionRows.filter((row) => row.statusType === 'coleta-prevista').length;

    return {
      todas: naoColetas,
      realizadas: 0,
      pendentes: 0,
      'nao-coletas': naoColetas,
      'coletas-previstas': coletasPrevistas
    };
  }, [collectionRows]);

  const launchedStats = useMemo(() => {
    const naoColetas = collectionRows.filter((row) => row.statusType === 'nao-coleta');
    const total = naoColetas.length;
    const launched = naoColetas.filter((row) => row.isAlreadyLaunched).length;
    const percent = total > 0 ? Number(((launched / total) * 100).toFixed(1)) : 0;
    return { total, launched, percent };
  }, [collectionRows]);

  const totalPages = useMemo(() => {
    if (filteredRows.length === 0) return 1;
    if (rowsPerPage === 'all') return 1;
    return Math.max(1, Math.ceil(filteredRows.length / rowsPerPage));
  }, [filteredRows.length, rowsPerPage]);

  const paginatedRows = useMemo(() => {
    if (rowsPerPage === 'all') return filteredRows;
    const start = (currentPage - 1) * rowsPerPage;
    return filteredRows.slice(start, start + rowsPerPage);
  }, [currentPage, filteredRows, rowsPerPage]);

  const pageRange = useMemo(() => {
    if (filteredRows.length === 0) return { start: 0, end: 0 };
    if (rowsPerPage === 'all') return { start: 1, end: filteredRows.length };
    const start = (currentPage - 1) * rowsPerPage + 1;
    const end = Math.min(currentPage * rowsPerPage, filteredRows.length);
    return { start, end };
  }, [currentPage, filteredRows.length, rowsPerPage]);

  useEffect(() => {
    setCurrentPage(1);
  }, [activeTab, filialFilter, launchStatusFilter, searchText, rowsPerPage, selectedFilters]);

  useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages);
    }
  }, [currentPage, totalPages]);

  const filterDebug = useMemo(() => {
    const statusCounts = collectionRows.reduce((acc: Record<string, number>, row) => {
      acc[row.statusType] = (acc[row.statusType] || 0) + 1;
      return acc;
    }, {});

    const typeNameCounts = collectionRows.reduce((acc: Record<string, number>, row) => {
      const key = String(row.typeName || 'SEM_TIPO').trim() || 'SEM_TIPO';
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    let afterTab = collectionRows;
    if (activeTab === 'nao-coletas') {
      afterTab = afterTab.filter((row) => row.statusType === 'nao-coleta');
    } else if (activeTab === 'coletas-previstas') {
      afterTab = afterTab.filter((row) => row.statusType === 'coleta-prevista');
    }

    let afterFilial = afterTab;
    if (filialFilter !== 'todas') {
      afterFilial = afterFilial.filter((row) => row.filial === filialFilter);
    }

    let afterLaunch = afterFilial;
    if (launchStatusFilter === 'nao-lancadas') {
      afterLaunch = afterLaunch.filter((row) => !row.isAlreadyLaunched);
    }

    let afterSearch = afterLaunch;
    const normalizedSearch = normalizeText(searchText);
    if (normalizedSearch) {
      afterSearch = afterSearch.filter((row) => {
        const combined = normalizeText(`${row.data} ${row.placa} ${row.produtor} ${row.codigoProdutor} ${row.filial} ${row.rota} ${row.motorista} ${row.motivo}`);
        return combined.includes(normalizedSearch);
      });
    }

    let afterColFilters = afterSearch;
    afterColFilters = afterColFilters.filter((row) => {
      for (const column of FILTER_COLUMNS) {
        const selectedValues = selectedFilters[column] || [];
        if (selectedValues.length === 0) continue;
        const cellValue = getRowColumnValue(row, column);
        if (!selectedValues.includes(cellValue)) return false;
      }
      return true;
    });

    return {
      activeTab,
      filialFilter,
      launchStatusFilter,
      selectedFilters,
      searchText,
      baseCount: collectionRows.length,
      statusCounts,
      typeNameCounts,
      afterTabCount: afterTab.length,
      afterFilialCount: afterFilial.length,
      afterLaunchCount: afterLaunch.length,
      afterSearchCount: afterSearch.length,
      afterColFiltersCount: afterColFilters.length,
      finalCount: filteredRows.length,
      sample: filteredRows.slice(0, 5).map((row) => ({
        eventId: row.eventRowId,
        routeId: row.routeId,
        filial: row.filial,
        typeName: row.typeName,
        statusType: row.statusType,
        isAlreadyLaunched: row.isAlreadyLaunched
      }))
    };
  }, [activeTab, collectionRows, filialFilter, filteredRows, launchStatusFilter, searchText, selectedFilters]);

  useEffect(() => {
    if (!result) return;
    if (isHydratingEvents) return;
  }, [filterDebug, isHydratingEvents, result]);

  return (
    <div className="h-full overflow-auto bg-slate-50 dark:bg-slate-950 p-3 lg:p-4">
      <div className="w-full max-w-none mx-auto space-y-4">
        <div className="bg-white/95 dark:bg-slate-900/95 backdrop-blur-sm border border-white/50 dark:border-slate-800 rounded-[2rem] p-4 shadow-2xl">
          <div className="flex items-start gap-4">
            <div className="w-12 h-12 rounded-2xl bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 flex items-center justify-center shrink-0">
              <Settings2 size={22} />
            </div>
            <div className="space-y-2">
              <div>
                <h1 className="text-2xl font-black uppercase tracking-tight text-slate-800 dark:text-white">Integração de eventos</h1>
              </div>
              <div className="flex items-center gap-2 flex-wrap pt-2">
                <span className="inline-flex items-center gap-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-100 dark:bg-slate-800 px-3 py-2 text-xs font-semibold text-slate-700 dark:text-slate-200">
                  Data de referência: {formatDateBR(referenceDate)}
                </span>
              </div>
            </div>
          </div>
        </div>

        <section className="relative bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-2xl overflow-hidden">
          <div className="p-5 border-b border-slate-200 dark:border-slate-700 flex flex-col xl:flex-row xl:items-center gap-3 justify-between">
            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={() => { setSelectedFilters(createEmptySelectedFilters()); setColFilters({} as ColFilters); setActiveFilterCol(null); setActiveTab('nao-coletas'); }}
                className={`px-3 py-1.5 rounded-lg text-sm font-semibold transition ${activeTab === 'nao-coletas' ? 'bg-slate-700 text-white' : 'text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800'}`}
              >
                Não Coletas ({tabCounts['nao-coletas']})
              </button>
              <button
                type="button"
                onClick={() => { setSelectedFilters(createEmptySelectedFilters()); setColFilters({} as ColFilters); setActiveFilterCol(null); setActiveTab('coletas-previstas'); }}
                className={`px-3 py-1.5 rounded-lg text-sm font-semibold transition ${activeTab === 'coletas-previstas' ? 'bg-amber-700 text-white' : 'text-amber-600 dark:text-amber-300 hover:bg-amber-50 dark:hover:bg-amber-900/40'}`}
              >
                Coletas Previstas ({tabCounts['coletas-previstas']})
              </button>
              <div className="px-3 py-1.5 rounded-lg border border-cyan-500/30 bg-cyan-50 dark:bg-cyan-500/10 min-w-[220px]">
                <div className="text-[10px] font-semibold uppercase tracking-[0.14em] text-cyan-600 dark:text-cyan-300/85">% não coletas lançadas</div>
                <div className="mt-0.5 flex items-baseline gap-2">
                  <span className="text-xl leading-none font-black text-cyan-700 dark:text-cyan-200">{launchedStats.percent}%</span>
                  <span className="text-xs font-semibold text-cyan-600 dark:text-cyan-100/80">{launchedStats.launched}/{launchedStats.total}</span>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2 flex-col sm:flex-row">
              <div className="relative w-full sm:w-[320px]">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  value={searchText}
                  onChange={(event) => setSearchText(event.target.value)}
                  placeholder="Buscar produtor, código, rota, motivo..."
                  className="w-full bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg pl-9 pr-3 py-2 text-sm text-slate-800 dark:text-slate-100 placeholder-slate-400 outline-none focus:border-slate-400 dark:focus:border-slate-500"
                />
              </div>
              <select
                value={String(rowsPerPage)}
                onChange={(event) => {
                  const value = String(event.target.value);
                  if (value === 'all') {
                    setRowsPerPage('all');
                    return;
                  }
                  const parsed = Number(value);
                  setRowsPerPage(Number.isFinite(parsed) && parsed > 0 ? parsed : 60);
                }}
                className="w-full sm:w-[160px] bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg px-3 py-2 text-sm text-slate-800 dark:text-slate-100 outline-none focus:border-slate-400 dark:focus:border-slate-500"
              >
                <option value="60">60 por página</option>
                <option value="100">100 por página</option>
                <option value="all">Todas</option>
              </select>
            </div>
          </div>

          {isLoading && (
            <div className="p-8 flex items-center justify-center">
              <Loader2 size={24} className="animate-spin text-slate-400" />
            </div>
          )}

          {fetchError && (
            <div className="m-5 rounded-xl border border-red-300 dark:border-red-700/40 bg-red-50 dark:bg-red-950/30 px-4 py-3 text-sm font-semibold text-red-700 dark:text-red-300 flex items-start gap-2">
              <AlertCircle size={16} className="mt-0.5 shrink-0" />
              <span>{fetchError}</span>
            </div>
          )}

          {result && !isLoading && (
            <>
              <div className="px-5 py-3 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between gap-3 text-xs font-semibold text-slate-600 dark:text-slate-300">
                <span>
                  Exibindo <strong className="text-slate-800 dark:text-white">{pageRange.start}-{pageRange.end}</strong> de <strong className="text-slate-800 dark:text-white">{filteredRows.length}</strong>
                </span>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setCurrentPage((prev) => Math.max(1, prev - 1))}
                    disabled={rowsPerPage === 'all' || currentPage <= 1}
                    className="px-3 py-1.5 rounded-md border border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-200 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 dark:hover:bg-slate-800"
                  >
                    Anterior
                  </button>
                  <span className="text-slate-600 dark:text-slate-300">Página <strong className="text-slate-800 dark:text-white">{currentPage}</strong>/{totalPages}</span>
                  <button
                    type="button"
                    onClick={() => setCurrentPage((prev) => Math.min(totalPages, prev + 1))}
                    disabled={rowsPerPage === 'all' || currentPage >= totalPages}
                    className="px-3 py-1.5 rounded-md border border-slate-200 dark:border-slate-700 text-slate-700 dark:text-slate-200 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-100 dark:hover:bg-slate-800"
                  >
                    Próxima
                  </button>
                </div>
              </div>

              <div className="overflow-auto">
                <table className="w-full border-collapse text-[10px]">
                  <thead className="sticky top-0 bg-slate-100 dark:bg-slate-800/80 z-10">
                    <tr>
                      {TABLE_HEADER_COLUMNS.filter((col) => !(activeTab === 'coletas-previstas' && (col.key === 'motivo' || col.key === 'status'))).map((column) => {
                        const filterColumn = column.key;
                        const allValues = filterOptionsByColumn[filterColumn] || [];
                        const isOpen = activeFilterCol === filterColumn;
                        const activeFilter = (selectedFilters[filterColumn] || []).length > 0;
                        const showFilterButton = filterColumn !== 'horario';

                        return (
                          <th key={column.key} className="p-2 border-b border-r border-slate-200 dark:border-slate-700 text-slate-500 dark:text-slate-400 font-bold text-left group">
                            <div className="flex items-center justify-between">
                              <span>{column.label}</span>
                              {showFilterButton && (
                              <button
                                ref={(el) => { if (isOpen && el) filterAnchorRef.current = el; }}
                                onClick={(e) => {
                                  if (activeFilterCol === filterColumn) {
                                    setActiveFilterCol(null);
                                    setFilterPos(null);
                                  } else {
                                    const rect = e.currentTarget.getBoundingClientRect();
                                    const dropdownWidth = 256;
                                    let left = rect.left;
                                    if (left + dropdownWidth > window.innerWidth) {
                                      left = window.innerWidth - dropdownWidth - 8;
                                    }
                                    setFilterPos({ top: rect.bottom + 4, left });
                                    setActiveFilterCol(filterColumn);
                                  }
                                }}
                                className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                  activeFilter
                                    ? 'opacity-100 bg-primary-600 text-white'
                                    : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                }`}
                              >
                                <Filter size={10} />
                              </button>
                              )}
                            </div>
                            {isOpen && (() => {
                              const selected = selectedFilters[filterColumn] || [];
                              const colFilter = colFilters[filterColumn] || '';
                              const filteredValues = allValues.filter((v) => v.toLowerCase().includes(colFilter.toLowerCase()));
                              const toggleValue = (val: string) => {
                                const next = selected.includes(val) ? selected.filter((v) => v !== val) : [...selected, val];
                                setSelectedFilters({ ...selectedFilters, [filterColumn]: next });
                              };

                              return (
                                <div
                                  ref={filterDropdownRef}
                                  className="fixed z-[9999] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-2xl rounded-xl w-64 flex flex-col"
                                  style={{ top: filterPos?.top ?? 0, left: filterPos?.left ?? 0, height: '380px' }}
                                >
                                  <div className="flex items-center gap-2 p-3 pb-2 bg-slate-50 dark:bg-slate-900 rounded-t-xl border-b border-slate-200 dark:border-slate-700 shrink-0">
                                    <Search size={14} className="text-slate-400 shrink-0" />
                                    <input type="text" placeholder="Filtrar..." autoFocus value={colFilter} onChange={(e) => {
                                      const val = e.target.value;
                                      setColFilters({ ...colFilters, [filterColumn]: val });
                                      if (val.trim()) {
                                        const matched = allValues.filter((v) => v.toLowerCase().includes(val.toLowerCase()));
                                        setSelectedFilters({ ...selectedFilters, [filterColumn]: matched });
                                      } else {
                                        setSelectedFilters({ ...selectedFilters, [filterColumn]: [] });
                                      }
                                    }} className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800 dark:text-white" />
                                  </div>
                                  <div className="overflow-y-auto p-2 space-y-1 flex-1">
                                    {filteredValues.length === 0 ? (
                                      <div className="px-3 py-3 text-xs text-slate-400">Nenhum valor encontrado.</div>
                                    ) : (
                                      filteredValues.map((v) => (
                                        <div key={v} onClick={() => toggleValue(v)} className="flex items-center gap-2 p-2.5 hover:bg-slate-50 dark:hover:bg-slate-700 rounded-lg cursor-pointer transition-all">
                                          {selected.includes(v) ? <CheckSquare size={14} className="text-blue-600 shrink-0" /> : <Square size={14} className="text-slate-300 shrink-0" />}
                                          <span className="text-[10px] font-bold uppercase truncate text-slate-700 dark:text-slate-300">{v || '(VAZIO)'}</span>
                                        </div>
                                      ))
                                    )}
                                  </div>
                                  <div className="p-2 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50 rounded-b-xl shrink-0">
                                    <button onClick={() => { setColFilters({ ...colFilters, [filterColumn]: '' }); setSelectedFilters({ ...selectedFilters, [filterColumn]: [] }); setActiveFilterCol(null); setFilterPos(null); }} className="w-full py-2.5 text-[10px] font-black uppercase text-red-600 bg-red-50 dark:bg-red-900/30 hover:bg-red-100 dark:hover:bg-red-900/50 rounded-lg border border-red-100 dark:border-red-900/50 transition-colors"> Limpar Filtro </button>
                                  </div>
                                </div>
                              );
                            })()}
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {paginatedRows.length === 0 ? (
                      <tr>
                        <td colSpan={activeTab === 'coletas-previstas' ? 7 : 9} className="p-4 text-center text-slate-400 font-medium">
                          Nenhum registro encontrado com os filtros atuais.
                        </td>
                      </tr>
                    ) : (
                      paginatedRows.map((row) => {
                        const isPrevista = activeTab === 'coletas-previstas';
                        return (
                        <tr key={row.key} className="hover:bg-slate-50 dark:hover:bg-slate-800/50 border-b border-slate-100 dark:border-slate-800">
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300">
                            <span className="font-bold">{row.rota}</span>
                          </td>
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300">
                            {row.codigoProdutor}
                          </td>
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300 font-semibold">
                            {row.produtor}
                          </td>
                          {!isPrevista && (
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300">
                            {row.motivo || '-'}
                          </td>
                          )}
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300">
                            {row.motorista}
                          </td>
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300">
                            {row.placa}
                          </td>
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300">
                            <div className="leading-tight">
                              <div>Prev: {row.horarioPrevisto}</div>
                              {row.statusType === 'coleta-prevista' ? (() => {
                                const waitMin = getMinutesWaiting(row.rawEvent?.expected_arrival);
                                return <div className={getWaitingColorClass(waitMin)}>{formatWaitingTime(waitMin)}</div>;
                              })() : (
                                <div className="text-emerald-600 dark:text-emerald-400">Justificado: {row.horarioRealizado}</div>
                              )}
                            </div>
                          </td>
                          <td className="p-2 border-r border-slate-100 dark:border-slate-800 text-slate-700 dark:text-slate-300">
                            {row.operacao}
                          </td>
                          {!isPrevista && (
                          <td className="p-2 text-slate-700 dark:text-slate-300">
                            <div className="flex items-center gap-2 flex-wrap">
                              <span className={`inline-flex items-center px-2 py-0.5 rounded text-[9px] font-bold ${getStatusBadgeClass(row.statusType)}`}>
                                {row.statusLabel}
                              </span>
                              {row.statusType === 'coleta-prevista' ? null : row.isAlreadyLaunched ? (
                                <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[9px] font-semibold bg-emerald-100 dark:bg-emerald-500/15 text-emerald-700 dark:text-emerald-300 border border-emerald-300 dark:border-emerald-500/30">
                                  <Check size={10} />
                                  Lançada
                                </span>
                              ) : (
                                <button
                                  type="button"
                                  onClick={() => void handleLaunchNonCollection(row)}
                                  disabled={addingRowKeys.has(row.key)}
                                  className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[9px] font-bold bg-sky-100 dark:bg-sky-500/20 text-sky-700 dark:text-sky-200 border border-sky-300 dark:border-sky-500/40 hover:bg-sky-200 dark:hover:bg-sky-500/30 disabled:opacity-60 disabled:cursor-not-allowed"
                                  title="Adicionar na tabela de não coletas"
                                >
                                  {addingRowKeys.has(row.key) ? <Loader2 size={10} className="animate-spin" /> : <Plus size={10} />}
                                  {addingRowKeys.has(row.key) ? 'Lançando' : 'Lançar'}
                                </button>
                              )}
                            </div>
                          </td>
                          )}
                        </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              {(result.failedPlantIds.length > 0 || result.failedRouteIds.length > 0) && (
                <div className="m-5 rounded-xl border border-amber-700/40 bg-amber-950/30 p-4 text-xs font-semibold text-amber-200 space-y-2">
                  {result.failedPlantIds.map((item) => (
                    <div key={`fail-plant-${item.plantId}`}>
                      plant_id={item.plantId}: {item.error}{item.status ? ` (status ${item.status})` : ''}
                    </div>
                  ))}
                  {result.failedRouteIds.map((item) => (
                    <div key={`fail-route-${item.routeId}`}>
                      route_id={item.routeId}: {item.error}{item.status ? ` (status ${item.status})` : ''}
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </section>
      </div>
    </div>
  );
};

export default RouteWebLabView;

