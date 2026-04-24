
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteDeparture, User, RouteOperationMapping, RouteConfig } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import * as XLSX from 'xlsx';
import { getBrazilDate, getBrazilHours, getBrazilMinutes, toBrazilDate, getWeekString, getRouteDateForCurrentTime } from '../utils/dateUtils';
import { isDealeUser, getDealeFilteredConfigs, getDealeAnchorOperation, getDealeRealOperations } from '../utils/dealeUtils';
import {
  Clock, X, Loader2, RefreshCw, ShieldCheck,
  CheckCircle2, ChevronDown,
  Filter, Search, CheckSquare, Square,
  ChevronRight, Maximize2, Minimize2,
  Archive, Database, Save, LinkIcon,
  Layers, Trash2, Settings2, Check, Table, SortAsc,
  Sun, Moon, AlertTriangle, Calendar, ArrowUpDown
} from 'lucide-react';

const MOTIVOS = [
  'Fábrica', 'Logística', 'Mão de obra', 'Manutenção', 'Divergência de Roteirização', 'Solicitado pelo Cliente', 'Infraestrutura'
];

const OBSERVATION_TEMPLATES: Record<string, string[]> = {
  'Fábrica': ["Atraso na descarga | Entrada **:**h - Saída **:**h"],
  'Logística': ["Atraso no lavador | Chegada da rota anterior às **:**h - Entrada na fábrica às **:**h", "Motorista adiantou a rota devido à desvios", "Atraso na rota anterior (nome da rota)", "Atraso na rota anterior | Chegada no lavador **:**h - Entrada na fábrica às **:**h", "Falta de material de coleta para realizar a rota"],
  'Mão de obra': ["Atraso do motorista", "Adiantamento do motorista", "A rota iniciou atrasada devido à interjornada do motorista | Atrasou na rota anterior devido à", "Troca do motorista previsto devido à saúde"],
  'Manutenção': ["Precisou realizar a troca de pneus | Início **:**h - Término **:**h", "Troca de mola | Início **:**h - Término **:**h", "Manutenção na parte elétrica | Início **:**h - Término **:**h", "Manutenção na parte elétrica | Início **:**h - Término **:**h", "Manutenção nos freios | Início **:**h - Término **:**h", "Manutenção na bomba de carregamento de leite | Início **:**h - Término **:**h"],
  'Divergência de Roteirização': ["Horário de saída da rota não atende os produtores", "Horário de saída da rota precisa ser alterado devido à entrada de produtores"],
  'Solicitado pelo Cliente': ["Rota saiu adiantada para realizar socorro", "Cliente solicitou para a rota sair adiantada"],
  'Infraestrutura': []
};

const CAUSAS_RAIZ_MANUTENCAO = [
  'Ar condicionado',
  'Cabo do cambio',
  'Cambio',
  'Transmissão',
  'Embreagem',
  'Retentor da Caixa Cambio',
  'Trinca no Chassi',
  'Tirante do Eixo',
  'Ponta de eixo',
  'Bomba de coleta',
  'Multiplicadora',
  'Cardan da bomba',
  'Tomada de força',
  'Cruzeta Bomba de Coleta',
  'Geladeira do caminhão',
  'Rastreador e camenras',
  'Cruzeta Diferencial',
  'Cardan do Diferencial',
  'Rolamento do cardã',
  'Diferencial',
  'Alternador',
  'Lanterna traseira',
  'Farol',
  'Correia do alternador',
  'Vidro Eletrico',
  'Sensor do filtro racor',
  'Motor de partida',
  'Fusiveis',
  'Fusiveis : Farois / Ventuinha',
  'Perca de força',
  'Luz vermelha no Painel',
  'Freio Estacionario',
  'Cuica de Freio',
  'Mangueira de ar Freio',
  'Vazamento Cuica',
  'Retentor de Cubo',
  'Borrachas do tanque',
  'DSPL',
  'Vazamento Tanque de leite',
  'Caminhão esquentando',
  'Vazamento de Agua',
  'Radiador',
  'Valcula termostática',
  'Motor Aquecendo',
  'Motor Vazamento',
  'Regeneração (DPF)',
  'Retentor do Motor',
  'Pneu furado',
  'Pneu Careca',
  'Revisão Atrasada',
  'Lubrificação',
  'Revisão do Cubo',
  'Revisão do Freio',
  'Valvula de engate',
  'Engate do reboque',
  'Mangueira de ar Suspensor',
  'Suspensor do Truck',
  'Suspensão',
  'Molas',
  'Balão do Truck',
  'Pino de centro',
  'Mola Tensora',
  'Barra de direção',
  'Levante do Trcuk',
  'Bolsa de ar Suspensor do truck',
  'Atrasos por manutenção anterior'
];

const normalizeCauseText = (value: string): string =>
  String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();

type ArchiveDivergentRoute = {
  route: RouteDeparture;
  missingFields: string[];
};

type ArchiveHardBlockType = 'resumo' | 'status';
type ArchiveStatusDivergence = {
  operacao: string;
  status: string;
};

const FilterDropdown = ({ col, routes, colFilters, setColFilters, selectedFilters, setSelectedFilters, onClose, dropdownRef }: any) => {
    // Mapeia o nome da coluna para o campo real no objeto RouteDeparture
    const fieldMap: Record<string, string> = {
        'geral': 'statusGeral',
        'status': 'statusOp',
        'observacao': 'observacao',
        'motivo': 'motivo',
        'operacao': 'operacao',
        'tempo': 'tempo',
        'rota': 'rota',
        'data': 'data',
        'inicio': 'inicio',
        'motorista': 'motorista',
        'placa': 'placa',
        'saida': 'saida'
    };

    const fieldName = fieldMap[col] || col;
    const values: string[] = Array.from(new Set(routes.map((r: any) => String(r[fieldName] || "")))).sort() as string[];
    const selected = (selectedFilters[col] as string[]) || [];
    const toggleValue = (val: string) => { const next = selected.includes(val) ? selected.filter(v => v !== val) : [...selected, val]; setSelectedFilters({ ...selectedFilters, [col]: next }); };
    return (
        <div ref={dropdownRef} className="absolute top-10 left-0 z-[100] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xl rounded-xl w-64 p-3 text-slate-700 dark:text-slate-300 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-center gap-2 mb-3 p-2 bg-slate-50 dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700">
                <Search size={14} className="text-slate-400" />
                <input type="text" placeholder="Filtrar..." autoFocus value={colFilters[col] || ""} onChange={e => setColFilters({ ...colFilters, [col]: e.target.value })} className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800 dark:text-white" />
            </div>
            <div className="max-h-56 overflow-y-auto space-y-1 scrollbar-thin border-t border-slate-100 dark:border-slate-700 py-2">
                {values.filter(v => v.toLowerCase().includes((colFilters[col] || "").toLowerCase())).map(v => (
                    <div key={v} onClick={() => toggleValue(v)} className="flex items-center gap-2 p-2 hover:bg-slate-50 dark:hover:bg-slate-700 rounded-lg cursor-pointer transition-all">
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} className="text-slate-300" />}
                        <span className="text-[10px] font-bold uppercase truncate">{v || "(VAZIO)"}</span>
                    </div>
                ))}
            </div>
            <button onClick={() => { setColFilters({ ...colFilters, [col]: "" }); setSelectedFilters({ ...selectedFilters, [col]: [] }); onClose(); }} className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-600 bg-red-50 dark:bg-red-900/30 hover:bg-red-100 rounded-lg border border-red-100 dark:border-red-900/50 transition-colors"> Limpar Filtro </button>
        </div>
    );
};

// Componente de filtro para o modal de Histórico (igual ao da tabela principal)
const HistoryFilterDropdown = ({ col, values, colFilter, setColFilter, selected, setSelected, onClose, dropdownRef, anchorEl }: any) => {
    const toggleValue = (val: string) => {
        const next = selected.includes(val) ? selected.filter((v: string) => v !== val) : [...selected, val];
        setSelected(next);
    };

    // Calcula posição fixed baseada no elemento âncora
    const [position, setPosition] = React.useState<{ top: number; left: number } | null>(null);

    React.useEffect(() => {
        if (anchorEl) {
            const rect = anchorEl.getBoundingClientRect();
            setPosition({
                top: rect.bottom + 4,
                left: rect.left
            });
        }
    }, [anchorEl]);

    if (!position) return null;

    return (
        <div
            ref={dropdownRef}
            className="fixed z-[9999] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-2xl rounded-xl w-64 animate-in fade-in zoom-in-95 duration-150"
            style={{
                top: `${position.top}px`,
                left: `${position.left}px`,
                maxHeight: 'min(450px, calc(100vh - 100px))'
            }}
        >
            <div className="flex items-center gap-2 p-3 pb-2 bg-slate-50 dark:bg-slate-900 rounded-t-xl border-b border-slate-200 dark:border-slate-700">
                <Search size={14} className="text-slate-400 shrink-0" />
                <input type="text" placeholder="Filtrar..." autoFocus value={colFilter || ""} onChange={e => setColFilter(e.target.value)} className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800 dark:text-white" />
            </div>
            <div className="overflow-y-auto p-2 space-y-1 scrollbar-thin" style={{ minHeight: '180px', maxHeight: 'min(400px, calc(100vh - 180px))' }}>
                {values.filter((v: string) => v.toLowerCase().includes((colFilter || "").toLowerCase())).map((v: string) => (
                    <div key={v} onClick={() => toggleValue(v)} className="flex items-center gap-2 p-2.5 hover:bg-slate-50 dark:hover:bg-slate-700 rounded-lg cursor-pointer transition-all">
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-600 shrink-0" /> : <Square size={14} className="text-slate-300 shrink-0" />}
                        <span className="text-[10px] font-bold uppercase truncate text-slate-700 dark:text-slate-300">{v || "(VAZIO)"}</span>
                    </div>
                ))}
            </div>
            <div className="p-2 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50 rounded-b-xl">
                <button onClick={() => { setColFilter(""); setSelected([]); onClose(); }} className="w-full py-2.5 text-[10px] font-black uppercase text-red-600 bg-red-50 dark:bg-red-900/30 hover:bg-red-100 dark:hover:bg-red-900/50 rounded-lg border border-red-100 dark:border-red-900/50 transition-colors"> Limpar Filtro </button>
            </div>
        </div>
    );
};

// Componente de Input de Emails com Pills (altura automática baseada no conteúdo)
interface EmailInputProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  onRemoveEmail: (email: string) => void;
  placeholder?: string;
}

const EmailInput: React.FC<EmailInputProps> = ({
  label,
  value,
  onChange,
  onRemoveEmail,
  placeholder = "Cole emails em massa...",
}) => {
  const emails = value ? value.split(';').map(e => e.trim()).filter(e => e.length > 0) : [];
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handlePaste = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    e.preventDefault();
    const pastedText = e.clipboardData.getData('text');
    const newEmails = pastedText.split(/[;\s,\n]+/).map(e => e.trim()).filter(e => e.length > 0);
    const existing = emails;
    const combined = [...existing, ...newEmails.filter(e => !existing.includes(e))];
    onChange(combined.join(';'));
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    // Adiciona email ao pressionar Enter, Espaço ou Vírgula
    if (e.key === 'Enter' || e.key === ',' || e.key === ' ') {
      e.preventDefault();
      const input = e.currentTarget;
      const currentText = input.value.trim();
      if (currentText && !emails.includes(currentText)) {
        onChange([...emails, currentText].join(';'));
        input.value = '';
      }
    }
  };

  return (
    <div className="space-y-2">
      <div className="flex items-center justify-between">
        <label className="text-[10px] font-black uppercase text-slate-400">{label}</label>
        <span className="text-[9px] font-bold text-slate-500 dark:text-slate-400">
          {emails.length} email(s)
        </span>
      </div>
      <div
        className="w-full border border-slate-200 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-900 shadow-sm focus-within:ring-2 focus-within:ring-blue-500 transition-all"
      >
        {/* Pills container */}
        <div className="p-3 min-h-[60px] max-h-[200px] overflow-y-auto scrollbar-thin">
          <div className="flex flex-wrap gap-2">
            {emails.map((email, index) => (
              <span
                key={index}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300 rounded-full text-[10px] font-bold uppercase tracking-wide border border-blue-200 dark:border-blue-800 flex-shrink-0"
              >
                {email}
                <button
                  onClick={() => onRemoveEmail(email)}
                  className="hover:bg-blue-200 dark:hover:bg-blue-800 rounded-full p-0.5 transition-colors"
                  title="Remover email"
                >
                  <X size={12} />
                </button>
              </span>
            ))}
          </div>
        </div>

        {/* Input area */}
        <div className="p-2 border-t border-slate-100 dark:border-slate-800 bg-slate-50 dark:bg-slate-900/50 rounded-b-2xl">
          <textarea
            ref={textareaRef}
            onPaste={handlePaste}
            onKeyDown={handleKeyDown}
            onChange={(e) => {
              const newVal = e.target.value.trim();
              if (newVal === '' || newVal.endsWith(';')) {
                const newEmail = newVal.replace(/;/g, '').trim();
                if (newEmail && !emails.includes(newEmail)) {
                  onChange([...emails, newEmail].join(';'));
                }
                e.target.value = '';
              }
            }}
            placeholder={placeholder}
            rows={1}
            className="w-full bg-transparent outline-none text-[11px] font-bold text-slate-700 dark:text-slate-300 placeholder-slate-400 resize-none"
          />
        </div>
      </div>
    </div>
  );
};

const RouteDepartureView: React.FC<{
  currentUser: User;
  isConfigModalOpen?: boolean;
  setIsConfigModalOpen?: (open: boolean) => void;
}> = ({ currentUser, isConfigModalOpen = false, setIsConfigModalOpen = () => {} }) => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [routeMappings, setRouteMappings] = useState<RouteOperationMapping[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [currentTime, setCurrentTime] = useState(new Date());
  const [bulkStatus, setBulkStatus] = useState<{ active: boolean, current: number, total: number } | null>(null);
  const [pendingBulkRoutes, setPendingBulkRoutes] = useState<string[]>([]);
  const [isBulkMappingModalOpen, setIsBulkMappingModalOpen] = useState(false);
  const [isDarkMode, setIsDarkMode] = useState(() => {
    const saved = localStorage.getItem('route_departure_dark_mode');
    return saved !== 'false'; // Default: true (dark mode)
  });

  // Estado para usuários DEALE (usado apenas no modal de configurar emails)
  const [isDeale, setIsDeale] = useState(false);

  const [ghostRow, setGhostRow] = useState<Partial<RouteDeparture>>({
    id: 'ghost', rota: '', data: getRouteDateForCurrentTime(), inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '', semana: '', causaRaiz: ''
  });

  // Controla qual célula da coluna SAÍDA está sendo editada (para mostrar valor completo com data)
  const [editingSaidaCell, setEditingSaidaCell] = useState<string | null>(null);

  // Armazena os últimos checklists de motorista por operação
  const [lastMotoristaChecklist, setLastMotoristaChecklist] = useState<Record<string, { data: string, porcentagem: string }>>({});

  const [isHistoryModalOpen, setIsHistoryModalOpen] = useState(false);
  const [isMappingModalOpen, setIsMappingModalOpen] = useState(false);
  const [pendingMappingRoute, setPendingMappingRoute] = useState<string | null>(null);

  // Estado para o popup de edição de horários
  const [isTimeEditModalOpen, setIsTimeEditModalOpen] = useState(false);
  const [timeEditData, setTimeEditData] = useState<{ routeId: string; template: string; startTime: string; endTime: string } | null>(null);

  // Estado para o modal de configuração de emails
  const [isEmailConfigModalOpen, setIsEmailConfigModalOpen] = useState(false);
  const [selectedOperacaoConfig, setSelectedOperacaoConfig] = useState<string>('');
  const [configEnvio, setConfigEnvio] = useState<string>('');
  const [configCopia, setConfigCopia] = useState<string>('');
  
  // Ref para rastrear se as configs de email já foram carregadas manualmente
  // Evita que o polling sobrescreva o que o usuário está digitando
  const emailConfigLoadedRef = useRef<Record<string, boolean>>({});

  // Estado para filtro de email - busca email em todas as operações
  const [emailFilter, setEmailFilter] = useState<string>('');
  const [isSavingConfig, setIsSavingConfig] = useState(false);

  // Funções auxiliares para manipular emails
  const parseEmails = (text: string): string[] => {
    if (!text || text.trim() === '') return [];
    // Separa por ;, vírgula, espaço ou newline, e filtra vazios
    return text.split(/[;\s,\n]+/).map(e => e.trim()).filter(e => e.length > 0);
  };

  const formatEmails = (emails: string[]): string => {
    return emails.join(';');
  };

  const addEmails = (currentText: string, newText: string): string => {
    const existing = parseEmails(currentText);
    const newEmails = parseEmails(newText);
    // Adiciona apenas emails que ainda não existem
    const combined = [...existing, ...newEmails.filter(e => !existing.includes(e))];
    return formatEmails(combined);
  };

  const removeEmail = (currentText: string, emailToRemove: string): string => {
    const emails = parseEmails(currentText);
    const filtered = emails.filter(e => e !== emailToRemove);
    return formatEmails(filtered);
  };

  // Busca um email em todas as operações (campos Envio e Copia)
  const findEmailInAllOperations = (email: string): { operacao: string; nomeExibicao: string; campo: 'Envio' | 'Copia' }[] => {
    if (!email.trim()) return [];
    const results: { operacao: string; nomeExibicao: string; campo: 'Envio' | 'Copia' }[] = [];
    userConfigs.forEach(config => {
      const emailsEnvio = parseEmails(config.Envio || '');
      const emailsCopia = parseEmails(config.Copia || '');
      if (emailsEnvio.includes(email.trim().toLowerCase()) || emailsEnvio.includes(email.trim())) {
        results.push({ operacao: config.operacao, nomeExibicao: config.nomeExibicao, campo: 'Envio' });
      }
      if (emailsCopia.includes(email.trim().toLowerCase()) || emailsCopia.includes(email.trim())) {
        results.push({ operacao: config.operacao, nomeExibicao: config.nomeExibicao, campo: 'Copia' });
      }
    });
    return results;
  };

  // Remove um email de uma operação específica (Envio ou Copia)
  const removeEmailFromOperation = async (operacao: string, campo: 'Envio' | 'Copia', email: string) => {
    const token = await getValidToken();
    if (!token) return;

    try {
      // Se a operação for "DEALE", usa a operação âncora (ALMIRANTE)
      const operacaoReal = (operacao === 'DEALE') ? getDealeAnchorOperation() : operacao;

      const config = userConfigs.find(c => c.operacao === operacaoReal);
      if (!config) return;

      const currentValue = campo === 'Envio' ? (config.Envio || '') : (config.Copia || '');
      const newValue = removeEmail(currentValue, email);

      await SharePointService.updateRouteConfigEmails(
        token,
        operacaoReal,
        campo === 'Envio' ? newValue : (config.Envio || ''),
        campo === 'Copia' ? newValue : (config.Copia || '')
      );

      // Atualiza estado local
      setUserConfigs(prev => prev.map(c =>
        c.operacao === operacaoReal
          ? { ...c, [campo]: newValue }
          : c
      ));

      // Se a operação atual do modal foi afetada, atualiza os campos
      if (operacao === selectedOperacaoConfig) {
        if (campo === 'Envio') setConfigEnvio(newValue);
        else setConfigCopia(newValue);
      }
    } catch (err: any) {
      console.error(`[EMAIL_FILTER] Erro ao remover email de ${operacao} (${campo}):`, err.message);
    }
  };

  // Estado para o popup de edição do checklist (GERAL)
  const [isChecklistEditModalOpen, setIsChecklistEditModalOpen] = useState(false);
  const [checklistEditData, setChecklistEditData] = useState<{ routeId: string; data: string; porcentagem: string; motivos: string; causaRaiz: string } | null>(null);

  const [histStart, setHistStart] = useState(getBrazilDate());
  const [histEnd, setHistEnd] = useState(getBrazilDate());
  const [archivedResults, setArchivedResults] = useState<RouteDeparture[]>([]);
  const [isSearchingArchive, setIsSearchingArchive] = useState(false);
  const [isHistoryFullscreen, setIsHistoryFullscreen] = useState(false);
  const archiveAbortRef = useRef<AbortController | null>(null);
  
  // Estado para edição em lote - armazena alterações pendentes
  const [pendingHistoryEdits, setPendingHistoryEdits] = useState<Record<string, Partial<RouteDeparture>>>({});
  const [editingHistoryId, setEditingHistoryId] = useState<string | null>(null);
  const [editingHistoryField, setEditingHistoryField] = useState<keyof RouteDeparture | null>(null);

  // Estados para modal de adicionar rota
  const [isAddRouteModalOpen, setIsAddRouteModalOpen] = useState(false);
  const [isAddingRoute, setIsAddingRoute] = useState(false);
  const [newRouteData, setNewRouteData] = useState<{ rota: string; inicio: string; motorista: string; placa: string; operacao: string }>({
    rota: '',
    inicio: '',
    motorista: '',
    placa: '',
    operacao: ''
  });

  // Estado para alertas de rotas com histórico de problemas
  const [routeAlerts, setRouteAlerts] = useState<Record<string, { count: number; history: RouteDeparture[] }>>({});
  const [selectedRouteAlert, setSelectedRouteAlert] = useState<{ rota: string; history: RouteDeparture[] } | null>(null);

  // Estado para alertas de motoristas com atrasos recorrentes por "Mão de obra"
  const [motoristAlerts, setMotoristAlerts] = useState<Record<string, { count: number; history: RouteDeparture[] }>>({});
  const [selectedMotoristAlert, setSelectedMotoristAlert] = useState<{ motorista: string; count: number; history: RouteDeparture[] } | null>(null);

  // Estado para modal de aviso quando tentar adicionar rota com filtros/ordenação ativos
  const [isFilterBlockModalOpen, setIsFilterBlockModalOpen] = useState(false);
  const [filterBlockReason, setFilterBlockReason] = useState<'single' | 'bulk' | 'ghost'>('ghost');
  const [isArchiveValidationModalOpen, setIsArchiveValidationModalOpen] = useState(false);
  const [archiveReadyRoutes, setArchiveReadyRoutes] = useState<RouteDeparture[]>([]);
  const [archiveDivergentRoutes, setArchiveDivergentRoutes] = useState<ArchiveDivergentRoute[]>([]);
  const [archiveHardBlockModal, setArchiveHardBlockModal] = useState<ArchiveHardBlockType | null>(null);
  const [archiveStatusDivergences, setArchiveStatusDivergences] = useState<ArchiveStatusDivergence[]>([]);
  
  // Estados para filtros do histórico
  const [historyColFilters, setHistoryColFilters] = useState<Record<string, string[]>>({});
  const [historySelectedFilters, setHistorySelectedFilters] = useState<Record<string, string[]>>({});
  const [historyActiveFilterCol, setHistoryActiveFilterCol] = useState<string | null>(null);
  const [historyFilterAnchorEl, setHistoryFilterAnchorEl] = useState<HTMLElement | null>(null);
  const historyFilterDropdownRef = useRef<HTMLDivElement>(null);
  const [historySortByOperacao, setHistorySortByOperacao] = useState(false);
  const [isSortByOperacao, setIsSortByOperacao] = useState(false);

  const [activeObsId, setActiveObsId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>(() => {
    const saved = sessionStorage.getItem('route_departure_col_filters');
    if (saved) {
        console.log('[ROUTE_DEPARTURE] Filtros de coluna restaurados:', JSON.parse(saved));
        return JSON.parse(saved);
    }
    return {};
  });
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>(() => {
    const saved = sessionStorage.getItem('route_departure_selected_filters');
    if (saved) {
        console.log('[ROUTE_DEPARTURE] Filtros selecionados restaurados:', JSON.parse(saved));
        return JSON.parse(saved);
    }
    return {};
  });
  const [isSortByTimeEnabled, setIsSortByTimeEnabled] = useState(() => {
    const saved = sessionStorage.getItem('route_departure_sort_by_time');
    if (saved) {
        console.log('[ROUTE_DEPARTURE] Ordenação por horário restaurada:', JSON.parse(saved));
        return JSON.parse(saved);
    }
    return true; // Padrão: ativado ao abrir a tela
  });
  const [colWidths, setColWidths] = useState<Record<string, number>>(() => {
    const saved = sessionStorage.getItem('route_departure_col_widths');
    if (saved) {
        console.log('[ROUTE_DEPARTURE] Larguras das colunas restauradas:', JSON.parse(saved));
        return JSON.parse(saved);
    }
    return { rota: 140, data: 125, inicio: 95, motorista: 230, placa: 100, saida: 95, motivo: 170, observacao: 400, geral: 120, operacao: 140, status: 90, tempo: 90 };
  });
  const [hiddenColumns, setHiddenColumns] = useState<Set<string>>(() => {
    const saved = sessionStorage.getItem('route_departure_hidden_cols');
    if (saved) {
        console.log('[ROUTE_DEPARTURE] Colunas ocultas restauradas:', new Set(JSON.parse(saved)));
        return new Set(JSON.parse(saved));
    }
    return new Set();
  });
  const [contextMenu, setContextMenu] = useState<{ visible: boolean; x: number; y: number; col: string | null }>({ visible: false, x: 0, y: 0, col: null });
  const [checklistTooltip, setChecklistTooltip] = useState<{ routeId: string; content: string } | null>(null);
  const [copiedGeralStatus, setCopiedGeralStatus] = useState<string | null>(null);
  const [hoveredGeralCell, setHoveredGeralCell] = useState<string | null>(null);

  // ⚠️ Estados para envio automático REMOVIDOS — envio agora é feito apenas pela tela "Resumo"
  // const [pendingSendOps, setPendingSendOps] = useState<Set<string>>(new Set());
  // const [countdowns, setCountdowns] = useState<Record<string, number>>({});
  // const [sendingOps, setSendingOps] = useState<Set<string>>(new Set());
  // const countdownTimersRef = useRef<Record<string, NodeJS.Timeout>>({});
  // const sentTodayRef = useRef<Set<string>>(new Set());
  // const blockedUntilRef = useRef<Record<string, number>>({});

  const obsDropdownRef = useRef<HTMLDivElement>(null);
  const obsTextareaRefs = useRef<Record<string, HTMLTextAreaElement>>({});
  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);
  const filterDropdownRef = useRef<HTMLDivElement>(null);
  const contextMenuRef = useRef<HTMLDivElement>(null);
  const tooltipTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  const getAccessToken = async (): Promise<string> => {
    // Tenta sempre obter o token mais fresco via MSAL (renova silenciosamente se necessário)
    const freshToken = await getValidToken();
    if (freshToken) return freshToken;
    // Fallback para o token em memória (pode estar próximo de expirar, mas evita quebrar a operação)
    const fallback = currentUser?.accessToken || (window as any).__access_token;
    if (fallback) return fallback;
    throw new Error('Sessão expirada. Por favor, renove sua sessão.');
  };


  // Analisa histórico dos últimos 7 dias e identifica rotas com problemas
  const analyzeRouteHistory = async (token: string) => {
    try {
      console.log('[ROUTE_ALERT] Analisando histórico dos últimos 7 dias...');

      // Calcula data de 7 dias atrás
      const sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
      const startDate = sevenDaysAgo.toISOString().split('T')[0];
      const endDate = getBrazilDate();

      // Busca histórico no SharePoint
      const history = await SharePointService.getArchivedDepartures(token, null, startDate, endDate);

      console.log(`[ROUTE_ALERT] ${history.length} registros encontrados, filtrando status Atrasada/Adiantada...`);
      
      // Log para depuração - verifica statusOp dos registros
      const statusCounts: Record<string, number> = {};
      history.forEach(r => {
        statusCounts[r.statusOp] = (statusCounts[r.statusOp] || 0) + 1;
      });
      console.log('[ROUTE_ALERT] Status encontrados:', statusCounts);

      // Filtra apenas rotas com status "Atrasada/Atrasado" ou "Adiantada/Adiantado"
      const problemRoutes = history.filter(r =>
        r.statusOp === 'Atrasada' || r.statusOp === 'Atrasado' ||
        r.statusOp === 'Adiantada' || r.statusOp === 'Adiantado'
      );

      console.log(`[ROUTE_ALERT] ${problemRoutes.length} registros com problemas`);

      // Agrupa por nome de rota
      const alerts: Record<string, { count: number; history: RouteDeparture[] }> = {};

      problemRoutes.forEach(route => {
        const rotaNome = route.rota;
        if (!alerts[rotaNome]) {
          alerts[rotaNome] = { count: 0, history: [] };
        }
        alerts[rotaNome].count++;
        alerts[rotaNome].history.push(route);
      });

      // Ordena histórico por data (mais recente primeiro)
      Object.keys(alerts).forEach(rota => {
        alerts[rota].history.sort((a, b) =>
          new Date(b.data).getTime() - new Date(a.data).getTime()
        );
      });

      setRouteAlerts(alerts);
      console.log(`[ROUTE_ALERT] ✅ ${Object.keys(alerts).length} rotas com alertas de problemas`);
      
      // Log para depuração - mostra primeiras 5 rotas com alertas
      const first5Routes = Object.keys(alerts).slice(0, 5);
      first5Routes.forEach(rota => {
        console.log(`[ROUTE_ALERT] Rota: ${rota} -> ${alerts[rota].count} ocorrências`);
      });
    } catch (e: any) {
      console.error('[ROUTE_ALERT] Erro ao analisar histórico:', e.message);
    }
  };

  // Analisa histórico dos últimos 30 dias e identifica motoristas com atrasos recorrentes por "Mão de obra"
  const analyzeMotoristHistory = async (motoristaNome: string, token: string): Promise<{ count: number; history: RouteDeparture[] } | null> => {
    try {
      if (!motoristaNome || motoristaNome.trim() === '') return null;

      console.log(`[MOTORIST_ALERT] Analisando histórico de "${motoristaNome}" nos últimos 30 dias...`);

      // Calcula data de 30 dias atrás
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      const startDate = thirtyDaysAgo.toISOString().split('T')[0];
      const endDate = getBrazilDate();

      // Busca histórico no SharePoint (sem filtrar por operação para pegar todas as ocorrências do motorista)
      const history = await SharePointService.getArchivedDepartures(token, null, startDate, endDate);

      // Filtra apenas registros com motivo "Mão de obra" E o motorista específico
      const maodeObraRecords = history.filter(r =>
        r.motivo === 'Mão de obra' &&
        r.motorista &&
        r.motorista.toLowerCase().trim() === motoristaNome.toLowerCase().trim()
      );

      console.log(`[MOTORIST_ALERT] ${maodeObraRecords.length} ocorrência(s) de "Mão de obra" para ${motoristaNome}`);

      if (maodeObraRecords.length === 0) return null;

      // Ordena por data (mais recente primeiro)
      maodeObraRecords.sort((a, b) => new Date(b.data).getTime() - new Date(a.data).getTime());

      return { count: maodeObraRecords.length, history: maodeObraRecords };
    } catch (e: any) {
      console.error(`[MOTORIST_ALERT] Erro ao analisar histórico de "${motoristaNome}":`, e.message);
      return null;
    }
  };

  // Reavalia alerta de motorista para uma linha específica (chamado quando o motivo muda)
  const reevaluateMotoristAlert = async (routeId: string, motoristaNome: string, motivoAtual: string) => {
    // Se o motivo NÃO é "Mão de obra", remove o alerta deste motorista se existir
    if (motivoAtual !== 'Mão de obra') {
      setMotoristAlerts(prev => {
        const next = { ...prev };
        delete next[motoristaNome];
        return next;
      });
      return;
    }

    // Se o motivo É "Mão de obra", busca no histórico
    const token = await getValidToken();
    if (!token) return;

    const result = await analyzeMotoristHistory(motoristaNome, token);
    setMotoristAlerts(prev => {
      const next = { ...prev };
      if (result && result.count > 0) {
        next[motoristaNome] = result;
      } else {
        delete next[motoristaNome];
      }
      return next;
    });
  };

  // Escaneia todas as rotas da tabela principal e carrega alertas de motoristas com "Mão de obra"
  const scanExistingMotoristAlerts = async (token: string, routesToScan?: RouteDeparture[]) => {
    try {
      const routesArray = routesToScan || routes;

      // Coleta motoristas únicos que têm motivo "Mão de obra" na tabela principal
      const motoristasWithMaDeObra = new Set<string>();
      routesArray.forEach(r => {
        if (r.motivo === 'Mão de obra' && r.motorista && r.motorista.trim() !== '') {
          motoristasWithMaDeObra.add(r.motorista.trim());
        }
      });

      if (motoristasWithMaDeObra.size === 0) {
        setMotoristAlerts({});
        return;
      }

      console.log(`[MOTORIST_SCAN] ${motoristasWithMaDeObra.size} motorista(s) com "Mão de obra" na tabela:`, Array.from(motoristasWithMaDeObra));

      // Busca histórico dos últimos 30 dias UMA VEZ (compartilhado entre todos)
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
      const startDate = thirtyDaysAgo.toISOString().split('T')[0];
      const endDate = getBrazilDate();

      const history = await SharePointService.getArchivedDepartures(token, null, startDate, endDate);

      // Filtra apenas "Mão de obra"
      const maodeObraHistory = history.filter(r => r.motivo === 'Mão de obra' && r.motorista && r.motorista.trim() !== '');

      // Agrupa por motorista
      const alerts: Record<string, { count: number; history: RouteDeparture[] }> = {};
      motoristasWithMaDeObra.forEach(motorista => {
        const records = maodeObraHistory.filter(r =>
          r.motorista.toLowerCase().trim() === motorista.toLowerCase().trim()
        );
        if (records.length > 0) {
          records.sort((a, b) => new Date(b.data).getTime() - new Date(a.data).getTime());
          alerts[motorista] = { count: records.length, history: records };
        }
      });

      setMotoristAlerts(alerts);
      console.log(`[MOTORIST_SCAN] ✅ ${Object.keys(alerts).length} motorista(s) com alertas`);
    } catch (e: any) {
      console.error('[MOTORIST_SCAN] Erro ao escanear alertas:', e.message);
    }
  };

  // Sistema de Lock Temporário para Edição de Linhas
  const LOCK_TIMEOUT = 30 * 1000; // 30 segundos

  /**
   * Tenta adquirir lock para editar uma linha
   * Retorna true se conseguiu, false se outra pessoa está editando
   */
  const tryAcquireLock = (routeId: string): boolean => {
    const route = routes.find(r => r.id === routeId);
    if (!route) return false;

    const now = Date.now();
    
    // Verifica se já tem lock válido de outro usuário
    if (route.editingUser && route.lockExpiresAt && now < route.lockExpiresAt) {
      if (route.editingUser !== currentUser.email) {
        console.warn(`[LOCK_BLOCKED] Linha ${routeId} está sendo editada por ${route.editingUser}`);
        return false;
      }
    }

    // Adquire o lock (ou renova se já era do usuário atual)
    setRoutes(prev => prev.map(r => {
      if (r.id === routeId) {
        return {
          ...r,
          editingUser: currentUser.email,
          lockExpiresAt: now + LOCK_TIMEOUT
        };
      }
      return r;
    }));

    return true;
  };

  /**
   * Libera o lock de uma linha
   */
  const releaseLock = (routeId: string) => {
    setRoutes(prev => prev.map(r => {
      if (r.id === routeId && r.editingUser === currentUser.email) {
        return { ...r, editingUser: undefined, lockExpiresAt: undefined };
      }
      return r;
    }));
  };

  /**
   * Libera todos os locks do usuário atual (ao sair da tela ou desmontar)
   */
  const releaseAllLocks = () => {
    setRoutes(prev => prev.map(r => {
      if (r.editingUser === currentUser.email) {
        return { ...r, editingUser: undefined, lockExpiresAt: undefined };
      }
      return r;
    }));
  };

  // ⚠️ checkOperationAllOK REMOVIDA — só era usada pelo envio automático desabilitado
  // const checkOperationAllOK = (operacao: string): boolean => { ... };

  // ⚠️ Funções de envio automática REMOVIDAS — envio agora é feito apenas pela tela "Resumo"
  // startSendCountdown, cancelSendCountdown, cleanupSendState, handleSendStatus

  // Limpa timers ao desmontar (removido pois countdowns não existem mais)
  // useEffect(() => {
  //   return () => {
  //     Object.values(countdownTimersRef.current).forEach(timer => clearInterval(timer));
  //   };
  // }, []);

  // Sincroniza com modal de configuração vindo do App.tsx
  useEffect(() => {
    if (isConfigModalOpen && userConfigs.length > 0) {
      setIsEmailConfigModalOpen(true);

      // Para usuários DEALE, usa configs filtradas (agrupa como DEALE)
      const modalConfigs = isDeale ? getDealeFilteredConfigs(userConfigs) : userConfigs;

      // Seleciona a primeira operação por padrão e carrega configs APENAS se ainda não foi carregado
      if (!selectedOperacaoConfig) {
        const firstOp = modalConfigs[0].operacao;
        setSelectedOperacaoConfig(firstOp);
        
        // Só carrega se ainda não foi carregado para esta operação
        if (!emailConfigLoadedRef.current[firstOp]) {
          const firstConfig = modalConfigs[0];
          setConfigEnvio(firstConfig.Envio || '');
          setConfigCopia(firstConfig.Copia || '');
          emailConfigLoadedRef.current[firstOp] = true;
        }
      }
    }
    
    // Quando o modal fecha, reseta o flag para permitir recarregar na próxima abertura
    if (!isConfigModalOpen) {
      emailConfigLoadedRef.current = {};
      setSelectedOperacaoConfig('');
      setConfigEnvio('');
      setConfigCopia('');
    }
  }, [isConfigModalOpen, userConfigs, isDeale]);

  // Carrega dados da configuração quando seleciona operação
  // SÓ executa quando o usuário TROCA a operação selecionada, NÃO quando userConfigs muda por polling
  useEffect(() => {
    if (!selectedOperacaoConfig || !isEmailConfigModalOpen) return;
    if (userConfigs.length === 0) return;
    
    // Se já foi carregado para esta operação, ignora (evita sobrescrever durante edição)
    if (emailConfigLoadedRef.current[selectedOperacaoConfig]) return;

    // Determina qual operação buscar no userConfigs
    // Se for "DEALE", busca na operação âncora (ALMIRANTE)
    const operacaoParaBuscar = (isDeale && selectedOperacaoConfig === 'DEALE')
      ? getDealeAnchorOperation()
      : selectedOperacaoConfig;

    const config = userConfigs.find(c =>
      c.operacao.toUpperCase() === operacaoParaBuscar.toUpperCase()
    );
    if (config) {
      console.log(`[EMAIL_CONFIG] Carregando dados iniciais para ${selectedOperacaoConfig} (buscando: ${operacaoParaBuscar}):`);
      console.log(`  Envio: ${config.Envio || '(vazio)'}`);
      console.log(`  Copia: ${config.Copia || '(vazio)'}`);
      setConfigEnvio(config.Envio || '');
      setConfigCopia(config.Copia || '');
      emailConfigLoadedRef.current[selectedOperacaoConfig] = true;
    }
  }, [selectedOperacaoConfig, isEmailConfigModalOpen, userConfigs, isDeale]);

  // Função para salvar configuração de emails
  const handleSaveEmailConfig = async () => {
    if (!selectedOperacaoConfig) {
      alert('Selecione uma operação para configurar.');
      return;
    }

    setIsSavingConfig(true);
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        alert('Erro de autenticação. Tente novamente.');
        return;
      }

      // Para usuários DEALE, se a operação selecionada é "DEALE", salva na operação âncora (ALMIRANTE)
      // Caso contrário, salva na operação selecionada normalmente
      let operacaoParaSalvar = selectedOperacaoConfig;
      if (isDeale && selectedOperacaoConfig === 'DEALE') {
        operacaoParaSalvar = getDealeAnchorOperation();
      }

      console.log(`[EMAIL_CONFIG] Salvando configuração para ${operacaoParaSalvar} (selecionado: ${selectedOperacaoConfig})`);
      console.log(`[EMAIL_CONFIG] Envio: ${configEnvio}`);
      console.log(`[EMAIL_CONFIG] Copia: ${configCopia}`);

      await SharePointService.updateRouteConfigEmails(token, operacaoParaSalvar, configEnvio, configCopia);

      // Atualiza estado local — atualiza a config selecionada com os novos valores
      setUserConfigs(prev => prev.map(c =>
        c.operacao.toUpperCase() === operacaoParaSalvar.toUpperCase()
          ? { ...c, Envio: configEnvio, Copia: configCopia }
          : c
      ));

      alert('Configuração salva com sucesso!');
      // Reseta o flag para a operação salva, permitindo recarregar dados atualizados na próxima abertura
      delete emailConfigLoadedRef.current[selectedOperacaoConfig];
      setIsEmailConfigModalOpen(false);
      setIsConfigModalOpen(false);
    } catch (error: any) {
      console.error('[EMAIL_CONFIG] Erro ao salvar configuração:', error);
      alert(`Erro ao salvar: ${error.message}`);
    } finally {
      setIsSavingConfig(false);
    }
  };

  // Cleanup automático de locks expirados (timeout de 30 segundos)
  useEffect(() => {
    const LOCK_TIMEOUT = 30 * 1000; // 30 segundos
    
    const cleanupExpiredLocks = () => {
      const now = Date.now();
      let hasChanges = false;
      
      setRoutes(prev => {
        const updated = prev.map(route => {
          // Se tem lock e expirou, remove
          if (route.lockExpiresAt && now > route.lockExpiresAt && route.editingUser) {
            console.log(`[LOCK_CLEANUP] Lock expirado para ${route.id} (era de ${route.editingUser})`);
            hasChanges = true;
            return { ...route, editingUser: undefined, lockExpiresAt: undefined };
          }
          return route;
        });
        return hasChanges ? updated : prev;
      });
    };

    // Verifica a cada 5 segundos
    const interval = setInterval(cleanupExpiredLocks, 5000);
    
    return () => clearInterval(interval);
  }, []);

  // ⚠️ ENVIO AUTOMÁTICO DESABILITADO — O envio agora é feito apenas manualmente pela tela "Resumo" (SendReportView)
  // O useEffect abaixo foi desabilitado intencionalmente. O envio automático por countdown
  // não é mais necessário pois o usuário controla os envios pela tela de "Envio de Saídas e Não Coletas".
  //
  // useEffect(() => {
  //   if (isHistoryModalOpen) return;
  //   if (routes.length === 0 || userConfigs.length === 0) return;
  //   ... lógica de checkOperationAllOK + startSendCountdown ...
  // }, [routes, userConfigs, pendingSendOps, sendingOps, isHistoryModalOpen]);

  useEffect(() => {
    // Atualiza o tempo a cada 30 segundos usando fuso de Brasília
    const timer = setInterval(() => {
      const now = new Date();
      const brazilTimeStr = now.toLocaleTimeString('pt-BR', { timeZone: 'America/Sao_Paulo' });
      const [hours, minutes, seconds] = brazilTimeStr.split(':').map(Number);
      const brazilDate = new Date();
      brazilDate.setHours(hours, minutes, seconds);
      setCurrentTime(brazilDate);
    }, 30000);
    return () => clearInterval(timer);
  }, []);

  // Limpar timeout do tooltip ao desmontar
  useEffect(() => {
    return () => {
      if (tooltipTimeoutRef.current) {
        clearTimeout(tooltipTimeoutRef.current);
      }
    };
  }, []);

  // Persistir preferência de tema
  useEffect(() => {
    localStorage.setItem('route_departure_dark_mode', String(isDarkMode));
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [isDarkMode]);

  // Fechar dropdown de filtro ao clicar fora
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (filterDropdownRef.current && !filterDropdownRef.current.contains(event.target as Node)) {
        setActiveFilterCol(null);
      }
      // Fechar dropdown de observação ao clicar fora
      if (obsDropdownRef.current && !obsDropdownRef.current.contains(event.target as Node)) {
        setActiveObsId(null);
      }
      // Fechar menu de contexto ao clicar fora
      if (contextMenuRef.current && !contextMenuRef.current.contains(event.target as Node)) {
        setContextMenu(prev => ({ ...prev, visible: false }));
      }
      // Fechar tooltip do checklist
      if (checklistTooltip) {
        setChecklistTooltip(null);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [checklistTooltip]);

  // Redimensionamento de colunas
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!resizingRef.current) return;
      const { col, startX, startWidth } = resizingRef.current;
      const diff = e.clientX - startX;
      const newWidth = Math.max(50, startWidth + diff); // Mínimo 50px
      setColWidths(prev => ({ ...prev, [col]: newWidth }));
    };

    const handleMouseUp = () => {
      resizingRef.current = null;
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  // Recalcula status e tempo quando currentTime mudar (a cada 30s)
  useEffect(() => {
    if (!routes.length || !userConfigs.length) return;
    
    setRoutes(prev => prev.map(route => {
      const config = userConfigs.find(c => c.operacao === route.operacao);
      const { status, gap } = calculateStatusWithTolerance(route.inicio || '', route.saida || '', config?.tolerancia || "00:00:00", route.data || '');
      
      // Só atualiza se mudou
      if (route.statusOp !== status || route.tempo !== gap) {
        return { ...route, statusOp: status, tempo: gap };
      }
      return route;
    }));
  }, [currentTime]);

  // Salvar preferências de colunas no localStorage
  useEffect(() => {
    const savedWidths = localStorage.getItem('route_departure_col_widths');
    const savedHidden = localStorage.getItem('route_departure_hidden_cols');
    if (savedWidths) {
      try {
        setColWidths(JSON.parse(savedWidths));
      } catch (e) {}
    }
    if (savedHidden) {
      try {
        setHiddenColumns(new Set(JSON.parse(savedHidden)));
      } catch (e) {}
    }
  }, []);

  useEffect(() => {
    localStorage.setItem('route_departure_col_widths', JSON.stringify(colWidths));
  }, [colWidths]);

  useEffect(() => {
    localStorage.setItem('route_departure_hidden_cols', JSON.stringify(Array.from(hiddenColumns)));
  }, [hiddenColumns]);

  // Atalho CTRL+SHIFT+L para limpar todos os filtros
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 'l') {
        event.preventDefault();
        setColFilters({});
        setSelectedFilters({});
        setActiveFilterCol(null);
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, []);

  // Scroll horizontal com Shift + Wheel
  useEffect(() => {
    const tableContainer = document.getElementById('table-container');
    if (!tableContainer) return;

    const handleWheel = (event: WheelEvent) => {
      if (event.shiftKey && !event.ctrlKey && !event.altKey) {
        event.preventDefault();
        // Scroll horizontal: delta Y vira scroll X
        tableContainer.scrollLeft += event.deltaY;
      }
    };

    tableContainer.addEventListener('wheel', handleWheel, { passive: false });
    return () => tableContainer.removeEventListener('wheel', handleWheel);
  }, []);

  // Handler para copiar/colar status GERAL com Ctrl+C e Ctrl+V
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      // Só processa se for Ctrl+C ou Ctrl+V
      if (!(event.ctrlKey && (event.key.toLowerCase() === 'c' || event.key.toLowerCase() === 'v'))) {
        return;
      }
      
      // Previne comportamento padrão SEMPRE quando tiver seleção
      if (selectedIds.size > 0) {
        event.preventDefault();
        event.stopPropagation();
      }
      
      // Ctrl+C para copiar status GERAL da primeira rota selecionada
      if (event.ctrlKey && event.key.toLowerCase() === 'c') {
        const selectedRoutes = routes.filter(r => selectedIds.has(r.id));
        if (selectedRoutes.length >= 1) {
          // Pega o valor da primeira rota selecionada
          const route = selectedRoutes[0];
          if (route.statusGeral) {
            setCopiedGeralStatus(route.statusGeral);
          }
        }
      }
      
      // Ctrl+V para colar status GERAL em todas as rotas selecionadas
      if (event.ctrlKey && event.key.toLowerCase() === 'v') {
        if (copiedGeralStatus && selectedIds.size > 0) {
          const routesToUpdate = routes.filter(r => selectedIds.has(r.id));
          routesToUpdate.forEach(route => {
            updateCell(route.id, 'statusGeral', copiedGeralStatus);
          });
        }
      }
    };
    document.addEventListener('keydown', handleKeyDown, { capture: true });
    return () => document.removeEventListener('keydown', handleKeyDown, { capture: true });
  }, [copiedGeralStatus, selectedIds, routes]);

  const timeToSeconds = (timeStr: string): number => {
    if (!timeStr || !timeStr.includes(':')) return 0;
    const parts = timeStr.split(':').map(Number);
    return (parts[0] || 0) * 3600 + (parts[1] || 0) * 60 + (parts[2] || 0);
  };

  const secondsToTime = (totalSeconds: number): string => {
    const isNegative = totalSeconds < 0;
    const absSeconds = Math.abs(totalSeconds);
    const h = Math.floor(absSeconds / 3600);
    const m = Math.floor((absSeconds % 3600) / 60);
    const s = absSeconds % 60;
    const formatted = [h, m, s].map(v => v.toString().padStart(2, '0')).join(':');
    return isNegative ? `-${formatted}` : formatted;
  };

  const calculateStatusWithTolerance = (inicio: string, saida: string, toleranceStr: string = "00:00:00", routeDate: string): { status: string, gap: string } => {
    // Se não tem horário de início, retorna Previsto
    // NOTA: "00:00:00" é horário válido (meia-noite) e NÃO deve ser tratado como vazio
    if (!inicio || inicio === '') return { status: 'Previsto', gap: '' };
    if (!routeDate) return { status: 'Previsto', gap: '' };

    // Se saida for "-", considera rota não saída (atrasada)
    if (saida === '-') {
        return { status: 'Atrasada', gap: 'Não saiu' };
    }

    // Usa data brasileira para comparação correta
    const todayBrazil = getBrazilDate();
    const [todayY, todayM, todayD] = todayBrazil.split('-').map(Number);
    const today = new Date(todayY, todayM - 1, todayD);
    today.setHours(0, 0, 0, 0);

    const [y, m, d] = routeDate.split('-').map(Number);
    const rDate = new Date(y, m - 1, d);
    rDate.setHours(0, 0, 0, 0);

    const toleranceSec = timeToSeconds(toleranceStr);
    const startSec = timeToSeconds(inicio);

    // IMPORTANTE: "00:00:00" é horário válido (meia-noite), deve ser considerado no cálculo
    if (saida && saida !== '') {
        // Verifica se saida tem data completa (DD/MM/AAAA HH:MM:SS)
        const dateTimeMatch = saida.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
        
        let endSec: number;
        
        if (dateTimeMatch) {
            // Tem data completa - usa essa data para o cálculo
            const [, day, month, year, hour, minute, second] = dateTimeMatch;
            const saidaDate = new Date(Number(year), Number(month) - 1, Number(day));
            const routeDateObj = new Date(y, m - 1, d);
            
            // Calcula diferença de dias entre a data da saída e a data da rota
            const diffDays = Math.floor((saidaDate.getTime() - routeDateObj.getTime()) / (1000 * 60 * 60 * 24));
            
            // Adiciona/subtrai dias ao cálculo do gap
            endSec = timeToSeconds(`${hour}:${minute}:${second}`) + (diffDays * 24 * 3600);
        } else {
            // Apenas horário - usa data da coluna DATA
            endSec = timeToSeconds(saida);
        }
        
        const diff = endSec - startSec;
        const gapFormatted = secondsToTime(diff);

        if (diff < -toleranceSec) return { status: 'Adiantada', gap: gapFormatted };
        if (diff > toleranceSec) return { status: 'Atrasada', gap: gapFormatted };

        // Está OK - não mostra tempo
        return { status: 'OK', gap: '' };
    }

    if (rDate > today) return { status: 'Programada', gap: '' };
    if (rDate < today) return { status: 'Atrasada', gap: '' };

    // Calcula atraso baseado no horário atual (Brasília) vs horário de início
    // Usa o currentTime que já está sincronizado com Brasília
    const nowSec = currentTime.getHours() * 3600 + currentTime.getMinutes() * 60 + currentTime.getSeconds();
    const diffAtual = nowSec - startSec;

    if (diffAtual > toleranceSec) {
        // Está atrasada - calcula quanto tempo passou do horário
        const gapFormatted = secondsToTime(diffAtual);
        return { status: 'Atrasada', gap: gapFormatted };
    }

    return { status: 'Previsto', gap: '' };
  };

  const formatTimeInput = (value: string): string => {
    // Se já tem formato HH:MM:SS, valida e retorna
    if (/^\d{2}:\d{2}:\d{2}$/.test(value)) {
        const [h, m, s] = value.split(':').map(n => parseInt(n) || 0);
        return `${String(Math.min(23, h)).padStart(2, '0')}:${String(Math.min(59, m)).padStart(2, '0')}:${String(Math.min(59, s)).padStart(2, '0')}`;
    }

    // Se tem formato parcial HH:MM, adiciona segundos
    if (/^\d{2}:\d{2}$/.test(value)) {
        const [h, m] = value.split(':').map(n => parseInt(n) || 0);
        return `${String(Math.min(23, h)).padStart(2, '0')}:${String(Math.min(59, m)).padStart(2, '0')}:00`;
    }

    // Remove tudo que não é número
    let clean = (value || "").replace(/[^0-9]/g, '');
    if (!clean) return '';

    let h = '00', m = '00', s = '00';

    // Interpreta os dígitos como HHMMSS
    if (clean.length === 1) {
        // Apenas 1 dígito (ex: 5 -> 05:00:00)
        h = '0' + clean;
    } else if (clean.length === 2) {
        // Apenas horas (ex: 05 -> 05:00:00)
        h = clean;
    } else if (clean.length === 3) {
        // Horas e minutos sem zero (ex: 530 -> 05:30:00)
        h = clean.slice(0, 1).padStart(2, '0');
        m = clean.slice(1, 3);
    } else if (clean.length === 4) {
        // Horas e minutos (ex: 0530 ou 1200 -> 05:30:00 ou 12:00:00)
        h = clean.slice(0, 2);
        m = clean.slice(2, 4);
    } else if (clean.length === 5) {
        // Horas, minutos e segundos sem zero (ex: 12305 -> 12:30:05)
        h = clean.slice(0, 2);
        m = clean.slice(2, 4);
        s = clean.slice(4, 5).padStart(2, '0');
    } else if (clean.length >= 6) {
        // Horas, minutos e segundos completos (ex: 123045 -> 12:30:45)
        h = clean.slice(0, 2);
        m = clean.slice(2, 4);
        s = clean.slice(4, 6);
    }

    // Valida os valores
    h = String(Math.min(23, parseInt(h) || 0)).padStart(2, '0');
    m = String(Math.min(59, parseInt(m) || 0)).padStart(2, '0');
    s = String(Math.min(59, parseInt(s) || 0)).padStart(2, '0');

    return `${h}:${m}:${s}`;
  };

  // Extrai horários do template e abre modal de edição
  const openTimeEditModal = (routeId: string, template: string) => {
    // Verifica se o template tem placeholders **:**h ou horários já preenchidos
    const hasPlaceholders = template.includes('**:**h');
    
    if (hasPlaceholders) {
      // Abre o modal com valores vazios para o usuário preencher
      setTimeEditData({ routeId, template, startTime: '', endTime: '' });
      setIsTimeEditModalOpen(true);
    } else {
      // Tenta extrair horários já existentes no formato HH:MM:SSh ou HH:MMh
      const timeMatches = template.match(/(\d{2}:\d{2}(?::\d{2})?)h/g);
      
      if (timeMatches && timeMatches.length >= 2) {
        const startTime = timeMatches[0].replace('h', '');
        const endTime = timeMatches[1].replace('h', '');
        setTimeEditData({ routeId, template, startTime, endTime });
        setIsTimeEditModalOpen(true);
      } else {
        // Se não tem horários, aplica o template diretamente
        updateCell(routeId, 'observacao', template);
        setActiveObsId(null);
      }
    }
  };

  // Valida tempo de descarga para MONTES CLAROS + FÁBRICA
  const validateDescargaTime = (route: RouteDeparture, observacao: string): boolean => {
    if (route.operacao !== 'MONTES CLAROS' || route.motivo !== 'Fábrica') {
      return true; // Passa sem validação
    }
    
    // Extrai horários do texto da observação (formato HH:MM:SSh ou HH:MMh)
    const timeMatches = observacao.match(/(\d{2}:\d{2}(?::\d{2})?)h/g);
    
    if (timeMatches && timeMatches.length >= 2) {
      const startTime = timeMatches[0].replace('h', '');
      const endTime = timeMatches[1].replace('h', '');
      
      const startSeconds = timeToSeconds(startTime);
      const endSeconds = timeToSeconds(endTime);
      const diffSeconds = endSeconds - startSeconds;
      const diffHours = diffSeconds / 3600;
      
      // Se tempo de descarga for menor que 5 horas
      if (diffHours < 5 && diffSeconds > 0) {
        const confirmacao = window.confirm(
          `⚠️ ATENÇÃO: Motivo Incorreto\n\n` +
          `Tempo de descarga: ${diffHours.toFixed(1)} horas\n` +
          `Tolerância mínima do cliente: 5 horas\n\n` +
          `O motivo "Fábrica" não é recomendado para tempo inferior de descarga.\n\n` +
          `Deseja continuar mesmo assim?`
        );
        
        return confirmacao; // true = continua, false = cancela
      }
    }
    
    return true;
  };

  // Aplica os horários editados ao template
  const applyTimeEdit = () => {
    if (!timeEditData) return;
    
    const { routeId, template, startTime, endTime } = timeEditData;
    const route = routes.find(r => r.id === routeId);
    
    // Formata os horários para HH:MM:SS se necessário
    const formatTime = (time: string) => {
      if (!time) return '00:00:00';
      // Se já tem formato HH:MM:SS
      if (time.split(':').length === 3) return time;
      // Se tem formato HH:MM, adiciona segundos
      if (time.split(':').length === 2) return time + ':00';
      return time;
    };
    
    const startFormatted = formatTime(startTime);
    const endFormatted = formatTime(endTime);
    
    // Verifica se é placeholder ou horário existente
    const hasPlaceholders = template.includes('**:**h');
    
    let result = template;
    if (hasPlaceholders) {
      // Substitui os placeholders **:**h pelos horários formatados
      let replaceCount = 0;
      result = result.replace(/\*\*:\*\*h/g, (match) => {
        replaceCount++;
        if (replaceCount === 1) return `${startFormatted}h`;
        if (replaceCount === 2) return `${endFormatted}h`;
        return match;
      });
    } else {
      // Substitui horários existentes
      let replaceCount = 0;
      result = result.replace(/(\d{2}:\d{2}(?::\d{2})?)h/g, (match) => {
        replaceCount++;
        if (replaceCount === 1) return `${startFormatted}h`;
        if (replaceCount === 2) return `${endFormatted}h`;
        return match;
      });
    }
    
    // Validação específica para MONTES CLAROS + FÁBRICA
    if (route) {
      const valid = validateDescargaTime(route, result);
      if (!valid) {
        setIsTimeEditModalOpen(false);
        setTimeEditData(null);
        setActiveObsId(null);
        return;
      }
    }
    
    updateCell(routeId, 'observacao', result);
    setIsTimeEditModalOpen(false);
    setTimeEditData(null);
    setActiveObsId(null);
  };

  // Abre o modal de edição do checklist
  const openChecklistEditModal = (routeId: string, currentText: string) => {
    // Tenta extrair data e porcentagem do texto atual
    // Formato esperado: "DD/MM/AAAA - **% - motivos" ou "AAAA-MM-DD - **% - motivos"
    const dateMatch = currentText.match(/(\d{2}\/\d{2}\/\d{4})|(\d{4}-\d{2}-\d{2})/);
    const percentMatch = currentText.match(/(\d+)%/);
    const motivosMatch = currentText.match(/- (.+)$/);

    let data = '';
    if (dateMatch) {
      const matchedDate = dateMatch[0];
      // Se estiver no formato AAAA-MM-DD, converte para DD/MM/AAAA
      if (matchedDate.includes('-')) {
        const [year, month, day] = matchedDate.split('-');
        data = `${day}/${month}/${year}`;
      } else {
        data = matchedDate;
      }
    } else {
      // Data padrão no fuso de Brasília
      const today = getBrazilDate();
      const [year, month, day] = today.split('-');
      data = `${day}/${month}/${year}`;
    }
    
    const porcentagem = percentMatch ? percentMatch[1] : '100';
    const motivos = motivosMatch && !currentText.includes(percentMatch ? percentMatch[0] : '') ? motivosMatch[1] : '';
    const currentRoute = routes.find(r => r.id === routeId);
    const causaRaiz = (currentRoute?.causaRaiz || '').trim();

    setChecklistEditData({ routeId, data, porcentagem, motivos, causaRaiz });
    setIsChecklistEditModalOpen(true);
  };

  // Aplica a edição do checklist
  const applyChecklistEdit = async () => {
    if (!checklistEditData) return;

    const { routeId, data, porcentagem, motivos, causaRaiz } = checklistEditData;

    // Salva apenas os dados do checklist (sem o texto "Último checklist realizado")
    let result = `${data} - ${porcentagem}%`;

    // Se tem motivos e porcentagem < 100%, adiciona com hífen
    if (motivos && motivos.trim() !== '' && parseInt(porcentagem) < 100) {
      result += ` - ${motivos}`;
    }

    const causaRaizSanitizada = (causaRaiz || '').trim();
    console.log('[CHECKLIST] Salvando:', { routeId, result, causaRaiz: causaRaizSanitizada });

    try {
      await updateCell(routeId, 'checklistMotorista', result);
      await updateCell(routeId, 'causaRaiz' as keyof RouteDeparture, causaRaizSanitizada);
    } catch (e) {
      console.error('[CHECKLIST] Erro ao salvar checklist/causa raiz:', e);
    }

    setIsChecklistEditModalOpen(false);
    setChecklistEditData(null);
  };

  // Extrai dados do checklist do texto atual
  const extractChecklistData = (currentText: string) => {
    if (!currentText) return { data: '', porcentagem: '', motivos: '' };

    // Formato: "DD/MM/AAAA - **%" ou "AAAA-MM-DD - **%"
    const dateMatch = currentText.match(/(\d{2}\/\d{2}\/\d{4})|(\d{4}-\d{2}-\d{2})/);
    const percentMatch = currentText.match(/(\d+)%/);
    
    // Extrai motivos: tudo após "XX% - "
    let motivos = '';
    if (percentMatch) {
      const afterPercent = currentText.substring(percentMatch.index! + percentMatch[0].length).trim();
      if (afterPercent.startsWith('-')) {
        motivos = afterPercent.substring(1).trim();
      }
    }

    const data = dateMatch ? dateMatch[0] : '';
    const porcentagem = percentMatch ? percentMatch[1] : '';

    return { data, porcentagem, motivos };
  };

  // Formata o texto do tooltip
  const formatChecklistTooltip = (currentText: string): string => {
    const { data, porcentagem, motivos } = extractChecklistData(currentText);
    if (!data || !porcentagem) return '';
    
    let tooltip = `Checklist: ${data} - ${porcentagem}%`;
    if (motivos && motivos.trim() !== '') {
      tooltip += `\n${motivos}`;
    }
    return tooltip;
  };

  // Verifica se o checklist está preenchido
  const isChecklistFilled = (currentText: string): boolean => {
    const { data, porcentagem } = extractChecklistData(currentText);
    return !!(data && porcentagem);
  };

  // Máscara que formata enquanto digita (adiciona : automaticamente)
  const applyTimeMask = (value: string): string => {
    // Se for apenas "-", mantém
    if (value === '-') return '-';
    
    let clean = (value || "").replace(/[^0-9]/g, '');
    if (!clean) return '';

    // Limita a 6 dígitos
    clean = clean.slice(0, 6);

    // Formata com dois pontos
    if (clean.length <= 2) return clean;
    if (clean.length <= 4) return `${clean.slice(0, 2)}:${clean.slice(2)}`;
    return `${clean.slice(0, 2)}:${clean.slice(2, 4)}:${clean.slice(4)}`;
  };

  const loadData = async (isBackgroundRefresh: boolean = false) => {
    let token: string;
    try {
      token = await getAccessToken();
    } catch (e: any) {
      console.error('[RouteDeparture] Erro ao obter token:', e.message);
      // Dispara o evento para o App.tsx exibir o modal de renovação de sessão (com debounce)
      const now = Date.now();
      if (now - (window as any).__lastTokenEventTime > 10000) {
        (window as any).__lastTokenEventTime = now;
        window.dispatchEvent(new CustomEvent('token-expired'));
      } else {
        console.warn('[RouteDeparture] token-expired ignorado (debounce)');
      }
      return;
    }

    // Só mostra loading se NÃO for refresh em segundo plano
    if (!isBackgroundRefresh) {
      setIsLoading(true);
    }

    try {
      console.log('[LOAD_DATA] Buscando dados atualizados...', isBackgroundRefresh ? '(segundo plano)' : '(inicial)');
      console.log('[LOAD_DATA] Usuário logado:', currentUser.email);

      const [configs, mappings, spData] = await Promise.all([
        SharePointService.getRouteConfigs(token, currentUser.email, true), // force refresh
        SharePointService.getRouteOperationMappings(token),
        SharePointService.getDepartures(token, true) // force refresh
      ]);

      console.log('[LOAD_DATA] Configurações carregadas:', configs?.length || 0);
      console.log('[LOAD_DATA] Operações do usuário:', configs?.map(c => c.operacao));
      console.log('[LOAD_DATA] Total de rotas brutas do SharePoint:', spData?.length || 0);

      // Detecta se é usuário DEALE (para o modal de configurar emails)
      const deale = isDealeUser(configs || []);
      setIsDeale(deale);

      setUserConfigs(configs || []);
      setRouteMappings(mappings || []);

      // Filtra rotas APENAS das operações do usuário logado
      const myOps = new Set((configs || []).map(c => c.operacao));

      // DEBUG: Log detalhe das operações para identificar problemas de comparação
      console.log('[LOAD_DATA] Operações configuradas (myOps):', Array.from(myOps));
      console.log('[LOAD_DATA] Total de rotas brutas do SharePoint:', spData?.length || 0);

      // Log das primeiras 5 operações únicas nas rotas brutas
      const uniqueOpsInRoutes = Array.from(new Set((spData || []).map(r => r.operacao)));
      console.log('[LOAD_DATA] Operações únicas nas rotas brutas:', uniqueOpsInRoutes.slice(0, 10));

      const filteredByUser = (spData || []).filter(route => {
        // Se não houver operações configuradas, retorna todas (fallback)
        if (myOps.size === 0) return true;
        const match = myOps.has(route.operacao);
        if (!match && myOps.size <= 3) {
          // Log apenas para poucos configs (debug)
          console.log(`[LOAD_DATA] Rota "${route.rota}" NÃO match: operacao="${route.operacao}" não está em myOps`);
        }
        return match;
      });

      console.log('[LOAD_DATA] Rotas filtradas por usuário:', filteredByUser.length);
      console.log('[LOAD_DATA] Operações nas rotas filtradas:', Array.from(new Set(filteredByUser.map(r => r.operacao))));

      // Recalcula status e tempo para todas as rotas FILTRADAS
      const recalculatedRoutes = filteredByUser.map(route => {
        const config = configs?.find(c => c.operacao === route.operacao);
        const { status, gap } = calculateStatusWithTolerance(route.inicio || '', route.saida || '', config?.tolerancia || "00:00:00", route.data || '');
        return { ...route, statusOp: status, tempo: gap };
      });
      if (isBackgroundRefresh) {
        // MERGE inteligente: preserva a ordem atual da tabela e evita re-renderização desnecessária
        // Só atualiza se houver mudança real nos dados
        setRoutes(prevRoutes => {
          const newMap = new Map(recalculatedRoutes.map(r => [r.id, r]));
          const prevMap = new Map(prevRoutes.map(r => [r.id, r]));

          // Detecta mudanças: IDs novos/removidos ou campos alterados
          const newIds = new Set(newMap.keys());
          const prevIds = new Set(prevMap.keys());
          const hasIdChanges = newIds.size !== prevIds.size || ![...newIds].every(id => prevIds.has(id));

          if (!hasIdChanges) {
            // Mesmos IDs — verifica se algum campo relevante mudou
            let hasChanges = false;
            for (const route of recalculatedRoutes) {
              const prev = prevMap.get(route.id);
              if (!prev) { hasChanges = true; break; }
              // Compara campos relevantes (ignora campos que mudam a cada refresh)
              if (
                prev.rota !== route.rota ||
                prev.data !== route.data ||
                prev.inicio !== route.inicio ||
                prev.saida !== route.saida ||
                prev.motorista !== route.motorista ||
                prev.placa !== route.placa ||
                prev.motivo !== route.motivo ||
                prev.observacao !== route.observacao ||
                prev.statusGeral !== route.statusGeral ||
                prev.statusOp !== route.statusOp ||
                prev.tempo !== route.tempo ||
                prev.aviso !== route.aviso ||
                prev.operacao !== route.operacao ||
                prev.checklistMotorista !== route.checklistMotorista ||
                prev.causaRaiz !== route.causaRaiz
              ) {
                hasChanges = true;
                break;
              }
            }
            if (!hasChanges) {
              console.log('[LOAD_DATA] Nenhuma mudança detectada — mantendo estado atual');
              return prevRoutes; // Retorna mesma referência — React não re-renderiza
            }
          }

          // Há mudanças: merge preservando a ordem das rotas existentes + adicionando novas
          const merged = prevRoutes
            .map(existing => {
              const updated = newMap.get(existing.id);
              return updated || existing; // Mantém existente se não veio do SP (não deve acontecer)
            })
            .filter(r => newMap.has(r.id)); // Remove rotas que sumiram do SP

          // Adiciona rotas novas que não existiam no estado anterior
          const existingIds = new Set(prevRoutes.map(r => r.id));
          for (const route of recalculatedRoutes) {
            if (!existingIds.has(route.id)) {
              merged.push(route);
            }
          }

          console.log('[LOAD_DATA] Merge concluído —', merged.length, 'rotas');
          return merged;
        });
      } else {
        // Primeira carga ou refresh manual: substituição total (com spinner)
        setRoutes(recalculatedRoutes);
      }

      console.log('[LOAD_DATA] Dados carregados com sucesso');

      // Analisa histórico dos últimos 7 dias para alertas (apenas no carregamento inicial)
      if (!isBackgroundRefresh) {
        analyzeRouteHistory(token);
      }
      // Escaneia motoristas com "Mão de obra" — apenas no carregamento inicial para evitar queries excessivas
      if (!isBackgroundRefresh) {
        scanExistingMotoristAlerts(token, recalculatedRoutes);
      }

      // Atualiza o último checklist de motorista após carregar as rotas
      if (spData && spData.length > 0) {
        const motoristaRecords = spData.filter(r => r.motorista && r.motorista.trim() !== '');
        const byOperation: Record<string, RouteDeparture[]> = {};
        motoristaRecords.forEach(r => {
          if (!byOperation[r.operacao]) byOperation[r.operacao] = [];
          byOperation[r.operacao].push(r);
        });

        const result: Record<string, { data: string, porcentagem: string }> = {};
        Object.entries(byOperation).forEach(([op, records]) => {
          records.sort((a, b) => new Date(b.data).getTime() - new Date(a.data).getTime());
          if (records.length > 0) {
            const latest = records[0];
            const totalOps = records.length;
            const okOps = records.filter(r => r.statusOp === 'OK').length;
            const percentage = totalOps > 0 ? ((okOps / totalOps) * 100).toFixed(2) : '0.00';
            result[op] = {
              data: new Date(latest.data).toLocaleDateString('pt-BR'),
              porcentagem: `${percentage}%`
            };
          }
        });
        setLastMotoristaChecklist(result);
      }
    } catch (e: any) {
      console.error('[RouteDeparture] Erro ao carregar dados:', e.message);
      if (!isBackgroundRefresh && (e.message.includes('expired') || e.message.includes('401'))) {
        alert('Sua sessão expirou. Você será redirecionado para o login.');
        window.location.href = '/';
      } else if (!isBackgroundRefresh) {
        alert('Erro ao carregar dados: ' + e.message);
      }
    } finally {
      if (!isBackgroundRefresh) {
        setIsLoading(false);
      }
    }
  };

  useEffect(() => { loadData(); }, [currentUser]);

  // Função para buscar histórico do SharePoint (usada no modal e no polling)
  const handleSearchArchive = async () => {
    // Validação de período máximo (90 dias)
    const startMs = new Date(histStart).getTime();
    const endMs = new Date(histEnd).getTime();
    if (isNaN(startMs) || isNaN(endMs)) {
      alert('Selecione datas válidas para a busca.');
      return;
    }
    const dayDiff = Math.ceil(Math.abs(endMs - startMs) / (1000 * 60 * 60 * 24));
    if (dayDiff > 90) {
      alert(`O intervalo máximo permitido é de 90 dias. O intervalo selecionado é de ${dayDiff} dias.`);
      return;
    }

    // Cancela requisição anterior se existir
    if (archiveAbortRef.current) {
      archiveAbortRef.current.abort();
    }
    const controller = new AbortController();
    archiveAbortRef.current = controller;

    setIsSearchingArchive(true);
    try {
        console.log('[SEARCH_ARCHIVE] Requesting history from SharePoint list {856bf9d5-6081-4360-bcad-e771cbabfda8}...');
        const results = await SharePointService.getArchivedDepartures(await getAccessToken(), null, histStart, histEnd, controller.signal);
        console.log('[SEARCH_ARCHIVE] Results received:', results.length);

        // Só atualiza state se esta requisição não foi cancelada
        if (!controller.signal.aborted) {
          const myOps = new Set(userConfigs.map(c => c.operacao));
          // If myOps is empty, show everything for the user to avoid blockage if config loading is slow
          const filtered = results && results.length > 0
            ? results.filter(r => !myOps.size || myOps.has(r.operacao))
            : [];

          setArchivedResults(filtered);
        }
    } catch (err: any) {
        if (err.name === 'AbortError') {
          console.log('[SEARCH_ARCHIVE] Requisição cancelada.');
          return;
        }
        console.error('[SEARCH_ARCHIVE] Error during search:', err);
        alert("Erro na busca: " + (err?.message || "Erro desconhecido ao acessar o SharePoint. Verifique se você tem permissão na lista de histórico."));
    } finally {
        if (!controller.signal.aborted) {
          setIsSearchingArchive(false);
        }
    }
  };

  // Estado para controlar se o usuário está editando alguma célula (pausa polling)
  const [isEditingCell, setIsEditingCell] = useState(false);

  // Polling para atualizar dados automaticamente a cada 30 segundos (OTIMIZADO)
  useEffect(() => {
    const refreshInterval = setInterval(() => {
      // Se o usuário está editando, NÃO faz polling para evitar lag
      if (isEditingCell) {
        console.log('[POLLING_ROUTE_DEPARTURE] Pulando atualização (usuário editando)');
        return;
      }
      console.log('[POLLING_ROUTE_DEPARTURE] Atualização automática de dados (segundo plano)');
      console.log('[POLLING_ROUTE_DEPARTURE] Usuário:', currentUser.email);
      loadData(true); // true = segundo plano (sem loading, sem spinner)
    }, 30000);

    return () => clearInterval(refreshInterval);
  }, [currentUser, isEditingCell]);

  // Handler de teclado para salvar edições no histórico (Enter)
  useEffect(() => {
    if (!isHistoryModalOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      // Verifica se Enter foi pressionado e não está em um input/textarea
      if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && !e.altKey) {
        const target = e.target as HTMLElement;
        // Se estiver em input/textarea/select, salva e fecha
        if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT') {
          // Salva as edições pendentes
          if (Object.keys(pendingHistoryEdits).length > 0) {
            e.preventDefault();
            savePendingHistoryEdits();
          }
          return;
        }
        
        // Se não estiver em input, apenas salva se houver edições
        if (Object.keys(pendingHistoryEdits).length > 0) {
          e.preventDefault();
          savePendingHistoryEdits();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isHistoryModalOpen, pendingHistoryEdits]);

  // Persistir configurações da tabela no sessionStorage
  useEffect(() => {
    sessionStorage.setItem('route_departure_col_filters', JSON.stringify(colFilters));
  }, [colFilters]);

  useEffect(() => {
    sessionStorage.setItem('route_departure_selected_filters', JSON.stringify(selectedFilters));
  }, [selectedFilters]);

  useEffect(() => {
    sessionStorage.setItem('route_departure_sort_by_time', JSON.stringify(isSortByTimeEnabled));
  }, [isSortByTimeEnabled]);

  useEffect(() => {
    sessionStorage.setItem('route_departure_col_widths', JSON.stringify(colWidths));
  }, [colWidths]);

  useEffect(() => {
    sessionStorage.setItem('route_departure_hidden_cols', JSON.stringify(Array.from(hiddenColumns)));
  }, [hiddenColumns]);

  const handleDeleteRoute = async (id: string) => {
    if (!confirm('Deseja excluir permanentemente esta rota do SharePoint?')) return;
    
    // VALIDAÇÃO CRÍTICA: Só permite excluir rotas que pertencem às operações do usuário logado
    const routeToDelete = routes.find(r => r.id === id);
    if (routeToDelete && routeToDelete.operacao) {
        const myOps = new Set(userConfigs.map(c => c.operacao));
        if (!myOps.has(routeToDelete.operacao)) {
            console.error('[DELETE_BLOCKED] Usuário tentou excluir rota de operação não pertencente:', routeToDelete.operacao);
            alert('⚠️ Erro: Você não tem permissão para excluir esta rota.');
            return;
        }
    }
    
    const token = await getAccessToken();
    setIsSyncing(true);
    try {
      await SharePointService.deleteDeparture(token, id);
      setRoutes(prev => prev.filter(r => r.id !== id));
      setSelectedIds(prev => { const next = new Set(prev); next.delete(id); return next; });
    } catch (e) { alert("Erro ao excluir item."); }
    finally { setIsSyncing(false); }
  };

  const handleDeleteSelected = async () => {
    if (selectedIds.size === 0) return;
    
    // VALIDAÇÃO CRÍTICA: Filtra apenas rotas que pertencem às operações do usuário logado
    const myOps = new Set(userConfigs.map(c => c.operacao));
    const validIdsToDelete = Array.from(selectedIds).filter(id => {
        const route = routes.find(r => r.id === id);
        if (!route || !route.operacao) return false;
        return myOps.has(route.operacao);
    });
    
    const blockedCount = selectedIds.size - validIdsToDelete.length;
    if (blockedCount > 0) {
        alert(`⚠️ Você só pode excluir rotas das suas operações. ${blockedCount} rota(s) serão ignoradas.`);
    }
    
    if (validIdsToDelete.length === 0) {
        alert('⚠️ Nenhuma rota selecionada pertence às suas operações.');
        return;
    }
    
    if (!confirm(`Deseja excluir as ${validIdsToDelete.length} rotas selecionadas do SharePoint?`)) return;
    
    const token = await getAccessToken();
    setIsSyncing(true);
    let success = 0;
    for (const id of validIdsToDelete) {
        try {
          await SharePointService.deleteDeparture(token, id);
          success++;
        } catch (e) {}
    }
    setRoutes(prev => prev.filter(r => !selectedIds.has(r.id!)));
    setSelectedIds(new Set());
    setIsSyncing(false);
    alert(`${success} rotas excluídas.`);
  };

  const isDelayedOrAdvancedStatus = (status: string): boolean => {
    const normalized = (status || '').trim();
    return normalized === 'Atrasada' || normalized === 'Adiantada' || normalized === 'Atrasado' || normalized === 'Adiantado';
  };

  const getArchiveMissingFields = (route: RouteDeparture): string[] => {
    const missing: string[] = [];
    const saida = (route.saida || '').trim();
    const motivo = (route.motivo || '').trim();
    const observacao = (route.observacao || '').trim();
    const causaRaiz = (route.causaRaiz || '').trim();
    const geral = (route.statusGeral || '').trim().toUpperCase();

    if (geral !== 'OK') {
      missing.push('GERAL');
    }

    if (!saida) {
      missing.push('SAÍDA');
    }

    if (isDelayedOrAdvancedStatus(route.statusOp || '')) {
      if (!motivo) {
        missing.push('MOTIVO');
      }
      if (!observacao) {
        missing.push('OBSERVAÇÃO');
      }
    }

    if (motivo === 'Manutenção' && !causaRaiz) {
      missing.push('CAUSA RAIZ');
    }

    return missing;
  };

  const addOneDayToRouteDate = (dateValue: string): string => {
    const trimmed = (dateValue || '').trim();
    if (!trimmed) return getBrazilDate();

    let dateObj: Date | null = null;
    const isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const brMatch = trimmed.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);

    if (isoMatch) {
      const [, y, m, d] = isoMatch;
      dateObj = new Date(Number(y), Number(m) - 1, Number(d));
    } else if (brMatch) {
      const [, d, m, y] = brMatch;
      dateObj = new Date(Number(y), Number(m) - 1, Number(d));
    } else {
      const tentative = new Date(trimmed);
      if (!Number.isNaN(tentative.getTime())) {
        dateObj = tentative;
      }
    }

    if (!dateObj) return trimmed;

    dateObj.setDate(dateObj.getDate() + 1);
    const yyyy = dateObj.getFullYear();
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const dd = String(dateObj.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };

  const validateArchiveHardBlocks = (routesCandidate: RouteDeparture[]): boolean => {
    // Trava 1: Resumo geral de saídas precisa estar OK
    const configsComResumo = userConfigs.filter(c => !!c.UltimoEnvioResumoSaida && !!c.StatusResumoSaida);
    const statusResumoGlobal = (() => {
      if (configsComResumo.length === 0) return '';
      const maisRecente = [...configsComResumo].sort((a, b) => {
        const dateA = new Date(a.UltimoEnvioResumoSaida || '').getTime();
        const dateB = new Date(b.UltimoEnvioResumoSaida || '').getTime();
        return dateB - dateA;
      })[0];
      return (maisRecente?.StatusResumoSaida || '').trim();
    })();

    if (statusResumoGlobal.toUpperCase() !== 'OK') {
      setArchiveStatusDivergences([]);
      setArchiveHardBlockModal('resumo');
      return false;
    }

    // Trava 2: Status individual (coluna Status) por operação precisa estar OK
    const opsInArchive = Array.from(new Set(routesCandidate.map(r => (r.operacao || '').trim()).filter(Boolean)));
    const statusDivergences: ArchiveStatusDivergence[] = opsInArchive
      .map(op => {
        const cfg = userConfigs.find(c => c.operacao === op);
        const rawStatus = (cfg?.Status || '').trim();
        return { operacao: op, status: rawStatus || 'Previsto' };
      })
      .filter(item => item.status.toUpperCase() !== 'OK');

    if (statusDivergences.length > 0) {
      setArchiveStatusDivergences(statusDivergences);
      setArchiveHardBlockModal('status');
      return false;
    }

    return true;
  };

  const executeArchiveWithDivergences = async (routesToArchive: RouteDeparture[], routesToPostpone: ArchiveDivergentRoute[]) => {
    const token = await getAccessToken();
    setIsSyncing(true);

    // Desmarca a opção HORÁRIO após arquivar
    setIsSortByTimeEnabled(false);
    setIsSortByOperacao(false);

    try {
      let archiveSuccess = 0;
      let archiveFailed = 0;
      let clearCount = 0;

      if (routesToArchive.length > 0) {
        console.log(`[ARCHIVE] Movendo ${routesToArchive.length} itens para o histórico...`);
        const archiveResult = await SharePointService.moveDeparturesToHistory(token, routesToArchive);
        archiveSuccess = archiveResult.success;
        archiveFailed = archiveResult.failed;
        console.log(`[ARCHIVE] Sucesso: ${archiveSuccess}, Falhas: ${archiveFailed}`);

        // Limpa status de envio apenas das operações arquivadas
        console.log('[ARCHIVE] Limpando status de envio nas configurações das operações arquivadas...');
        const opsToClear = Array.from(new Set(routesToArchive.map(r => r.operacao).filter(Boolean)));
        for (const operacao of opsToClear) {
          try {
            await SharePointService.updateUltimoEnvioSaida(token, operacao, '');
            await SharePointService.updateStatusOperacao(token, operacao, '');
            await SharePointService.updateUltimoEnvioResumoSaida(token, operacao, '');
            await SharePointService.updateStatusResumoSaida(token, operacao, '');
            clearCount++;
            console.log(`[ARCHIVE] ✅ Status limpo para ${operacao}`);
          } catch (e: any) {
            console.error(`[ARCHIVE] Erro ao limpar status de ${operacao}:`, e.message);
          }
        }
      }

      // Rotas divergentes: mantêm no painel e data vai para o próximo dia
      let postponedSuccess = 0;
      for (const divergence of routesToPostpone) {
        try {
          const nextDate = addOneDayToRouteDate(divergence.route.data || '');
          const cfg = userConfigs.find(c => c.operacao === divergence.route.operacao);
          const statusCalc = calculateStatusWithTolerance(
            divergence.route.inicio || '',
            divergence.route.saida || '',
            cfg?.tolerancia || '00:00:00',
            nextDate
          );

          const updatedRoute: RouteDeparture = {
            ...divergence.route,
            data: nextDate,
            semana: getWeekString(nextDate),
            statusOp: statusCalc.status,
            tempo: statusCalc.gap
          };

          await SharePointService.updateDeparture(token, updatedRoute);
          postponedSuccess++;
        } catch (e: any) {
          console.error(`[ARCHIVE] Erro ao postergar rota ${divergence.route.rota}:`, e.message);
        }
      }

      // Recarrega dados e configs
      await loadData(true);
      try {
        const refreshedConfigs = await SharePointService.getRouteConfigs(token, currentUser.email, true);
        setUserConfigs(refreshedConfigs);
      } catch (e: any) {
        console.error('[ARCHIVE] Erro ao atualizar configs:', e.message);
      }

      if (routesToArchive.length > 0 && routesToPostpone.length > 0) {
        alert(
          `${archiveSuccess} rota(s) arquivada(s) com sucesso (falhas: ${archiveFailed}).\n` +
          `${postponedSuccess} rota(s) incompleta(s) foram mantidas no painel e ajustadas para o próximo dia.\n` +
          `Status de envio limpo para ${clearCount} operação(ões).`
        );
      } else if (routesToArchive.length > 0) {
        alert(`${archiveSuccess} rota(s) arquivada(s) com sucesso!\nStatus de envio limpo para ${clearCount} operação(ões).`);
      } else {
        alert(`${postponedSuccess} rota(s) incompleta(s) foram mantidas no painel e ajustadas para o próximo dia.`);
      }
    } catch (e: any) {
      console.error('[ARCHIVE] Erro geral:', e.message);
      alert(`Erro ao arquivar: ${e.message}`);
    } finally {
      setIsSyncing(false);
      setArchiveReadyRoutes([]);
      setArchiveDivergentRoutes([]);
      setIsArchiveValidationModalOpen(false);
    }
  };

  const handleArchiveAll = async () => {
    // VALIDAÇÃO CRÍTICA: Filtra apenas rotas que pertencem às operações do usuário logado
    const myOps = new Set(userConfigs.map(c => c.operacao));
    const validFilteredRoutes = filteredRoutes.filter(r => !r.operacao || myOps.has(r.operacao));

    if (validFilteredRoutes.length === 0) {
      alert("Não há rotas das suas operações para arquivar.");
      return;
    }

    // Verifica se há alguma rota de operação não pertencente
    const blockedCount = filteredRoutes.length - validFilteredRoutes.length;
    if (blockedCount > 0) {
      console.warn(`[ARCHIVE] ${blockedCount} rotas de outras operações serão ignoradas`);
    }

    // Verifica se há alguma rota com status "Previsto"
    const rotasPrevistas = validFilteredRoutes.filter(r => r.statusOp === 'Previsto');
    if (rotasPrevistas.length > 0) {
      alert(
        `⚠️ Atenção!\n\n` +
        `Existem ${rotasPrevistas.length} rota(s) com status "Previsto":\n\n` +
        rotasPrevistas.map(r => `• ${r.rota} (${r.operacao})`).join('\n') +
        `\n\nPor favor, ajuste todas as rotas antes de arquivar.`
      );
      return;
    }

    const divergentRoutes: ArchiveDivergentRoute[] = [];
    const readyToArchive: RouteDeparture[] = [];

    validFilteredRoutes.forEach(route => {
      const missingFields = getArchiveMissingFields(route);
      if (missingFields.length > 0) {
        divergentRoutes.push({ route, missingFields });
      } else {
        readyToArchive.push(route);
      }
    });

    if (divergentRoutes.length > 0) {
      setArchiveReadyRoutes(readyToArchive);
      setArchiveDivergentRoutes(divergentRoutes);
      setIsArchiveValidationModalOpen(true);
      return;
    }

    // Travas bloqueantes: resumo geral OK + status individual OK
    if (!validateArchiveHardBlocks(readyToArchive)) {
      return;
    }

    await executeArchiveWithDivergences(readyToArchive, []);
  };

  const handleAddRoute = async () => {
    // Validação básica
    if (!newRouteData.rota || !newRouteData.inicio || !newRouteData.motorista || !newRouteData.placa || !newRouteData.operacao) {
      alert('Preencha todos os campos obrigatórios!');
      return;
    }

    // Validação do formato do horário (HH:MM:SS)
    const timeRegex = /^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$/;
    if (!timeRegex.test(newRouteData.inicio)) {
      alert('⚠️ Horário de início deve estar no formato HH:MM:SS (ex: 08:30:00)');
      return;
    }

    // NORMALIZAÇÃO DA PLACA: Remove espaços, hífens e converte para maiúsculo
    const cleanPlaca = newRouteData.placa.replace(/[\s-]/g, '').toUpperCase();

    // VALIDAÇÃO CRÍTICA: Só permite adicionar rotas nas operações do usuário logado
    const myOps = new Set(userConfigs.map(c => c.operacao));
    if (!myOps.has(newRouteData.operacao)) {
        console.error('[ADD_ROUTE_BLOCKED] Usuário tentou adicionar rota em operação não pertencente:', newRouteData.operacao);
        alert('⚠️ Erro: Você não tem permissão para adicionar rotas desta operação.');
        return;
    }

    setIsAddingRoute(true);
    const token = await getAccessToken();

    try {
      const config = userConfigs.find(c => c.operacao === newRouteData.operacao);
      // Usa a data correta baseada no horário atual (D+1 após 21:00h)
      const routeDate = getRouteDateForCurrentTime();
      const { status, gap } = calculateStatusWithTolerance(newRouteData.inicio, '', config?.tolerancia || "00:00:00", routeDate);

      const newRoute: RouteDeparture = {
        id: '',
        semana: getWeekString(routeDate),
        rota: newRouteData.rota,
        data: routeDate,
        inicio: newRouteData.inicio,
        motorista: newRouteData.motorista,
        placa: cleanPlaca,
        saida: '',
        motivo: '',
        observacao: '',
        statusGeral: '',
        aviso: 'NÃO',
        operacao: newRouteData.operacao,
        statusOp: status,
        tempo: gap,
        createdAt: new Date().toISOString(),
        causaRaiz: ''
      };

      const newId = await SharePointService.updateDeparture(token, newRoute);

      // Recarrega as rotas para mostrar a nova
      await loadData(true);

      // Limpa o formulário e fecha o modal
      setNewRouteData({ rota: '', inicio: '', motorista: '', placa: '', operacao: '' });
      setIsAddRouteModalOpen(false);

      // DESABILITA o filtro por horário para não misturar a nova rota na ordenação
      setIsSortByTimeEnabled(false);
      setIsSortByOperacao(false);

      alert('Rota adicionada com sucesso!');
    } catch (e: any) {
      console.error('[ADD_ROUTE] Erro:', e.message);
      alert(`Erro ao adicionar rota: ${e.message}`);
    } finally {
      setIsAddingRoute(false);
    }
  };

  const toggleSelection = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const handleBulkCreateSave = async (operacao: string) => {
    // VALIDAÇÃO CRÍTICA: Só permite criar rotas em massa nas operações do usuário logado
    const myOps = new Set(userConfigs.map(c => c.operacao));
    if (!myOps.has(operacao)) {
        console.error('[BULK_CREATE_BLOCKED] Usuário tentou criar rotas em massa em operação não pertencente:', operacao);
        alert('⚠️ Erro: Você não tem permissão para adicionar rotas desta operação.');
        setIsBulkMappingModalOpen(false);
        setPendingBulkRoutes([]);
        return;
    }

    // BLOQUEIO: Não permite criar rotas em massa se há filtros ou ordenação ativos
    if (hasActiveFiltersOrSort) {
        console.warn('[BULK_CREATE_BLOCKED] Filtros ou ordenação ativos');
        setFilterBlockReason('bulk');
        setIsBulkMappingModalOpen(false);
        setPendingBulkRoutes([]);
        setIsFilterBlockModalOpen(true);
        return;
    }

    const token = await getAccessToken();
    const total = pendingBulkRoutes.length;
    setIsBulkMappingModalOpen(false);
    setBulkStatus({ active: true, current: 0, total });
    const newRoutes: RouteDeparture[] = [];
    const config = userConfigs.find(c => c.operacao === operacao);
    for (let i = 0; i < total; i++) {
        const rotaName = pendingBulkRoutes[i];
        setBulkStatus((prev: any) => prev ? { ...prev, current: i + 1 } : null);
        const { status, gap } = calculateStatusWithTolerance(ghostRow.inicio || '', ghostRow.saida || '', config?.tolerancia || "00:00:00", ghostRow.data || "");
        const payload: RouteDeparture = { ...ghostRow, id: '', rota: rotaName, operacao: operacao, statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture;
        try { const newId = await SharePointService.updateDeparture(token, payload); newRoutes.push({ ...payload, id: newId }); } catch (e) {}
    }
    setRoutes(prev => [...prev, ...newRoutes]);
    setBulkStatus(null);
    setPendingBulkRoutes([]);
    setGhostRow({ id: 'ghost', rota: '', data: getRouteDateForCurrentTime(), inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '', causaRaiz: '' });

    // DESABILITA o filtro por horário para não misturar as novas rotas na ordenação
    setIsSortByTimeEnabled(false);
    setIsSortByOperacao(false);
  };

  const handleMultilinePaste = async (field: keyof RouteDeparture, startRowIndex: number, value: string) => {
    const lines = value.split(/[\n\r]/).map(l => l.trim()).filter(Boolean);
    if (lines.length <= 1) return;
    const token = await getAccessToken();
    setIsSyncing(true);

    // CORREÇÃO CRÍTICA DO BUG:
    // O startRowIndex vem do filteredRoutes.map(), então devemos usar filteredRoutes
    // para identificar as linhas afetadas, NÃO routes (que contém todas as rotas)
    const targetRoutes = filteredRoutes.slice(startRowIndex, startRowIndex + lines.length);

    if (targetRoutes.length === 0) {
        setIsSyncing(false);
        return;
    }

    // VALIDAÇÃO CRÍTICA: Verifica TODAS as linhas afetadas antes de aplicar o paste
    // Se houver QUALQUER linha de outra operação, REJEITA o paste inteiro
    const myOps = new Set(userConfigs.map(c => c.operacao));
    const invalidRoutes = targetRoutes.filter(r => r.operacao && !myOps.has(r.operacao));

    if (invalidRoutes.length > 0) {
        // REJEITA O PASTE INTEIRO - não aplica em nenhuma linha
        console.error(`[PASTE_BLOCKED] ${invalidRoutes.length} linhas são de outras operações. Paste rejeitado.`);
        alert(`❌ Paste bloqueado: ${invalidRoutes.length} linha(s) pertencem a outras operações.\n\nIsso previne edição acidental de dados de outros usuários.`);
        setIsSyncing(false);
        return;
    }

    // VALIDAÇÃO CRÍTICA 2: Verifica se alguma linha está com lock de outro usuário
    const now = Date.now();
    const lockedRoutes = targetRoutes.filter(r =>
      r.editingUser &&
      r.lockExpiresAt &&
      now < r.lockExpiresAt &&
      r.editingUser !== currentUser.email
    );

    if (lockedRoutes.length > 0) {
        console.error(`[PASTE_BLOCKED] ${lockedRoutes.length} linhas estão sendo editadas por outros usuários.`);
        const lockedBy = lockedRoutes.map(r => `${r.rota} (${r.editingUser})`).join(', ');
        alert(`🔒 Paste bloqueado: ${lockedRoutes.length} linha(s) estão sendo editadas por outros usuários.\n\nLinhas bloqueadas: ${lockedBy}`);
        setIsSyncing(false);
        return;
    }

    // Todas as linhas são válidas, prossegue com o paste
    const updatePromises = targetRoutes.map(async (route, i) => {
        let finalValue = lines[i];
        if (field === 'inicio' || field === 'saida') {
            finalValue = formatTimeInput(finalValue);
        }
        // NORMALIZAÇÃO DA PLACA: Remove espaços, hífens e converte para maiúsculo
        if (field === 'placa') {
            finalValue = finalValue.replace(/[\s-]/g, '').toUpperCase();
        }
        const updatedRoute: RouteDeparture = { ...route, [field]: finalValue };
        if (field === 'motivo' && (finalValue || '').trim() !== 'Manutenção') {
            updatedRoute.causaRaiz = '';
        }
        const config = userConfigs.find(c => c.operacao === updatedRoute.operacao);
        const { status, gap } = calculateStatusWithTolerance(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia || "00:00:00", updatedRoute.data);
        updatedRoute.statusOp = status;
        updatedRoute.tempo = gap;

        try {
            await SharePointService.updateDeparture(token, updatedRoute);
            return { id: route.id, updatedRoute };
        } catch (err) {
            console.error('[PASTE] Error updating route:', route.rota, err);
            return null;
        }
    });

    // Executa todas em paralelo
    const results = await Promise.all(updatePromises);

    // Atualiza o estado com todos os resultados de uma vez
    setRoutes(prev => {
        const newRoutes = [...prev];
        results.forEach(result => {
            if (!result) return;
            const index = newRoutes.findIndex(r => r.id === result.id);
            if (index !== -1) {
                newRoutes[index] = result.updatedRoute;
            }
        });
        return newRoutes;
    });

    // DESABILITA o filtro por horário para não misturar as rotas atualizadas na ordenação
    setIsSortByTimeEnabled(false);
    setIsSortByOperacao(false);

    setIsSyncing(false);
  };

  const updateCell = async (id: string, field: keyof RouteDeparture, value: string) => {
    // Normalização automática da placa: remove espaços e hífens
    if (field === 'placa') {
      value = value.replace(/[\s-]/g, '').toUpperCase();
    }

    if (id === 'ghost') {
        let updatedGhost = { ...ghostRow, [field]: value };

        // Verifica se é campo 'rota' e se tem múltiplas linhas (paste)
        if (field === 'rota' && (value.includes('\n') || value.includes(';'))) {
            const lines = value.split(/[\n;]/).map(l => l.trim()).filter(Boolean);
            console.log('[GHOST_ROTA] Múltiplas linhas detectadas:', lines);
            if (lines.length > 1) {
                setPendingBulkRoutes(lines);
                setIsBulkMappingModalOpen(true);
                return;
            }
            // Se chegou aqui, é uma única linha com newline no final - remove o newline
            value = lines[0] || '';
            updatedGhost = { ...ghostRow, [field]: value };
        }

        // Se é campo 'rota' e tem valor, abre popup de mapeamento SEMPRE
        if (field === 'rota' && value !== "" && value.trim() !== "") {
            console.log('[GHOST_ROTA] Buscando mapeamento para:', value, 'Mappings disponíveis:', routeMappings.map(m => m.Title));
            
            // BLOQUEIO: Não permite adicionar rota se há filtros ou ordenação ativos
            if (hasActiveFiltersOrSort) {
                console.warn('[GHOST_ROTA] Bloqueado - filtros ou ordenação ativos');
                setFilterBlockReason('ghost');
                setGhostRow({ id: 'ghost', rota: '', data: getRouteDateForCurrentTime(), inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '', causaRaiz: '' });
                setIsFilterBlockModalOpen(true);
                return;
            }
            
            const mapping = routeMappings.find(m => m.Title === value);
            if (mapping) {
                // Já tem mapeamento, aplica diretamente
                console.log('[GHOST_ROTA] Mapeamento encontrado:', mapping);
                updatedGhost.operacao = mapping.OPERACAO;
                setGhostRow(updatedGhost);
            } else {
                // Não tem mapeamento, abre popup
                console.log('[GHOST_ROTA] Sem mapeamento, abrindo modal para:', value);
                // Primeiro atualiza o ghostRow para manter o valor da rota
                setGhostRow(updatedGhost);
                setPendingMappingRoute(value);
                setIsMappingModalOpen(true);
            }
            return;
        }

        // Para outros campos da ghost row - VALIDA se a operação pertence ao usuário antes de salvar
        if (field !== 'rota' && updatedGhost.rota) {
            // VALIDAÇÃO CRÍTICA: Só permite salvar se a operação estiver nas configurações do usuário logado
            const myOps = new Set(userConfigs.map(c => c.operacao));
            if (updatedGhost.operacao && !myOps.has(updatedGhost.operacao)) {
                console.error('[UPDATE_BLOCKED] Usuário tentou salvar rota com operação não pertencente:', updatedGhost.operacao);
                alert('⚠️ Erro: Você não tem permissão para adicionar rotas desta operação.');
                // Reseta a ghost row para evitar dados inconsistentes
                setGhostRow({ id: 'ghost', rota: '', data: getRouteDateForCurrentTime(), inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '', causaRaiz: '' });
                return;
            }

            setIsSyncing(true);
            try {
                const config = userConfigs.find(c => c.operacao === updatedGhost.operacao);
                const { status, gap } = calculateStatusWithTolerance(updatedGhost.inicio || '', updatedGhost.saida || '', config?.tolerancia || "00:00:00", updatedGhost.data || "");
                const payload = { ...updatedGhost, statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture;
                const newId = await SharePointService.updateDeparture(await getAccessToken(), payload);
                setRoutes(prev => [...prev, { ...payload, id: newId }]);

                // Limpa filtros após criar nova rota para garantir que ela seja visível
                setColFilters({});
                setSelectedFilters({});

                // DESABILITA o filtro por horário para não misturar a nova rota na ordenação
                setIsSortByTimeEnabled(false);
                setIsSortByOperacao(false);

                setGhostRow({ id: 'ghost', rota: '', data: getRouteDateForCurrentTime(), inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '', causaRaiz: '' });
            } catch (e) {} finally { setIsSyncing(false); }
        } else {
            setGhostRow(updatedGhost);
        }
        return;
    }

    const route = routes.find(r => r.id === id);
    if (!route) return;

    // VALIDAÇÃO CRÍTICA 1: Só permite editar rotas que pertencem às operações do usuário logado
    const myOps = new Set(userConfigs.map(c => c.operacao));
    if (route.operacao && !myOps.has(route.operacao)) {
        console.error('[UPDATE_BLOCKED] Usuário tentou editar rota de operação não pertencente:', route.operacao);
        alert('⚠️ Erro: Você não tem permissão para editar esta rota.');
        return;
    }

    // VALIDAÇÃO CRÍTICA 2: Verifica se outra pessoa está editando esta linha (lock temporário)
    const now = Date.now();
    if (route.editingUser && route.lockExpiresAt && now < route.lockExpiresAt) {
      if (route.editingUser !== currentUser.email) {
        console.warn(`[UPDATE_BLOCKED] Linha ${id} está sendo editada por ${route.editingUser} (lock até ${new Date(route.lockExpiresAt).toLocaleTimeString()})`);
        alert(`🔒 Esta linha está sendo editada por ${route.editingUser}.\n\nAguarde alguns segundos e tente novamente.`);
        return;
      }
    }

    // Tenta adquirir o lock para esta edição
    if (!tryAcquireLock(id)) {
      console.error('[UPDATE_BLOCKED] Não foi possível adquirir lock para', id);
      return;
    }

    let updatedRoute = { ...route, [field]: value };

    // Validação específica para MONTES CLAROS + FÁBRICA quando editar observação
    if (field === 'observacao' && value) {
      const valid = validateDescargaTime(updatedRoute, value);
      if (!valid) {
        return; // Cancela a atualização
      }
    }

    // Se o campo alterado é 'motivo', reavalia o alerta de motorista
    if (field === 'motivo' && updatedRoute.motorista) {
      reevaluateMotoristAlert(id, updatedRoute.motorista, value);
      if ((value || '').trim() !== 'Manutenção') {
        updatedRoute.causaRaiz = '';
      }
    }

    // O status GERAL é apenas um marcador visual, NÃO altera o statusOp da rota
    // Calcula o status automaticamente baseado nos horários (isso só afeta a exibição da coluna STATUS)
    const config = userConfigs.find(c => c.operacao === updatedRoute.operacao);
    const { status, gap } = calculateStatusWithTolerance(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia || "00:00:00", updatedRoute.data);
    updatedRoute.statusOp = status;
    updatedRoute.tempo = gap;

    // Limpa motivo e observação se o status não for de atraso/adiantamento e não for manutenção
    if (status !== 'Atrasada' && status !== 'Adiantada' && status !== 'Programada' && status !== 'Previsto') {
      if (updatedRoute.motivo !== 'Manutenção') {
        updatedRoute.motivo = "";
        updatedRoute.observacao = "";
        updatedRoute.causaRaiz = "";
      }
    }

    setRoutes(prev => prev.map(r => r.id === id ? updatedRoute : r));
    setIsSyncing(true);

    try {
        await SharePointService.updateDeparture(await getAccessToken(), updatedRoute);
    } catch (e) {
        console.error('[UPDATE] Error:', e);
    } finally {
        releaseLock(id);
        setIsSyncing(false);
    }
  };

  const handleUpdateHistoryCell = (id: string, field: keyof RouteDeparture, value: string) => {
    // Normalização automática da placa: remove espaços e hífens
    if (field === 'placa') {
      value = value.replace(/[\s-]/g, '').toUpperCase();
    }

    // Apenas armazena a edição pendente (sem salvar no SharePoint ainda)
    setPendingHistoryEdits(prev => {
      const current = prev[id] || {};
      return {
        ...prev,
        [id]: { ...current, [field]: value }
      };
    });
  };

  // Salva todas as edições pendentes de uma vez (chamado ao pressionar Enter)
  const savePendingHistoryEdits = async () => {
    const editIds = Object.keys(pendingHistoryEdits);
    if (editIds.length === 0) return;

    console.log(`[HISTORY_BATCH_SAVE] Salvando ${editIds.length} edições pendentes...`);
    setIsSyncing(true);

    let successCount = 0;
    let errorCount = 0;

    try {
      const token = await getAccessToken();
      
      for (const id of editIds) {
        const edits = pendingHistoryEdits[id];
        const route = archivedResults.find(r => r.id === id);
        
        if (!route) {
          console.warn(`[HISTORY_SAVE] Rota ${id} não encontrada, pulando...`);
          continue;
        }

        try {
          // Cria rota atualizada com todas as edições
          const updatedRoute = { ...route, ...edits };

          // Recalcula status baseado nos horários se necessário
          if (edits.inicio || edits.saida || edits.data) {
            const config = userConfigs.find(c => c.operacao === updatedRoute.operacao);
            const { status, gap } = calculateStatusWithTolerance(
              updatedRoute.inicio, 
              updatedRoute.saida, 
              config?.tolerancia || "00:00:00", 
              updatedRoute.data
            );
            updatedRoute.statusOp = status;
            updatedRoute.tempo = gap;
          }

          // Salva no SharePoint
          await SharePointService.updateArchivedDeparture(token, updatedRoute);
          successCount++;
          console.log(`[HISTORY_SAVE] ✅ Rota ${id} atualizada com sucesso`);
        } catch (e: any) {
          errorCount++;
          console.error(`[HISTORY_SAVE] Erro ao atualizar ${id}:`, e.message);
        }
      }

      // Atualiza archivedResults localmente com as edições salvas (feedback imediato)
      setArchivedResults(prev => prev.map(r => {
        const edits = pendingHistoryEdits[r.id!];
        if (!edits) return r;

        const updated = { ...r, ...edits };

        // Recalcula status se horários foram alterados
        if (edits.inicio || edits.saida || edits.data) {
          const config = userConfigs.find(c => c.operacao === updated.operacao);
          const { status, gap } = calculateStatusWithTolerance(
            updated.inicio,
            updated.saida,
            config?.tolerancia || "00:00:00",
            updated.data
          );
          updated.statusOp = status;
          updated.tempo = gap;
        }

        return updated;
      }));

      // Limpa edições pendentes
      setPendingHistoryEdits({});
      setEditingHistoryId(null);
      setEditingHistoryField(null);

      // Feedback para o usuário
      if (errorCount === 0) {
        console.log(`[HISTORY_SAVE] ✅ ${successCount} edições salvas com sucesso!`);
      } else {
        alert(`⚠️ ${successCount} edições salvas, ${errorCount} falharam.`);
      }
    } catch (e: any) {
      console.error('[HISTORY_SAVE] Erro crítico:', e);
      alert('Erro ao salvar edições: ' + (e.message || 'Erro desconhecido'));
    } finally {
      setIsSyncing(false);
    }
  };

  // Funções de filtro do histórico (igual à tabela principal)
  const getHistoryColUniqueValues = (col: string) => {
    if (col === 'operacao') {
      return Array.from(new Set(archivedResults.map(r => r.operacao).filter(Boolean))).sort();
    }
    if (col === 'motorista') {
      return Array.from(new Set(archivedResults.map(r => r.motorista).filter(Boolean))).sort();
    }
    if (col === 'rota') {
      return Array.from(new Set(archivedResults.map(r => r.rota).filter(Boolean))).sort();
    }
    if (col === 'status') {
      return Array.from(new Set(archivedResults.map(r => r.statusOp).filter(Boolean))).sort();
    }
    if (col === 'placa') {
      return Array.from(new Set(archivedResults.map(r => r.placa).filter(Boolean))).sort();
    }
    if (col === 'motivo') {
      return Array.from(new Set(archivedResults.map(r => r.motivo).filter(Boolean))).sort();
    }
    if (col === 'semana') {
      return Array.from(new Set(archivedResults.map(r => r.semana).filter(Boolean))).sort();
    }
    return [];
  };

  const toggleHistoryColFilter = (col: string, value: string) => {
    setHistorySelectedFilters(prev => {
      const current = prev[col] || [];
      const updated = current.includes(value)
        ? current.filter(v => v !== value)
        : [...current, value];
      return { ...prev, [col]: updated };
    });
  };

  // Wrapper para toggle com stopPropagation (evita fechar dropdown ao clicar)
  const handleHistoryFilterClick = (e: React.MouseEvent, col: string, value: string) => {
    e.stopPropagation();
    toggleHistoryColFilter(col, value);
  };

  const clearHistoryColFilters = () => {
    setHistorySelectedFilters({});
    setHistoryColFilters({});
  };

  const hasHistoryActiveColFilters = Object.keys(historySelectedFilters).some(col => (historySelectedFilters[col] || []).length > 0);

  // Fecha dropdown ao clicar fora
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (historyActiveFilterCol && historyFilterDropdownRef.current && !historyFilterDropdownRef.current.contains(e.target as Node)) {
        setHistoryActiveFilterCol(null);
        setHistoryFilterAnchorEl(null);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [historyActiveFilterCol]);

  // Aplica filtros aos resultados do histórico e ordena por data/hora (da mais antiga para a mais recente)
  const filteredArchivedResults = useMemo(() => {
    let result = archivedResults;

    // Aplica os filtros selecionados
    if (hasHistoryActiveColFilters) {
      result = result.filter(r => {
        // Filtro por operação
        if (historySelectedFilters['operacao']?.length > 0) {
          if (!historySelectedFilters['operacao'].includes(r.operacao)) {
            return false;
          }
        }

        // Filtro por motorista
        if (historySelectedFilters['motorista']?.length > 0) {
          if (!historySelectedFilters['motorista'].includes(r.motorista || '')) {
            return false;
          }
        }

        // Filtro por rota
        if (historySelectedFilters['rota']?.length > 0) {
          if (!historySelectedFilters['rota'].includes(r.rota)) {
            return false;
          }
        }

        // Filtro por status
        if (historySelectedFilters['status']?.length > 0) {
          if (!historySelectedFilters['status'].includes(r.statusOp || '')) {
            return false;
          }
        }

        // Filtro por placa
        if (historySelectedFilters['placa']?.length > 0) {
          if (!historySelectedFilters['placa'].includes(r.placa || '')) {
            return false;
          }
        }

        // Filtro por motivo
        if (historySelectedFilters['motivo']?.length > 0) {
          if (!historySelectedFilters['motivo'].includes(r.motivo || '')) {
            return false;
          }
        }

        // Filtro por semana (vigência)
        if (historySelectedFilters['semana']?.length > 0) {
          if (!historySelectedFilters['semana'].includes(r.semana || '')) {
            return false;
          }
        }

        return true;
      });
    }

    // Ordena por data e hora de início (da mais antiga para a mais recente)
    let sorted = [...result].sort((a, b) => {
      // Converte data (pode vir em YYYY-MM-DD ou DD/MM/AAAA) para timestamp
      const parseDate = (dateStr: string) => {
        if (!dateStr) return 0;

        // Tenta formato YYYY-MM-DD (vem do SharePoint)
        const matchISO = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (matchISO) {
          const [, year, month, day] = matchISO;
          return new Date(Number(year), Number(month) - 1, Number(day)).getTime();
        }

        // Tenta formato DD/MM/AAAA
        const matchBR = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
        if (matchBR) {
          const [, day, month, year] = matchBR;
          return new Date(Number(year), Number(month) - 1, Number(day)).getTime();
        }

        return 0;
      };

      // Converte hora HH:MM:SS para segundos do dia
      const parseTime = (timeStr: string) => {
        if (!timeStr) return 0;
        const parts = timeStr.split(':');
        return Number(parts[0] || 0) * 3600 + Number(parts[1] || 0) * 60 + Number(parts[2] || 0);
      };

      // Compara primeiro por data
      const dateA = parseDate(a.data);
      const dateB = parseDate(b.data);
      if (dateA !== dateB) {
        return dateA - dateB;
      }

      // Se mesma data, compara por horário de início
      const timeA = parseTime(a.inicio || '');
      const timeB = parseTime(b.inicio || '');
      return timeA - timeB;
    });

    // Se ordenação por operação estiver ativada, reordena por operação (alfabética)
    if (historySortByOperacao) {
      sorted.sort((a, b) => {
        const opA = (a.operacao || '').toUpperCase();
        const opB = (b.operacao || '').toUpperCase();
        return opA.localeCompare(opB, 'pt-BR');
      });
    }

    return sorted;
  }, [archivedResults, historySelectedFilters, hasHistoryActiveColFilters, historySortByOperacao]);

  // Converte data YYYY-MM-DD para DD/MM/AAAA (parse manual para evitar fuso)
  const formatDateToBR = (dateString: string) => {
    if (!dateString) return '';
    const match = dateString.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
      const [, year, month, day] = match;
      return `${day}/${month}/${year}`;
    }
    return dateString;
  };

  // Converte data DD/MM/AAAA para YYYY-MM-DD (para input type="date")
  const formatDateToInput = (dateString: string) => {
    if (!dateString) return '';
    const match = dateString.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (match) {
      const [, day, month, year] = match;
      return `${year}-${month}-${day}`;
    }
    // Se já estiver no formato YYYY-MM-DD, retorna como está
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
      return dateString;
    }
    return dateString;
  };

  const handleExportToExcel = () => {
    if (archivedResults.length === 0) return;

    // Formata datas para o padrão brasileiro
    const formatDateBR = (dateString: string) => {
        if (!dateString) return '';
        // Se já estiver no formato DD/MM/AAAA, retorna como está
        if (dateString.includes('/') && /^\d{2}\/\d{2}\/\d{4}$/.test(dateString)) {
            return dateString;
        }
        // Se for ISO completo (com T e hora), usa new Date normal
        if (dateString.includes('T')) {
            try {
                const date = new Date(dateString);
                if (isNaN(date.getTime())) return dateString;
                return date.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
            } catch {
                return dateString;
            }
        }
        // Se for apenas data (YYYY-MM-DD), parse manualmente para evitar problema de fuso
        const match = dateString.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (match) {
            const [, year, month, day] = match;
            return `${day}/${month}/${year}`;
        }
        // Fallback
        try {
            const date = new Date(dateString);
            if (isNaN(date.getTime())) return dateString;
            return date.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
        } catch {
            return dateString;
        }
    };

    // Prepara os dados formatados (na mesma ordem da tabela: SEMANA, DATA, ROTA)
    const data = archivedResults.map(r => ({
      'Semana': r.semana || '',
      'Data': formatDateBR(r.data || ''),
      'Rota': r.rota,
      'Início': r.inicio || '',
      'Motorista': r.motorista || '',
      'Placa': r.placa || '',
      'Saída': r.saida || '',
      'Motivo': r.motivo || '',
      'Observação': r.observacao || '',
      'Operação': r.operacao,
      'Status': r.statusOp,
      'Tempo': r.tempo || ''
    }));

    // Cria workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);

    // Configura largura das colunas (na ordem: Semana, Data, Rota, ...)
    const colWidths = [
        { wch: 10 }, // Semana
        { wch: 12 }, // Data
        { wch: 25 }, // Rota
        { wch: 10 }, // Início
        { wch: 25 }, // Motorista
        { wch: 12 }, // Placa
        { wch: 10 }, // Saída
        { wch: 30 }, // Motivo
        { wch: 50 }, // Observação
        { wch: 20 }, // Operação
        { wch: 12 }, // Status
        { wch: 10 }  // Tempo
    ];
    ws['!cols'] = colWidths;

    // Estiliza o header (negrito, fundo colorido)
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_cell({ r: 0, c: C });
        if (!ws[address]) continue;
        ws[address].s = {
            font: { bold: true, color: { rgb: 'FFFFFF' } },
            fill: { fgColor: { rgb: '2563EB' } }, // Azul primary
            alignment: { horizontal: 'center', vertical: 'center' }
        };
    }

    // Adiciona tabela ao workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Histórico');

    // Gera nome do arquivo com datas no formato brasileiro
    const fileName = `Historico_CCO_${formatDateBR(histStart).replace(/\//g, '-')}_ate_${formatDateBR(histEnd).replace(/\//g, '-')}.xlsx`;

    // Download
    XLSX.writeFile(wb, fileName, {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    });
  };

  // Funções para menu de contexto (clique direito)
  const handleContextMenu = (e: React.MouseEvent, col: string) => {
    e.preventDefault();
    setContextMenu({
      visible: true,
      x: e.clientX,
      y: e.clientY,
      col
    });
  };

  const toggleColumnVisibility = (col: string) => {
    setHiddenColumns(prev => {
      const next = new Set(prev);
      if (next.has(col)) {
        next.delete(col);
      } else {
        next.add(col);
      }
      return next;
    });
    setContextMenu(prev => ({ ...prev, visible: false }));
  };

  const resetColumnSettings = () => {
    setColWidths({ rota: 140, data: 125, inicio: 95, motorista: 230, placa: 100, saida: 95, motivo: 170, observacao: 400, geral: 250, operacao: 140, status: 90, tempo: 90 });
    setHiddenColumns(new Set());
    setColFilters({});
    setSelectedFilters({});
    setIsSortByTimeEnabled(false);
    setIsSortByOperacao(false);
    sessionStorage.removeItem('route_departure_col_widths');
    sessionStorage.removeItem('route_departure_hidden_cols');
    sessionStorage.removeItem('route_departure_col_filters');
    sessionStorage.removeItem('route_departure_selected_filters');
    sessionStorage.removeItem('route_departure_sort_by_time');
  };

  const getRowStyle = (route: RouteDeparture | Partial<RouteDeparture>) => {
    if (route.id === 'ghost') return isDarkMode ? "bg-slate-800 italic text-slate-400" : "bg-slate-50 italic text-slate-500 border-l-4 border-dashed border-slate-300";
    const status = route.statusOp;
    const geralOK = route.statusGeral === 'OK';

    // Se a saída for "-", aplica estilo de atrasado crítico (não saiu)
    if (route.saida === '-') {
      return isDarkMode ? "bg-red-700/40 text-white font-bold border-l-[12px] border-red-800 shadow-lg" : "bg-red-200 text-red-900 font-bold border-l-[12px] border-red-600 shadow-sm";
    }

    // Atrasada sem saída (amarelo) — prioridade sobre GERAL
    // IMPORTANTE: "00:00:00" é horário válido (meia-noite), não considera como vazio
    if (status === 'Atrasada' && (!route.saida || route.saida === '')) {
      return isDarkMode ? "bg-yellow-500/30 text-yellow-100 font-bold border-l-[12px] border-yellow-500 shadow-lg" : "bg-amber-300 text-amber-950 font-bold border-l-[12px] border-amber-600 shadow-sm";
    }

    // Atrasada/Adiantada com saída (laranja) — prioridade sobre GERAL
    if (status === 'Atrasada' || status === 'Adiantada') {
      return isDarkMode ? "bg-orange-500/30 text-orange-100 font-bold border-l-[12px] border-orange-500 shadow-lg" : "bg-orange-300 text-orange-950 font-bold border-l-[12px] border-orange-600 shadow-sm";
    }

    // GERAL = OK + statusOp = OK → verde mais vivo (saiu dentro da tolerância)
    if (geralOK && status === 'OK') {
      return isDarkMode
        ? "bg-emerald-800/40 border-l-4 border-emerald-400 text-emerald-100"
        : "bg-emerald-200 border-l-4 border-emerald-500 text-slate-800";
    }

    if (status === 'Previsto') return isDarkMode ? "bg-slate-800 border-l-4 border-slate-600 text-slate-400" : "bg-white border-l-4 border-slate-300 text-slate-700";
    if (status === 'Programada') return isDarkMode ? "bg-slate-700 border-l-4 border-slate-500 text-slate-400" : "bg-slate-50 border-l-4 border-slate-400 text-slate-700";
    if (status === 'OK') return isDarkMode ? "bg-emerald-900/20 border-l-4 border-emerald-600" : "bg-emerald-50 border-l-4 border-emerald-500 text-slate-800";
    return isDarkMode ? "bg-slate-800 border-l-4 border-transparent" : "bg-white border-l-4 border-transparent text-slate-800";
  };

  // Verifica se há filtros ou ordenação ativos (para bloquear adição de rotas)
  const hasActiveFiltersOrSort = useMemo(() => {
    // Verifica se há filtros de texto ativos
    const hasTextFilters = Object.values(colFilters).some(val => val && val.length > 0);
    // Verifica se há filtros de seleção ativos
    const hasSelectionFilters = Object.values(selectedFilters).some(vals => vals && vals.length > 0);
    // Verifica se ordenação por data/horário está ativada
    const hasSortEnabled = isSortByTimeEnabled;

    return hasTextFilters || hasSelectionFilters || hasSortEnabled;
  }, [colFilters, selectedFilters, isSortByTimeEnabled]);

  const filteredRoutes = useMemo(() => {
    // Mapeia o nome da coluna para o campo real no objeto RouteDeparture
    const fieldMap: Record<string, string> = {
        'geral': 'statusGeral',
        'status': 'statusOp',
        'observacao': 'observacao',
        'motivo': 'motivo',
        'operacao': 'operacao',
        'tempo': 'tempo',
        'rota': 'rota',
        'data': 'data',
        'inicio': 'inicio',
        'motorista': 'motorista',
        'placa': 'placa',
        'saida': 'saida'
    };

    // Filtra primeiro pelas operações do usuário logado
    const myOps = new Set(userConfigs.map(c => c.operacao));
    let result = routes.filter(r => {
        // Se não houver operações configuradas, mostra todas (evita bloqueio se config não carregou)
        if (myOps.size === 0) return true;
        return myOps.has(r.operacao);
    });

    // Aplica filtros de coluna
    result = result.filter(r => {
        return (Object.entries(colFilters) as [string, string][]).every(([col, val]) => {
            if (!val) return true;
            const fieldName = fieldMap[col] || col;
            return r[fieldName as keyof RouteDeparture]?.toString().toLowerCase().includes(val.toLowerCase());
        }) && (Object.entries(selectedFilters) as [string, string[]][]).every(([col, vals]) => {
            if (!vals || vals.length === 0) return true;
            const fieldName = fieldMap[col] || col;
            return vals.includes(r[fieldName as keyof RouteDeparture]?.toString() || "");
        });
    });

    // Ordenação por data + horário (início)
    if (isSortByTimeEnabled) {
        result = [...result].sort((a, b) => {
            // Converte data + início em timestamp para comparação
            const getTimestamp = (route: RouteDeparture) => {
                if (!route.data || !route.inicio) return 0;
                const [year, month, day] = route.data.split('-').map(Number);
                const timeParts = route.inicio.split(':').map(Number);
                const date = new Date(year, month - 1, day, timeParts[0] || 0, timeParts[1] || 0, timeParts[2] || 0);
                return date.getTime();
            };
            return getTimestamp(a) - getTimestamp(b);
        });
    }

    // Se ordenação por operação estiver ativada, reordena por operação (alfabética)
    if (isSortByOperacao) {
        result = [...result].sort((a, b) => {
            const opA = (a.operacao || '').toUpperCase();
            const opB = (b.operacao || '').toUpperCase();
            return opA.localeCompare(opB, 'pt-BR');
        });
    }

    return result;
  }, [routes, colFilters, selectedFilters, isSortByTimeEnabled, isSortByOperacao, userConfigs]);

  // Cálculo dos indicadores GERAL e INTERNO - memoizado para evitar re-renderização desnecessária
  const [performanceIndicators, setPerformanceIndicators] = useState({ geral: '0.00', interno: '0.00' });
  
  useEffect(() => {
    // Usa TODAS as rotas do usuário, ignorando filtros de coluna
    const myOps = new Set(userConfigs.map(c => c.operacao));
    const allUserRoutes = routes.filter(r => {
      if (myOps.size === 0) return true;
      return myOps.has(r.operacao);
    });

    const total = allUserRoutes.length;
    if (total === 0) {
      setPerformanceIndicators({ geral: '0.00', interno: '0.00' });
      return;
    }

    // GERAL: (OK + PREVISTO) / total * 100
    const okPrevistoCount = allUserRoutes.filter(r =>
      r.statusOp === 'OK' || r.statusOp === 'Previsto'
    ).length;
    const geral = ((okPrevistoCount / total) * 100).toFixed(2);

    // INTERNO: (total - justificativas) / total * 100
    // Justificativas: Manutenção, Mão de obra, Logística
    const justificativas = ['Manutenção', 'Mão de obra', 'Logística'];
    const rotasComJustificativa = allUserRoutes.filter(r =>
      justificativas.includes(r.motivo)
    ).length;
    const rotasSemJustificativa = total - rotasComJustificativa;
    const interno = ((rotasSemJustificativa / total) * 100).toFixed(2);

    // Só atualiza se os valores mudaram (evita re-renderização desnecessária)
    setPerformanceIndicators(prev => {
      if (prev.geral === geral && prev.interno === interno) {
        return prev; // Sem mudança, não re-renderiza
      }
      return { geral, interno };
    });
  }, [routes.length, userConfigs.length]); // Apenas quantidade importa para estabilidade

  const tableColumns = [
    { id: 'rota', label: 'ROTA' }, { id: 'data', label: 'DATA' }, { id: 'inicio', label: 'INÍCIO' },
    { id: 'motorista', label: 'MOTORISTA' }, { id: 'placa', label: 'PLACA' }, { id: 'saida', label: 'SAÍDA' },
    { id: 'motivo', label: 'MOTIVO' }, { id: 'observacao', label: 'OBSERVAÇÃO' }, { id: 'geral', label: 'GERAL' },
    { id: 'operacao', label: 'OPERAÇÃO' }, { id: 'status', label: 'STATUS' }, { id: 'tempo', label: 'TEMPO' }
  ];

  if (isLoading) return <div className="h-full flex flex-col items-center justify-center text-primary-500 gap-4"><Loader2 size={48} className="animate-spin" /><p className="font-bold text-[10px] uppercase tracking-widest">Carregando Grid...</p></div>;

  return (
    <div className={`flex flex-col h-full p-4 overflow-hidden select-none font-sans animate-fade-in relative ${isDarkMode ? 'bg-[#020617]' : 'bg-gradient-to-br from-white via-slate-50/50 to-slate-50'}`}>
      {/* Header Section */}
      <div className="flex justify-between items-center mb-6 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className={`p-3 rounded-2xl shadow-lg ${isDarkMode ? 'bg-primary-600 text-white' : 'bg-primary-500 text-white'}`}><Clock size={20} /></div>
          <div>
            <h2 className={`text-xl font-black uppercase tracking-tight flex items-center gap-3 ${isDarkMode ? 'text-white' : 'text-slate-800'}`}>
              Controle de Saídas 
              <span className="inline-flex w-4 h-4">
                {isSyncing && <Loader2 size={16} className={`animate-spin ${isDarkMode ? 'text-primary-500' : 'text-primary-600'}`}/>}
              </span>
            </h2>
            <div className="flex items-center gap-2 mt-1">
              <p className={`text-[9px] font-bold uppercase tracking-widest flex items-center gap-2 ${isDarkMode ? 'text-slate-400' : 'text-slate-600'}`}>
                <ShieldCheck size={12} className="text-emerald-500"/> Operador: {currentUser.name}
              </p>
            </div>
          </div>
          {/* Indicadores GERAL, INTERNO e MINHAS ROTAS */}
          <div className="flex items-center gap-3 ml-8">
            <div className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[140px] ${isDarkMode ? 'bg-emerald-900/30 border border-emerald-700/50' : 'bg-emerald-100 border border-emerald-300'}`}>
              <div className="text-center flex-1">
                <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? 'text-emerald-400' : 'text-emerald-700'}`}>Geral</p>
                <p className={`text-2xl font-black leading-none ${isDarkMode ? 'text-emerald-400' : 'text-emerald-700'}`}>{performanceIndicators.geral}%</p>
              </div>
              <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse shrink-0"></div>
            </div>
            <div className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[140px] ${isDarkMode ? 'bg-blue-900/30 border border-blue-700/50' : 'bg-blue-100 border border-blue-300'}`}>
              <div className="text-center flex-1">
                <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>Interno</p>
                <p className={`text-2xl font-black leading-none ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>{performanceIndicators.interno}%</p>
              </div>
              <div className="w-2 h-2 bg-blue-500 rounded-full animate-pulse shrink-0"></div>
            </div>
            <div className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[140px] ${isDarkMode ? 'bg-purple-900/30 border border-purple-700/50' : 'bg-purple-100 border border-purple-300'}`}>
              <div className="text-center flex-1">
                <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? 'text-purple-400' : 'text-purple-700'}`}>Total Rotas</p>
                <p className={`text-2xl font-black leading-none ${isDarkMode ? 'text-purple-400' : 'text-purple-700'}`}>{filteredRoutes.length}</p>
              </div>
              <div className="w-2 h-2 bg-purple-500 rounded-full animate-pulse shrink-0"></div>
            </div>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <button onClick={() => setIsDarkMode(!isDarkMode)} className={`p-2 rounded-lg font-bold border transition-all shadow-sm ${isDarkMode ? 'bg-slate-800 text-yellow-400 border-slate-700 hover:bg-slate-700' : 'bg-white text-slate-700 border-slate-400 hover:bg-slate-50 hover:border-slate-500'}`} title={isDarkMode ? 'Modo Claro' : 'Modo Escuro'}>
            {isDarkMode ? <Sun size={18} /> : <Moon size={18} />}
          </button>
          <button onClick={() => setIsSortByTimeEnabled(!isSortByTimeEnabled)} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] transition-all shadow-sm ${isSortByTimeEnabled ? 'bg-primary-600 text-white border-primary-600' : isDarkMode ? 'bg-slate-800 text-slate-300 border-slate-700' : 'bg-white text-slate-800 border-slate-400 hover:bg-slate-50 hover:border-slate-500'}`}><SortAsc size={16} /> Horário</button>
          <button
            onClick={() => {
              if (hasActiveFiltersOrSort) {
                setFilterBlockReason('single');
                setIsFilterBlockModalOpen(true);
              } else {
                setIsAddRouteModalOpen(true);
              }
            }}
            className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] tracking-wide transition-all shadow-sm ${isDarkMode ? 'bg-slate-800 text-slate-300 hover:bg-slate-700 border-slate-700' : 'bg-white text-slate-800 hover:bg-slate-50 hover:border-slate-500 border-slate-400'}`}
          >
            <CheckCircle2 size={16} /> Adicionar Rota
          </button>
          <button onClick={() => setIsHistoryModalOpen(true)} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] tracking-wide transition-all shadow-sm ${isDarkMode ? 'bg-slate-800 text-slate-300 hover:bg-slate-700 border-slate-700' : 'bg-white text-slate-800 hover:bg-slate-50 hover:border-slate-500 border-slate-400'}`}><Database size={16} /> Histórico</button>
          <button onClick={handleArchiveAll} disabled={isSyncing || filteredRoutes.length === 0} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] tracking-wide transition-all shadow-sm disabled:opacity-50 disabled:cursor-not-allowed ${isDarkMode ? 'bg-slate-800 text-slate-300 hover:bg-slate-700 border-slate-700' : 'bg-white text-slate-800 hover:bg-slate-50 hover:border-slate-500 border-slate-400'}`}><Archive size={16} /> Arquivar</button>
        </div>
      </div>

      {bulkStatus?.active && (
          <div className="fixed inset-0 z-[500] bg-slate-950/60 backdrop-blur-sm flex items-center justify-center animate-in fade-in duration-300">
              <div className="bg-white dark:bg-slate-900 p-8 rounded-[2.5rem] border border-primary-500 shadow-2xl flex flex-col items-center gap-6 max-w-sm w-full">
                  <div className="relative"><Loader2 size={64} className="text-primary-600 animate-spin" /><Layers size={24} className="absolute inset-0 m-auto text-primary-400" /></div>
                  <div className="text-center"><h3 className="text-lg font-black uppercase text-slate-800 dark:text-white">Criando Lote</h3><p className="text-xs text-slate-400 font-bold uppercase mt-1 tracking-widest">{bulkStatus.current} de {bulkStatus.total}</p></div>
                  <div className="w-full bg-slate-200 dark:bg-slate-800 h-2 rounded-full overflow-hidden"><div className="h-full bg-primary-600 transition-all duration-300" style={{ width: `${(bulkStatus.current / bulkStatus.total) * 100}%` }}></div></div>
              </div>
          </div>
      )}

      {/* Table Section - flex-1 para ocupar espaço restante */}
      <div
        className={`flex-1 overflow-y-auto overflow-x-auto rounded-2xl border shadow-2xl relative scrollbar-thin ${isDarkMode ? 'bg-slate-900 border-slate-700/50' : 'bg-white border-slate-400 shadow-xl'}`}
        id="table-container"
        onFocusCapture={() => setIsEditingCell(true)}
        onBlurCapture={(e) => {
          // Só marca como não-editando se o novo foco está FORA da tabela
          // Evita flicker quando o usuário navega entre inputs dentro da mesma tabela
          const related = e.relatedTarget as HTMLElement | null;
          if (!related || !e.currentTarget.contains(related)) {
            setIsEditingCell(false);
          }
        }}
      >
        <div className="min-w-max" style={{ overflow: 'visible' }}>
        <table className="border-collapse" style={{ width: `${tableColumns.filter(col => !hiddenColumns.has(col.id)).reduce((acc, col) => acc + colWidths[col.id], 0) + 60}px`, tableLayout: 'fixed' }}>
          <thead className={`sticky top-0 z-[100] shadow-md ${isDarkMode ? 'bg-[#1e293b] text-white' : 'bg-gradient-to-r from-slate-800 to-slate-900 text-white shadow-slate-900/30'}`} style={{ position: 'sticky', top: 0, left: 0 }}>
            <tr className={`${isDarkMode ? 'bg-[#1e293b]' : 'bg-slate-800'}`}>
              {/* Célula extra na esquerda para cobrir vão */}
              <th className={`sticky left-0 z-[102] ${isDarkMode ? 'bg-[#1e293b]' : 'bg-slate-800'} w-[8px] p-0 m-0 border-none`} style={{ position: 'sticky', left: 0, minWidth: '8px', maxWidth: '8px' }}></th>
              {tableColumns.filter(col => !hiddenColumns.has(col.id)).map(col => (
                <th key={col.id} data-col={col.id} style={{ width: colWidths[col.id] }} className={`relative p-0 border ${isDarkMode ? 'border-slate-700/50' : 'border-slate-600/60'} text-[10px] font-black uppercase tracking-wider text-left group`}>
                  <div className="flex items-center justify-between px-3 h-[48px]">
                    <span onContextMenu={(e) => handleContextMenu(e, col.id)}>{col.label}</span>
                    <div className="flex items-center gap-1">
                      {col.id === 'operacao' && (
                        <button
                          onClick={(e) => { e.stopPropagation(); setIsSortByOperacao(!isSortByOperacao); }}
                          className={`p-1 rounded transition-all ${
                            isSortByOperacao
                              ? 'text-emerald-400'
                              : isDarkMode ? 'text-white/40 hover:text-white/60' : 'text-white/60 hover:text-white/80'
                          }`}
                          title={isSortByOperacao ? 'Ordenação por operação ativada' : 'Ordenar por operação (A-Z)'}
                        >
                          <ArrowUpDown size={11} />
                        </button>
                      )}
                      <button onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }} className={`p-1 rounded ${!!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0 ? 'text-yellow-400' : isDarkMode ? 'text-white/40' : 'text-white/60'}`}><Filter size={11} /></button>
                    </div>
                  </div>
                  {activeFilterCol === col.id && <FilterDropdown col={col.id} routes={routes} colFilters={colFilters} setColFilters={setColFilters} selectedFilters={selectedFilters} setSelectedFilters={setSelectedFilters} onClose={() => setActiveFilterCol(null)} dropdownRef={filterDropdownRef} />}
                  <div onMouseDown={(e) => { e.preventDefault(); resizingRef.current = { col: col.id, startX: e.clientX, startWidth: colWidths[col.id] }; }} className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10" />
                </th>
              ))}
              <th style={{ width: 60 }} className={`relative p-0 border ${isDarkMode ? 'border-slate-700/50' : 'border-slate-600/60'} text-[10px] font-black uppercase text-center ${isDarkMode ? 'bg-slate-900/50' : 'bg-slate-700/60'}`}>
                  {selectedIds.size > 0 ? (
                      <button onClick={handleDeleteSelected} className="p-1 text-red-500 hover:text-red-400 transition-colors" title="Deletar Selecionados"><Trash2 size={16} /></button>
                  ) : <Settings2 size={14} className="mx-auto opacity-40" />}
              </th>
            </tr>
          </thead>
          <tbody>
            {/* Renderiza rotas filtradas primeiro */}
            {filteredRoutes.map((route, rowIndex) => {
              const rowStyle = getRowStyle(route);
              const isSelected = selectedIds.has(route.id!);
              const isDelayed = route.statusOp === 'Atrasada' || route.statusOp === 'Adiantada';
              // IMPORTANTE: "00:00:00" é horário válido (meia-noite), considera como preenchido
              const isDelayedFilled = isDelayed && route.saida !== '';
              const isGhost = route.id === 'ghost';
              const inputClass = `w-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold uppercase transition-all ${isDelayedFilled ? (isDarkMode ? 'text-white placeholder-white/50' : 'text-slate-900 placeholder-slate-700') : isDarkMode ? 'text-slate-200 placeholder-slate-500' : 'text-slate-900 placeholder-slate-400'}`;

              return (
                <tr key={route.id} className={`${isSelected ? 'bg-primary-600/20' : rowStyle} group transition-all`} style={{ height: 'auto', minHeight: '48px', verticalAlign: 'top' }}>
                  {/* Célula extra na esquerda para alinhar com o header */}
                  <td className={`sticky left-0 z-[99] ${isDarkMode ? 'bg-slate-800' : 'bg-white'} w-[8px] p-0 m-0 border-none`} style={{ position: 'sticky', left: 0, minWidth: '8px', maxWidth: '8px' }}></td>
                  {tableColumns.filter(col => !hiddenColumns.has(col.id)).map(col => {
                    const cellKey = `${route.id}-${col.id}`;

                    if (col.id === 'rota') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle' }}>
                            {isGhost ? (
                                <textarea
                                  rows={1}
                                  value={route.rota}
                                  placeholder="Digite p/ criar..."
                                  onChange={(e) => {
                                      const val = e.target.value;
                                      updateCell(route.id!, 'rota', val);
                                      setTimeout(() => {
                                          e.target.style.height = 'auto';
                                          e.target.style.height = Math.max(e.target.scrollHeight, 48) + 'px';
                                      }, 0);
                                  }}
                                  className={`${inputClass} font-black resize-none whitespace-pre-wrap break-words min-h-[48px] text-center`}
                                />
                            ) : (
                                <div className="relative flex items-center justify-center p-2">
                                    <input type="text" value={route.rota} onChange={(e) => updateCell(route.id!, 'rota', e.target.value)} className={`${inputClass} font-black text-center w-full`} />
                                    {/* Indicador de alerta para rotas com histórico de problemas */}
                                    {routeAlerts[route.rota] && routeAlerts[route.rota].count > 0 && (
                                        <span
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setSelectedRouteAlert({ rota: route.rota, history: routeAlerts[route.rota].history });
                                          }}
                                          className="absolute right-2 top-1/2 -translate-y-1/2 inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 bg-red-500 hover:bg-red-600 text-white text-[9px] font-black rounded-full cursor-pointer transition-colors z-10"
                                          title={`${routeAlerts[route.rota].count} ocorrência(s) de atraso/adiantamento nos últimos 7 dias. Clique para ver histórico.`}
                                        >
                                          {routeAlerts[route.rota].count}
                                        </span>
                                    )}
                                </div>
                            )}
                        </td>
                      );
                    }

                    if (col.id === 'data') {
                      // Converte a data para exibição brasileira
                      const displayDate = formatDateToBR(route.data || '');
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                            type="text"
                            value={displayDate}
                            onChange={(e) => {
                              // Converte de DD/MM/AAAA para YYYY-MM-DD ao salvar
                              let val = e.target.value;
                              // Aplica máscara DD/MM/AAAA
                              val = val.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              updateCell(route.id!, 'data', val);
                            }}
                            onBlur={(e) => {
                              // Garante que o formato seja DD/MM/AAAA e converte para YYYY-MM-DD no SharePoint
                              let val = e.target.value;
                              const match = val.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                              if (match) {
                                const [, day, month, year] = match;
                                updateCell(route.id!, 'data', `${year}-${month}-${day}`);
                              }
                            }}
                            placeholder="DD/MM/AAAA"
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'inicio') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                              type="text"
                              key={route.id + '-inicio'}
                              value={route.inicio || ''}
                              placeholder="--:--:--"
                              onChange={(e) => {
                                  const masked = applyTimeMask(e.target.value);
                                  updateCell(route.id!, 'inicio', masked);
                              }}
                              onPaste={(e: any) => {
                                  const val = e.clipboardData.getData('text');
                                  if (val.includes('\n')) {
                                      e.preventDefault();
                                      handleMultilinePaste('inicio', rowIndex, val);
                                  }
                              }}
                              onBlur={(e) => {
                                  const formatted = formatTimeInput(e.target.value);
                                  updateCell(route.id!, 'inicio', formatted);
                              }}
                              className={`${inputClass} font-mono text-center`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'motorista') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="relative flex items-center">
                            <input
                                type="text"
                                value={route.motorista}
                                onChange={(e) => updateCell(route.id!, 'motorista', e.target.value)}
                                onPaste={(e: any) => {
                                    const val = e.clipboardData.getData('text');
                                    if (val.includes('\n')) {
                                        e.preventDefault();
                                        handleMultilinePaste('motorista', rowIndex, val);
                                    }
                                }}
                                className={`${inputClass} text-center w-full`}
                            />
                            {/* Indicador de alerta para motoristas com atrasos recorrentes por "Mão de obra" */}
                            {route.motorista && route.motivo === 'Mão de obra' && motoristAlerts[route.motorista] && motoristAlerts[route.motorista].count > 0 && (
                                <span
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setSelectedMotoristAlert({
                                      motorista: route.motorista,
                                      count: motoristAlerts[route.motorista].count,
                                      history: motoristAlerts[route.motorista].history
                                    });
                                  }}
                                  className="absolute right-2 top-1/2 -translate-y-1/2 inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 bg-orange-500 hover:bg-orange-600 text-white text-[9px] font-black rounded-full cursor-pointer transition-colors z-10"
                                  title={`${motoristAlerts[route.motorista].count} atraso(s) por "Mão de obra" nos últimos 30 dias. Clique para ver histórico.`}
                                >
                                  {motoristAlerts[route.motorista].count}
                                </span>
                            )}
                          </div>
                        </td>
                      );
                    }

                    if (col.id === 'placa') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                              type="text"
                              value={route.placa}
                              onChange={(e) => updateCell(route.id!, 'placa', e.target.value)}
                              onPaste={(e: any) => {
                                  const val = e.clipboardData.getData('text');
                                  if (val.includes('\n')) {
                                      e.preventDefault();
                                      handleMultilinePaste('placa', rowIndex, val);
                                  }
                              }}
                              className={`${inputClass} font-mono text-center`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'saida') {
                      // Verifica se esta célula está sendo editada
                      const isEditing = editingSaidaCell === route.id;

                      // Extrai apenas o horário para exibição se houver data completa, senão mostra o valor completo
                      const displayValue = (() => {
                        if (!route.saida || route.saida === '-') return route.saida || '';
                        // Se está editando, mostra o valor completo (com data)
                        if (isEditing) return route.saida;
                        // Se não está editando e tem data completa, mostra apenas horário
                        const dateTimeMatch = route.saida.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
                        if (dateTimeMatch) {
                          return `${dateTimeMatch[4]}:${dateTimeMatch[5]}:${dateTimeMatch[6]}`;
                        }
                        return route.saida;
                      })();

                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                            type="text"
                            value={displayValue}
                            placeholder="--:--:--"
                            onFocus={() => setEditingSaidaCell(route.id!)}
                            onBlur={(e) => {
                                setEditingSaidaCell(null);
                                const val = e.target.value;
                                if (val === '-') {
                                    updateCell(route.id!, 'saida', '-');
                                } else if (!val.trim()) {
                                    // Campo vazio - limpa
                                    updateCell(route.id!, 'saida', '');
                                } else {
                                    // Verifica se usuário digitou data completa (DD/MM/AAAA HH:MM:SS)
                                    const fullDateTimeMatch = val.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
                                    if (fullDateTimeMatch) {
                                        // Salva data e hora completas
                                        updateCell(route.id!, 'saida', val);
                                    } else {
                                        // Apenas horário - formata como HH:MM:SS
                                        const formatted = formatTimeInput(val);
                                        updateCell(route.id!, 'saida', formatted);
                                    }
                                }
                            }}
                            onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                    (e.target as HTMLInputElement).blur();
                                }
                            }}
                            onChange={(e) => {
                                const val = e.target.value;
                                // Permite digitação livre sem formatação automática
                                // A formatação só ocorre no onBlur
                                if (val === '-') {
                                    updateCell(route.id!, 'saida', '-');
                                } else {
                                    // Atualiza diretamente para permitir digitação fluida
                                    updateCell(route.id!, 'saida', val);
                                }
                            }}
                            onPaste={(e: any) => {
                                const pastedText = e.clipboardData.getData('text').trim();
                                e.preventDefault();

                                // Verifica se é paste de múltiplas linhas
                                if (pastedText.includes('\n')) {
                                    handleMultilinePaste('saida', rowIndex, pastedText);
                                } else {
                                    // Paste de valor único - insere diretamente
                                    updateCell(route.id!, 'saida', pastedText);
                                }
                            }}
                            className={`${inputClass} font-mono text-center`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'motivo') {
                      const isMaintenance = route.motivo === 'Manutenção';

                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          {(isDelayed || route.statusOp === 'Programada' || route.statusOp === 'Previsto') && !isGhost && (
                            <select
                              value={route.motivo}
                              onChange={(e) => updateCell(route.id!, 'motivo', e.target.value)}
                              className="w-full bg-white/20 dark:bg-slate-800/20 border-none px-2 py-1 text-[10px] font-bold text-inherit outline-none appearance-none cursor-pointer text-center"
                              disabled={!isMaintenance && route.motivo !== '' && route.statusOp === 'OK'}
                            >
                                <option value="" className="text-slate-800">---</option>
                                {MOTIVOS.map(m => (<option key={m} value={m} className="text-slate-800">{m}</option>))}
                            </select>
                          )}

                          {/* Campo vazio ou OK quando não é manutenção e já tem valor */}
                          {!isMaintenance && route.motivo !== '' && route.statusOp === 'OK' && !isGhost && (
                            <div className="w-full h-full flex items-center justify-center px-3 text-[10px] font-bold uppercase text-slate-400">
                              {route.motivo}
                            </div>
                          )}
                        </td>
                      );
                    }

                    if (col.id === 'observacao') {
                      const isMaintenance = route.motivo === 'Manutenção';
                      const canEdit = isMaintenance || (isDelayed || route.statusOp === 'Programada' || route.statusOp === 'Previsto');
                      const hasTemplates = route.motivo && OBSERVATION_TEMPLATES[route.motivo] && OBSERVATION_TEMPLATES[route.motivo].length > 0;
                      const checklistData = route.checklistMotorista || '';
                      const checklistFilled = isChecklistFilled(checklistData);
                      const tooltipContent = checklistData ? formatChecklistTooltip(checklistData) : '';

                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} relative align-top`} style={{ minHeight: '48px', height: 'auto', overflow: 'visible' }}>
                          {canEdit && !isGhost && (
                            <div className="flex items-start w-full relative p-0" style={{ minHeight: '44px', height: 'auto', overflow: 'visible' }}>
                              <textarea
                                ref={(el) => {
                                    obsTextareaRefs.current[route.id!] = el;
                                    // Ajusta altura inicial e quando largura mudar
                                    if (el) {
                                        setTimeout(() => {
                                            el.style.height = 'auto';
                                            el.style.height = el.scrollHeight + 'px';
                                        }, 0);
                                    }
                                }}
                                value={route.observacao || ""}
                                onChange={(e) => {
                                    updateCell(route.id!, 'observacao', e.target.value);
                                    e.target.style.height = 'auto';
                                    e.target.style.height = e.target.scrollHeight + 'px';
                                }}
                                onFocus={() => setActiveObsId(route.id!)}
                                placeholder="..."
                                disabled={!isMaintenance && route.motivo !== '' && route.statusOp === 'OK'}
                                className={`w-full bg-transparent outline-none border-none px-1 py-2 text-[11px] font-normal resize-none whitespace-pre-wrap break-words pr-20 text-left ${!isMaintenance && route.motivo !== '' && route.statusOp === 'OK' ? 'text-slate-400 cursor-not-allowed' : ''}`}
                                onInput={(e: any) => {
                                    e.target.style.height = 'auto';
                                    e.target.style.height = e.target.scrollHeight + 'px';
                                }}
                                style={{ wordBreak: 'break-word', overflowWrap: 'break-word', minHeight: '44px', height: 'auto', overflow: 'hidden' }}
                              />
                              <div className="absolute right-1 top-1/2 -translate-y-1/2 flex items-center gap-1">
                                {hasTemplates && (
                                  <button onClick={(e) => { e.stopPropagation(); setActiveObsId(activeObsId === route.id ? null : route.id!); }} className="p-1 opacity-60 hover:opacity-100"><ChevronDown size={14} /></button>
                                )}
                                {isMaintenance && (
                                  <div className="flex items-center gap-1">
                                    {/* Ícone de Check (V) - Checklist */}
                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        openChecklistEditModal(route.id!, checklistData);
                                      }}
                                      onMouseEnter={() => {
                                        if (tooltipTimeoutRef.current) {
                                          clearTimeout(tooltipTimeoutRef.current);
                                          tooltipTimeoutRef.current = null;
                                        }
                                        if (checklistFilled && tooltipContent) {
                                          setChecklistTooltip({ routeId: route.id!, content: tooltipContent });
                                        }
                                      }}
                                      onMouseLeave={() => {
                                        tooltipTimeoutRef.current = setTimeout(() => {
                                          setChecklistTooltip(null);
                                        }, 200);
                                      }}
                                      className="p-1 opacity-60 hover:opacity-100"
                                      title={checklistFilled ? 'Clique para editar checklist' : 'Clique para preencher checklist'}
                                    >
                                      <svg
                                        width="16"
                                        height="16"
                                        viewBox="0 0 24 24"
                                        fill="none"
                                        stroke={checklistFilled ? '#10b981' : '#ef4444'}
                                        strokeWidth="3"
                                        strokeLinecap="round"
                                        strokeLinejoin="round"
                                        className="transition-colors"
                                      >
                                        <polyline points="20 6 9 17 4 12" />
                                      </svg>
                                    </button>

                                    {/* Ícone de Ajuste de Horário */}
                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        // Abre o modal de edição de horários
                                        setIsTimeEditModalOpen(true);
                                        setTimeEditData({ routeId: route.id!, template: route.observacao || '', startTime: '', endTime: '' });
                                      }}
                                      className="p-1 opacity-60 hover:opacity-100 text-blue-500"
                                      title="Editar horários"
                                    >
                                      <Settings2 size={14} />
                                    </button>
                                  </div>
                                )}
                              </div>

                              {/* Tooltip do Checklist */}
                              {checklistTooltip && checklistTooltip.routeId === route.id && (
                                <div className="absolute bottom-full right-0 mb-2 z-[110] bg-slate-900 text-white text-[10px] font-bold px-3 py-2 rounded-xl shadow-2xl whitespace-pre-line max-w-[250px] text-left animate-in fade-in slide-in-from-bottom-1">
                                  {tooltipContent}
                                  <div className="absolute top-full right-4 -mt-1 border-4 border-transparent border-t-slate-900"></div>
                                </div>
                              )}

                              {activeObsId === route.id && hasTemplates && (
                                <div ref={obsDropdownRef} className="absolute top-full left-0 w-full z-[110] bg-white dark:bg-slate-800 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} rounded-xl shadow-2xl overflow-hidden animate-in fade-in slide-in-from-top-1">
                                  <div className="max-h-48 overflow-y-auto scrollbar-thin">{(route.motivo ? (OBSERVATION_TEMPLATES[route.motivo] || []) : []).map((template, tIdx) => {
                                    const hasTimes = template.match(/(\d{2}:\d{2}(?::\d{2})?)h/g);
                                    return (
                                      <div
                                        key={tIdx}
                                        onClick={() => openTimeEditModal(route.id!, template)}
                                        className="p-3 text-[10px] text-slate-700 dark:text-slate-300 hover:bg-primary-100 dark:hover:bg-slate-700 cursor-pointer border-b border-slate-100 dark:border-slate-700 flex items-center justify-between gap-2"
                                      >
                                        <div className="flex items-center gap-2 flex-1">
                                          <ChevronRight size={12} className="shrink-0 text-primary-500" />
                                          <span className="flex-1">{template}</span>
                                        </div>
                                        {hasTimes && hasTimes.length >= 2 && (
                                          <Settings2 size={12} className="text-blue-500 shrink-0" />
                                        )}
                                      </div>
                                    );
                                  })}</div>
                                </div>
                              )}
                            </div>
                          )}

                          {/* Mostra apenas leitura quando não é manutenção e já tem valor */}
                          {!isMaintenance && route.observacao && route.statusOp === 'OK' && !isGhost && (
                            <div className="w-full px-3 py-2 text-[11px] text-slate-400 whitespace-pre-wrap break-words" style={{ wordBreak: 'break-word', overflowWrap: 'break-word', minHeight: '44px', height: 'auto' }}>
                              {route.observacao}
                            </div>
                          )}
                        </td>
                      );
                    }

                    if (col.id === 'geral') {
                      const hasCopiedValue = copiedGeralStatus && copiedGeralStatus !== '';
                      const isHovered = hoveredGeralCell === route.id;
                      return (
                        <td key={cellKey} data-col-cell="geral" className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} relative`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <button
                            onClick={() => {
                              const newValue = route.statusGeral === 'OK' ? '' : 'OK';
                              updateCell(route.id!, 'statusGeral', newValue);
                            }}
                            onMouseEnter={() => setHoveredGeralCell(route.id!)}
                            onMouseLeave={() => setHoveredGeralCell(null)}
                            className="absolute inset-0 w-full h-full flex items-center justify-center font-bold text-[10px] transition-all border-none outline-none"
                            style={{
                              backgroundColor: route.statusGeral === 'OK' ? '#059669' : isHovered ? (isDarkMode ? '#334155' : '#f1f5f9') : 'transparent',
                              color: route.statusGeral === 'OK' ? '#ffffff' : isDarkMode ? '#94a3b8' : '#475569'
                            }}
                            title={hasCopiedValue ? `Valor copiado: "${copiedGeralStatus}" - Selecione rotas e pressione Ctrl+V para colar` : 'Clique para alternar OK/vazio'}
                          >
                            {route.statusGeral || '---'}
                          </button>
                          {/* Indicador visual de valor copiado (ponto verde sem animação) */}
                          {hasCopiedValue && (
                            <div className="absolute top-1 right-1 w-2 h-2 bg-emerald-500 rounded-full z-10" title={`Valor copiado: "${copiedGeralStatus}" - Selecione rotas e pressione Ctrl+V`}></div>
                          )}
                        </td>
                      );
                    }

                    if (col.id === 'operacao') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <select value={route.operacao} onChange={(e) => updateCell(route.id!, 'operacao', e.target.value)} className="w-full h-full bg-transparent border-none text-[9px] font-black text-center uppercase">
                            <option value="">OP...</option>
                            {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                          </select>
                        </td>
                      );
                    }

                    if (col.id === 'status') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="w-full flex items-center justify-center">
                            <span className={`px-2 py-0.5 rounded-full text-[8px] font-black border whitespace-nowrap ${route.statusOp === 'OK' ? 'bg-emerald-100 border-emerald-400 text-emerald-800' : route.statusOp === 'Atrasada' ? 'bg-yellow-100 border-yellow-400 text-yellow-800' : route.statusOp === 'Programada' ? 'bg-slate-200 border-slate-400 text-slate-600' : route.statusOp === 'Previsto' ? 'bg-slate-100 border-slate-400 text-slate-500' : 'bg-red-100 border-red-400 text-red-800'}`}>
                              {route.statusOp}
                            </span>
                          </div>
                        </td>
                      );
                    }

                    if (col.id === 'tempo') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="w-full flex items-center justify-center text-[10px] font-bold">
                            {route.tempo}
                          </div>
                        </td>
                      );
                    }

                    return null;
                  })}
                  <td className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} flex items-center justify-center gap-1`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                    {!isGhost && (
                      <>
                        <button onClick={() => toggleSelection(route.id!)} className={`p-1.5 rounded-lg transition-colors ${isSelected ? 'text-primary-500 bg-primary-500/10' : isDarkMode ? 'text-slate-400 hover:bg-slate-700' : 'text-slate-500 hover:bg-slate-200'}`}>
                          {isSelected ? <CheckSquare size={16}/> : <Square size={16}/>}
                        </button>
                        <button onClick={() => handleDeleteRoute(route.id!)} className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-red-500/10 rounded-lg transition-colors">
                          <Trash2 size={16} />
                        </button>
                      </>
                    )}
                  </td>
                </tr>
              );
            })}
            
            {/* Ghost Row - SEMPRE no final da tabela */}
            {(() => {
              const route = ghostRow;
              const rowStyle = getRowStyle(route);
              const isSelected = selectedIds.has(route.id!);
              const isDelayed = route.statusOp === 'Atrasada' || route.statusOp === 'Adiantada';
              // IMPORTANTE: "00:00:00" é horário válido (meia-noite), considera como preenchido
              const isDelayedFilled = isDelayed && route.saida !== '';
              const inputClass = `w-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold uppercase transition-all ${isDelayedFilled ? (isDarkMode ? 'text-white placeholder-white/50' : 'text-slate-900 placeholder-slate-700') : isDarkMode ? 'text-slate-200 placeholder-slate-500' : 'text-slate-900 placeholder-slate-400'}`;

              return (
                <tr key={route.id} className={`${isSelected ? 'bg-primary-600/20' : rowStyle} group transition-all`} style={{ height: 'auto', minHeight: '48px', verticalAlign: 'top' }}>
                  {/* Célula extra na esquerda para alinhar com o header */}
                  <td className={`sticky left-0 z-[99] ${isDarkMode ? 'bg-slate-800' : 'bg-white'} w-[8px] p-0 m-0 border-none`} style={{ position: 'sticky', left: 0, minWidth: '8px', maxWidth: '8px' }}></td>
                  {tableColumns.filter(col => !hiddenColumns.has(col.id)).map(col => {
                    const cellKey = `${route.id}-${col.id}`;

                    if (col.id === 'rota') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle' }}>
                            <textarea
                              rows={1}
                              value={route.rota}
                              placeholder="Digite p/ criar..."
                              onChange={(e) => {
                                  const val = e.target.value;
                                  updateCell(route.id!, 'rota', val);
                                  setTimeout(() => {
                                      e.target.style.height = 'auto';
                                      e.target.style.height = Math.max(e.target.scrollHeight, 48) + 'px';
                                  }, 0);
                              }}
                              className={`${inputClass} font-black resize-none whitespace-pre-wrap break-words min-h-[48px] text-center`}
                            />
                        </td>
                      );
                    }

                    if (col.id === 'data') {
                      // Converte a data para exibição brasileira
                      const displayDate = formatDateToBR(route.data || '');
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                            type="text"
                            value={displayDate}
                            onChange={(e) => {
                              // Converte de DD/MM/AAAA para YYYY-MM-DD ao salvar
                              let val = e.target.value;
                              // Aplica máscara DD/MM/AAAA
                              val = val.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              updateCell(route.id!, 'data', val);
                            }}
                            onBlur={(e) => {
                              // Garante que o formato seja DD/MM/AAAA e converte para YYYY-MM-DD no SharePoint
                              let val = e.target.value;
                              const match = val.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                              if (match) {
                                const [, day, month, year] = match;
                                updateCell(route.id!, 'data', `${year}-${month}-${day}`);
                              }
                            }}
                            placeholder="DD/MM/AAAA"
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'inicio') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                              type="text"
                              key={`${route.id}-inicio`}
                              value={route.inicio || ''}
                              placeholder="HH:MM:SS"
                              onChange={(e) => {
                                  const masked = applyTimeMask(e.target.value);
                                  updateCell(route.id!, 'inicio', masked);
                              }}
                              onBlur={(e) => {
                                  const formatted = formatTimeInput(e.target.value);
                                  updateCell(route.id!, 'inicio', formatted);
                              }}
                              className={`${inputClass} font-mono text-center`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'motorista') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="relative flex items-center">
                            <input
                                type="text"
                                value={route.motorista}
                                onChange={(e) => updateCell(route.id!, 'motorista', e.target.value)}
                                className={`${inputClass} text-center w-full`}
                            />
                            {/* Indicador de alerta para motoristas com atrasos recorrentes por "Mão de obra" */}
                            {route.motorista && route.motivo === 'Mão de obra' && motoristAlerts[route.motorista] && motoristAlerts[route.motorista].count > 0 && (
                                <span
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setSelectedMotoristAlert({
                                      motorista: route.motorista,
                                      count: motoristAlerts[route.motorista].count,
                                      history: motoristAlerts[route.motorista].history
                                    });
                                  }}
                                  className="absolute right-2 top-1/2 -translate-y-1/2 inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 bg-orange-500 hover:bg-orange-600 text-white text-[9px] font-black rounded-full cursor-pointer transition-colors z-10"
                                  title={`${motoristAlerts[route.motorista].count} atraso(s) por "Mão de obra" nos últimos 30 dias. Clique para ver histórico.`}
                                >
                                  {motoristAlerts[route.motorista].count}
                                </span>
                            )}
                          </div>
                        </td>
                      );
                    }

                    if (col.id === 'placa') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                              type="text"
                              value={route.placa}
                              onChange={(e) => {
                                // Normalização: remove espaços, hífens e converte para maiúsculo
                                const cleanValue = e.target.value.replace(/[\s-]/g, '').toUpperCase();
                                updateCell(route.id!, 'placa', cleanValue);
                              }}
                              className={`${inputClass} font-mono text-center uppercase`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'saida') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <input
                              type="text"
                              key={`${route.id}-saida`}
                              value={route.saida || ''}
                              placeholder="HH:MM:SS"
                              onChange={(e) => {
                                  const val = e.target.value;
                                  if (val === '-') {
                                      updateCell(route.id!, 'saida', '-');
                                  } else {
                                      const masked = applyTimeMask(val);
                                      updateCell(route.id!, 'saida', masked);
                                  }
                              }}
                              onBlur={(e) => {
                                  const val = e.target.value;
                                  if (val === '-') {
                                      updateCell(route.id!, 'saida', '-');
                                  } else {
                                      const formatted = formatTimeInput(val);
                                      updateCell(route.id!, 'saida', formatted);
                                  }
                              }}
                              className={`${inputClass} font-mono text-center`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'motivo') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="w-full h-full flex items-center justify-center px-3 py-2 text-[10px] text-slate-400 italic">---</div>
                        </td>
                      );
                    }

                    if (col.id === 'observacao') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'top', minHeight: '48px', height: 'auto' }}>
                          <div className="w-full px-1 py-2 text-[10px] text-slate-400 italic whitespace-pre-wrap break-words text-left" style={{ wordBreak: 'break-word', overflowWrap: 'break-word' }}>---</div>
                        </td>
                      );
                    }

                    if (col.id === 'geral') {
                      return (
                        <td key={cellKey} data-col-cell="geral" className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="w-full h-full flex items-center justify-center px-3 py-2 text-[10px] text-slate-400 italic">---</div>
                        </td>
                      );
                    }

                    if (col.id === 'operacao') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <select value={route.operacao} onChange={(e) => updateCell(route.id!, 'operacao', e.target.value)} className="w-full h-full bg-transparent border-none text-[9px] font-black text-center uppercase">
                            <option value="">OP...</option>
                            {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                          </select>
                        </td>
                      );
                    }

                    if (col.id === 'status') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="w-full flex items-center justify-center">
                            <span className={`px-2 py-0.5 rounded-full text-[8px] font-black border whitespace-nowrap ${route.statusOp === 'OK' ? 'bg-emerald-100 border-emerald-400 text-emerald-800' : route.statusOp === 'Atrasada' ? 'bg-yellow-100 border-yellow-400 text-yellow-800' : route.statusOp === 'Programada' ? 'bg-slate-200 border-slate-400 text-slate-600' : route.statusOp === 'Previsto' ? 'bg-slate-100 border-slate-400 text-slate-500' : 'bg-red-100 border-red-400 text-red-800'}`}>
                              {route.statusOp}
                            </span>
                          </div>
                        </td>
                      );
                    }

                    if (col.id === 'tempo') {
                      return (
                        <td key={cellKey} className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                          <div className="w-full flex items-center justify-center text-[10px] font-bold">
                            {route.tempo}
                          </div>
                        </td>
                      );
                    }

                    return null;
                  })}
                  <td className={`p-0 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} flex items-center justify-center gap-1`} style={{ verticalAlign: 'middle', minHeight: '48px' }}>
                    {/* Ghost row não tem botões de ação */}
                  </td>
                </tr>
              );
            })()}
          </tbody>
        </table>
        </div>
      </div>

      {/* Menu de Contexto (Clique Direito) */}
      {contextMenu.visible && (
        <div
          ref={contextMenuRef}
          className="fixed z-[1000] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-2xl py-2 min-w-[200px] animate-in fade-in zoom-in-95 duration-150"
          style={{ left: contextMenu.x, top: contextMenu.y }}
        >
          <div className="px-3 py-2 border-b border-slate-100 dark:border-slate-700">
            <p className="text-[10px] font-black uppercase text-slate-400">Coluna: {contextMenu.col?.toUpperCase()}</p>
          </div>
          <button
            onClick={() => toggleColumnVisibility(contextMenu.col!)}
            className="w-full px-4 py-2 text-left text-[11px] font-bold text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center gap-2"
          >
            {hiddenColumns.has(contextMenu.col!) ? <Check size={14} className="text-green-500" /> : <Square size={14} />}
            {hiddenColumns.has(contextMenu.col!) ? 'Mostrar Coluna' : 'Ocultar Coluna'}
          </button>
          <div className="border-t border-slate-100 dark:border-slate-700 my-1"></div>
          <button
            onClick={resetColumnSettings}
            className="w-full px-4 py-2 text-left text-[11px] font-bold text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors flex items-center gap-2"
          >
            <RefreshCw size={14} /> Resetar Configurações
          </button>
        </div>
      )}

      {isBulkMappingModalOpen && (
          <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[300] flex items-center justify-center p-4">
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-primary-500 shadow-2xl animate-in zoom-in">
                  <div className="flex items-center gap-3 text-primary-500 mb-6 font-black uppercase text-xs"><Layers size={24} /> Atribuir Planta para Lote</div>
                  <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">Você colou <span className="text-primary-500 font-black">{pendingBulkRoutes.length} rotas</span>. Escolha a operação:</p>
                  <div className="grid grid-cols-2 gap-3 max-h-64 overflow-y-auto pr-2 scrollbar-thin">{userConfigs.map(c => ( <button key={c.operacao} onClick={() => handleBulkCreateSave(c.operacao)} className="p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-2xl hover:bg-primary-600 hover:text-white transition-all font-black text-xs uppercase">{c.operacao}</button> ))}</div>
                  <button onClick={() => setIsBulkMappingModalOpen(false)} className="w-full mt-6 py-4 text-[10px] font-black uppercase text-slate-400">Cancelar</button>
              </div>
          </div>
      )}

      {isMappingModalOpen && (
          <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[300] flex items-center justify-center p-4">
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-primary-500 animate-in zoom-in">
                  <div className="flex items-center gap-3 text-primary-500 mb-6 font-black uppercase text-xs"><LinkIcon size={24} /> Vínculo Necessário</div>
                  <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">A rota <span className="text-primary-500 font-black">{pendingMappingRoute}</span> não possui planta vinculada:</p>
                  <div className="grid grid-cols-2 gap-3">{userConfigs.map(c => ( <button key={c.operacao} onClick={async () => { const tok = await getAccessToken(); SharePointService.addRouteOperationMapping(tok, pendingMappingRoute!, c.operacao); setGhostRow(prev => ({...prev, operacao: c.operacao})); setIsMappingModalOpen(false); }} className="p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 rounded-2xl hover:bg-primary-600 hover:text-white transition-all font-black text-xs uppercase">{c.operacao}</button> ))}</div>
                  <button onClick={() => { setIsMappingModalOpen(false); setGhostRow(prev => ({...prev, rota: ''})); }} className="w-full mt-6 py-4 text-[10px] font-black uppercase text-slate-400">Cancelar</button>
              </div>
          </div>
      )}

      {/* Modal de Bloqueio - Filtros/Ordenação Ativos */}
      {isFilterBlockModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[250] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-lg">
            <div className="bg-gradient-to-r from-amber-500 to-orange-500 p-6 flex justify-between items-center text-white">
              <div className="flex items-center gap-3">
                <AlertTriangle size={28} />
                <h3 className="font-black uppercase tracking-widest text-base">Atenção</h3>
              </div>
              <button
                onClick={() => setIsFilterBlockModalOpen(false)}
                className="p-2 hover:bg-white/20 rounded-lg transition-colors"
              >
                <X size={24} />
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div className="flex items-start gap-4">
                <div className="w-14 h-14 rounded-full bg-amber-100 dark:bg-amber-900/30 flex items-center justify-center shrink-0">
                  <Filter size={28} className="text-amber-600 dark:text-amber-400" />
                </div>
                <div className="flex-1">
                  <h4 className="text-sm font-black uppercase text-slate-800 dark:text-white mb-2">
                    {filterBlockReason === 'single' && 'Adição de Rota Bloqueada'}
                    {filterBlockReason === 'bulk' && 'Criação em Lote Bloqueada'}
                    {filterBlockReason === 'ghost' && 'Criação de Rota Bloqueada'}
                  </h4>
                  <p className="text-xs text-slate-600 dark:text-slate-300 font-medium leading-relaxed">
                    {filterBlockReason === 'single' && 'O botão "Adicionar Rota" está desabilitado porque há filtros ou ordenação ativos.'}
                    {filterBlockReason === 'bulk' && 'A criação de múltiplas rotas foi bloqueada porque há filtros ou ordenação ativos.'}
                    {filterBlockReason === 'ghost' && 'A criação automática de rota foi bloqueada porque há filtros ou ordenação ativos.'}
                  </p>
                </div>
              </div>

              <div className="bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700 rounded-xl p-4">
                <p className="text-[10px] font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wide mb-3">
                  Filtros/Ordenação detectados:
                </p>
                <div className="space-y-2">
                  {Object.values(colFilters).some(val => val && val.length > 0) && (
                    <div className="flex items-center gap-2 text-xs text-red-600 dark:text-red-400">
                      <X size={14} className="shrink-0" />
                      <span className="font-medium">Filtros de texto ativos nas colunas</span>
                    </div>
                  )}
                  {Object.values(selectedFilters).some(vals => vals && vals.length > 0) && (
                    <div className="flex items-center gap-2 text-xs text-red-600 dark:text-red-400">
                      <X size={14} className="shrink-0" />
                      <span className="font-medium">Filtros de seleção ativos nas colunas</span>
                    </div>
                  )}
                  {isSortByTimeEnabled && (
                    <div className="flex items-center gap-2 text-xs text-red-600 dark:text-red-400">
                      <X size={14} className="shrink-0" />
                      <span className="font-medium">Ordenação por Horário ativada</span>
                    </div>
                  )}
                </div>
              </div>

              <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-xl p-4">
                <p className="text-[10px] font-bold text-amber-700 dark:text-amber-300 uppercase tracking-wide mb-3">
                  Como liberar:
                </p>
                <ul className="text-[11px] font-medium text-amber-800 dark:text-amber-200 space-y-2">
                  <li className="flex items-start gap-2">
                    <span className="bg-amber-200 dark:bg-amber-800 text-amber-800 dark:text-amber-200 w-5 h-5 rounded-full flex items-center justify-center text-[9px] font-black shrink-0">1</span>
                    <span>Clique nos ícones amarelos de filtro no cabeçalho da tabela e limpe os filtros</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="bg-amber-200 dark:bg-amber-800 text-amber-800 dark:text-amber-200 w-5 h-5 rounded-full flex items-center justify-center text-[9px] font-black shrink-0">2</span>
                    <span>Clique no botão "HORÁRIO" para desativar a ordenação por data/horário</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="bg-amber-200 dark:bg-amber-800 text-amber-800 dark:text-amber-200 w-5 h-5 rounded-full flex items-center justify-center text-[9px] font-black shrink-0">3</span>
                    <span>Tente adicionar a rota novamente</span>
                  </li>
                </ul>
              </div>

              <button
                onClick={() => setIsFilterBlockModalOpen(false)}
                className="w-full py-4 bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 text-white rounded-xl font-black uppercase text-xs tracking-wider transition-all shadow-lg shadow-amber-500/20"
              >
                Entendi, vou remover os filtros
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Adicionar Rota */}
      {isAddRouteModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[200] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-lg">
            <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white">
              <div className="flex items-center gap-3">
                <CheckCircle2 size={24} />
                <h3 className="font-black uppercase tracking-widest text-base">Adicionar Rota</h3>
              </div>
              <button onClick={() => setIsAddRouteModalOpen(false)} className="p-2 hover:bg-slate-700 rounded-lg transition-colors">
                <X size={24} />
              </button>
            </div>
            <div className="p-6 space-y-4">
              {/* Operação */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Operação *
                </label>
                <select
                  value={newRouteData.operacao}
                  onChange={e => setNewRouteData({ ...newRouteData, operacao: e.target.value })}
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                >
                  <option value="">Selecione a operação</option>
                  {userConfigs.map(config => (
                    <option key={config.operacao} value={config.operacao}>{config.operacao}</option>
                  ))}
                </select>
              </div>

              {/* Rota */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Rota *
                </label>
                <input
                  type="text"
                  value={newRouteData.rota}
                  onChange={e => setNewRouteData({ ...newRouteData, rota: e.target.value })}
                  placeholder="Ex: ROTA 01"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                />
              </div>

              {/* Início */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Horário de Início *
                </label>
                <input
                  type="text"
                  value={newRouteData.inicio}
                  onChange={e => {
                    let value = e.target.value.replace(/\D/g, ''); // Remove não dígitos
                    if (value.length > 6) value = value.slice(0, 6); // Limita a 6 dígitos (HHMMSS)
                    
                    // Aplica máscara HH:MM:SS
                    if (value.length >= 6) {
                      value = `${value.slice(0, 2)}:${value.slice(2, 4)}:${value.slice(4, 6)}`;
                    } else if (value.length >= 4) {
                      value = `${value.slice(0, 2)}:${value.slice(2, 4)}:${value.slice(4)}`;
                    } else if (value.length >= 2) {
                      value = `${value.slice(0, 2)}:${value.slice(2)}`;
                    }
                    
                    setNewRouteData({ ...newRouteData, inicio: value });
                  }}
                  placeholder="HH:MM:SS"
                  maxLength={8}
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors font-mono"
                />
              </div>

              {/* Motorista */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Motorista *
                </label>
                <input
                  type="text"
                  value={newRouteData.motorista}
                  onChange={e => setNewRouteData({ ...newRouteData, motorista: e.target.value })}
                  placeholder="Nome do motorista"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                />
              </div>

              {/* Placa */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Placa *
                </label>
                <input
                  type="text"
                  value={newRouteData.placa}
                  onChange={e => {
                    // Normalização: remove espaços, hífens e converte para maiúsculo
                    const cleanValue = e.target.value.replace(/[\s-]/g, '').toUpperCase();
                    setNewRouteData({ ...newRouteData, placa: cleanValue });
                  }}
                  placeholder="Ex: ABC1D23"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors uppercase"
                />
              </div>

              {/* Botões de ação */}
              <div className="flex gap-3 pt-4">
                <button
                  onClick={() => setIsAddRouteModalOpen(false)}
                  className="flex-1 py-3 bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-black uppercase text-[11px] rounded-xl hover:bg-slate-300 dark:hover:bg-slate-700 transition-colors"
                >
                  Cancelar
                </button>
                <button
                  onClick={handleAddRoute}
                  disabled={isAddingRoute}
                  className="flex-1 py-3 bg-primary-600 text-white font-black uppercase text-[11px] rounded-xl hover:bg-primary-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                >
                  {isAddingRoute ? <Loader2 size={16} className="animate-spin" /> : <CheckCircle2 size={16} />}
                  {isAddingRoute ? 'Adicionando...' : 'Lançar Rota'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {isHistoryModalOpen && (
          <div className={`fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[200] flex items-center justify-center p-4 ${isHistoryFullscreen ? 'p-0' : ''}`}>
              <div className={`bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-[2.5rem] shadow-2xl w-full flex flex-col ${isHistoryFullscreen ? 'max-w-none w-full h-full rounded-none' : 'max-w-7xl max-h-[90vh]'}`}>
                  <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white shrink-0">
                      <div className="flex items-center gap-4">
                          <Database size={24} />
                          <h3 className="font-black uppercase tracking-widest text-base">Histórico Definitivo</h3>
                          {archivedResults.length > 0 && (
                              <span className="text-[10px] font-bold text-slate-400 bg-slate-800 px-3 py-1 rounded-full">
                                  {archivedResults.length} registro(s)
                              </span>
                          )}
                          {/* Indicador de edições pendentes */}
                          {Object.keys(pendingHistoryEdits).length > 0 && (
                              <span className="text-[10px] font-black uppercase tracking-widest text-amber-400 bg-amber-900/30 px-3 py-1 rounded-full border border-amber-600 animate-pulse">
                                  {Object.keys(pendingHistoryEdits).length} alteração(ões) pendente(s) - Pressione ENTER para salvar
                              </span>
                          )}
                      </div>
                      <div className="flex items-center gap-2">
                          {/* Botão de salvar edições pendentes */}
                          {Object.keys(pendingHistoryEdits).length > 0 && (
                              <button
                                  onClick={savePendingHistoryEdits}
                                  disabled={isSyncing}
                                  className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-black uppercase text-[10px] rounded-xl shadow-lg disabled:opacity-50 disabled:cursor-not-allowed"
                                  title="Salvar alterações (Enter)"
                              >
                                  {isSyncing ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                                  SALVAR ({Object.keys(pendingHistoryEdits).length})
                              </button>
                          )}
                          <button
                              onClick={() => setIsHistoryFullscreen(!isHistoryFullscreen)}
                              className="p-2 hover:bg-slate-700 rounded-lg transition-colors"
                              title={isHistoryFullscreen ? 'Sair da tela cheia' : 'Tela cheia'}
                          >
                              {isHistoryFullscreen ? <Minimize2 size={20} /> : <Maximize2 size={20} />}
                          </button>
                          <button onClick={() => setIsHistoryModalOpen(false)} className="p-2 hover:bg-slate-700 rounded-lg transition-colors">
                              <X size={28} />
                          </button>
                      </div>
                  </div>
                  <div className="p-6 bg-slate-50 dark:bg-slate-900 border-b dark:border-slate-800 grid grid-cols-4 gap-4 shrink-0">
                      <input type="date" value={histStart} onChange={e => setHistStart(e.target.value)} className="p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white" />
                      <input type="date" value={histEnd} onChange={e => setHistEnd(e.target.value)} className="p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white" />
                      <button onClick={handleSearchArchive} disabled={isSearchingArchive} className="py-3 bg-primary-600 text-white font-black uppercase text-[11px] rounded-xl flex items-center justify-center gap-2 hover:bg-primary-700 shadow-lg">
                          {isSearchingArchive ? <Loader2 size={16} className="animate-spin" /> : <Search size={16} />} BUSCAR
                      </button>
                      {archivedResults.length > 0 && (
                          <button
                              onClick={handleExportToExcel}
                              className="py-3 bg-emerald-600 text-white font-black uppercase text-[11px] rounded-xl flex items-center justify-center gap-2 hover:bg-emerald-700 shadow-lg"
                              title="Exportar para Excel (.xlsx)"
                          >
                              <Table size={16} /> EXCEL
                          </button>
                      )}
                  </div>
                  {/* Cards de Desempenho (FILTRADO) - Atualizam conforme filtros aplicados */}
                  {filteredArchivedResults.length > 0 && (
                      <div className="px-6 py-4 bg-slate-100 dark:bg-slate-800/50 border-b dark:border-slate-800 flex items-center gap-4 shrink-0">
                          <div className="flex items-center gap-2">
                              <Database size={18} className="text-slate-400" />
                              <span className="text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400">Desempenho do Período</span>
                              {hasHistoryActiveColFilters && (
                                  <span className="text-[8px] font-bold text-amber-600 dark:text-amber-400 bg-amber-100 dark:bg-amber-900/30 px-2 py-0.5 rounded-full border border-amber-300 dark:border-amber-700">
                                      FILTRADO ({filteredArchivedResults.length} de {archivedResults.length})
                                  </span>
                              )}
                          </div>
                          <div className="flex items-center gap-3 ml-auto">
                              <div className={`flex items-center gap-3 px-5 py-2 rounded-xl min-w-[130px] ${isDarkMode ? 'bg-emerald-900/30 border border-emerald-700/50' : 'bg-emerald-100 border border-emerald-300'}`}>
                                <div className="text-center flex-1">
                                  <p className={`text-[8px] font-black uppercase tracking-wider mb-0.5 ${isDarkMode ? 'text-emerald-400' : 'text-emerald-700'}`}>Geral</p>
                                  <p className={`text-xl font-black leading-none ${isDarkMode ? 'text-emerald-400' : 'text-emerald-700'}`}>{
                                    (() => {
                                      const total = filteredArchivedResults.length;
                                      const okPrevistoCount = filteredArchivedResults.filter(r => r.statusOp === 'OK' || r.statusOp === 'Previsto').length;
                                      return total > 0 ? ((okPrevistoCount / total) * 100).toFixed(2) : '0.00';
                                    })()
                                  }%</p>
                                </div>
                                <div className="w-1.5 h-1.5 bg-emerald-500 rounded-full shrink-0"></div>
                              </div>
                              <div className={`flex items-center gap-3 px-5 py-2 rounded-xl min-w-[130px] ${isDarkMode ? 'bg-blue-900/30 border border-blue-700/50' : 'bg-blue-100 border border-blue-300'}`}>
                                <div className="text-center flex-1">
                                  <p className={`text-[8px] font-black uppercase tracking-wider mb-0.5 ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>Interno</p>
                                  <p className={`text-xl font-black leading-none ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>{
                                    (() => {
                                      const total = filteredArchivedResults.length;
                                      const justificativas = ['Manutenção', 'Mão de obra', 'Logística'];
                                      const rotasComJustificativa = filteredArchivedResults.filter(r => justificativas.includes(r.motivo)).length;
                                      const rotasSemJustificativa = total - rotasComJustificativa;
                                      return total > 0 ? ((rotasSemJustificativa / total) * 100).toFixed(2) : '0.00';
                                    })()
                                  }%</p>
                                </div>
                                <div className="w-1.5 h-1.5 bg-blue-500 rounded-full shrink-0"></div>
                              </div>
                          </div>
                      </div>
                  )}
                  <div className="flex-1 overflow-auto bg-slate-50 dark:bg-slate-950 p-4">
                      {archivedResults.length > 0 ? (
                          <div className="bg-white dark:bg-slate-900 rounded-2xl border dark:border-slate-800 overflow-hidden">
                              <table className="w-full border-collapse text-[10px]">
                                  <thead className="sticky top-0 bg-slate-200 dark:bg-slate-800 text-slate-600 font-black uppercase z-10">
                                      <tr>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center relative group">
                                              <div className="flex items-center justify-center gap-1">
                                                  <span>Semana</span>
                                                  <button
                                                    ref={(el) => { if (historyActiveFilterCol === 'semana') setHistoryFilterAnchorEl(el); }}
                                                    onClick={() => {
                                                        if (historyActiveFilterCol === 'semana') {
                                                            setHistoryActiveFilterCol(null);
                                                            setHistoryFilterAnchorEl(null);
                                                        } else {
                                                            setHistoryActiveFilterCol('semana');
                                                        }
                                                    }}
                                                    className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                                      historyActiveFilterCol === 'semana' || historySelectedFilters['semana']?.length > 0
                                                        ? 'opacity-100 bg-primary-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                  >
                                                    <Filter size={10} />
                                                  </button>
                                              </div>
                                              {historyActiveFilterCol === 'semana' && (
                                                <HistoryFilterDropdown
                                                    col="semana"
                                                    values={getHistoryColUniqueValues('semana')}
                                                    colFilter={historyColFilters['semana'] || ''}
                                                    setColFilter={(v: string) => setHistoryColFilters({ ...historyColFilters, semana: v })}
                                                    selected={historySelectedFilters['semana'] || []}
                                                    setSelected={(v: string[]) => setHistorySelectedFilters({ ...historySelectedFilters, semana: v })}
                                                    onClose={() => { setHistoryActiveFilterCol(null); setHistoryFilterAnchorEl(null); }}
                                                    dropdownRef={historyFilterDropdownRef}
                                                    anchorEl={historyFilterAnchorEl}
                                                />
                                              )}
                                          </th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center">Data</th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-left relative group">
                                              <div className="flex items-center justify-between">
                                                  <span>Rota</span>
                                                  <button
                                                    ref={(el) => { if (historyActiveFilterCol === 'rota') setHistoryFilterAnchorEl(el); }}
                                                    onClick={() => {
                                                        if (historyActiveFilterCol === 'rota') {
                                                            setHistoryActiveFilterCol(null);
                                                            setHistoryFilterAnchorEl(null);
                                                        } else {
                                                            setHistoryActiveFilterCol('rota');
                                                        }
                                                    }}
                                                    className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                                      historyActiveFilterCol === 'rota' || historySelectedFilters['rota']?.length > 0
                                                        ? 'opacity-100 bg-primary-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                  >
                                                    <Filter size={10} />
                                                  </button>
                                              </div>
                                              {historyActiveFilterCol === 'rota' && (
                                                <HistoryFilterDropdown
                                                    col="rota"
                                                    values={getHistoryColUniqueValues('rota')}
                                                    colFilter={historyColFilters['rota'] || ''}
                                                    setColFilter={(v: string) => setHistoryColFilters({ ...historyColFilters, rota: v })}
                                                    selected={historySelectedFilters['rota'] || []}
                                                    setSelected={(v: string[]) => setHistorySelectedFilters({ ...historySelectedFilters, rota: v })}
                                                    onClose={() => { setHistoryActiveFilterCol(null); setHistoryFilterAnchorEl(null); }}
                                                    dropdownRef={historyFilterDropdownRef}
                                                    anchorEl={historyFilterAnchorEl}
                                                />
                                              )}
                                          </th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center">Início</th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-left relative group">
                                              <div className="flex items-center justify-between">
                                                  <span>Motorista</span>
                                                  <button
                                                    ref={(el) => { if (historyActiveFilterCol === 'motorista') setHistoryFilterAnchorEl(el); }}
                                                    onClick={() => {
                                                        if (historyActiveFilterCol === 'motorista') {
                                                            setHistoryActiveFilterCol(null);
                                                            setHistoryFilterAnchorEl(null);
                                                        } else {
                                                            setHistoryActiveFilterCol('motorista');
                                                        }
                                                    }}
                                                    className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                                      historyActiveFilterCol === 'motorista' || historySelectedFilters['motorista']?.length > 0
                                                        ? 'opacity-100 bg-primary-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                  >
                                                    <Filter size={10} />
                                                  </button>
                                              </div>
                                              {historyActiveFilterCol === 'motorista' && (
                                                <HistoryFilterDropdown
                                                    col="motorista"
                                                    values={getHistoryColUniqueValues('motorista')}
                                                    colFilter={historyColFilters['motorista'] || ''}
                                                    setColFilter={(v: string) => setHistoryColFilters({ ...historyColFilters, motorista: v })}
                                                    selected={historySelectedFilters['motorista'] || []}
                                                    setSelected={(v: string[]) => setHistorySelectedFilters({ ...historySelectedFilters, motorista: v })}
                                                    onClose={() => { setHistoryActiveFilterCol(null); setHistoryFilterAnchorEl(null); }}
                                                    dropdownRef={historyFilterDropdownRef}
                                                    anchorEl={historyFilterAnchorEl}
                                                />
                                              )}
                                          </th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center relative group">
                                              <div className="flex items-center justify-center gap-1">
                                                  <span>Placa</span>
                                                  <button
                                                    ref={(el) => { if (historyActiveFilterCol === 'placa') setHistoryFilterAnchorEl(el); }}
                                                    onClick={() => {
                                                        if (historyActiveFilterCol === 'placa') {
                                                            setHistoryActiveFilterCol(null);
                                                            setHistoryFilterAnchorEl(null);
                                                        } else {
                                                            setHistoryActiveFilterCol('placa');
                                                        }
                                                    }}
                                                    className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                                      historyActiveFilterCol === 'placa' || historySelectedFilters['placa']?.length > 0
                                                        ? 'opacity-100 bg-primary-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                  >
                                                    <Filter size={10} />
                                                  </button>
                                              </div>
                                              {historyActiveFilterCol === 'placa' && (
                                                <HistoryFilterDropdown
                                                    col="placa"
                                                    values={getHistoryColUniqueValues('placa')}
                                                    colFilter={historyColFilters['placa'] || ''}
                                                    setColFilter={(v: string) => setHistoryColFilters({ ...historyColFilters, placa: v })}
                                                    selected={historySelectedFilters['placa'] || []}
                                                    setSelected={(v: string[]) => setHistorySelectedFilters({ ...historySelectedFilters, placa: v })}
                                                    onClose={() => { setHistoryActiveFilterCol(null); setHistoryFilterAnchorEl(null); }}
                                                    dropdownRef={historyFilterDropdownRef}
                                                    anchorEl={historyFilterAnchorEl}
                                                />
                                              )}
                                          </th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center">Saída</th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-left relative group">
                                              <div className="flex items-center justify-between">
                                                  <span>Motivo</span>
                                                  <button
                                                    ref={(el) => { if (historyActiveFilterCol === 'motivo') setHistoryFilterAnchorEl(el); }}
                                                    onClick={() => {
                                                        if (historyActiveFilterCol === 'motivo') {
                                                            setHistoryActiveFilterCol(null);
                                                            setHistoryFilterAnchorEl(null);
                                                        } else {
                                                            setHistoryActiveFilterCol('motivo');
                                                        }
                                                    }}
                                                    className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                                      historyActiveFilterCol === 'motivo' || historySelectedFilters['motivo']?.length > 0
                                                        ? 'opacity-100 bg-primary-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                  >
                                                    <Filter size={10} />
                                                  </button>
                                              </div>
                                              {historyActiveFilterCol === 'motivo' && (
                                                <HistoryFilterDropdown
                                                    col="motivo"
                                                    values={getHistoryColUniqueValues('motivo')}
                                                    colFilter={historyColFilters['motivo'] || ''}
                                                    setColFilter={(v: string) => setHistoryColFilters({ ...historyColFilters, motivo: v })}
                                                    selected={historySelectedFilters['motivo'] || []}
                                                    setSelected={(v: string[]) => setHistorySelectedFilters({ ...historySelectedFilters, motivo: v })}
                                                    onClose={() => { setHistoryActiveFilterCol(null); setHistoryFilterAnchorEl(null); }}
                                                    dropdownRef={historyFilterDropdownRef}
                                                    anchorEl={historyFilterAnchorEl}
                                                />
                                              )}
                                          </th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-left">Observação</th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center relative group">
                                              <div className="flex items-center justify-center gap-1">
                                                  <span>Operação</span>
                                                  <button
                                                    onClick={() => setHistorySortByOperacao(!historySortByOperacao)}
                                                    className={`p-1 rounded transition-all ${
                                                      historySortByOperacao
                                                        ? 'bg-emerald-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                    title={historySortByOperacao ? 'Ordenação por operação ativada' : 'Ordenar por operação (A-Z)'}
                                                  >
                                                    <ArrowUpDown size={10} />
                                                  </button>
                                                  <button
                                                    ref={(el) => { if (historyActiveFilterCol === 'operacao') setHistoryFilterAnchorEl(el); }}
                                                    onClick={() => {
                                                        if (historyActiveFilterCol === 'operacao') {
                                                            setHistoryActiveFilterCol(null);
                                                            setHistoryFilterAnchorEl(null);
                                                        } else {
                                                            setHistoryActiveFilterCol('operacao');
                                                        }
                                                    }}
                                                    className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                                      historyActiveFilterCol === 'operacao' || historySelectedFilters['operacao']?.length > 0
                                                        ? 'opacity-100 bg-primary-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                  >
                                                    <Filter size={10} />
                                                  </button>
                                              </div>
                                              {historyActiveFilterCol === 'operacao' && (
                                                <HistoryFilterDropdown
                                                    col="operacao"
                                                    values={getHistoryColUniqueValues('operacao')}
                                                    colFilter={historyColFilters['operacao'] || ''}
                                                    setColFilter={(v: string) => setHistoryColFilters({ ...historyColFilters, operacao: v })}
                                                    selected={historySelectedFilters['operacao'] || []}
                                                    setSelected={(v: string[]) => setHistorySelectedFilters({ ...historySelectedFilters, operacao: v })}
                                                    onClose={() => { setHistoryActiveFilterCol(null); setHistoryFilterAnchorEl(null); }}
                                                    dropdownRef={historyFilterDropdownRef}
                                                    anchorEl={historyFilterAnchorEl}
                                                />
                                              )}
                                          </th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center relative group">
                                              <div className="flex items-center justify-center gap-1">
                                                  <span>Status</span>
                                                  <button
                                                    ref={(el) => { if (historyActiveFilterCol === 'status') setHistoryFilterAnchorEl(el); }}
                                                    onClick={() => {
                                                        if (historyActiveFilterCol === 'status') {
                                                            setHistoryActiveFilterCol(null);
                                                            setHistoryFilterAnchorEl(null);
                                                        } else {
                                                            setHistoryActiveFilterCol('status');
                                                        }
                                                    }}
                                                    className={`p-1 rounded transition-all opacity-0 group-hover:opacity-100 ${
                                                      historyActiveFilterCol === 'status' || historySelectedFilters['status']?.length > 0
                                                        ? 'opacity-100 bg-primary-600 text-white'
                                                        : 'hover:bg-slate-300 dark:hover:bg-slate-600'
                                                    }`}
                                                  >
                                                    <Filter size={10} />
                                                  </button>
                                              </div>
                                              {historyActiveFilterCol === 'status' && (
                                                <HistoryFilterDropdown
                                                    col="status"
                                                    values={getHistoryColUniqueValues('status')}
                                                    colFilter={historyColFilters['status'] || ''}
                                                    setColFilter={(v: string) => setHistoryColFilters({ ...historyColFilters, status: v })}
                                                    selected={historySelectedFilters['status'] || []}
                                                    setSelected={(v: string[]) => setHistorySelectedFilters({ ...historySelectedFilters, status: v })}
                                                    onClose={() => { setHistoryActiveFilterCol(null); setHistoryFilterAnchorEl(null); }}
                                                    dropdownRef={historyFilterDropdownRef}
                                                    anchorEl={historyFilterAnchorEl}
                                                />
                                              )}
                                          </th>
                                          <th className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center">Tempo</th>
                                      </tr>
                                  </thead>
                                  <tbody>
                                      {filteredArchivedResults.map((r, i) => {
                                          // Verifica se esta linha tem edições pendentes
                                          const pendingEdits = pendingHistoryEdits[r.id!];
                                          
                                          return (
                                              <tr key={i} className={`hover:bg-slate-50 dark:hover:bg-slate-800 border-b border-slate-200 dark:border-slate-800 group ${pendingEdits ? 'bg-amber-50 dark:bg-amber-900/10' : ''}`}>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center font-mono">
                                                  {editingHistoryId === r.id && editingHistoryField === 'semana' ? (
                                                      <input
                                                          type="text"
                                                          defaultValue={r.semana || ''}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'semana', e.target.value)}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-bold outline-none"
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('semana'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 ${pendingEdits?.semana ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {pendingEdits?.semana || r.semana || '---'}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center">
                                                  {editingHistoryId === r.id && editingHistoryField === 'data' ? (
                                                      <input
                                                          type="date"
                                                          defaultValue={r.data}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'data', e.target.value)}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-bold outline-none"
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('data'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 ${pendingEdits?.data ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {(() => {
                                                            if (!r.data) return '';
                                                            // Converte de AAAA-MM-DD para DD/MM/AAAA
                                                            const [ano, mes, dia] = r.data.split('-');
                                                            return `${dia}/${mes}/${ano}`;
                                                          })()}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} relative">
                                                  {(() => {
                                                    // Debug: log para verificar se routeAlerts está acessível
                                                    if (r.rota && routeAlerts[r.rota] && routeAlerts[r.rota].count > 0 && Math.random() < 0.01) {
                                                      console.log(`[ROUTE_CELL_DEBUG] Rota ${r.rota} tem ${routeAlerts[r.rota].count} alertas`);
                                                    }
                                                    return (
                                                      <>
                                                  {editingHistoryId === r.id && editingHistoryField === 'rota' ? (
                                                      <input
                                                          type="text"
                                                          defaultValue={r.rota}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'rota', e.target.value)}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-bold outline-none"
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('rota'); }}
                                                          className="font-bold text-primary-700 dark:text-primary-400 cursor-pointer hover:bg-primary-50 dark:hover:bg-primary-900/20 rounded px-1"
                                                      >
                                                          {r.rota}
                                                      </div>
                                                  )}
                                                      </>
                                                    );
                                                  })()}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center font-mono">
                                                  {editingHistoryId === r.id && editingHistoryField === 'inicio' ? (
                                                      <input
                                                          type="text"
                                                          defaultValue={r.inicio}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'inicio', applyTimeMask(e.target.value))}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-mono font-bold outline-none"
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('inicio'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 ${pendingEdits?.inicio ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {pendingEdits?.inicio || r.inicio || '---'}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}">
                                                  {editingHistoryId === r.id && editingHistoryField === 'motorista' ? (
                                                      <input
                                                          type="text"
                                                          defaultValue={r.motorista}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'motorista', e.target.value)}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-bold outline-none"
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('motorista'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 ${pendingEdits?.motorista ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {pendingEdits?.motorista || r.motorista || '---'}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center font-mono">
                                                  {editingHistoryId === r.id && editingHistoryField === 'placa' ? (
                                                      <input
                                                          type="text"
                                                          defaultValue={r.placa}
                                                          onChange={(e) => {
                                                            const cleanValue = e.target.value.replace(/[\s-]/g, '').toUpperCase();
                                                            handleUpdateHistoryCell(r.id!, 'placa', cleanValue);
                                                          }}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-mono font-bold outline-none uppercase"
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('placa'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 ${pendingEdits?.placa ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {pendingEdits?.placa || r.placa || '---'}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center font-mono">
                                                  {editingHistoryId === r.id && editingHistoryField === 'saida' ? (
                                                      <input
                                                          type="text"
                                                          defaultValue={r.saida}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'saida', applyTimeMask(e.target.value))}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-mono font-bold outline-none"
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('saida'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 ${pendingEdits?.saida ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {pendingEdits?.saida || r.saida || '---'}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'}">
                                                  {editingHistoryId === r.id && editingHistoryField === 'motivo' ? (
                                                      <select
                                                          defaultValue={r.motivo}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'motivo', e.target.value)}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-bold outline-none"
                                                          autoFocus
                                                      >
                                                          <option value="">---</option>
                                                          {MOTIVOS.map(m => <option key={m} value={m}>{m}</option>)}
                                                      </select>
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('motivo'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 ${pendingEdits?.motivo ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {pendingEdits?.motivo || r.motivo || '---'}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} max-w-xs">
                                                  {editingHistoryId === r.id && editingHistoryField === 'observacao' ? (
                                                      <textarea
                                                          defaultValue={r.observacao}
                                                          onChange={(e) => handleUpdateHistoryCell(r.id!, 'observacao', e.target.value)}
                                                          onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                                          onInput={(e: any) => {
                                                              e.target.style.height = 'auto';
                                                              e.target.style.height = Math.max(e.target.scrollHeight, 80) + 'px';
                                                          }}
                                                          className="w-full bg-primary-100 dark:bg-primary-900/30 border-2 border-primary-500 px-2 py-1 font-normal outline-none resize-none whitespace-pre-wrap break-words"
                                                          rows={3}
                                                          style={{ wordBreak: 'break-word', overflowWrap: 'break-word' }}
                                                          autoFocus
                                                      />
                                                  ) : (
                                                      <div
                                                          onClick={() => { setEditingHistoryId(r.id!); setEditingHistoryField('observacao'); }}
                                                          className={`cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800 rounded px-1 whitespace-pre-wrap break-words ${pendingEdits?.observacao ? 'bg-amber-200 dark:bg-amber-800 font-bold' : ''}`}
                                                      >
                                                          {pendingEdits?.observacao || r.observacao || '---'}
                                                      </div>
                                                  )}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center font-black">
                                                  {r.operacao}
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center">
                                                  <span className={`px-2 py-0.5 rounded-full text-[8px] font-black border ${
                                                      r.statusOp === 'OK' ? 'bg-emerald-100 border-emerald-400 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-400' :
                                                      r.statusOp === 'Atrasada' ? 'bg-yellow-100 border-yellow-400 text-yellow-800 dark:bg-yellow-900/30 dark:text-yellow-400' :
                                                      r.statusOp === 'Adiantada' ? 'bg-blue-100 border-blue-400 text-blue-800 dark:bg-blue-900/30 dark:text-blue-400' :
                                                      r.statusOp === 'Programada' ? 'bg-slate-200 border-slate-400 text-slate-600 dark:bg-slate-700 dark:text-slate-300' :
                                                      r.statusOp === 'Previsto' ? 'bg-slate-100 border-slate-400 text-slate-500 dark:bg-slate-800 dark:text-slate-400' :
                                                      'bg-red-100 border-red-400 text-red-800 dark:bg-red-900/30 dark:text-red-400'
                                                  }`}>
                                                      {r.statusOp}
                                                  </span>
                                              </td>
                                              <td className="p-2 border ${isDarkMode ? 'border-slate-700' : 'border-slate-400'} text-center font-mono font-bold">
                                                  {r.tempo}
                                              </td>
                                          </tr>
                                          );
                                      })}
                                  </tbody>
                              </table>
                          </div>
                      ) : (
                          <div className="h-full flex flex-col items-center justify-center text-slate-400 italic font-bold">
                              {isSearchingArchive ? (
                                  <>
                                      <Loader2 size={48} className="animate-spin mb-4 text-primary-500" />
                                      <p>Buscando...</p>
                                  </>
                              ) : (
                                  <>
                                      <Database size={48} className="mb-4 opacity-50" />
                                      <p>Nenhum dado retornado para este período</p>
                                  </>
                              )}
                          </div>
                      )}
                  </div>
                  <div className="p-4 bg-slate-100 dark:bg-slate-800 border-t dark:border-slate-700 shrink-0">
                      <p className="text-[10px] font-bold text-slate-500 dark:text-slate-400 text-center">
                          💡 Clique em qualquer célula para editar • Os dados são sincronizados com o SharePoint
                      </p>
                  </div>
              </div>
          </div>
      )}

      {/* Modal de Alerta de Rota com Histórico de Problemas */}
      {selectedRouteAlert && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[300] flex items-center justify-center p-4" onClick={() => setSelectedRouteAlert(null)}>
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-2xl border dark:border-slate-800 overflow-hidden" onClick={(e) => e.stopPropagation()}>
                  <div className="bg-red-600 p-6 flex justify-between items-center text-white">
                      <div className="flex items-center gap-4">
                          <div className="w-12 h-12 bg-white/20 rounded-2xl flex items-center justify-center">
                              <AlertTriangle size={28} />
                          </div>
                          <div>
                              <h3 className="font-black uppercase tracking-widest text-lg">Histórico de Problemas</h3>
                              <p className="text-[10px] font-bold text-white/80 uppercase tracking-wide">Rota: {selectedRouteAlert.rota}</p>
                          </div>
                      </div>
                      <button onClick={() => setSelectedRouteAlert(null)} className="p-2 hover:bg-white/20 rounded-xl transition-colors">
                          <X size={28} />
                      </button>
                  </div>
                  <div className="p-6 max-h-[60vh] overflow-y-auto scrollbar-thin">
                      <div className="mb-4 p-4 bg-red-50 dark:bg-red-900/20 rounded-2xl border border-red-200 dark:border-red-800">
                          <p className="text-[11px] font-black uppercase text-red-700 dark:text-red-400 text-center">
                              ⚠️ {selectedRouteAlert.history.length} ocorrência(s) de atraso/adiantamento nos últimos 7 dias
                          </p>
                      </div>
                      <div className="space-y-3">
                          {selectedRouteAlert.history.map((item, idx) => (
                              <div key={idx} className="p-5 bg-white dark:bg-slate-900 rounded-2xl border-2 border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all">
                                  {/* Cabeçalho: Status + Data */}
                                  <div className="flex items-center gap-2 mb-3">
                                      {/* Badge de Status com ícone */}
                                      <span className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[10px] font-black uppercase border-2 ${
                                          item.statusOp === 'Atrasada' || item.statusOp === 'Atrasado'
                                            ? 'bg-red-100 border-red-400 text-red-800 dark:bg-red-900/40 dark:border-red-700 dark:text-red-300'
                                            : 'bg-blue-100 border-blue-400 text-blue-800 dark:bg-blue-900/40 dark:border-blue-700 dark:text-blue-300'
                                      }`}>
                                          {item.statusOp === 'Atrasada' || item.statusOp === 'Atrasado' ? (
                                              <><AlertTriangle size={12} /> ATRASADA</>
                                          ) : (
                                              <><Clock size={12} /> ADIANTADA</>
                                          )}
                                      </span>
                                      {/* Data com ícone */}
                                      <span className="inline-flex items-center gap-1.5 text-[10px] font-bold text-slate-500 dark:text-slate-400 bg-slate-100 dark:bg-slate-800 px-3 py-1.5 rounded-full">
                                          <Calendar size={12} />
                                          {new Date(item.data).toLocaleDateString('pt-BR')}
                                      </span>
                                  </div>
                                  {/* Motivo em destaque */}
                                  {item.motivo && (
                                      <div className="mb-3 p-3 bg-amber-50 dark:bg-amber-900/20 rounded-xl border border-amber-200 dark:border-amber-800">
                                          <div className="flex items-start gap-2">
                                              <span className="text-[10px] font-black uppercase text-amber-700 dark:text-amber-400 whitespace-nowrap mt-0.5">
                                                  📌 Motivo:
                                              </span>
                                              <p className="text-[11px] font-bold text-amber-900 dark:text-amber-200 leading-relaxed">
                                                  {item.motivo}
                                              </p>
                                          </div>
                                      </div>
                                  )}
                                  {/* Observação em destaque */}
                                  {item.observacao && (
                                      <div className="p-3 bg-slate-50 dark:bg-slate-800/50 rounded-xl border border-slate-200 dark:border-slate-700">
                                          <div className="flex items-start gap-2">
                                              <span className="text-[10px] font-black uppercase text-slate-500 dark:text-slate-400 whitespace-nowrap mt-0.5">
                                                  📝 Observação:
                                              </span>
                                              <p className="text-[10px] font-normal text-slate-700 dark:text-slate-300 leading-relaxed whitespace-pre-wrap">
                                                  {item.observacao}
                                              </p>
                                          </div>
                                      </div>
                                  )}
                              </div>
                          ))}
                      </div>
                  </div>
              </div>
          </div>
      )}

      {/* Modal de Bloqueio (Resumo/Status) */}
      {archiveHardBlockModal && (
        <div className="fixed inset-0 bg-slate-950/85 backdrop-blur-md z-[455] flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.2rem] w-full max-w-2xl border border-red-300 dark:border-red-800 shadow-2xl overflow-hidden">
            <div className="bg-red-500/10 dark:bg-red-900/30 border-b border-red-200 dark:border-red-800 px-6 py-5">
              <div className="flex items-center gap-3">
                <AlertTriangle className="text-red-600 dark:text-red-400" size={24} />
                <div>
                  <h3 className="text-base font-black uppercase tracking-widest text-red-700 dark:text-red-300">
                    {archiveHardBlockModal === 'resumo' ? 'Resumo de Saídas Pendente' : 'Status Individual Pendente'}
                  </h3>
                  <p className="text-[11px] font-bold text-red-700/80 dark:text-red-300/80 mt-1">
                    {archiveHardBlockModal === 'resumo'
                      ? 'Não é possível arquivar enquanto o resumo de saídas não estiver como OK.'
                      : 'Não é possível arquivar enquanto houver operação com Status individual diferente de OK.'}
                  </p>
                </div>
              </div>
            </div>

            <div className="px-6 py-5">
              {archiveHardBlockModal === 'status' && (
                <div className="mb-4">
                  <p className="text-[11px] font-bold text-slate-600 dark:text-slate-300 mb-2">
                    Operações com divergência no campo Status:
                  </p>
                  <div className="space-y-2 max-h-[220px] overflow-y-auto scrollbar-thin">
                    {archiveStatusDivergences.map((item, idx) => (
                      <div
                        key={`${item.operacao}-${idx}`}
                        className="p-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/50"
                      >
                        <p className="text-[11px] font-black uppercase text-slate-700 dark:text-slate-200">
                          {item.operacao}
                        </p>
                        <p className="text-[10px] font-bold uppercase text-red-600 dark:text-red-400 mt-1">
                          Status: {item.status}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <div className="p-3 rounded-xl border border-blue-200 dark:border-blue-800 bg-blue-50 dark:bg-blue-900/20">
                <p className="text-[10px] font-black uppercase text-blue-700 dark:text-blue-300 leading-relaxed">
                  Coloque "-" nas saídas que não serão realizadas no dia, envie novamente o resumo e o status, e depois tente arquivar de novo.
                </p>
              </div>
            </div>

            <div className="px-6 pb-6 pt-2">
              <button
                onClick={() => {
                  setArchiveHardBlockModal(null);
                  setArchiveStatusDivergences([]);
                }}
                className="w-full py-3 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 font-black uppercase text-[10px] tracking-wider hover:bg-slate-200 dark:hover:bg-slate-700 transition-all"
              >
                Entendi
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Validação para Arquivamento Parcial */}
      {isArchiveValidationModalOpen && (
        <div className="fixed inset-0 bg-slate-950/85 backdrop-blur-md z-[450] flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.2rem] w-full max-w-3xl border border-amber-300 dark:border-amber-800 shadow-2xl overflow-hidden">
            <div className="bg-amber-500/10 dark:bg-amber-900/30 border-b border-amber-200 dark:border-amber-800 px-6 py-5">
              <div className="flex items-center gap-3">
                <AlertTriangle className="text-amber-600 dark:text-amber-400" size={24} />
                <div>
                  <h3 className="text-base font-black uppercase tracking-widest text-amber-700 dark:text-amber-300">
                    Rotas Incompletas para Arquivamento
                  </h3>
                  <p className="text-[11px] font-bold text-amber-700/80 dark:text-amber-300/80 mt-1">
                    Você pode arquivar as demais rotas. As incompletas ficarão no painel com data ajustada para o próximo dia.
                  </p>
                </div>
              </div>
            </div>

            <div className="px-6 py-5 max-h-[50vh] overflow-y-auto scrollbar-thin">
              <div className="mb-4 text-[11px] font-bold text-slate-600 dark:text-slate-300">
                {archiveDivergentRoutes.length} rota(s) com divergência encontrada(s):
              </div>
              <div className="space-y-2">
                {archiveDivergentRoutes.map((item, idx) => (
                  <div
                    key={`${item.route.id}-${idx}`}
                    className="p-3 rounded-xl border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/50"
                  >
                    <div className="text-[11px] font-black uppercase text-slate-700 dark:text-slate-200">
                      {item.route.rota || '(ROTA SEM NOME)'} {item.route.operacao ? `- ${item.route.operacao}` : ''}
                    </div>
                    <div className="mt-1 text-[10px] font-bold text-red-600 dark:text-red-400 uppercase">
                      Faltando: {item.missingFields.join(', ')}
                    </div>
                  </div>
                ))}
              </div>
              <div className="mt-5 p-3 rounded-xl border border-blue-200 dark:border-blue-800 bg-blue-50 dark:bg-blue-900/20">
                <p className="text-[10px] font-black uppercase text-blue-700 dark:text-blue-300">
                  Prontas para arquivar agora: {archiveReadyRoutes.length}
                </p>
              </div>
            </div>

            <div className="px-6 pb-6 pt-2 flex items-center gap-3">
              <button
                onClick={() => {
                  setIsArchiveValidationModalOpen(false);
                  setArchiveReadyRoutes([]);
                  setArchiveDivergentRoutes([]);
                }}
                disabled={isSyncing}
                className="flex-1 py-3 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 font-black uppercase text-[10px] tracking-wider hover:bg-slate-200 dark:hover:bg-slate-700 transition-all disabled:opacity-60"
              >
                Cancelar
              </button>
              <button
                onClick={async () => {
                  if (!validateArchiveHardBlocks(archiveReadyRoutes)) {
                    return;
                  }
                  await executeArchiveWithDivergences(archiveReadyRoutes, archiveDivergentRoutes);
                }}
                disabled={isSyncing}
                className="flex-1 py-3 rounded-xl bg-amber-500 hover:bg-amber-600 text-white font-black uppercase text-[10px] tracking-wider transition-all disabled:opacity-60 flex items-center justify-center gap-2"
              >
                {isSyncing ? <Loader2 size={16} className="animate-spin" /> : <Archive size={16} />}
                Estou ciente, arquivar mesmo assim
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Edição de Horários */}
      {isTimeEditModalOpen && timeEditData && (
        <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[400] flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-primary-500 shadow-2xl animate-in zoom-in">
            <div className="flex items-center gap-3 text-primary-500 mb-6 font-black uppercase text-xs">
              <Clock size={24} />
              Editar Horários
            </div>
            
            <div className="mb-6">
              <p className="text-sm text-slate-500 dark:text-slate-400 mb-4 font-medium">
                {timeEditData.template}
              </p>
            </div>
            
            <div className="grid grid-cols-2 gap-4 mb-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase text-slate-400">Início</label>
                <input
                  type="text"
                  value={timeEditData.startTime}
                  onChange={(e) => {
                    const masked = applyTimeMask(e.target.value);
                    setTimeEditData({ ...timeEditData, startTime: masked });
                  }}
                  placeholder="HH:MM:SS"
                  className="w-full p-3 border border-slate-200 dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white text-center font-mono"
                  autoFocus
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase text-slate-400">Término</label>
                <input
                  type="text"
                  value={timeEditData.endTime}
                  onChange={(e) => {
                    const masked = applyTimeMask(e.target.value);
                    setTimeEditData({ ...timeEditData, endTime: masked });
                  }}
                  placeholder="HH:MM:SS"
                  className="w-full p-3 border border-slate-200 dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white text-center font-mono"
                />
              </div>
            </div>
            
            <div className="flex gap-3">
              <button
                onClick={applyTimeEdit}
                className="flex-1 py-3 bg-primary-600 text-white rounded-xl font-black uppercase text-[10px] hover:bg-primary-700 transition-all flex items-center justify-center gap-2"
              >
                <Check size={16} /> Aplicar
              </button>
              <button
                onClick={() => {
                  setIsTimeEditModalOpen(false);
                  setTimeEditData(null);
                }}
                className="flex-1 py-3 bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-300 rounded-xl font-black uppercase text-[10px] hover:bg-slate-200 dark:hover:bg-slate-700 transition-all flex items-center justify-center gap-2"
              >
                <X size={16} /> Cancelar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Edição do Checklist */}
      {isChecklistEditModalOpen && checklistEditData && (
        <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[400] flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-primary-500 shadow-2xl animate-in zoom-in">
            <div className="flex items-center gap-3 mb-6">
              <svg
                width="28"
                height="28"
                viewBox="0 0 24 24"
                fill="none"
                stroke={checklistEditData.data && checklistEditData.porcentagem ? '#10b981' : '#ef4444'}
                strokeWidth="3"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <polyline points="20 6 9 17 4 12" />
              </svg>
              <div className="text-primary-500 font-black uppercase text-xs">
                {checklistEditData.data && checklistEditData.porcentagem ? 'Editar Checklist de Motorista' : 'Preencher Checklist de Motorista'}
              </div>
            </div>

            <div className="space-y-4 mb-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase text-slate-400">Data do Checklist</label>
                <input
                  type="text"
                  placeholder="DD/MM/AAAA"
                  value={checklistEditData.data}
                  onChange={(e) => {
                    let value = e.target.value.replace(/\D/g, '');
                    if (value.length > 8) value = value.slice(0, 8);
                    if (value.length >= 8) {
                      value = `${value.slice(0, 2)}/${value.slice(2, 4)}/${value.slice(4)}`;
                    }
                    setChecklistEditData({ ...checklistEditData, data: value });
                  }}
                  className="w-full p-3 border border-slate-200 dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white text-center"
                />
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase text-slate-400">Porcentagem</label>
                <input
                  type="text"
                  value={checklistEditData.porcentagem}
                  onChange={(e) => {
                    const value = e.target.value.replace(/\D/g, '');
                    setChecklistEditData({ ...checklistEditData, porcentagem: value });
                  }}
                  placeholder="100"
                  className="w-full p-3 border border-slate-200 dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white text-center"
                />
              </div>

              {(!checklistEditData.porcentagem || parseInt(checklistEditData.porcentagem) < 100) && (
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase text-slate-400">Motivos Apontados</label>
                  <textarea
                    value={checklistEditData.motivos}
                    onChange={(e) => {
                      setChecklistEditData({ ...checklistEditData, motivos: e.target.value });
                    }}
                    placeholder="Descreva os motivos..."
                    rows={3}
                    className="w-full p-3 border border-slate-200 dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white resize-none"
                  />
                </div>
              )}

              <div className="space-y-2">
                <label className="text-[10px] font-black uppercase text-slate-400">Causa Raiz (Manutenção)</label>
                <input
                  type="text"
                  value={checklistEditData.causaRaiz}
                  onChange={(e) => setChecklistEditData({ ...checklistEditData, causaRaiz: e.target.value })}
                  placeholder="Digite para filtrar e selecione a causa..."
                  className="w-full p-3 border border-slate-200 dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white"
                />
                <div className="max-h-32 overflow-y-auto scrollbar-thin border border-slate-200 dark:border-slate-700 rounded-xl bg-white dark:bg-slate-900">
                  {CAUSAS_RAIZ_MANUTENCAO
                    .filter(item => {
                      const query = normalizeCauseText(checklistEditData.causaRaiz || '');
                      if (!query) return true;
                      return normalizeCauseText(item).includes(query);
                    })
                    .slice(0, 10)
                    .map((item) => (
                      <button
                        key={item}
                        type="button"
                        onClick={() => setChecklistEditData({ ...checklistEditData, causaRaiz: item })}
                        className={`w-full text-left px-3 py-2 text-[10px] font-bold border-b border-slate-100 dark:border-slate-800 last:border-b-0 hover:bg-primary-50 dark:hover:bg-slate-800 transition-colors ${
                          checklistEditData.causaRaiz === item
                            ? 'text-primary-700 dark:text-primary-300 bg-primary-50/60 dark:bg-primary-900/20'
                            : 'text-slate-700 dark:text-slate-300'
                        }`}
                      >
                        {item}
                      </button>
                    ))}
                </div>
              </div>
            </div>

            <div className="flex gap-3">
              <button
                onClick={applyChecklistEdit}
                className="flex-1 py-3 bg-primary-600 text-white rounded-xl font-black uppercase text-[10px] hover:bg-primary-700 transition-all flex items-center justify-center gap-2"
              >
                <Check size={16} /> Aplicar
              </button>
              <button
                onClick={() => {
                  setIsChecklistEditModalOpen(false);
                  setChecklistEditData(null);
                }}
                className="flex-1 py-3 bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-300 rounded-xl font-black uppercase text-[10px] hover:bg-slate-200 dark:hover:bg-slate-700 transition-all flex items-center justify-center gap-2"
              >
                <X size={16} /> Cancelar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Alerta de Motorista com Histórico de Atrasos "Mão de obra" */}
      {selectedMotoristAlert && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[300] flex items-center justify-center p-4" onClick={() => setSelectedMotoristAlert(null)}>
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-2xl border dark:border-slate-800 overflow-hidden" onClick={(e) => e.stopPropagation()}>
                  <div className="bg-orange-600 p-6 flex justify-between items-center text-white">
                      <div className="flex items-center gap-4">
                          <div className="w-12 h-12 bg-white/20 rounded-2xl flex items-center justify-center">
                              <AlertTriangle size={28} />
                          </div>
                          <div>
                              <h3 className="font-black uppercase tracking-widest text-lg">Histórico de Atrasos — Mão de Obra</h3>
                              <p className="text-[10px] font-bold text-white/80 uppercase tracking-wide">Motorista: {selectedMotoristAlert.motorista}</p>
                          </div>
                      </div>
                      <button onClick={() => setSelectedMotoristAlert(null)} className="p-2 hover:bg-white/20 rounded-xl transition-colors">
                          <X size={28} />
                      </button>
                  </div>
                  <div className="p-6 max-h-[60vh] overflow-y-auto scrollbar-thin">
                      <div className="mb-4 p-4 bg-orange-50 dark:bg-orange-900/20 rounded-2xl border border-orange-200 dark:border-orange-800">
                          <p className="text-[11px] font-black uppercase text-orange-700 dark:text-orange-400 text-center">
                              ⚠️ {selectedMotoristAlert.count} ocorrência(s) de atraso por "Mão de obra" nos últimos 30 dias
                          </p>
                      </div>
                      <div className="space-y-3">
                          {selectedMotoristAlert.history.map((item, idx) => (
                              <div key={idx} className="p-5 bg-white dark:bg-slate-900 rounded-2xl border-2 border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all">
                                  {/* Cabeçalho: Data + Operação */}
                                  <div className="flex items-center gap-2 mb-3">
                                      <span className="inline-flex items-center gap-1.5 text-[10px] font-bold text-slate-500 dark:text-slate-400 bg-slate-100 dark:bg-slate-800 px-3 py-1.5 rounded-full">
                                          <Calendar size={12} />
                                          {new Date(item.data).toLocaleDateString('pt-BR')}
                                      </span>
                                      {item.operacao && (
                                          <span className="inline-flex items-center gap-1.5 text-[10px] font-bold text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/30 px-3 py-1.5 rounded-full border border-blue-200 dark:border-blue-800">
                                              {item.operacao}
                                          </span>
                                      )}
                                      {item.rota && (
                                          <span className="inline-flex items-center gap-1.5 text-[10px] font-bold text-slate-600 dark:text-slate-300 bg-slate-50 dark:bg-slate-800 px-3 py-1.5 rounded-full border border-slate-200 dark:border-slate-700">
                                              Rota: {item.rota}
                                          </span>
                                      )}
                                  </div>
                                  {/* Observação em destaque */}
                                  {item.observacao && (
                                      <div className="p-3 bg-slate-50 dark:bg-slate-800/50 rounded-xl border border-slate-200 dark:border-slate-700">
                                          <div className="flex items-start gap-2">
                                              <span className="text-[10px] font-black uppercase text-slate-500 dark:text-slate-400 whitespace-nowrap mt-0.5">
                                                  📝 Observação:
                                              </span>
                                              <p className="text-[10px] font-normal text-slate-700 dark:text-slate-300 leading-relaxed whitespace-pre-wrap">
                                                  {item.observacao}
                                              </p>
                                          </div>
                                      </div>
                                  )}
                                  {/* Informações adicionais */}
                                  <div className="mt-3 flex flex-wrap gap-3 text-[10px] text-slate-500 dark:text-slate-400">
                                      {item.inicio && (
                                          <span className="flex items-center gap-1">
                                              <Clock size={12} /> Início: {item.inicio}
                                          </span>
                                      )}
                                      {item.saida && (
                                          <span className="flex items-center gap-1">
                                              <Clock size={12} /> Saída: {item.saida}
                                          </span>
                                      )}
                                      {item.placa && (
                                          <span className="flex items-center gap-1">
                                              Placa: {item.placa}
                                          </span>
                                      )}
                                  </div>
                              </div>
                          ))}
                      </div>
                  </div>
                  <div className="p-4 bg-slate-100 dark:bg-slate-800 border-t dark:border-slate-700 shrink-0">
                      <p className="text-[10px] font-bold text-slate-500 dark:text-slate-400 text-center">
                          📊 Dados dos últimos 30 dias (fuso de Brasília)
                      </p>
                  </div>
              </div>
          </div>
      )}

      {/* Modal de Configuração de Emails */}
      {isEmailConfigModalOpen && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-[110] flex items-center justify-center p-4 animate-in zoom-in duration-300">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden border border-blue-500/30 flex flex-col max-h-[90vh]">
            <div className="bg-blue-600 text-white p-6 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-4">
                <Settings2 size={32} className="animate-spin-slow" />
                <div>
                  <h3 className="text-xl font-black uppercase tracking-tight">Configurar Emails de Envio</h3>
                  <p className="text-xs text-blue-200 font-bold uppercase tracking-widest">Selecione a operação e edite os emails</p>
                </div>
              </div>
              <button
                onClick={() => {
                  setIsEmailConfigModalOpen(false);
                  setIsConfigModalOpen(false);
                }}
                className="hover:bg-white/10 p-2 rounded-full transition-all"
              >
                <X size={24} />
              </button>
            </div>

            <div className="p-6 bg-slate-50 dark:bg-slate-950 overflow-y-auto flex-1 scrollbar-thin">
              <div className="space-y-6">
                {/* 🔍 Filtro por Email - busca email em todas as operações */}
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Filter size={16} className="text-blue-500" />
                    <label className="text-[10px] font-black uppercase text-slate-400">Buscar Email em Todas as Operações</label>
                  </div>
                  <input
                    type="email"
                    value={emailFilter}
                    onChange={(e) => setEmailFilter(e.target.value)}
                    placeholder="Digite um email para buscar..."
                    className="w-full p-4 border border-slate-200 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-900 text-sm font-bold outline-none dark:text-white shadow-sm focus:ring-2 focus:ring-blue-500"
                  />

                  {/* Resultados da busca */}
                  {emailFilter.trim().length > 0 && (
                    <div className="space-y-2">
                      {(() => {
                        const results = findEmailInAllOperations(emailFilter.trim());
                        if (results.length === 0) {
                          return (
                            <div className="p-4 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-2xl text-center">
                              <p className="text-sm font-bold text-amber-600 dark:text-amber-400">
                                Email não encontrado em nenhuma operação
                              </p>
                            </div>
                          );
                        }

                        return (
                          <div className="space-y-2">
                            <p className="text-[10px] font-black uppercase text-slate-400">
                              Encontrado em {results.length} operação(ões):
                            </p>
                            {results.map((result, index) => (
                              <div
                                key={`${result.operacao}-${result.campo}-${index}`}
                                className="flex items-center justify-between p-3 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-2xl shadow-sm"
                              >
                                <div className="flex items-center gap-3 flex-1 min-w-0">
                                  <div className="flex flex-col flex-1 min-w-0">
                                    <span className="text-sm font-bold dark:text-white truncate">
                                      {result.nomeExibicao}
                                    </span>
                                    <span className="text-[10px] font-bold uppercase text-slate-400">
                                      {emailFilter.trim()}
                                    </span>
                                  </div>
                                  <span className={`inline-flex px-2 py-1 rounded-full text-[9px] font-black uppercase tracking-wide ${
                                    result.campo === 'Envio'
                                      ? 'bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300 border border-blue-200 dark:border-blue-800'
                                      : 'bg-purple-100 dark:bg-purple-900/40 text-purple-700 dark:text-purple-300 border border-purple-200 dark:border-purple-800'
                                  }`}>
                                    {result.campo === 'Envio' ? '📤 Envio' : '📋 Cópia'}
                                  </span>
                                </div>
                                <button
                                  onClick={async () => {
                                    await removeEmailFromOperation(result.operacao, result.campo, emailFilter.trim());
                                    // Não limpa o filtro para permitir ver o resultado da remoção
                                  }}
                                  className="ml-3 p-2 hover:bg-red-100 dark:hover:bg-red-900/30 rounded-full transition-colors group"
                                  title={`Remover email de ${result.nomeExibicao} (${result.campo})`}
                                >
                                  <X size={14} className="text-red-500 group-hover:text-red-700 dark:group-hover:text-red-400" />
                                </button>
                              </div>
                            ))}
                          </div>
                        );
                      })()}
                    </div>
                  )}
                </div>

                {/* Separador */}
                <div className="border-t border-slate-200 dark:border-slate-700"></div>

                {/* Seleção de Operação */}
                <div className="space-y-2">
                  <label className="text-[10px] font-black uppercase text-slate-400">Operação</label>
                  <select
                    value={selectedOperacaoConfig}
                    onChange={(e) => {
                      const novaOperacao = e.target.value;
                      // Reseta o flag da operação anterior para permitir recarregar se voltar a ela
                      delete emailConfigLoadedRef.current[selectedOperacaoConfig];
                      setSelectedOperacaoConfig(novaOperacao);
                    }}
                    className="w-full p-4 border border-slate-200 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-900 text-sm font-bold outline-none dark:text-white shadow-sm focus:ring-2 focus:ring-blue-500"
                  >
                    {(isDeale ? getDealeFilteredConfigs(userConfigs) : userConfigs).map(config => (
                      <option key={config.operacao} value={config.operacao}>
                        {config.nomeExibicao}
                      </option>
                    ))}
                  </select>
                </div>

                {/* Campos de Email */}
                <div className="space-y-4">
                  {/* Campo Envio com Pills */}
                  <EmailInput
                    label="Emails para Envio (Principal)"
                    value={configEnvio}
                    onChange={setConfigEnvio}
                    onRemoveEmail={(email) => setConfigEnvio(removeEmail(configEnvio, email))}
                    placeholder="Cole emails em massa..."
                  />

                  {/* Campo Cópia com Pills */}
                  <EmailInput
                    label="Emails para Cópia"
                    value={configCopia}
                    onChange={setConfigCopia}
                    onRemoveEmail={(email) => setConfigCopia(removeEmail(configCopia, email))}
                    placeholder="Cole emails em massa..."
                  />
                </div>
              </div>
            </div>

            <div className="p-6 bg-white dark:bg-slate-900 border-t dark:border-slate-800 shrink-0 flex gap-4">
              <button
                onClick={handleSaveEmailConfig}
                disabled={isSavingConfig}
                className="flex-1 py-4 bg-blue-600 hover:bg-blue-700 text-white rounded-2xl font-black uppercase text-[11px] tracking-widest flex items-center justify-center gap-3 transition-all disabled:opacity-60 shadow-lg shadow-blue-500/20"
              >
                {isSavingConfig ? (
                  <><Loader2 size={18} className="animate-spin" /> Salvando...</>
                ) : (
                  <><Check size={18} /> Salvar Configuração</>
                )}
              </button>
              <button
                onClick={() => {
                  setIsEmailConfigModalOpen(false);
                  setIsConfigModalOpen(false);
                }}
                disabled={isSavingConfig}
                className="flex-1 py-4 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-600 dark:text-slate-300 rounded-2xl font-black uppercase text-[11px] tracking-widest flex items-center justify-center gap-3 transition-all disabled:opacity-60 border border-slate-200 dark:border-slate-700"
              >
                <X size={18} /> Cancelar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ⚠️ Popup de Envio Automático REMOVIDO — Envio agora é feito apenas pela tela "Resumo" */}
    </div>
  );
};

export default RouteDepartureView;
