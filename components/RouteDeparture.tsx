
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteDeparture, User, RouteOperationMapping, RouteConfig } from '../types';
import { SharePointService } from '../services/sharepointService';
import {
  Clock, X, Loader2, RefreshCw, ShieldCheck,
  AlertTriangle, CheckCircle2, ChevronDown,
  Filter, Search, CheckSquare, Square,
  BarChart3, TrendingUp, SortAsc,
  Activity, ChevronRight,
  Archive, Database, Save, Link as LinkIcon,
  Layers, Trash2, Settings2, Check
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

const RouteDepartureView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [routeMappings, setRouteMappings] = useState<RouteOperationMapping[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [currentTime, setCurrentTime] = useState(new Date());
  const [bulkStatus, setBulkStatus] = useState<{ active: boolean, current: number, total: number } | null>(null);
  const [pendingBulkRoutes, setPendingBulkRoutes] = useState<string[]>([]);
  const [isBulkMappingModalOpen, setIsBulkMappingModalOpen] = useState(false);

  const [ghostRow, setGhostRow] = useState<Partial<RouteDeparture>>({
    id: 'ghost', rota: '', data: new Date().toISOString().split('T')[0], inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '', semana: ''
  });

  // Armazena os últimos checklists de motorista por operação
  const [lastMotoristaChecklist, setLastMotoristaChecklist] = useState<Record<string, { data: string, porcentagem: string }>>({});

  const [isStatsModalOpen, setIsStatsModalOpen] = useState(false);
  const [isHistoryModalOpen, setIsHistoryModalOpen] = useState(false);
  const [isMappingModalOpen, setIsMappingModalOpen] = useState(false);
  const [pendingMappingRoute, setPendingMappingRoute] = useState<string | null>(null);

  // Estado para o popup de edição de horários
  const [isTimeEditModalOpen, setIsTimeEditModalOpen] = useState(false);
  const [timeEditData, setTimeEditData] = useState<{ routeId: string; template: string; startTime: string; endTime: string } | null>(null);

  // Estado para o popup de edição do checklist (GERAL)
  const [isChecklistEditModalOpen, setIsChecklistEditModalOpen] = useState(false);
  const [checklistEditData, setChecklistEditData] = useState<{ routeId: string; data: string; porcentagem: string; motivos: string } | null>(null);
  
  const [histStart, setHistStart] = useState(new Date().toISOString().split('T')[0]);
  const [histEnd, setHistEnd] = useState(new Date().toISOString().split('T')[0]);
  const [archivedResults, setArchivedResults] = useState<RouteDeparture[]>([]);
  const [isSearchingArchive, setIsSearchingArchive] = useState(false);

  const [activeObsId, setActiveObsId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});
  const [isSortByTimeEnabled, setIsSortByTimeEnabled] = useState(false);
  const [colWidths, setColWidths] = useState<Record<string, number>>({ rota: 140, data: 125, inicio: 95, motorista: 230, placa: 100, saida: 95, motivo: 170, observacao: 400, geral: 120, operacao: 140, status: 90, tempo: 90 });
  const [hiddenColumns, setHiddenColumns] = useState<Set<string>>(new Set());
  const [contextMenu, setContextMenu] = useState<{ visible: boolean; x: number; y: number; col: string | null }>({ visible: false, x: 0, y: 0, col: null });
  const [isHiddenColsMenuOpen, setIsHiddenColsMenuOpen] = useState(false);
  const [checklistTooltip, setChecklistTooltip] = useState<{ routeId: string; content: string } | null>(null);
  const [copiedGeralStatus, setCopiedGeralStatus] = useState<string | null>(null);

  // Estados para popup de envio de status
  const [pendingSendOps, setPendingSendOps] = useState<Set<string>>(new Set());
  const [countdowns, setCountdowns] = useState<Record<string, number>>({});
  const [sendingOps, setSendingOps] = useState<Set<string>>(new Set());
  const countdownTimersRef = useRef<Record<string, NodeJS.Timeout>>({});
  const sentTodayRef = useRef<Set<string>>(new Set()); // Evita envio duplicado no mesmo dia
  const blockedUntilRef = useRef<Record<string, number>>({}); // Bloqueio temporário após cancelamento

  const obsDropdownRef = useRef<HTMLDivElement>(null);
  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);
  const filterDropdownRef = useRef<HTMLDivElement>(null);
  const contextMenuRef = useRef<HTMLDivElement>(null);
  const hiddenColsMenuRef = useRef<HTMLDivElement>(null);
  const tooltipTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  const getAccessToken = (): string => {
    const token = currentUser?.accessToken || (window as any).__access_token;
    if (!token) {
      console.error('[RouteDeparture] Token não encontrado!');
      throw new Error('Token de autenticação não encontrado. Por favor, faça login novamente.');
    }
    return token;
  };

  // Verifica se todas as rotas de uma operação estão com statusGeral = 'OK'
  const checkOperationAllOK = (operacao: string): boolean => {
    const operationRoutes = routes.filter(r => r.operacao === operacao);
    if (operationRoutes.length === 0) return false;

    return operationRoutes.every(r => r.statusGeral === 'OK');
  };

  // Inicia o countdown para envio de status
  const startSendCountdown = (operacao: string) => {
    // Se já tem countdown em andamento, não inicia outro
    if (countdownTimersRef.current[operacao]) {
      console.warn(`[COUNTDOWN] ⚠️ Já existe countdown ativo para ${operacao}`);
      return;
    }

    console.log(`[COUNTDOWN] Iniciando 20s para ${operacao}`);
    setCountdowns(prev => ({ ...prev, [operacao]: 20 }));

    countdownTimersRef.current[operacao] = setInterval(() => {
      setCountdowns(prev => {
        const newValue = (prev[operacao] || 1) - 1;
        if (newValue <= 0) {
          // Countdown chegou a zero - chama o webhook
          console.log(`[COUNTDOWN] Tempo esgotado para ${operacao}, chamando webhook`);
          clearInterval(countdownTimersRef.current[operacao]);
          delete countdownTimersRef.current[operacao];
          handleSendStatus(operacao);
          return { ...prev, [operacao]: 0 };
        }
        return { ...prev, [operacao]: newValue };
      });
    }, 1000);
  };

  // Cancela o countdown para uma operação e adiciona bloqueio de 10 minutos
  const cancelSendCountdown = (operacao: string) => {
    if (countdownTimersRef.current[operacao]) {
      clearInterval(countdownTimersRef.current[operacao]);
      delete countdownTimersRef.current[operacao];
    }
    
    // Adiciona bloqueio de 10 minutos (600000 ms)
    const blockedUntil = Date.now() + 10 * 60 * 1000; // 10 minutos
    blockedUntilRef.current[operacao] = blockedUntil;
    
    console.log(`[CANCEL_SEND] ${operacao} bloqueada até ${new Date(blockedUntil).toLocaleTimeString()}`);
    
    setPendingSendOps(prev => {
      const next = new Set(prev);
      next.delete(operacao);
      return next;
    });
    setCountdowns(prev => {
      const next = { ...prev };
      delete next[operacao];
      return next;
    });
  };

  // Envia o status da operação para o webhook ao final do countdown
  const handleSendStatus = async (operacao: string) => {
    const token = getAccessToken();
    const config = userConfigs.find(c => c.operacao === operacao);

    // Filtra rotas da operação
    const operationRoutes = routes.filter(r => r.operacao === operacao);
    if (operationRoutes.length === 0) {
      console.warn(`[WEBHOOK] Nenhuma rota encontrada para ${operacao}`);
      return;
    }

    // Verifica se já enviou hoje para evitar duplicidade
    const today = new Date().toISOString().split('T')[0];
    const sentKey = `${operacao}_${today}`;
    
    if (sentTodayRef.current.has(sentKey)) {
      console.log(`[WEBHOOK] ⚠️ Já enviado hoje para ${operacao}, ignorando envio duplicado`);
      // Limpa estados mesmo assim
      setSendingOps(prev => {
        const next = new Set(prev);
        next.delete(operacao);
        return next;
      });
      setPendingSendOps(prev => {
        const next = new Set(prev);
        next.delete(operacao);
        return next;
      });
      setCountdowns(prev => {
        const next = { ...prev };
        delete next[operacao];
        return next;
      });
      return;
    }

    // Marca como enviado para evitar duplicidade
    sentTodayRef.current.add(sentKey);

    // Adiciona aos estados de envio
    setSendingOps(prev => new Set(prev).add(operacao));

    const payload = {
      tipo: "SAIDAS_AUTO",
      operacao: operacao,
      nomeExibicao: config?.nomeExibicao || operacao,
      tolerancia: config?.tolerancia || "00:00:00",
      atualizacao: "não",
      usuario: currentUser.email,
      dataEnvio: new Date().toISOString(),
      envio: config?.Envio || "", // Emails para envio principal
      copia: config?.Copia || "", // Emails para cópia
      saidas: operationRoutes.map(r => ({
        rota: r.rota,
        data: r.data,
        inicio: r.inicio,
        motorista: r.motorista,
        placa: r.placa,
        saida: r.saida,
        motivo: r.motivo,
        observacao: r.observacao,
        status: r.statusOp
      }))
    };

    const WEBHOOK_URL = import.meta.env.VITE_WEBHOOK_SAIDAS_URL || "https://n8n.datastack.viagroup.com.br/webhook/8cb1f3e1-833d-42a7-a3f0-2f959ea390d6";

    try {
      console.log(`[WEBHOOK_AUTO] 🚀 Enviando status de ${operacao}:`, payload);
      const response = await fetch(WEBHOOK_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        let responseData;
        try {
          responseData = await response.json();
        } catch {
          console.warn("[WEBHOOK_AUTO] Resposta não é JSON válido");
          responseData = { sucesso: true };
        }

        console.log(`[WEBHOOK_AUTO] ✅ Resposta recebida:`, responseData);
        console.log(`[WEBHOOK_AUTO] Campos na resposta:`, Object.keys(responseData));
        console.log(`[WEBHOOK_AUTO] responseData[0]:`, responseData[0]);
        console.log(`[WEBHOOK_AUTO] status:`, responseData[0]?.status || responseData.status);
        console.log(`[WEBHOOK_AUTO] dataEnvioEmail:`, responseData[0]?.dataEnvioEmail || responseData.dataEnvioEmail);
        console.log(`[WEBHOOK_AUTO] horarioEnvioEmail:`, responseData[0]?.horarioEnvioEmail || responseData.horarioEnvioEmail);

        // Atualiza UltimoEnvioSaida se o webhook retornar data/hora
        const dataEnvio = responseData[0]?.dataEnvioEmail || responseData.dataEnvioEmail;
        const horarioEnvio = responseData[0]?.horarioEnvioEmail || responseData.horarioEnvioEmail;

        if (dataEnvio && horarioEnvio) {
          const dataHoraEnvio = `${dataEnvio} ${horarioEnvio}`;
          console.log(`[WEBHOOK_AUTO] Atualizando UltimoEnvioSaida: ${dataHoraEnvio}`);
          await SharePointService.updateUltimoEnvioSaida(token, operacao, dataHoraEnvio);

          // Atualiza estado local IMEDIATAMENTE
          setUserConfigs(prev => prev.map(c =>
            c.operacao === operacao
              ? { ...c, ultimoEnvioSaida: dataHoraEnvio }
              : c
          ));
        } else {
          console.warn(`[WEBHOOK_AUTO] ⚠️ Webhook não retornou dataEnvioEmail/horarioEnvioEmail`);
        }

        // Atualiza Status se o webhook retornar
        const statusRetorno = responseData[0]?.status || responseData.status;
        if (statusRetorno) {
          const statusFormatado = statusRetorno.toLowerCase() === "atualizar" ? "Atualizar" : 
                                  statusRetorno.toLowerCase() === "ok" ? "OK" : statusRetorno;
          console.log(`[WEBHOOK_AUTO] Atualizando Status: ${statusFormatado}`);
          await SharePointService.updateStatusOperacao(token, operacao, statusFormatado);
          
          // Atualiza estado local IMEDIATAMENTE
          setUserConfigs(prev => prev.map(c =>
            c.operacao === operacao
              ? { ...c, Status: statusFormatado }
              : c
          ));
        } else {
          console.warn(`[WEBHOOK_AUTO] ⚠️ Webhook não retornou campo status`);
        }

        console.log(`[WEBHOOK_AUTO] ✅ Status de ${operacao} enviado com sucesso!`);
        
        // Força refresh dos dados após 2 segundos para garantir que o SharePoint replicou (segundo plano)
        setTimeout(() => {
          console.log(`[WEBHOOK_AUTO] 🔄 Forçando refresh dos dados após webhook`);
          loadData(true); // background refresh
        }, 2000);
      } else {
        throw new Error(`Erro na resposta do webhook: ${response.status}`);
      }
    } catch (e: any) {
      console.error(`[WEBHOOK_AUTO] ❌ Erro ao enviar status de ${operacao}:`, e.message);
      // Remove do sentTodayRef em caso de erro para permitir retry
      sentTodayRef.current.delete(sentKey);
      alert(`Erro ao enviar status automático: ${e.message}`);
    } finally {
      // Remove dos estados
      setSendingOps(prev => {
        const next = new Set(prev);
        next.delete(operacao);
        return next;
      });
      setPendingSendOps(prev => {
        const next = new Set(prev);
        next.delete(operacao);
        return next;
      });
      setCountdowns(prev => {
        const next = { ...prev };
        delete next[operacao];
        return next;
      });
    }
  };

  // Limpa timers ao desmontar
  useEffect(() => {
    return () => {
      Object.values(countdownTimersRef.current).forEach(timer => clearInterval(timer));
    };
  }, []);

  // Verifica mudanças nas rotas para disparar o popup
  useEffect(() => {
    if (routes.length === 0 || userConfigs.length === 0) return;

    const today = new Date().toISOString().split('T')[0];
    const now = Date.now();

    // Para cada operação do usuário, verifica se todas as rotas estão OK
    userConfigs.forEach(config => {
      const { operacao, ultimoEnvioSaida } = config;

      // Se já está na lista de pendentes ou enviando, ignora
      if (pendingSendOps.has(operacao) || sendingOps.has(operacao)) return;

      // Verifica se está BLOQUEADA temporariamente (após cancelamento)
      const blockedUntil = blockedUntilRef.current[operacao];
      if (blockedUntil && now < blockedUntil) {
        const remainingMinutes = Math.ceil((blockedUntil - now) / 60000);
        console.log(`[SKIP_AUTO_SEND] ${operacao} está bloqueada por mais ${remainingMinutes} min (cancelamento do usuário)`);
        return; // Pula para a próxima operação
      }

      // Verifica se JÁ FOI ENVIADO HOJE usando o campo ultimoEnvioSaida da config
      if (ultimoEnvioSaida && ultimoEnvioSaida.trim() !== "") {
        let envioDateStr = "";

        // Extrai a data do ultimoEnvioSaida (pode ser ISO ou DD/MM/YYYY HH:MM:SS)
        if (ultimoEnvioSaida.includes('T')) {
          // Formato ISO: 2026-03-13T08:57:00Z
          envioDateStr = ultimoEnvioSaida.split('T')[0];
        } else if (ultimoEnvioSaida.includes('/')) {
          // Formato brasileiro: 13/03/2026 08:57:00
          const [data] = ultimoEnvioSaida.split(' ');
          const [dia, mes, ano] = data.split('/');
          envioDateStr = `${ano}-${mes}-${dia}`;
        }

        // Compara com hoje
        if (envioDateStr === today) {
          console.log(`[SKIP_AUTO_SEND] ${operacao} já foi enviado hoje (${ultimoEnvioSaida}), ignorando`);
          return; // Pula para a próxima operação
        }
      }

      // Verifica se todas as rotas estão OK
      if (checkOperationAllOK(operacao)) {
        console.log(`[COUNTDOWN_TRIGGER] Iniciando countdown para ${operacao}`);
        // Adiciona aos pendentes e inicia countdown
        setPendingSendOps(prev => {
          const next = new Set(prev);
          next.add(operacao);
          return next;
        });
        startSendCountdown(operacao);
      }
    });
  }, [routes, userConfigs, pendingSendOps, sendingOps]);

  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 30000);
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
      // Fechar menu de colunas ocultas ao clicar fora
      if (hiddenColsMenuRef.current && !hiddenColsMenuRef.current.contains(event.target as Node)) {
        setIsHiddenColsMenuOpen(false);
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
    if (!inicio || inicio === '' || inicio === '00:00:00') return { status: 'Previsto', gap: '' };
    if (!routeDate) return { status: 'Previsto', gap: '' };

    // Se saida for "-", considera rota não saída (atrasada)
    if (saida === '-') {
        return { status: 'Atrasada', gap: 'Não saiu' };
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const [y, m, d] = routeDate.split('-').map(Number);
    const rDate = new Date(y, m - 1, d);
    rDate.setHours(0, 0, 0, 0);

    const toleranceSec = timeToSeconds(toleranceStr);
    const startSec = timeToSeconds(inicio);

    if (saida && saida !== '00:00:00' && saida !== '') {
        const endSec = timeToSeconds(saida);
        const diff = endSec - startSec;
        const gapFormatted = secondsToTime(diff);

        if (diff < -toleranceSec) return { status: 'Adiantada', gap: gapFormatted };
        if (diff > toleranceSec) return { status: 'Atrasada', gap: gapFormatted };

        return { status: 'OK', gap: gapFormatted };
    }

    if (rDate > today) return { status: 'Programada', gap: '' };
    if (rDate < today) return { status: 'Atrasada', gap: '' };

    const nowSec = currentTime.getHours() * 3600 + currentTime.getMinutes() * 60 + currentTime.getSeconds();
    if (nowSec > (startSec + toleranceSec)) {
        return { status: 'Atrasada', gap: '' };
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
      // Data padrão no formato DD/MM/AAAA
      const today = new Date();
      data = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
    }
    
    const porcentagem = percentMatch ? percentMatch[1] : '100';
    const motivos = motivosMatch && !currentText.includes(percentMatch ? percentMatch[0] : '') ? motivosMatch[1] : '';

    setChecklistEditData({ routeId, data, porcentagem, motivos });
    setIsChecklistEditModalOpen(true);
  };

  // Aplica a edição do checklist
  const applyChecklistEdit = () => {
    if (!checklistEditData) return;

    const { routeId, data, porcentagem, motivos } = checklistEditData;

    // Salva apenas os dados do checklist (sem o texto "Último checklist realizado")
    let result = `${data} - ${porcentagem}%`;

    // Se tem motivos e porcentagem < 100%, adiciona com hífen
    if (motivos && motivos.trim() !== '' && parseInt(porcentagem) < 100) {
      result += ` - ${motivos}`;
    }

    console.log('[CHECKLIST] Salvando:', { routeId, result });
    setIsSyncing(true);
    try {
      updateCell(routeId, 'checklistMotorista', result);
    } finally {
      setIsSyncing(false);
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
      token = getAccessToken();
    } catch (e: any) {
      console.error('[RouteDeparture] Erro ao obter token:', e.message);
      alert('Sessão expirada. Você será redirecionado para o login.');
      window.location.href = '/';
      return;
    }

    // Só mostra loading se NÃO for refresh em segundo plano
    if (!isBackgroundRefresh) {
      setIsLoading(true);
    }
    
    try {
      console.log('[LOAD_DATA] Buscando dados atualizados...', isBackgroundRefresh ? '(segundo plano)' : '(inicial)');
      const [configs, mappings, spData] = await Promise.all([
        SharePointService.getRouteConfigs(token, currentUser.email, true), // force refresh
        SharePointService.getRouteOperationMappings(token),
        SharePointService.getDepartures(token, true) // force refresh
      ]);
      setUserConfigs(configs || []);
      setRouteMappings(mappings || []);
      setRoutes(spData || []);
      console.log('[LOAD_DATA] Dados carregados com sucesso');

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
            const percentage = totalOps > 0 ? Math.round((okOps / totalOps) * 100) : 0;
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

  // Polling para atualizar dados automaticamente a cada 10 segundos (SEGUNDO PLANO)
  useEffect(() => {
    const refreshInterval = setInterval(() => {
      console.log('[POLLING_ROUTE_DEPARTURE] Atualização automática de dados (segundo plano)');
      loadData(true); // true = background refresh
    }, 10000);

    return () => clearInterval(refreshInterval);
  }, []);

  const handleDeleteRoute = async (id: string) => {
    if (!confirm('Deseja excluir permanentemente esta rota do SharePoint?')) return;
    const token = getAccessToken();
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
    if (!confirm(`Deseja excluir as ${selectedIds.size} rotas selecionadas do SharePoint?`)) return;
    const token = getAccessToken();
    setIsSyncing(true);
    let success = 0;
    const idsToProcess = Array.from(selectedIds) as string[];
    for (const id of idsToProcess) {
        try { await SharePointService.deleteDeparture(token, id); success++; } catch (e) {}
    }
    setRoutes(prev => prev.filter(r => !selectedIds.has(r.id!)));
    setSelectedIds(new Set());
    setIsSyncing(false);
    alert(`${success} rotas excluídas.`);
  };

  const handleArchiveAll = async () => {
    if (filteredRoutes.length === 0) {
      alert("Não há rotas para arquivar.");
      return;
    }

    if (!confirm(`Arquivar ${filteredRoutes.length} itens no histórico e limpar status de envio?`)) return;

    const token = getAccessToken();
    setIsSyncing(true);

    try {
      // Passo 1: Mover rotas para o histórico
      console.log(`[ARCHIVE] Movendo ${filteredRoutes.length} itens para o histórico...`);
      const archiveResult = await SharePointService.moveDeparturesToHistory(token, filteredRoutes as RouteDeparture[]);
      console.log(`[ARCHIVE] Sucesso: ${archiveResult.success}, Falhas: ${archiveResult.failed}`);

      // Passo 2: Limpar status de envio (OK, ATUALIZAR) na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
      console.log('[ARCHIVE] Limpando status de envio nas configurações...');
      const opsToClear = Array.from(new Set(filteredRoutes.map(r => r.operacao)));
      let clearCount = 0;

      for (const operacao of opsToClear) {
        try {
          // Limpa UltimoEnvioSaida
          await SharePointService.updateUltimoEnvioSaida(token, operacao, '');
          // Limpa Status
          await SharePointService.updateStatusOperacao(token, operacao, '');
          // Limpa UltimoEnvioResumoSaida
          await SharePointService.updateUltimoEnvioResumoSaida(token, operacao, '');
          // Limpa StatusResumoSaida
          await SharePointService.updateStatusResumoSaida(token, operacao, '');
          clearCount++;
          console.log(`[ARCHIVE] ✅ Status limpo para ${operacao}`);
        } catch (e: any) {
          console.error(`[ARCHIVE] Erro ao limpar status de ${operacao}:`, e.message);
        }
      }

      // Passo 3: Recarregar dados com force refresh
      await loadData(true);
      
      // Passo 4: Forçar refresh extra nas configs para garantir que o status "Previsto" apareça
      try {
        const refreshedConfigs = await SharePointService.getRouteConfigs(token, currentUser.email, true);
        setUserConfigs(refreshedConfigs);
        console.log('[ARCHIVE] ✅Configs atualizadas após arquivamento');
      } catch (e: any) {
        console.error('[ARCHIVE] Erro ao atualizar configs:', e.message);
      }

      alert(`${archiveResult.success} rotas arquivadas com sucesso!\nStatus de envio limpo para ${clearCount} operações.`);
    } catch (e: any) {
      console.error('[ARCHIVE] Erro geral:', e.message);
      alert(`Erro ao arquivar: ${e.message}`);
    } finally {
      setIsSyncing(false);
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
    const token = getAccessToken();
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
    setGhostRow({ id: 'ghost', rota: '', data: new Date().toISOString().split('T')[0], inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '' });
  };

  const handleMultilinePaste = async (field: keyof RouteDeparture, startRowIndex: number, value: string) => {
    const lines = value.split(/[\n\r]/).map(l => l.trim()).filter(Boolean);
    if (lines.length <= 1) return;
    const token = getAccessToken();
    setIsSyncing(true);
    
    // Usa routes diretamente em vez de filteredRoutes para evitar problemas de sincronização
    const targetRoutes = routes.slice(startRowIndex, startRowIndex + lines.length);
    
    if (targetRoutes.length === 0) {
        setIsSyncing(false);
        return;
    }
    
    // Prepara todas as atualizações
    const updatePromises = targetRoutes.map(async (route, i) => {
        let finalValue = lines[i];
        if (field === 'inicio' || field === 'saida') {
            finalValue = formatTimeInput(finalValue);
        }
        const updatedRoute: RouteDeparture = { ...route, [field]: finalValue };
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
            if (result !== null) {
                const index = newRoutes.findIndex(r => r.id === result.id);
                if (index !== -1) {
                    newRoutes[index] = result.updatedRoute;
                }
            }
        });
        return newRoutes;
    });
    
    setIsSyncing(false);
  };

  const updateCell = async (id: string, field: keyof RouteDeparture, value: string) => {
    if (id === 'ghost') {
        const updatedGhost = { ...ghostRow, [field]: value };
        if (field === 'rota' && (value.includes('\n') || value.includes(';'))) {
            const lines = value.split(/[\n;]/).map(l => l.trim()).filter(Boolean);
            if (lines.length > 1) { setPendingBulkRoutes(lines); setIsBulkMappingModalOpen(true); return; }
        }
        if (field === 'rota' && value !== "") {
            const mapping = routeMappings.find(m => m.Title === value);
            if (mapping) updatedGhost.operacao = mapping.OPERACAO;
            else { setPendingMappingRoute(value); setIsMappingModalOpen(true); }
        }
        if (field !== 'rota' && updatedGhost.rota) {
            setIsSyncing(true);
            try {
                const config = userConfigs.find(c => c.operacao === updatedGhost.operacao);
                const { status, gap } = calculateStatusWithTolerance(updatedGhost.inicio || '', updatedGhost.saida || '', config?.tolerancia || "00:00:00", updatedGhost.data || "");
                const payload = { ...updatedGhost, statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture;
                const newId = await SharePointService.updateDeparture(getAccessToken(), payload);
                setRoutes(prev => [...prev, { ...payload, id: newId }]);
                setGhostRow({ id: 'ghost', rota: '', data: new Date().toISOString().split('T')[0], inicio: '', saida: '', motorista: '', placa: '', statusGeral: '', aviso: 'NÃO', operacao: '', statusOp: 'Previsto', tempo: '' });
            } catch (e) {} finally { setIsSyncing(false); }
        } else { setGhostRow(updatedGhost); }
        return;
    }

    const route = routes.find(r => r.id === id);
    if (!route) return;

    let updatedRoute = { ...route, [field]: value };

    // Validação específica para MONTES CLAROS + FÁBRICA quando editar observação
    if (field === 'observacao' && value) {
      const valid = validateDescargaTime(updatedRoute, value);
      if (!valid) {
        return; // Cancela a atualização
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
      }
    }

    setRoutes(prev => prev.map(r => r.id === id ? updatedRoute : r));
    setIsSyncing(true);

    try {
        await SharePointService.updateDeparture(getAccessToken(), updatedRoute);
    } catch (e) {
        console.error('[UPDATE] Error:', e);
    } finally {
        setIsSyncing(false);
    }
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
    localStorage.removeItem('route_departure_col_widths');
    localStorage.removeItem('route_departure_hidden_cols');
  };

  const getRowStyle = (route: RouteDeparture | Partial<RouteDeparture>) => {
    if (route.id === 'ghost') return "bg-slate-50 dark:bg-slate-900 italic text-slate-400";
    const status = route.statusOp;
    
    // Se a saída for "-", aplica estilo de atrasado crítico (não saiu)
    if (route.saida === '-') {
      return "bg-red-600 dark:bg-red-700/40 text-white font-bold border-l-[12px] border-red-800 shadow-lg";
    }
    
    if (status === 'Previsto') return "bg-slate-50 dark:bg-slate-900 border-l-4 border-slate-300 text-slate-400 dark:text-slate-500";
    if (status === 'Programada') return "bg-slate-100 dark:bg-slate-800 border-l-4 border-slate-400 text-slate-500 dark:text-slate-400";
    if (status === 'OK') return "bg-emerald-50 dark:bg-emerald-900/10 border-l-4 border-emerald-600";
    if (status === 'Atrasada' && (!route.saida || route.saida === '00:00:00' || route.saida === '')) {
      return "bg-yellow-300 dark:bg-yellow-500/30 text-slate-900 dark:text-yellow-100 font-bold border-l-[12px] border-yellow-600 shadow-lg";
    }
    if (status === 'Atrasada' || status === 'Adiantada') {
      return "bg-orange-500 dark:bg-orange-600/30 text-white font-bold border-l-[12px] border-orange-700 shadow-lg";
    }
    return "bg-white dark:bg-slate-900 border-l-4 border-transparent";
  };

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

    return result;
  }, [routes, colFilters, selectedFilters, isSortByTimeEnabled, userConfigs]);

  const dashboardStats = useMemo(() => {
    const total = filteredRoutes.length; if (total === 0) return null;
    const okCount = filteredRoutes.filter(r => r.statusOp === 'OK').length;
    const delayedCount = filteredRoutes.filter(r => r.statusOp === 'Atrasada').length;
    return { total, okCount, delayedCount };
  }, [filteredRoutes]);

  // Cálculo dos indicadores GERAL e INTERNO
  const performanceIndicators = useMemo(() => {
    const total = filteredRoutes.length;
    if (total === 0) return { geral: 0, interno: 0 };

    // GERAL: (OK + PREVISTO) / total * 100
    const okPrevistoCount = filteredRoutes.filter(r => 
      r.statusOp === 'OK' || r.statusOp === 'Previsto'
    ).length;
    const geral = Math.round((okPrevistoCount / total) * 100);

    // INTERNO: (total - justificativas) / total * 100
    // Justificativas: Manutenção, Mão de obra, Logística
    const justificativas = ['Manutenção', 'Mão de obra', 'Logística'];
    const rotasComJustificativa = filteredRoutes.filter(r => 
      justificativas.includes(r.motivo)
    ).length;
    const rotasSemJustificativa = total - rotasComJustificativa;
    const interno = Math.round((rotasSemJustificativa / total) * 100);

    return { geral, interno };
  }, [filteredRoutes]);

  const handleSearchArchive = async () => {
    setIsSearchingArchive(true);
    try {
        console.log('[SEARCH_ARCHIVE] Requesting history from SharePoint list {856bf9d5-6081-4360-bcad-e771cbabfda8}...');
        const results = await SharePointService.getArchivedDepartures(getAccessToken(), null, histStart, histEnd);
        console.log('[SEARCH_ARCHIVE] Results received:', results.length);
        
        const myOps = new Set(userConfigs.map(c => c.operacao));
        // If myOps is empty, show everything for the user to avoid blockage if config loading is slow
        const filtered = results && results.length > 0 
          ? results.filter(r => !myOps.size || myOps.has(r.operacao))
          : [];
        
        setArchivedResults(filtered);
    } catch (err: any) { 
        console.error('[SEARCH_ARCHIVE] Error during search:', err);
        alert("Erro na busca: " + (err?.message || "Erro desconhecido ao acessar o SharePoint. Verifique se você tem permissão na lista de histórico.")); 
    } finally { 
        setIsSearchingArchive(false); 
    }
  };

  const tableColumns = [
    { id: 'rota', label: 'ROTA' }, { id: 'data', label: 'DATA' }, { id: 'inicio', label: 'INÍCIO' },
    { id: 'motorista', label: 'MOTORISTA' }, { id: 'placa', label: 'PLACA' }, { id: 'saida', label: 'SAÍDA' },
    { id: 'motivo', label: 'MOTIVO' }, { id: 'observacao', label: 'OBSERVAÇÃO' }, { id: 'geral', label: 'GERAL' },
    { id: 'operacao', label: 'OPERAÇÃO' }, { id: 'status', label: 'STATUS' }, { id: 'tempo', label: 'TEMPO' }
  ];

  if (isLoading) return <div className="h-full flex flex-col items-center justify-center text-primary-500 gap-4"><Loader2 size={48} className="animate-spin" /><p className="font-bold text-[10px] uppercase tracking-widest">Carregando Grid...</p></div>;

  return (
    <div className="flex flex-col h-full bg-[#020617] p-4 overflow-hidden select-none font-sans animate-fade-in relative">
      {/* Header Section */}
      <div className="flex justify-between items-center mb-6 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-primary-600 text-white rounded-2xl shadow-lg"><Clock size={20} /></div>
          <div>
            <h2 className="text-xl font-black text-white uppercase tracking-tight flex items-center gap-3">
              Controle de Saídas {isSyncing && <Loader2 size={16} className="animate-spin text-primary-500"/>}
            </h2>
            <p className="text-[9px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
              <ShieldCheck size={12} className="text-emerald-500"/> Operador: {currentUser.name}
            </p>
          </div>
          {/* Indicadores GERAL e INTERNO */}
          <div className="flex items-center gap-3 ml-8">
            <div className="flex items-center gap-2 px-4 py-2 bg-emerald-900/30 border border-emerald-700/50 rounded-xl">
              <div className="text-right">
                <p className="text-[8px] font-black text-emerald-400 uppercase tracking-wider">Geral</p>
                <p className="text-lg font-black text-emerald-400 leading-none">{performanceIndicators.geral}%</p>
              </div>
              <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></div>
            </div>
            <div className="flex items-center gap-2 px-4 py-2 bg-blue-900/30 border border-blue-700/50 rounded-xl">
              <div className="text-right">
                <p className="text-[8px] font-black text-blue-400 uppercase tracking-wider">Interno</p>
                <p className="text-lg font-black text-blue-400 leading-none">{performanceIndicators.interno}%</p>
              </div>
              <div className="w-2 h-2 bg-blue-500 rounded-full animate-pulse"></div>
            </div>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <button onClick={() => setIsSortByTimeEnabled(!isSortByTimeEnabled)} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] transition-all ${isSortByTimeEnabled ? 'bg-primary-600 text-white border-primary-600' : 'bg-slate-800 text-slate-300 border-slate-700'}`}><SortAsc size={16} /> Horário</button>
          <div className="relative">
            <button
              onClick={() => setIsHiddenColsMenuOpen(!isHiddenColsMenuOpen)}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] transition-all relative ${hiddenColumns.size > 0 ? 'bg-amber-600 text-white border-amber-600' : 'bg-slate-800 text-slate-300 border-slate-700'}`}
            >
              <Settings2 size={16} /> Colunas
              {hiddenColumns.size > 0 && (
                <span className="absolute -top-1 -right-1 w-4 h-4 bg-red-500 text-white text-[8px] font-black rounded-full flex items-center justify-center">
                  {hiddenColumns.size}
                </span>
              )}
            </button>
            {isHiddenColsMenuOpen && (
              <div
                ref={hiddenColsMenuRef}
                className="absolute right-0 top-full mt-2 z-[1000] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-2xl py-2 min-w-[200px] animate-in fade-in zoom-in-95 duration-150"
              >
                <div className="px-3 py-2 border-b border-slate-100 dark:border-slate-700">
                  <p className="text-[10px] font-black uppercase text-slate-400">Colunas Ocultas ({hiddenColumns.size})</p>
                </div>
                {hiddenColumns.size === 0 ? (
                  <div className="px-4 py-3 text-[11px] text-slate-400 font-bold text-center">Todas as colunas visíveis</div>
                ) : (
                  Array.from(hiddenColumns).map(col => (
                    <button
                      key={col}
                      onClick={() => toggleColumnVisibility(col)}
                      className="w-full px-4 py-2 text-left text-[11px] font-bold text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center justify-between"
                    >
                      <span className="uppercase">{col}</span>
                      <Check size={14} className="text-green-500" />
                    </button>
                  ))
                )}
                {hiddenColumns.size > 0 && (
                  <div className="border-t border-slate-100 dark:border-slate-700 mt-1 pt-1">
                    <button
                      onClick={resetColumnSettings}
                      className="w-full px-4 py-2 text-left text-[11px] font-bold text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors flex items-center gap-2"
                    >
                      <RefreshCw size={14} /> Resetar Tudo
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>
          <button onClick={() => setIsHistoryModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-bold border border-slate-700 uppercase text-[10px] tracking-wide transition-all shadow-sm"><Database size={16} /> Histórico</button>
          <button onClick={() => setIsStatsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-bold border border-slate-700 uppercase text-[10px] tracking-wide transition-all shadow-sm"><BarChart3 size={16} /> Dashboard</button>
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-all border border-slate-700 bg-slate-900"><RefreshCw size={18} /></button>
          <button onClick={handleArchiveAll} disabled={isSyncing || filteredRoutes.length === 0} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-bold border border-slate-700 uppercase text-[10px] tracking-wide transition-all shadow-sm disabled:opacity-50 disabled:cursor-not-allowed"><Archive size={16} /> Arquivar</button>
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
      <div className="flex-1 overflow-auto bg-white dark:bg-slate-900 rounded-2xl border border-slate-700/50 shadow-2xl relative scrollbar-thin" id="table-container">
        <table className="border-collapse" style={{ width: `${tableColumns.filter(col => !hiddenColumns.has(col.id)).reduce((acc, col) => acc + colWidths[col.id], 0) + 60}px` }}>
          <thead className="sticky top-0 z-50 bg-[#1e293b] text-white shadow-md">
            <tr>
              {tableColumns.filter(col => !hiddenColumns.has(col.id)).map(col => (
                <th key={col.id} data-col={col.id} style={{ width: colWidths[col.id] }} className="relative p-0 border border-slate-700/50 text-[10px] font-black uppercase tracking-wider text-left group">
                  <div className="flex items-center justify-between px-3 h-[48px]">
                    <span onContextMenu={(e) => handleContextMenu(e, col.id)}>{col.label}</span>
                    <button onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }} className={`p-1 rounded ${!!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0 ? 'text-yellow-400' : 'text-white/40'}`}><Filter size={11} /></button>
                  </div>
                  {activeFilterCol === col.id && <FilterDropdown col={col.id} routes={routes} colFilters={colFilters} setColFilters={setColFilters} selectedFilters={selectedFilters} setSelectedFilters={setSelectedFilters} onClose={() => setActiveFilterCol(null)} dropdownRef={filterDropdownRef} />}
                  <div onMouseDown={(e) => { e.preventDefault(); resizingRef.current = { col: col.id, startX: e.clientX, startWidth: colWidths[col.id] }; }} className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10" />
                </th>
              ))}
              <th style={{ width: 60 }} className="relative p-0 border border-slate-700/50 text-[10px] font-black uppercase text-center bg-slate-900/50">
                  {selectedIds.size > 0 ? (
                      <button onClick={handleDeleteSelected} className="p-1 text-red-500 hover:text-red-400 transition-colors" title="Deletar Selecionados"><Trash2 size={16} /></button>
                  ) : <Settings2 size={14} className="mx-auto opacity-40" />}
              </th>
            </tr>
          </thead>
          <tbody>
            {[...filteredRoutes, ghostRow].map((route, rowIndex) => {
              const rowStyle = getRowStyle(route);
              const isGhost = route.id === 'ghost';
              const isSelected = selectedIds.has(route.id!);
              const isDelayed = route.statusOp === 'Atrasada' || route.statusOp === 'Adiantada';
              const isDelayedFilled = isDelayed && (route.saida !== '' && route.saida !== '00:00:00');
              const inputClass = `w-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold uppercase transition-all ${isDelayedFilled ? 'text-white placeholder-white/50' : 'text-slate-800 dark:text-slate-200 placeholder-slate-400'}`;

              return (
                <tr key={route.id} className={`${isSelected ? 'bg-primary-600/20' : rowStyle} group transition-all`} style={{ height: 'auto', minHeight: '48px' }}>
                  {tableColumns.filter(col => !hiddenColumns.has(col.id)).map(col => {
                    const cellKey = `${route.id}-${col.id}`;

                    if (col.id === 'rota') {
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700" style={{ verticalAlign: 'top' }}>
                            {isGhost ? (
                                <textarea rows={1} value={route.rota} placeholder="Digite p/ criar..." onChange={(e) => updateCell(route.id!, 'rota', e.target.value)} onInput={(e: any) => { e.target.style.height = 'auto'; e.target.style.height = (e.target.scrollHeight) + 'px'; }} className={`${inputClass} font-black resize-none overflow-hidden min-h-[48px]`} />
                            ) : (
                                <input type="text" value={route.rota} onChange={(e) => updateCell(route.id!, 'rota', e.target.value)} className={`${inputClass} font-black`} />
                            )}
                        </td>
                      );
                    }

                    if (col.id === 'data') {
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px]">
                          <input type="date" value={route.data} onChange={(e) => updateCell(route.id!, 'data', e.target.value)} className={`${inputClass} text-center`} />
                        </td>
                      );
                    }

                    if (col.id === 'inicio') {
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px]">
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
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px]">
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
                              className={`${inputClass}`}
                          />
                        </td>
                      );
                    }

                    if (col.id === 'placa') {
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px]">
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
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px]">
                          <input
                              type="text"
                              key={route.id + '-saida'}
                              value={route.saida || ''}
                              placeholder="--:--:--"
                              onChange={(e) => {
                                  const val = e.target.value;
                                  // Se digitou "-", mantém como está (rota não saiu)
                                  if (val === '-') {
                                      updateCell(route.id!, 'saida', '-');
                                  } else {
                                      const masked = applyTimeMask(val);
                                      updateCell(route.id!, 'saida', masked);
                                  }
                              }}
                              onPaste={(e: any) => {
                                  const val = e.clipboardData.getData('text');
                                  if (val.includes('\n')) {
                                      e.preventDefault();
                                      handleMultilinePaste('saida', rowIndex, val);
                                  }
                              }}
                              onBlur={(e) => {
                                  const val = e.target.value;
                                  // Se for "-", mantém, caso contrário formata como horário
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
                      const isMaintenance = route.motivo === 'Manutenção';
                      
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px]">
                          {(isDelayed || route.statusOp === 'Programada' || route.statusOp === 'Previsto') && !isGhost && (
                            <select 
                              value={route.motivo} 
                              onChange={(e) => updateCell(route.id!, 'motivo', e.target.value)} 
                              className="w-full bg-white/20 dark:bg-slate-800/20 border-none px-2 py-1 text-[10px] font-bold text-inherit outline-none appearance-none cursor-pointer"
                              disabled={!isMaintenance && route.motivo !== '' && route.statusOp === 'OK'}
                            >
                                <option value="" className="text-slate-800">---</option>
                                {MOTIVOS.map(m => (<option key={m} value={m} className="text-slate-800">{m}</option>))}
                            </select>
                          )}
                          
                          {/* Campo vazio ou OK quando não é manutenção e já tem valor */}
                          {!isMaintenance && route.motivo !== '' && route.statusOp === 'OK' && !isGhost && (
                            <div className="w-full h-full flex items-center px-3 text-[10px] font-bold uppercase text-slate-400">
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
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px] relative align-top">
                          {canEdit && !isGhost && (
                            <div className="flex items-start w-full h-full relative p-0 min-h-[44px]">
                              <textarea
                                value={route.observacao || ""}
                                onChange={(e) => updateCell(route.id!, 'observacao', e.target.value)}
                                onFocus={() => setActiveObsId(route.id!)}
                                placeholder="..."
                                disabled={!isMaintenance && route.motivo !== '' && route.statusOp === 'OK'}
                                className={`w-full h-full min-h-[44px] bg-transparent outline-none border-none px-3 py-2 text-[11px] font-normal resize-none overflow-hidden whitespace-normal pr-28 ${!isMaintenance && route.motivo !== '' && route.statusOp === 'OK' ? 'text-slate-400 cursor-not-allowed' : ''}`}
                                onInput={(e: any) => { e.target.style.height = 'auto'; e.target.style.height = e.target.scrollHeight + 'px'; }}
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
                                <div ref={obsDropdownRef} className="absolute top-full left-0 w-full z-[110] bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-xl shadow-2xl overflow-hidden animate-in fade-in slide-in-from-top-1">
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
                            <div className="w-full h-full px-3 py-2 text-[11px] text-slate-400 whitespace-normal">
                              {route.observacao}
                            </div>
                          )}
                        </td>
                      );
                    }

                    if (col.id === 'geral') {
                      const hasCopiedValue = copiedGeralStatus && copiedGeralStatus !== '';
                      return (
                        <td key={cellKey} data-col-cell="geral" className="p-0 border border-slate-300 dark:border-slate-700 h-[48px] relative">
                          <button
                            onClick={() => {
                              const newValue = route.statusGeral === 'OK' ? '' : 'OK';
                              updateCell(route.id!, 'statusGeral', newValue);
                            }}
                            className={`w-full h-full font-bold text-[10px] transition-all ${
                              route.statusGeral === 'OK'
                                ? 'bg-emerald-600 text-white'
                                : 'bg-transparent text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800'
                            }`}
                            title={hasCopiedValue ? `Valor copiado: "${copiedGeralStatus}" - Selecione rotas e pressione Ctrl+V para colar` : 'Clique para alternar OK/vazio'}
                          >
                            {route.statusGeral || '---'}
                          </button>
                          {/* Indicador visual de valor copiado (ponto verde sem animação) */}
                          {hasCopiedValue && (
                            <div className="absolute top-1 right-1 w-2 h-2 bg-emerald-500 rounded-full" title={`Valor copiado: "${copiedGeralStatus}" - Selecione rotas e pressione Ctrl+V`}></div>
                          )}
                        </td>
                      );
                    }

                    if (col.id === 'operacao') {
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px]">
                          <select value={route.operacao} onChange={(e) => updateCell(route.id!, 'operacao', e.target.value)} className="w-full h-full bg-transparent border-none text-[9px] font-black text-center uppercase">
                            <option value="">OP...</option>
                            {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                          </select>
                        </td>
                      );
                    }

                    if (col.id === 'status') {
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px] text-center">
                          <span className={`px-2 py-0.5 rounded-full text-[8px] font-black border ${route.statusOp === 'OK' ? 'bg-emerald-100 border-emerald-400 text-emerald-800' : route.statusOp === 'Atrasada' ? 'bg-yellow-100 border-yellow-400 text-yellow-800' : route.statusOp === 'Programada' ? 'bg-slate-200 border-slate-400 text-slate-600' : route.statusOp === 'Previsto' ? 'bg-slate-100 border-slate-300 text-slate-500' : 'bg-red-100 border-red-400 text-red-800'}`}>
                            {route.statusOp}
                          </span>
                        </td>
                      );
                    }

                    if (col.id === 'tempo') {
                      return (
                        <td key={cellKey} className="p-0 border border-slate-300 dark:border-slate-700 h-[48px] text-center font-mono font-bold text-[10px]">
                          {route.tempo}
                        </td>
                      );
                    }

                    return null;
                  })}
                  <td className="p-0 border border-slate-300 dark:border-slate-700 flex items-center justify-center gap-1 h-[48px]">
                    {!isGhost && (
                      <>
                        <button onClick={() => toggleSelection(route.id!)} className={`p-1.5 rounded-lg transition-colors ${isSelected ? 'text-primary-500 bg-primary-500/10' : 'text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
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
          </tbody>
        </table>
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
                  <div className="grid grid-cols-2 gap-3">{userConfigs.map(c => ( <button key={c.operacao} onClick={() => { SharePointService.addRouteOperationMapping(getAccessToken(), pendingMappingRoute!, c.operacao); setGhostRow(prev => ({...prev, operacao: c.operacao})); setIsMappingModalOpen(false); }} className="p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 rounded-2xl hover:bg-primary-600 hover:text-white transition-all font-black text-xs uppercase">{c.operacao}</button> ))}</div>
                  <button onClick={() => { setIsMappingModalOpen(false); setGhostRow(prev => ({...prev, rota: ''})); }} className="w-full mt-6 py-4 text-[10px] font-black uppercase text-slate-400">Cancelar</button>
              </div>
          </div>
      )}

      {isHistoryModalOpen && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[200] flex items-center justify-center p-4">
              <div className="bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-6xl max-h-[90vh] overflow-hidden flex flex-col">
                  <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white"><div className="flex items-center gap-4"><Database size={24} /><h3 className="font-black uppercase tracking-widest text-base">Histórico Definitivo</h3></div><button onClick={() => setIsHistoryModalOpen(false)}><X size={28} /></button></div>
                  <div className="p-6 bg-slate-50 dark:bg-slate-900 border-b dark:border-slate-800 grid grid-cols-3 gap-4"><input type="date" value={histStart} onChange={e => setHistStart(e.target.value)} className="p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white" /><input type="date" value={histEnd} onChange={e => setHistEnd(e.target.value)} className="p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white" /><button onClick={handleSearchArchive} disabled={isSearchingArchive} className="py-3 bg-primary-600 text-white font-black uppercase text-[11px] rounded-xl flex items-center justify-center gap-2 hover:bg-primary-700 shadow-lg">{isSearchingArchive ? <Loader2 size={16} className="animate-spin" /> : <Search size={16} />} BUSCAR</button></div>
                  <div className="flex-1 overflow-auto p-4 bg-slate-50 dark:bg-slate-950">{archivedResults.length > 0 ? ( <table className="w-full border-collapse text-[10px]"><thead className="sticky top-0 bg-slate-200 dark:bg-slate-800 text-slate-600 font-black uppercase"><tr><th className="p-2 border border-slate-300 dark:border-slate-700 text-left">Rota</th><th className="p-2 border border-slate-300 text-center">Data</th><th className="p-2 border border-slate-300 text-center">Saída</th><th className="p-2 border border-slate-300 text-left">Motivo</th><th className="p-2 border border-slate-300 text-center">OP</th></tr></thead><tbody>{archivedResults.map((r, i) => (<tr key={i} className="hover:bg-slate-50 dark:hover:bg-slate-800 border-b border-slate-200 dark:border-slate-800"><td className="p-2 font-bold text-primary-700">{r.rota}</td><td className="p-2 text-center">{r.data}</td><td className="p-2 text-center font-mono">{r.saida}</td><td className="p-2">{r.motivo || "---"}</td><td className="p-2 text-center font-black">{r.operacao}</td></tr>))}</tbody></table> ) : <div className="h-full flex flex-col items-center justify-center text-slate-400 italic font-bold">{isSearchingArchive ? "Buscando..." : "Nenhum dado retornado para este período"}</div>}</div>
              </div>
          </div>
      )}

      {isStatsModalOpen && dashboardStats && (
        <div className="fixed inset-0 bg-slate-950/70 backdrop-blur-md z-[200] flex items-center justify-center p-4">
            <div className="bg-white dark:bg-slate-900 rounded-[2rem] shadow-2xl w-full max-w-4xl overflow-hidden border dark:border-slate-800 animate-in zoom-in">
                <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white"><div className="flex items-center gap-4"><TrendingUp size={24} /><h3 className="font-black uppercase tracking-widest text-base">Dashboard Operacional</h3></div><button onClick={() => setIsStatsModalOpen(false)}><X size={28} /></button></div>
                <div className="p-8 grid grid-cols-3 gap-6 bg-slate-50 dark:bg-slate-950">{[{ label: 'Total', value: dashboardStats.total, icon: Activity, color: 'text-slate-700 bg-white' }, { label: 'OK', value: `${Math.round((dashboardStats.okCount / dashboardStats.total) * 100)}%`, icon: CheckCircle2, color: 'text-emerald-600 bg-emerald-50' }, { label: 'Atrasos', value: `${Math.round((dashboardStats.delayedCount / dashboardStats.total) * 100)}%`, icon: AlertTriangle, color: 'text-orange-600 bg-orange-50' }].map((stat: any, idx) => ( <div key={idx} className={`p-6 rounded-2xl border dark:border-slate-800 flex flex-col gap-2 ${stat.color}`}><stat.icon size={20} /><span className="text-[10px] font-black uppercase text-slate-400 mt-2">{stat.label}</span><div className="text-3xl font-black">{stat.value}</div></div> ))}</div>
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

      {/* Popup de Envio de Status - Parte Inferior */}
      <div className="fixed bottom-0 left-0 right-0 z-[100] flex flex-col items-center gap-3 pb-4 pointer-events-none">
        {Array.from(pendingSendOps).map(operacao => {
          const countdown = countdowns[operacao] || 20;
          const isSending = sendingOps.has(operacao);
          return (
            <div
              key={operacao}
              className="pointer-events-auto bg-emerald-600 hover:bg-emerald-700 transition-colors text-white px-6 py-3 rounded-t-2xl shadow-2xl flex items-center gap-4 animate-in slide-in-from-bottom duration-300 border-t border-x border-emerald-500"
            >
              <div className="flex items-center gap-3">
                {isSending ? (
                  <Loader2 size={20} className="animate-spin" />
                ) : (
                  <CheckCircle2 size={20} className="animate-pulse" />
                )}
                <span className="text-[11px] font-black uppercase tracking-widest">
                  {isSending ? (
                    <>Enviando status de <span className="text-emerald-200 font-black">{operacao}</span> para webhook...</>
                  ) : (
                    <>Enviando status de saída de <span className="text-emerald-200 font-black">{operacao}</span> em <span className="text-emerald-200 font-black w-6 text-center inline-block">{countdown}</span> segundos</>
                  )}
                </span>
              </div>
              {!isSending && (
                <button
                  onClick={() => cancelSendCountdown(operacao)}
                  className="px-4 py-1.5 bg-emerald-800 hover:bg-emerald-900 text-white rounded-xl text-[10px] font-black uppercase tracking-wider transition-all border border-emerald-700 shadow-lg"
                >
                  Cancelar
                </button>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default RouteDepartureView;
