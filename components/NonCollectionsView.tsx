import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteConfig, User, ColetaPrevista } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import * as XLSX from 'xlsx';
import { getWeekString, getNonCollectionDateForCurrentTime } from '../utils/dateUtils';
import {
  Clock, X, Loader2, RefreshCw, ShieldCheck,
  CheckCircle2, ChevronDown,
  Filter, Search, CheckSquare, Square,
  ChevronRight, Maximize2, Minimize2,
  Archive, Database, Save, LinkIcon,
  Layers, Trash2, Settings2, Check, Table, SortAsc,
  Sun, Moon, AlertTriangle, Plus, Milk
} from 'lucide-react';

interface NonCollection {
  id: string;
  semana: string;
  rota: string;
  data: string;
  codigo: string;
  produtor: string;
  motivo: string;
  observacao: string;
  acao: string;
  dataAcao: string;
  ultimaColeta: string;
  Culpabilidade: string;
  operacao: string;
}

// Lista de MOTIVOS no padrão "MOTIVO - Culpabilidade"
const MOTIVOS_CulpabilidadeS: Record<string, string> = {
  'Rota Atrasada/Fábrica': 'Cliente',
  'Rota Atrasada/Logística': 'VIA',
  'Rota Atrasada/Manutenção': 'VIA',
  'Rota Atrasada/Mão De Obra': 'VIA',
  'Eqp. Cheio/Coleta Extra': 'VIA',
  'Eqp. Cheio/Eqp Menor': 'VIA',
  'Eqp. Cheio/Rota Atrasada': 'VIA',
  'Alizarol Positivo': 'Outros',
  'Leite Com Antibiótico': 'Outros',
  'Leite Congelado': 'Outros',
  'Leite Descartado': 'Outros',
  'Leite Quente': 'Outros',
  'Não Coletado - Saúde Do Motorista': 'VIA',
  'Não Coletado - Greve': 'Outros',
  'Não Coletado - Solicitado Pelo SDL': 'Cliente',
  'Objeto Ou Sujeira No Leite': 'Outros',
  'Parou De Fornecer': 'Outros',
  'Produtor Suspenso': 'Cliente',
  'Passou Para 48Hrs': 'VIA',
  'Problemas Mecânicos Equipamento': 'VIA',
  'Produtor Solicitou A Não Coleta': 'Outros',
  'Resfriador Vazio': 'Outros',
  'Volume Insuficiente Para Medida': 'Outros',
  'Coletado Por Outra Transportadora': 'VIA',
  'Descumprimento de roteirização': 'VIA',
  'A rota Não foi realizada': 'Outros',
  'Falta De Acesso': 'Outros',
  'Jornada Excedida': 'Outros',
  'Eqp. Cheio/Aumento de Vol.': 'Outros',
  'Correção de Roteirização': 'VIA',
  'Crioscopia': 'Outros',
  'Rota Atrasada/Infraestrutura': 'Outros',
  'Eqp. Cheio/Manutenção': 'VIA'
};

// Motivos que mostram popup de causa raiz
const MOTIVOS_COM_CAUSA_RAIZ = [
  'Eqp. Cheio/Coleta Extra',
  'Eqp. Cheio/Eqp Menor',
  'Eqp. Cheio/Rota Atrasada',
  'Coletado Por Outra Transportadora',
  'Descumprimento de roteirização'
];

const Culpabilidade_OPCOES = ['VIA', 'Cliente', 'Outros'];

const NonCollectionsView: React.FC<{
  currentUser: User;
}> = ({ currentUser }) => {
  const [nonCollections, setNonCollections] = useState<NonCollection[]>([]);
  const [coletasPrevistas, setColetasPrevistas] = useState<ColetaPrevista[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isDarkMode, setIsDarkMode] = useState(() => {
    const saved = localStorage.getItem('non_collections_dark_mode');
    return saved !== 'false';
  });

  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [isSortByDataEnabled, setIsSortByDataEnabled] = useState(true);
  const [colWidths, setColWidths] = useState<Record<string, number>>({
    semana: 120,
    rota: 100,
    data: 120,
    codigo: 90,
    produtor: 220,
    motivo: 180,
    observacao: 300,
    acao: 200,
    dataAcao: 120,
    ultimaColeta: 120,
    Culpabilidade: 130,
    operacao: 120
  });

  const [hiddenColumns, setHiddenColumns] = useState<Set<string>>(() => {
    const saved = localStorage.getItem('non_collections_hidden_cols');
    if (saved) {
      return new Set(JSON.parse(saved));
    }
    return new Set(['semana']);
  });

  const [contextMenu, setContextMenu] = useState<{ visible: boolean; x: number; y: number; col: string | null }>({ visible: false, x: 0, y: 0, col: null });
  const filterDropdownRef = useRef<HTMLDivElement>(null);
  const contextMenuRef = useRef<HTMLDivElement>(null);
  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

  // Estado para o popup de atendimento detalhado
  const [isAttendanceModalOpen, setIsAttendanceModalOpen] = useState(false);

  // Estados para modal de adicionar Não coleta
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [newNonCollectionData, setNewNonCollectionData] = useState<{
    rota: string;
    data: string;
    codigo: string;
    produtor: string;
    operacao: string;
  }>({
    rota: '',
    data: getNonCollectionDateForCurrentTime(),
    codigo: '',
    produtor: '',
    operacao: ''
  });

  // Ghost Row para adição rápida via paste
  const [ghostRow, setGhostRow] = useState<Partial<NonCollection>>({
    id: 'ghost',
    semana: '',
    rota: '',
    data: getNonCollectionDateForCurrentTime(),
    codigo: '',
    produtor: '',
    motivo: '',
    observacao: '',
    acao: '',
    dataAcao: '',
    ultimaColeta: '',
    Culpabilidade: '',
    operacao: ''
  });

  const [showCausaRaiz, setShowCausaRaiz] = useState(false);

  // Estados para modal de seleção de Operação (bulk paste)
  const [isOperationModalOpen, setIsOperationModalOpen] = useState(false);
  const [pendingBulkRoutes, setPendingBulkRoutes] = useState<string[]>([]);
  const [isCreatingRecords, setIsCreatingRecords] = useState(false); // Trava contra cliques duplos

  // Estados para arquivamento e histórico
  const [isHistoryModalOpen, setIsHistoryModalOpen] = useState(false);
  const [histStart, setHistStart] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 7);
    return d.toISOString().split('T')[0];
  });
  const [histEnd, setHistEnd] = useState(() => new Date().toISOString().split('T')[0]);
  const [archivedResults, setArchivedResults] = useState<NonCollection[]>([]);
  const [isSearchingArchive, setIsSearchingArchive] = useState(false);
  const [isArchiving, setIsArchiving] = useState(false);
  const archiveAbortRef = useRef<AbortController | null>(null);
  const [historyColFilters, setHistoryColFilters] = useState<Record<string, string>>({});
  const [historySelectedFilters, setHistorySelectedFilters] = useState<Record<string, string[]>>({});
  const [historyActiveFilterCol, setHistoryActiveFilterCol] = useState<string | null>(null);
  const historyFilterDropdownRef = useRef<HTMLDivElement>(null);
  const [pendingHistoryEdits, setPendingHistoryEdits] = useState<Record<string, Partial<NonCollection>>>({});
  const [editingHistoryId, setEditingHistoryId] = useState<string | null>(null);
  const [editingHistoryField, setEditingHistoryField] = useState<keyof NonCollection | null>(null);
  const [isSavingHistoryEdits, setIsSavingHistoryEdits] = useState(false);

  const updateGhostCell = (field: keyof NonCollection, value: string) => {
    const updatedGhost = { ...ghostRow, [field]: value };

    if (field === 'motivo') {
      // Preenche Culpabilidade automaticamente baseado no motivo
      const CulpabilidadeAuto = MOTIVOS_CulpabilidadeS[value];
      if (CulpabilidadeAuto) {
        updatedGhost.Culpabilidade = CulpabilidadeAuto;
      }
      // Verifica se é motivo com causa raiz
      setShowCausaRaiz(MOTIVOS_COM_CAUSA_RAIZ.includes(value));
    }

    setGhostRow(updatedGhost);
  };

  const handleAddFromGhost = async () => {
    if (!ghostRow.rota || !ghostRow.data || !ghostRow.produtor || !ghostRow.operacao) {
      alert('Preencha todos os campos obrigatórios na linha de criação!');
      return;
    }

    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        alert('Erro: Token Não encontrado');
        return;
      }

      const dataParaSemana = ghostRow.data!.split('/').reverse().join('-');
      const semana = getWeekString(dataParaSemana);

      // Calcula dataAcao automática: data + 2 dias se motivo gerar "Será coletado na rota..."
      const calcularDataAcao = () => {
        const motivo = (ghostRow.motivo || '').trim();
        const dataNaoColeta = ghostRow.data!;

        if (motivo && dataNaoColeta) {
          // Motivos que geram "Leite Descartado" ou "Aguardando autorização" ficam com hífen
          const motivoLower = motivo.toLowerCase();
          if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') {
            return '-';
          }

          // Se tem rota e data, calcula data + 2 dias
          if (ghostRow.rota) {
            const match = dataNaoColeta.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
            if (match) {
              const [, day, month, year] = match;
              const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
              dateObj.setDate(dateObj.getDate() + 2);
              const d = String(dateObj.getDate()).padStart(2, '0');
              const m = String(dateObj.getMonth() + 1).padStart(2, '0');
              const y = dateObj.getFullYear();
              return `${d}/${m}/${y}`;
            }
          }
        }

        return ''; // Sem motivo definido, fica vazio
      };

      const newRecord: NonCollection = {
        id: Date.now().toString(),
        semana,
        rota: ghostRow.rota!,
        data: ghostRow.data!,
        codigo: ghostRow.codigo || `P${String(nonCollections.length + 1).padStart(3, '0')}`,
        produtor: ghostRow.produtor!,
        motivo: ghostRow.motivo || '',
        observacao: ghostRow.observacao || '',
        acao: ghostRow.acao || '',
        dataAcao: calcularDataAcao(),
        ultimaColeta: '',
        Culpabilidade: ghostRow.Culpabilidade || 'Não se aplica',
        operacao: ghostRow.operacao!
      };

      // Salva no SharePoint e obtém o ID real
      const spId = await SharePointService.saveNonCollection(token, newRecord);

      // Adiciona localmente com o ID real do SharePoint
      setNonCollections(prev => [...prev, { ...newRecord, id: spId }]);

      // Busca coletas previstas para a data inserida
      await fetchColetasPrevistas(ghostRow.data!);

      // Limpa ghost row
      setGhostRow({
        id: 'ghost',
        semana: '',
        rota: '',
        data: getNonCollectionDateForCurrentTime(),
        codigo: '',
        produtor: '',
        motivo: '',
        observacao: '',
        acao: '',
        dataAcao: '',
        ultimaColeta: '',
        Culpabilidade: '',
        operacao: ''
      });
      setShowCausaRaiz(false);
    } catch (e: any) {
      console.error('[ADD_NON_COLLECTION] Error:', e);
      alert(`Erro ao adicionar Não coleta: ${e.message}`);
    }
  };

  /**
   * Cria novas linhas de Não coleta após o usuário selecionar a Operação no modal.
   * SALVA CADA REGISTRO NO SHAREPOINT IMEDIATAMENTE.
   */
  const createBulkRecordsWithOperation = async (operacao: string) => {
    // Trava contra cliques duplos
    if (isCreatingRecords) {
      console.warn('[BULK_PASTE] Ignorando clique duplicado - criação em andamento');
      return;
    }

    setIsCreatingRecords(true);
    const dataFormatada = getNonCollectionDateForCurrentTime();
    const dataParaSemana = dataFormatada.split('/').reverse().join('-');
    const semana = getWeekString(dataParaSemana);

    // Tenta salvar no SharePoint imediatamente
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        console.error('[BULK_PASTE] Token Não encontrado, criando apenas localmente');
        // Fallback: cria localmente (Não será salvo no SharePoint)
        const newRecords: NonCollection[] = pendingBulkRoutes.map((rota, i) => ({
          id: Date.now().toString() + i,
          semana,
          rota,
          data: dataFormatada,
          codigo: '',
          produtor: '',
          motivo: '',
          observacao: '',
          acao: '',
          dataAcao: '',
          ultimaColeta: '',
          Culpabilidade: 'Não se aplica',
          operacao
        }));
        setNonCollections(prev => [...prev, ...newRecords]);

        // Busca coletas previstas para a data inserida
        await fetchColetasPrevistas(dataFormatada);
        return;
      }

      console.log('[BULK_PASTE] Criando', pendingBulkRoutes.length, 'registros no SharePoint...');
      const savedRecords: NonCollection[] = [];

      // Salva cada registro no SharePoint sequencialmente
      for (let i = 0; i < pendingBulkRoutes.length; i++) {
        const rota = pendingBulkRoutes[i];
        const tempRecord: NonCollection = {
          id: 'temp', // ID temporário, será substituído pelo ID real do SharePoint
          semana,
          rota,
          data: dataFormatada,
          codigo: '',
          produtor: '',
          motivo: '',
          observacao: '',
          acao: '',
          dataAcao: '',
          ultimaColeta: '',
          Culpabilidade: 'Não se aplica',
          operacao
        };

        try {
          const spId = await SharePointService.saveNonCollection(token, tempRecord);
          console.log('[BULK_PASTE] ? Criado no SharePoint:', rota, 'ID:', spId);
          savedRecords.push({ ...tempRecord, id: spId }); // Usa o ID real do SharePoint
        } catch (e: any) {
          console.error('[BULK_PASTE] Erro ao criar registro no SharePoint:', rota, e.message);
          // Adiciona localmente com ID temporário mesmo com erro
          savedRecords.push({ ...tempRecord, id: (Date.now() + i).toString() });
        }
      }

      setNonCollections(prev => [...prev, ...savedRecords]);
      console.log('[BULK_PASTE] -', savedRecords.length, 'linhas criadas e salvas com Operação:', operacao);

      // Busca coletas previstas para a data inserida
      await fetchColetasPrevistas(dataFormatada);
    } catch (e: any) {
      console.error('[BULK_PASTE] Erro crítico ao salvar no SharePoint:', e.message);
      // Fallback: cria localmente
      const newRecords: NonCollection[] = pendingBulkRoutes.map((rota, i) => ({
        id: Date.now().toString() + i,
        semana,
        rota,
        data: dataFormatada,
        codigo: '',
        produtor: '',
        motivo: '',
        observacao: '',
        acao: '',
        dataAcao: '',
        ultimaColeta: '',
        Culpabilidade: 'Não se aplica',
        operacao
      }));
      setNonCollections(prev => [...prev, ...newRecords]);
    } finally {
      // Sempre libera a trava
      setIsCreatingRecords(false);
      // Limpa estados do modal
      setIsOperationModalOpen(false);
      setPendingBulkRoutes([]);
    }
  };

  /**
   * Cancela o modal de seleção de Operação sem criar linhas.
   */
  const cancelBulkPaste = () => {
    setIsCreatingRecords(false);
    setIsOperationModalOpen(false);
    setPendingBulkRoutes([]);
  };

  // Fecha dropdown ao clicar fora
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (filterDropdownRef.current && !filterDropdownRef.current.contains(e.target as Node)) {
        setActiveFilterCol(null);
      }
      if (historyFilterDropdownRef.current && !historyFilterDropdownRef.current.contains(e.target as Node)) {
        setHistoryActiveFilterCol(null);
      }
      if (contextMenuRef.current && !contextMenuRef.current.contains(e.target as Node)) {
        setContextMenu(prev => ({ ...prev, visible: false }));
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Atalho CTRL+SHIFT+L para limpar todos os filtros
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.ctrlKey && event.shiftKey && event.key.toLowerCase() === 'l') {
        event.preventDefault();
        clearFilters();
        setActiveFilterCol(null);
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, []);

  // Redimensionamento de colunas
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!resizingRef.current) return;
      const { col, startX, startWidth } = resizingRef.current;
      const diff = e.clientX - startX;
      const newWidth = Math.max(50, startWidth + diff);
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

  // Salvar preferências de colunas
  useEffect(() => {
    localStorage.setItem('non_collections_hidden_cols', JSON.stringify(Array.from(hiddenColumns)));
  }, [hiddenColumns]);


  // Carrega dados ao montar e quando usuário muda
  useEffect(() => {
    loadData();
  }, [currentUser]);

  /**
   * Busca coletas previstas para uma data específica e atualiza o estado.
   * @param dataDDMMYYYY Data no formato DD/MM/YYYY
   */
  const fetchColetasPrevistas = async (dataDDMMYYYY: string, operationsOverride: string[] = []): Promise<boolean> => {
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) return false;

      const dataISO = dataDDMMYYYY.split('/').reverse().join('-');
      const userOperations = operationsOverride.length > 0
        ? operationsOverride
        : userConfigs.map(c => c.operacao);

      console.log('[COLETAS_PREVISTAS] Buscando para data:', dataISO);
      console.log('[COLETAS_PREVISTAS] Email do usuário:', currentUser.email);
      console.log('[COLETAS_PREVISTAS] Operações do usuário (userConfigs):', userOperations);

      const coletas = await SharePointService.getColetasPrevistas(token, dataISO, currentUser.email, userOperations);
      setColetasPrevistas(coletas);
      console.log('[COLETAS_PREVISTAS] Total retornado:', coletas.length);
      console.log('[COLETAS_PREVISTAS] Detalhes:', coletas.map(c => `${c.Title}=${c.QntColeta}`));
      console.log('[COLETAS_PREVISTAS] Soma QntColeta:', coletas.reduce((sum, c) => sum + c.QntColeta, 0));
      return true;
    } catch (e: any) {
      console.error('[COLETAS_PREVISTAS] Erro ao buscar:', e.message);
      return false;
    }
  };

  const loadData = async (forceRefresh = false) => {
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        console.error('[NonCollections] Token Não encontrado');
        return;
      }

      console.log('[NonCollections] Carregando dados...', currentUser.email);

      // Carrega configurações do usuário
      const configs = await SharePointService.getRouteConfigs(token, currentUser.email, forceRefresh);
      setUserConfigs(configs || []);
      console.log('[NonCollections] operações do usuário:', configs?.map(c => c.operacao));

      // Carrega Não coletas do SharePoint
      const spNonCollections = await SharePointService.getNonCollections(token, currentUser.email);
      console.log('[NonCollections] Total bruto do SharePoint:', spNonCollections.length);

      // Filtra APENAS Não coletas das operações do usuário logado
      const myOps = new Set((configs || []).map(c => c.operacao));
      const filtered = (spNonCollections || []).filter(nc => {
        if (myOps.size === 0) return true; // Fallback se config Não carregou
        return myOps.has(nc.operacao);
      });

      console.log('[NonCollections] Não coletas filtradas por usuário:', filtered.length);
      console.log('[NonCollections] operações nos dados filtrados:', Array.from(new Set(filtered.map(r => r.operacao))));

      setNonCollections(filtered);

      // Busca coletas previstas da data das não coletas
      if (filtered.length > 0) {
        // Pega a data da primeira não coleta (todas são da mesma data)
        const dataNC = filtered[0].data;
        await fetchColetasPrevistas(dataNC, (configs || []).map(c => c.operacao));
      } else {
        setColetasPrevistas([]);
      }

      console.log('[NonCollections] ? Dados carregados com sucesso');
    } catch (e: any) {
      console.error('[NonCollections] Erro ao carregar dados:', e.message);
    } finally {
      setIsLoading(false);
    }
  };

  const persistNonCollectionRow = async (rowId: string, fieldLabel: string) => {
    if (!rowId || rowId === 'ghost' || rowId.startsWith('temp')) return;

    const currentRow = nonCollections.find(r => r.id === rowId);
    if (!currentRow) return;

    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) return;

      await SharePointService.updateNonCollection(token, currentRow);
      console.log(`[NC_SAVE] ${fieldLabel} salvo:`, currentRow.rota);
    } catch (e: any) {
      console.error(`[NC_SAVE] Erro ao salvar ${fieldLabel.toLowerCase()}:`, e?.message || e);
    }
  };

  // Dark mode effect
  useEffect(() => {
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    localStorage.setItem('non_collections_dark_mode', String(isDarkMode));
  }, [isDarkMode]);

  const getUniqueValues = (field: keyof NonCollection) => {
    return Array.from(new Set(nonCollections.map(r => String(r[field] || '')))).filter(v => v).sort();
  };

  const toggleFilterValue = (field: keyof NonCollection, value: string) => {
    const current = selectedFilters[field] || [];
    const updated = current.includes(value)
      ? current.filter(v => v !== value)
      : [...current, value];
    setSelectedFilters({ ...selectedFilters, [field]: updated });
  };

  const clearFilters = () => {
    setColFilters({});
    setSelectedFilters({});
  };

  const hasActiveFilters = Object.keys(selectedFilters).some(col => (selectedFilters[col] || []).length > 0);

  const filteredData = useMemo(() => {
    let result = [...nonCollections];

    // Aplica filtros de texto
    result = result.filter(r => {
      return (Object.entries(colFilters) as [string, string][]).every(([col, val]) => {
        if (!val) return true;
        const field = col as keyof NonCollection;
        return String(r[field] || '').toLowerCase().includes(val.toLowerCase());
      });
    });

    // Aplica filtros selecionados
    result = result.filter(r => {
      return (Object.entries(selectedFilters) as [string, string[]][]).every(([col, vals]) => {
        if (!vals || vals.length === 0) return true;
        const field = col as keyof NonCollection;
        return vals.includes(String(r[field] || ''));
      });
    });

    // Ordenação por data
    if (isSortByDataEnabled) {
      result.sort((a, b) => {
        const parseDate = (dateStr: string) => {
          if (!dateStr) return 0;
          const match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
          if (match) {
            const [, day, month, year] = match;
            return new Date(Number(year), Number(month) - 1, Number(day)).getTime();
          }
          return 0;
        };
        return parseDate(a.data) - parseDate(b.data);
      });
    }

    return result;
  }, [nonCollections, colFilters, selectedFilters, isSortByDataEnabled]);

  const getHistoryUniqueValues = (field: keyof NonCollection) => {
    return Array.from(new Set(archivedResults.map(r => String(r[field] || '').trim())))
      .filter(v => v)
      .sort();
  };

  const toggleHistoryFilterValue = (field: keyof NonCollection, value: string) => {
    const current = historySelectedFilters[field] || [];
    const updated = current.includes(value)
      ? current.filter(v => v !== value)
      : [...current, value];
    setHistorySelectedFilters({ ...historySelectedFilters, [field]: updated });
  };

  const clearHistoryFilters = () => {
    setHistoryColFilters({});
    setHistorySelectedFilters({});
    setHistoryActiveFilterCol(null);
  };

  const hasHistoryActiveFilters = useMemo(() => {
    const hasSelected = Object.keys(historySelectedFilters).some(col => (historySelectedFilters[col] || []).length > 0);
    const hasTyped = Object.keys(historyColFilters).some(col => !!historyColFilters[col]);
    return hasSelected || hasTyped;
  }, [historySelectedFilters, historyColFilters]);

  const filteredArchivedResults = useMemo(() => {
    let result = [...archivedResults];

    // Filtros de texto
    result = result.filter(r => {
      return (Object.entries(historyColFilters) as [string, string][]).every(([col, val]) => {
        if (!val) return true;
        const field = col as keyof NonCollection;
        return String(r[field] || '').toLowerCase().includes(val.toLowerCase());
      });
    });

    // Filtros selecionados
    result = result.filter(r => {
      return (Object.entries(historySelectedFilters) as [string, string[]][]).every(([col, vals]) => {
        if (!vals || vals.length === 0) return true;
        const field = col as keyof NonCollection;
        return vals.includes(String(r[field] || ''));
      });
    });

    return result;
  }, [archivedResults, historyColFilters, historySelectedFilters]);

  // Dados detalhados de atendimento por operação
  const attendanceDetails = useMemo(() => {
    const myOps = userConfigs.map(c => c.operacao);

    return myOps.map(op => {
      // Coletas previstas para esta operação
      const prevOp = coletasPrevistas.find(c => c.Title === op);
      const previstas = prevOp ? prevOp.QntColeta : 0;

      // Não coletas internas (apenas desta operação)
      const ncInternas = nonCollections.filter(nc => nc.operacao === op);
      const ncCount = ncInternas.length;

      // Não coletas VIA (para atendimento interno)
      const ncVia = ncInternas.filter(nc => nc.Culpabilidade === 'VIA').length;

      // Atendimento interno: (previstas - ncVIA) / previstas
      const pctInterno = previstas > 0 ? ((previstas - ncVia) / previstas * 100) : 100;

      // Atendimento geral: (previstas - todas NCs da op) / previstas
      const pctGeral = previstas > 0 ? ((previstas - ncCount) / previstas * 100) : 100;

      return {
        operacao: op,
        nomeExibicao: userConfigs.find(c => c.operacao === op)?.nomeExibicao || op,
        previstas,
        ncCount,
        ncVia,
        pctInterno,
        pctGeral
      };
    }).sort((a, b) => b.pctInterno - a.pctInterno); // Ordena por atendimento interno (maior -> menor)
  }, [userConfigs, coletasPrevistas, nonCollections]);

  const handleAddNonCollection = async () => {
    if (!newNonCollectionData.rota || !newNonCollectionData.data || !newNonCollectionData.produtor || !newNonCollectionData.operacao) {
      alert('Preencha todos os campos obrigatórios!');
      return;
    }

    try {
      const dataParaSemana = newNonCollectionData.data.split('/').reverse().join('-');
      const semana = getWeekString(dataParaSemana);

      const newRecord: NonCollection = {
        id: Date.now().toString(),
        semana,
        rota: newNonCollectionData.rota,
        data: newNonCollectionData.data,
        codigo: newNonCollectionData.codigo || `P${String(nonCollections.length + 1).padStart(3, '0')}`,
        produtor: newNonCollectionData.produtor,
        motivo: '',
        observacao: '',
        acao: '',
        dataAcao: '',
        ultimaColeta: '',
        Culpabilidade: 'Não se aplica',
        operacao: newNonCollectionData.operacao
      };

      // Salva no SharePoint e obtém o ID real
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        alert('Erro: Token não encontrado');
        return;
      }

      const spId = await SharePointService.saveNonCollection(token, newRecord);

      // Adiciona localmente com o ID real do SharePoint
      setNonCollections(prev => [...prev, { ...newRecord, id: spId }]);

      setIsAddModalOpen(false);
      setNewNonCollectionData({
        rota: '',
        data: getNonCollectionDateForCurrentTime(),
        codigo: '',
        produtor: '',
        operacao: ''
      });

      // Busca coletas previstas para a data inserida
      await fetchColetasPrevistas(newRecord.data);
    } catch (e: any) {
      console.error('[ADD_NON_COLLECTION] Error:', e);
      alert(`Erro ao adicionar Não coleta: ${e.message}`);
    }
  };

  // ============================================================
  // ARQUIVAR NÃO COLETAS
  // ============================================================

  const getAccessToken = async (): Promise<string> => {
    return await getValidToken() || currentUser.accessToken || '';
  };

  const handleArchiveAll = async () => {
    // VALIDAÇÃO CRÍTICA: Filtra apenas não coletas que pertencem às operações do usuário logado
    const myOps = new Set(userConfigs.map(c => c.operacao));
    const validNonCollections = nonCollections.filter(nc => !nc.operacao || myOps.has(nc.operacao));

    if (validNonCollections.length === 0) {
      alert("Não há não coletas das suas operações para arquivar.");
      return;
    }

    if (!confirm(`Arquivar ${validNonCollections.length} não coleta(s) no histórico?`)) return;

    const token = await getAccessToken();
    setIsArchiving(true);

    try {
      console.log(`[NC_ARCHIVE] Movendo ${validNonCollections.length} itens para o histórico...`);
      const archiveResult = await SharePointService.moveNonCollectionsToHistory(token, validNonCollections);
      console.log(`[NC_ARCHIVE] Sucesso: ${archiveResult.success}, Falhas: ${archiveResult.failed}`);

      // Limpa o campo UltimoEnvioNcoletas de cada operação do usuário
      console.log('[NC_ARCHIVE] Limpando UltimoEnvioNcoletas das operações do usuário...');
      const operacoes = userConfigs.map(c => c.operacao);
      for (const operacao of operacoes) {
        try {
          await SharePointService.updateUltimoEnvioNaoColetas(token, operacao, '');
          console.log(`[NC_ARCHIVE] ?o. UltimoEnvioNcoletas limpo para ${operacao}`);
        } catch (e: any) {
          console.error(`[NC_ARCHIVE] Erro ao limpar UltimoEnvioNcoletas para ${operacao}:`, e.message);
        }
      }

      // Recarregar dados com force refresh para pegar configs atualizadas
      await loadData(true);

      alert(`${archiveResult.success} não coleta(s) arquivada(s) com sucesso!`);
    } catch (e: any) {
      console.error('[NC_ARCHIVE] Erro geral:', e.message);
      alert(`Erro ao arquivar: ${e.message}`);
    } finally {
      setIsArchiving(false);
    }
  };

  // ============================================================
  // BUSCAR NO HISTéRICO
  // ============================================================

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
      console.log('[NC_SEARCH_ARCHIVE] Requesting history from SharePoint list nao_coletas_web_hist...');
      const results = await SharePointService.getArchivedNonCollections(await getAccessToken(), currentUser.email, histStart, histEnd, controller.signal);
      console.log('[NC_SEARCH_ARCHIVE] Results received:', results.length);

      // Só atualiza state se esta requisição não foi cancelada
      if (!controller.signal.aborted) {
        const myOps = new Set(userConfigs.map(c => c.operacao));
        const filtered = results && results.length > 0
          ? results.filter(r => !myOps.size || myOps.has(r.operacao))
          : [];

        setArchivedResults(filtered);
        setPendingHistoryEdits({});
        setEditingHistoryId(null);
        setEditingHistoryField(null);
      }
    } catch (err: any) {
      if (err.name === 'AbortError') {
        console.log('[NC_SEARCH_ARCHIVE] Requisição cancelada.');
        return;
      }
      console.error('[NC_SEARCH_ARCHIVE] Error during search:', err);
      alert("Erro na busca: " + (err?.message || "Erro desconhecido ao acessar o SharePoint."));
    } finally {
      if (!controller.signal.aborted) {
        setIsSearchingArchive(false);
      }
    }
  };

  // Auto-busca quando o modal de histórico abre
  useEffect(() => {
    if (isHistoryModalOpen && userConfigs.length > 0) {
      // NÃO faz auto-busca - usuário precisa filtrar manualmente
      setArchivedResults([]);
      setPendingHistoryEdits({});
      setEditingHistoryId(null);
      setEditingHistoryField(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isHistoryModalOpen]);

  // Estado para tela cheia no histórico
  const [isHistoryFullscreen, setIsHistoryFullscreen] = useState(false);

  // Função auxiliar para formatar data YYYY-MM-DD -> DD/MM/YYYY
  const formatDisplayDate = (dateStr: string | undefined): string => {
    if (!dateStr || dateStr.trim() === '') return '-';
    // Jé está em formato DD/MM/YYYY
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) return dateStr;
    // Formato YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      const [y, m, d] = dateStr.split('-');
      return `${d}/${m}/${y}`;
    }
    // Formato ISO com tempo
    if (dateStr.includes('T')) {
      const datePart = dateStr.split('T')[0];
      const [y, m, d] = datePart.split('-');
      return `${d}/${m}/${y}`;
    }
    return dateStr;
  };

  // Converte data DD/MM/AAAA ou AAAA-MM-DD para ISO (AAAA-MM-DD)
  const normalizeDateToISO = (dateStr: string | undefined): string => {
    if (!dateStr || dateStr.trim() === '' || dateStr.trim() === '-') return '';
    const raw = dateStr.trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
      return raw;
    }

    if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
      const [d, m, y] = raw.split('/');
      return `${y}-${m}-${d}`;
    }

    if (raw.includes('T')) {
      return raw.split('T')[0];
    }

    return '';
  };

  const applyDateMask = (value: string): string => {
    if (value === '-') return value;
    let digits = value.replace(/\D/g, '').slice(0, 8);
    if (digits.length > 4) return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
    if (digits.length > 2) return `${digits.slice(0, 2)}/${digits.slice(2)}`;
    return digits;
  };

  const handleUpdateHistoryCell = (id: string, field: keyof NonCollection, value: string) => {
    setPendingHistoryEdits(prev => {
      const current = prev[id] || {};
      const next: Partial<NonCollection> = { ...current, [field]: value };

      // Se alterar a data, recalcula semana automaticamente
      if (field === 'data') {
        const isoDate = normalizeDateToISO(value);
        if (isoDate) {
          next.semana = getWeekString(isoDate);
        }
      }

      return {
        ...prev,
        [id]: next
      };
    });
  };

  const savePendingHistoryEdits = async () => {
    const editIds = Object.keys(pendingHistoryEdits);
    if (editIds.length === 0) return;

    setIsSavingHistoryEdits(true);
    let successCount = 0;
    let errorCount = 0;

    try {
      const token = await getAccessToken();
      if (!token) {
        alert('Token não encontrado para salvar o histórico.');
        return;
      }

      for (const id of editIds) {
        const edits = pendingHistoryEdits[id];
        const current = archivedResults.find(r => r.id === id);
        if (!current) continue;

        const rowToSave: NonCollection = { ...current, ...edits };

        // Garante semana coerente com a data informada
        const dataISO = normalizeDateToISO(rowToSave.data);
        if (dataISO) {
          rowToSave.semana = getWeekString(dataISO);
        }

        try {
          await SharePointService.updateArchivedNonCollection(token, rowToSave);
          successCount++;
        } catch (e) {
          errorCount++;
        }
      }

      // Atualiza UI local após persistência
      setArchivedResults(prev => prev.map(r => {
        const edits = pendingHistoryEdits[r.id];
        if (!edits) return r;
        const updated = { ...r, ...edits };
        const dataISO = normalizeDateToISO(updated.data);
        if (dataISO) {
          updated.semana = getWeekString(dataISO);
        }
        return updated;
      }));

      setPendingHistoryEdits({});
      setEditingHistoryId(null);
      setEditingHistoryField(null);

      if (errorCount > 0) {
        alert(`Salvas ${successCount} edição(ões), com ${errorCount} erro(s).`);
      }
    } catch (e: any) {
      alert(`Erro ao salvar alterações do histórico: ${e?.message || 'erro desconhecido'}`);
    } finally {
      setIsSavingHistoryEdits(false);
    }
  };

  const closeHistoryModal = () => {
    setIsHistoryModalOpen(false);
    setIsHistoryFullscreen(false);
    setArchivedResults([]);
    setHistoryColFilters({});
    setHistorySelectedFilters({});
    setHistoryActiveFilterCol(null);
    setPendingHistoryEdits({});
    setEditingHistoryId(null);
    setEditingHistoryField(null);
  };

  // Enter salva todas as alterações pendentes no histórico
  useEffect(() => {
    if (!isHistoryModalOpen) return;

    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key !== 'Enter' || e.shiftKey || e.ctrlKey || e.altKey) return;
      if (Object.keys(pendingHistoryEdits).length === 0 || isSavingHistoryEdits) return;

      e.preventDefault();
      savePendingHistoryEdits();
    };

    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [isHistoryModalOpen, pendingHistoryEdits, isSavingHistoryEdits]);

  /**
   * Tenta separar Código e produtor de uma string colada.
   * Padrão esperado: Código alfanumérico no início seguido de texto (ex: "202520769MARCELO DINIZ COUTO" ou "P0274001RODRIGO ALVES PEREIRA")
   * Retorna null se Não conseguir identificar o padrão.
   */
  const parseCodigoProdutor = (line: string): { codigo: string; produtor: string } | null => {
    const trimmed = line.trim();
    
    // Regex: captura caracteres alfanuméricos no início (Código) seguidos de texto (produtor)
    // O Código pode ter letras e números (ex: "P0274001", "202520769")
    // O produtor é todo o restante da string
    // Ex: "P0274001RODRIGO ALVES PEREIRA (IN )" é.??T codigo: "P0274001", produtor: "RODRIGO ALVES PEREIRA (IN )"
    const match = trimmed.match(/^([A-Za-z0-9]+)(.+)$/);
    
    if (match) {
      return {
        codigo: match[1].toUpperCase(),
        produtor: match[2].trim()
      };
    }
    
    return null;
  };

  const handleBulkPaste = async (field: keyof NonCollection, value: string) => {
    const lines = value.split(/[\n\r]/).map(l => l.trim()).filter(Boolean);
    if (lines.length === 0) return;

    console.log('[BULK_PASTE] Campo:', field, 'Valores:', lines);
    console.log('[BULK_PASTE] nonCollections.length:', nonCollections.length);

    // Colunas que SEMPRE criam novas linhas (mesmo se já houver dados)
    const colunasQueCriamLinhas: (keyof NonCollection)[] = ['rota'];
    const criaNovasLinhas = colunasQueCriamLinhas.includes(field);

    console.log('[BULK_PASTE] criaNovasLinhas:', criaNovasLinhas);

    if (criaNovasLinhas) {
      // COMPORTAMENTO 1: Criar novas linhas (apenas para ROTA)
      // Abre modal para selecionar Operação antes de criar as linhas
      setPendingBulkRoutes(lines);
      setIsOperationModalOpen(true);
      setIsCreatingRecords(false); // Garante que a trava esteja liberada ao abrir
    } else {
      // COMPORTAMENTO 2: Atualizar linhas existentes ou vazias
      // Para outras colunas (CÓDIGO, PRODUTOR, MOTIVO, etc.)
      console.log('[BULK_PASTE] Preenchendo linhas existentes e/ou vazias...');

      // Colunas que devem ser tratadas individualmente (sem tentar separar Código/produtor)
      const colunasIndividuais: (keyof NonCollection)[] = ['codigo', 'produtor'];
      const isColunaIndividual = colunasIndividuais.includes(field);

      // ============================================================
      // IDENTIFICA LINHAS VAZIAS vs PREENCHIDAS
      // ============================================================
      // Uma linha é considerada "vazia" se NÃO tem rota OU NÃO tem produtor
      // Isso evita sobrescrever dados que o usuário já preencheu
      const emptyIndices: number[] = [];
      const filledIndices: number[] = [];

      nonCollections.forEach((nc, idx) => {
        const isEmpty = !nc.rota?.trim() || !nc.produtor?.trim();
        if (isEmpty) {
          emptyIndices.push(idx);
        } else {
          filledIndices.push(idx);
        }
      });

      console.log('[BULK_PASTE] Linhas vazias detectadas:', emptyIndices.length);
      console.log('[BULK_PASTE] Linhas preenchidas detectadas:', filledIndices.length);

      // ============================================================
      // DISTRIBUI VALORES: PRIORIZA LINHAS VAZIAS
      // ============================================================
      const updatedRecords: NonCollection[] = [...nonCollections]; // Cópia completa
      let lineIndex = 0; // índice do valor being processed

      // Passo 1: Preenche primeiro nas linhas vazias
      for (let i = 0; i < emptyIndices.length && lineIndex < lines.length; i++) {
        const recordIndex = emptyIndices[i];
        const record = updatedRecords[recordIndex];
        let finalValue = lines[lineIndex];
        let codigo = '';
        let produtor = '';

        // Tenta separar Código e produtor SOMENTE se NÃO for coluna individual
        if (!isColunaIndividual) {
          const parsed = parseCodigoProdutor(lines[lineIndex]);
          if (parsed) {
            codigo = parsed.codigo;
            produtor = parsed.produtor;
            console.log(`[BULK_PASTE] -o. Linha vazia ${recordIndex}: Código/produtor separados:`, parsed);
          }
        }

        // Formata se for data
        if (field === 'data' || field === 'dataAcao' || field === 'ultimaColeta') {
          if (finalValue.includes('-')) {
            const [year, month, day] = finalValue.split('-');
            finalValue = `${day}/${month}/${year}`;
          }
        }

        // CÓDIGO: converte para maiúsculo
        if (field === 'codigo') {
          finalValue = finalValue.toUpperCase();
        }

        // Aplica o valor na linha vazia
        updatedRecords[recordIndex] = {
          ...record,
          ...(codigo && produtor ? { codigo, produtor } : { [field]: finalValue })
        };
        console.log(`[BULK_PASTE] Preenchida linha vazia ${recordIndex} (campo: ${field})`);
        lineIndex++;
      }

      // Passo 2: Se ainda há valores e todas linhas vazias foram usadas, cria novas linhas
      const remainingLines = lines.slice(lineIndex);
      if (remainingLines.length > 0) {
        console.log('[BULK_PASTE] Criando', remainingLines.length, 'novas linhas para valores restantes...');

        const dataFormatada = getNonCollectionDateForCurrentTime();
        const dataParaSemana = dataFormatada.split('/').reverse().join('-');
        const semana = getWeekString(dataParaSemana);

        for (let i = 0; i < remainingLines.length; i++) {
          let finalValue = remainingLines[i];
          let codigo = '';
          let produtor = '';

          // Tenta separar Código e produtor
          if (!isColunaIndividual) {
            const parsed = parseCodigoProdutor(remainingLines[i]);
            if (parsed) {
              codigo = parsed.codigo;
              produtor = parsed.produtor;
            }
          }

          // Formata se for data
          if (field === 'data' || field === 'dataAcao' || field === 'ultimaColeta') {
            if (finalValue.includes('-')) {
              const [year, month, day] = finalValue.split('-');
              finalValue = `${day}/${month}/${year}`;
            }
          }

          // CÓDIGO: converte para maiúsculo
          if (field === 'codigo') {
            finalValue = finalValue.toUpperCase();
          }

          const newRecord: NonCollection = {
            id: (Date.now() + i).toString(), // ID temporário
            semana,
            rota: '', // Será preenchido depois pelo usuário
            data: dataFormatada,
            codigo: codigo || (field === 'codigo' ? finalValue : ''),
            produtor: produtor || (field === 'produtor' ? finalValue : ''),
            motivo: field === 'motivo' ? finalValue : '',
            observacao: field === 'observacao' ? finalValue : '',
            acao: field === 'acao' ? finalValue : '',
            dataAcao: field === 'dataAcao' ? finalValue : '',
            ultimaColeta: field === 'ultimaColeta' ? finalValue : '',
            Culpabilidade: field === 'Culpabilidade' ? finalValue : 'Não se aplica',
            operacao: '' // Será definido pelo modal de operação
          };

          updatedRecords.push(newRecord);
          console.log(`[BULK_PASTE] Criada nova linha ${updatedRecords.length - 1} para valor restante`);
        }
      }

      // ============================================================
      // ATUALIZA ESTADO E SALVA NO SHAREPOINT
      // ============================================================
      setNonCollections(updatedRecords);
      console.log('[BULK_PASTE] -o.', lineIndex, 'valores distribuídos em linhas existentes/vazias,', remainingLines.length, 'novas linhas criadas');

      // Salva no SharePoint APENAS registros que já existem no SharePoint
      try {
        const token = await getValidToken() || currentUser.accessToken;
        if (!token) {
          console.error('[BULK_PASTE] Token não encontrado, pulando salvamento no SharePoint');
          return;
        }

        // Filtra apenas registros que já existem no SharePoint (IDs numéricos válidos)
        const recordsToSave = updatedRecords.filter(r => {
          const id = parseInt(r.id);
          // SharePoint IDs são inteiros pequenos (< 1 milhão), locais são timestamps grandes
          return !isNaN(id) && id < 1000000;
        });

        if (recordsToSave.length === 0) {
          console.log('[BULK_PASTE] Nenhum registro para salvar no SharePoint (todos são locais)');
          return;
        }

        console.log('[BULK_PASTE] Salvando', recordsToSave.length, 'registros no SharePoint...');
        const savePromises = recordsToSave.map(async (record) => {
          try {
            await SharePointService.updateNonCollection(token, record);
            console.log('[BULK_PASTE] -o. Salvo:', record.rota, '-', record.codigo);
          } catch (e: any) {
            console.error('[BULK_PASTE] Erro ao salvar registro:', record.rota, e.message);
          }
        });
        await Promise.all(savePromises);
        console.log('[BULK_PASTE] -o. Todos os registros salvos no SharePoint');
      } catch (e: any) {
        console.error('[BULK_PASTE] Erro ao salvar no SharePoint:', e.message);
      }

      // Busca coletas previstas se criou novas linhas
      if (remainingLines.length > 0 && nonCollections.length > 0) {
        // Pega a data do último registro adicionado (todos tém a mesma data)
        const lastRecord = updatedRecords[updatedRecords.length - 1];
        if (lastRecord.data) {
          await fetchColetasPrevistas(lastRecord.data);
        }
      }
    }
  };

  const handleColumnResize = (col: string, startX: number, startWidth: number) => {
    resizingRef.current = { col, startX, startWidth };
  };

  const handleContextMenu = (e: React.MouseEvent, col: string) => {
    e.preventDefault();
    setContextMenu({ visible: true, x: e.clientX, y: e.clientY, col });
  };

  const toggleColumn = (col: string) => {
    setHiddenColumns(prev => {
      const next = new Set(prev);
      if (next.has(col)) {
        next.delete(col);
      } else {
        next.add(col);
      }
      return next;
    });
  };

  const formatDateToBR = (dateString: string) => {
    if (!dateString) return '--/--/----';
    if (dateString.includes('/')) return dateString;
    try {
      const [year, month, day] = dateString.split('-');
      return `${day}/${month}/${year}`;
    } catch {
      return dateString;
    }
  };

  const columns: { key: keyof NonCollection; label: string }[] = [
    { key: 'semana', label: 'SEMANA' },
    { key: 'rota', label: 'ROTA' },
    { key: 'data', label: 'DATA' },
    { key: 'codigo', label: 'CÓDIGO' },
    { key: 'produtor', label: 'PRODUTOR' },
    { key: 'motivo', label: 'MOTIVO' },
    { key: 'observacao', label: 'OBSERVAÇÃO' },
    { key: 'acao', label: 'AÇÃO' },
    { key: 'dataAcao', label: 'DATA AÇÃO' },
    { key: 'ultimaColeta', label: 'ÚLTIMA COLETA' },
    { key: 'Culpabilidade', label: 'CULPABILIDADE' },
    { key: 'operacao', label: 'OPERAÇÃO' }
  ];

  const historyColumns: { key: keyof NonCollection; label: string }[] = [
    { key: 'semana', label: 'Semana' },
    { key: 'rota', label: 'Rota' },
    { key: 'data', label: 'Data' },
    { key: 'codigo', label: 'Código' },
    { key: 'produtor', label: 'Produtor' },
    { key: 'motivo', label: 'Motivo' },
    { key: 'acao', label: 'Ação' },
    { key: 'dataAcao', label: 'Data Ação' },
    { key: 'ultimaColeta', label: 'Última Coleta' },
    { key: 'Culpabilidade', label: 'Culpabilidade' },
    { key: 'operacao', label: 'Operação' }
  ];

  if (isLoading) {
    return (
      <div className="h-full flex items-center justify-center">
        <div className="flex flex-col items-center gap-4">
          <Loader2 size={40} className="animate-spin text-blue-600" />
          <p className="font-bold uppercase text-sm tracking-widest text-slate-500">Carregando...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col bg-gradient-to-br from-white via-slate-50 to-slate-100 dark:from-slate-950 dark:via-slate-900 dark:to-slate-950 overflow-hidden">
      {/* Header */}
      <div className="flex items-center justify-between p-6 pb-4">
        <div className="flex items-center gap-6">
          <div>
            <h1 className="text-2xl font-black uppercase text-slate-800 dark:text-white tracking-tight">
              Não coletas
            </h1>
            <p className="text-xs font-bold text-slate-500 dark:text-slate-400 mt-1 uppercase tracking-widest">
              Acompanhamento de ocorrências
            </p>
          </div>

          {/* Cards de Indicadores */}
          <div className="flex items-center gap-3 ml-8">
            {/* Card Total */}
            <div className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[140px] ${
              isDarkMode ? 'bg-blue-900/30 border border-blue-700/50' : 'bg-blue-100 border border-blue-300'
            }`}>
              <div className="text-center flex-1">
                <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>Total</p>
                <p className={`text-2xl font-black leading-none ${isDarkMode ? 'text-blue-400' : 'text-blue-700'}`}>{nonCollections.length}</p>
              </div>
              <div className="w-2 h-2 bg-blue-500 rounded-full animate-pulse shrink-0"></div>
            </div>

            {/* Card Coletas Previstas */}
            <div
              className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[160px] ${
                isDarkMode ? 'bg-emerald-900/30 border border-emerald-700/50' : 'bg-emerald-100 border border-emerald-300'
              }`}
            >
              <div className="text-center flex-1">
                <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? 'text-emerald-400' : 'text-emerald-700'}`}>Coletas Previstas</p>
                <p className={`text-2xl font-black leading-none ${isDarkMode ? 'text-emerald-400' : 'text-emerald-700'}`}>{coletasPrevistas.reduce((sum, c) => sum + c.QntColeta, 0)}</p>
              </div>
              <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse shrink-0"></div>
            </div>

            {/* Card Atendimento Interno */}
            {(() => {
              const totalPrevistas = coletasPrevistas.reduce((sum, c) => sum + c.QntColeta, 0);
              const ncVia = nonCollections.filter(nc => nc.Culpabilidade === 'VIA').length;
              const pct = totalPrevistas > 0 ? ((totalPrevistas - ncVia) / totalPrevistas * 100) : 0;
              const cor = pct >= 90 ? 'text-amber-400' : pct >= 70 ? 'text-orange-400' : 'text-red-400';
              const corBorder = pct >= 90 ? 'border-amber-700/50' : pct >= 70 ? 'border-orange-700/50' : 'border-red-700/50';
              const corBg = pct >= 90 ? 'bg-amber-900/30' : pct >= 70 ? 'bg-orange-900/30' : 'bg-red-900/30';
              const corBgLight = pct >= 90 ? 'bg-amber-100' : pct >= 70 ? 'bg-orange-100' : 'bg-red-100';
              const corBorderLight = pct >= 90 ? 'border-amber-300' : pct >= 70 ? 'border-orange-300' : 'border-red-300';
              const corText = pct >= 90 ? 'text-amber-400' : pct >= 70 ? 'text-orange-400' : 'text-red-400';
              const corTextLight = pct >= 90 ? 'text-amber-700' : pct >= 70 ? 'text-orange-700' : 'text-red-700';
              const corDot = pct >= 90 ? 'bg-amber-500' : pct >= 70 ? 'bg-orange-500' : 'bg-red-500';

              return (
                <div
                  onClick={() => setIsAttendanceModalOpen(true)}
                  className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[160px] cursor-pointer hover:scale-105 transition-all ${
                    isDarkMode ? `${corBg} border ${corBorder}` : `${corBgLight} border ${corBorderLight}`
                  }`}>
                  <div className="text-center flex-1">
                    <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? corText : corTextLight}`}>Atend. Interno</p>
                    <p className={`text-2xl font-black leading-none ${isDarkMode ? corText : corTextLight}`}>{pct.toFixed(1)}%</p>
                  </div>
                  <div className={`w-2 h-2 ${corDot} rounded-full animate-pulse shrink-0`}></div>
                </div>
              );
            })()}

            {/* Card Atendimento Geral */}
            {(() => {
              const totalPrevistas = coletasPrevistas.reduce((sum, c) => sum + c.QntColeta, 0);
              const todasNc = nonCollections.length;
              const pct = totalPrevistas > 0 ? ((totalPrevistas - todasNc) / totalPrevistas * 100) : 0;
              const cor = pct >= 90 ? 'text-amber-400' : pct >= 70 ? 'text-orange-400' : 'text-red-400';
              const corBorder = pct >= 90 ? 'border-amber-700/50' : pct >= 70 ? 'border-orange-700/50' : 'border-red-700/50';
              const corBg = pct >= 90 ? 'bg-amber-900/30' : pct >= 70 ? 'bg-orange-900/30' : 'bg-red-900/30';
              const corBgLight = pct >= 90 ? 'bg-amber-100' : pct >= 70 ? 'bg-orange-100' : 'bg-red-100';
              const corBorderLight = pct >= 90 ? 'border-amber-300' : pct >= 70 ? 'border-orange-300' : 'border-red-300';
              const corText = pct >= 90 ? 'text-amber-400' : pct >= 70 ? 'text-orange-400' : 'text-red-400';
              const corTextLight = pct >= 90 ? 'text-amber-700' : pct >= 70 ? 'text-orange-700' : 'text-red-700';
              const corDot = pct >= 90 ? 'bg-amber-500' : pct >= 70 ? 'bg-orange-500' : 'bg-red-500';

              return (
                <div
                  onClick={() => setIsAttendanceModalOpen(true)}
                  className={`flex items-center gap-3 px-6 py-3 rounded-2xl min-w-[160px] cursor-pointer hover:scale-105 transition-all ${
                    isDarkMode ? `${corBg} border ${corBorder}` : `${corBgLight} border ${corBorderLight}`
                  }`}>
                  <div className="text-center flex-1">
                    <p className={`text-[9px] font-black uppercase tracking-wider mb-1 ${isDarkMode ? corText : corTextLight}`}>Atend. Geral</p>
                    <p className={`text-2xl font-black leading-none ${isDarkMode ? corText : corTextLight}`}>{pct.toFixed(1)}%</p>
                  </div>
                  <div className={`w-2 h-2 ${corDot} rounded-full animate-pulse shrink-0`}></div>
                </div>
              );
            })()}
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={() => setIsAddModalOpen(true)}
            className="flex items-center gap-2 px-4 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-black uppercase text-[10px] tracking-widest transition-all shadow-lg shadow-blue-500/20"
          >
            <Plus size={16} />
            Adicionar Não Coleta
          </button>

          <button
            onClick={handleArchiveAll}
            disabled={isArchiving || nonCollections.length === 0}
            className={`flex items-center gap-2 px-4 py-2.5 rounded-xl font-black uppercase text-[10px] tracking-widest transition-all shadow-lg disabled:opacity-50 disabled:cursor-not-allowed ${
              isDarkMode
                ? 'bg-slate-800 text-slate-300 hover:bg-slate-700 border border-slate-700'
                : 'bg-white text-slate-800 hover:bg-slate-50 hover:border-slate-500 border border-slate-400'
            }`}
            title="Arquivar todas as não coletas no histórico"
          >
            {isArchiving ? (
              <><Loader2 size={16} className="animate-spin" /> Arquivando...</>
            ) : (
              <><Archive size={16} /> Arquivar</>
            )}
          </button>

          <button
            onClick={() => setIsHistoryModalOpen(true)}
            className={`flex items-center gap-2 px-4 py-2.5 rounded-xl font-black uppercase text-[10px] tracking-widest transition-all shadow-lg ${
              isDarkMode
                ? 'bg-slate-800 text-slate-300 hover:bg-slate-700 border border-slate-700'
                : 'bg-white text-slate-800 hover:bg-slate-50 hover:border-slate-500 border border-slate-400'
            }`}
            title="Buscar não coletas no histórico"
          >
            <Database size={16} /> Histórico
          </button>

          <button
            onClick={loadData}
            className="p-3 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all"
            title="Recarregar"
          >
            <RefreshCw size={20} className="text-slate-600 dark:text-slate-400" />
          </button>

          <button
            onClick={() => setIsDarkMode(!isDarkMode)}
            className="p-3 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-sm hover:shadow-md transition-all"
          >
            {isDarkMode ? (
              <Sun size={20} className="text-amber-400" />
            ) : (
              <Moon size={20} className="text-slate-600" />
            )}
          </button>

          <button
            onClick={() => setIsSortByDataEnabled(!isSortByDataEnabled)}
            className={`p-3 rounded-xl border shadow-sm hover:shadow-md transition-all ${
              isSortByDataEnabled
                ? 'bg-blue-600 border-blue-700 text-white'
                : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-400'
            }`}
            title="Ordenar por Data"
          >
            <SortAsc size={20} />
          </button>

          {hasActiveFilters && (
            <button
              onClick={clearFilters}
              className="px-4 py-3 rounded-xl bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-800 text-red-600 dark:text-red-400 font-black uppercase text-[10px] tracking-widest hover:bg-red-100 dark:hover:bg-red-900/50 transition-all"
            >
              Limpar Filtros
            </button>
          )}
        </div>
      </div>

      {/* Table Container */}
      <div className="flex-1 mx-6 mt-4 mb-6 bg-white/95 dark:bg-slate-900/95 backdrop-blur-sm rounded-[2rem] shadow-2xl border border-white/50 dark:border-slate-800 overflow-hidden flex flex-col">
        <div className="overflow-x-auto flex-1 scrollbar-thin" id="table-container">
          <table className="w-full border-collapse">
            <thead className="sticky top-0 z-20">
              <tr className="bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900">
                {columns.map(({ key, label }) => (
                  <th
                    key={key}
                    className={`relative px-4 py-4 text-left border-r border-slate-700/50 last:border-r-0 select-none ${
                      hiddenColumns.has(key) ? 'hidden' : ''
                    }`}
                    style={{ width: colWidths[key] }}
                    onContextMenu={(e) => handleContextMenu(e, key)}
                  >
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-black text-slate-300 uppercase tracking-widest whitespace-nowrap">
                        {label}
                      </span>

                      <button
                        onClick={() => setActiveFilterCol(activeFilterCol === key ? null : key)}
                        className={`p-1 rounded transition-colors ${
                          (selectedFilters[key] || []).length > 0
                            ? 'bg-blue-600 text-white'
                            : 'hover:bg-slate-700 text-slate-400'
                        }`}
                      >
                        <Filter size={12} />
                      </button>

                      <div
                        className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize hover:bg-blue-500/50 transition-colors"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          handleColumnResize(key, e.clientX, colWidths[key]);
                        }}
                      />
                    </div>

                    {activeFilterCol === key && (
                      <div
                        ref={filterDropdownRef}
                        className="absolute top-full left-0 mt-2 z-50 w-56 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-2xl p-3 animate-in fade-in zoom-in-95 duration-150"
                      >
                        <div className="flex items-center gap-2 mb-3 p-2 bg-slate-50 dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700">
                          <Search size={14} className="text-slate-400" />
                          <input
                            type="text"
                            placeholder="Filtrar..."
                            autoFocus
                            value={colFilters[key] || ''}
                            onChange={(e) => setColFilters({ ...colFilters, [key]: e.target.value })}
                            className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800 dark:text-white"
                          />
                        </div>

                        <div className="max-h-48 overflow-y-auto space-y-1 scrollbar-thin border-t border-slate-100 dark:border-slate-700 py-2">
                          {getUniqueValues(key).map((value) => (
                            <div
                              key={value}
                              onClick={() => toggleFilterValue(key, value)}
                              className="flex items-center gap-2 p-2 hover:bg-slate-50 dark:hover:bg-slate-700 rounded-lg cursor-pointer transition-all"
                            >
                              {(selectedFilters[key] || []).includes(value) ? (
                                <CheckSquare size={14} className="text-blue-600" />
                              ) : (
                                <Square size={14} className="text-slate-300" />
                              )}
                              <span className="text-[10px] font-bold uppercase truncate text-slate-700 dark:text-slate-300">
                                {value}
                              </span>
                            </div>
                          ))}
                        </div>

                        <button
                          onClick={() => {
                            setColFilters({ ...colFilters, [key]: '' });
                            setSelectedFilters({ ...selectedFilters, [key]: [] });
                            setActiveFilterCol(null);
                          }}
                          className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-600 bg-red-50 dark:bg-red-900/30 hover:bg-red-100 dark:hover:bg-red-900/50 rounded-lg border border-red-100 dark:border-red-900/50 transition-colors"
                        >
                          Limpar Filtro
                        </button>
                      </div>
                    )}
                  </th>
                ))}
                <th className="relative p-0 border border-slate-700/50 text-[10px] font-black uppercase text-center bg-slate-900/50" style={{ width: 60 }}>
                  <Settings2 size={14} className="mx-auto opacity-40" />
                </th>
              </tr>
            </thead>

            <tbody>
              {filteredData.map((row, index) => (
                <tr
                  key={row.id}
                  className={`border-b border-slate-200/50 dark:border-slate-800/50 transition-colors ${
                    index % 2 === 0 ? '' : 'bg-black/[0.02] dark:bg-white/[0.02]'
                  }`}
                >
                  {columns.map(({ key }) => {
                    const inputClass = `w-full bg-transparent outline-none border-none px-3 py-2 text-[11px] transition-all ${
                      isDarkMode ? 'text-slate-200' : 'text-slate-900'
                    }`;

                    // ROTA - Input editável
                    if (key === 'rota') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.rota}
                            onChange={(e) => {
                              const updated = { ...row, rota: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('rota', val);
                              }
                            }}
                            className={`${inputClass} font-black text-center`}
                          />
                        </td>
                      );
                    }

                    // DATA - Input editável com máscara
                    if (key === 'data') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.data}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              const updated = { ...row, data: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('data', val);
                              }
                            }}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // PRODUTOR - Input editável
                    if (key === 'produtor') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.produtor}
                            onChange={(e) => {
                              const updated = { ...row, produtor: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              console.log('[PASTE PRODUTOR]', val);
                              
                              // Múltiplas linhas: bulk paste
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('produtor', val);
                                return;
                              }
                              
                              // Linha única com Código + produtor: separa automaticamente
                              const parsed = parseCodigoProdutor(val);
                              if (parsed) {
                                e.preventDefault();
                                const updated = { ...row, codigo: parsed.codigo, produtor: parsed.produtor };
                                setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                                console.log('[PASTE PRODUTOR] ? Código e produtor separados:', parsed);
                                return;
                              }
                              
                              // Se Não tiver padrão Código+produtor, deixa colar normalmente
                            }}
                            className={`${inputClass} font-bold`}
                          />
                        </td>
                      );
                    }

                    // CÓDIGO - Input editável
                    if (key === 'codigo') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.codigo}
                            onChange={(e) => {
                              const updated = { ...row, codigo: e.target.value.toUpperCase() };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              console.log('[PASTE CODIGO ROW]', val);
                              
                              // Múltiplas linhas: bulk paste
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('codigo', val);
                                return;
                              }
                              
                              // Linha única com Código + produtor: separa automaticamente
                              const parsed = parseCodigoProdutor(val);
                              if (parsed) {
                                e.preventDefault();
                                const updated = { ...row, codigo: parsed.codigo, produtor: parsed.produtor };
                                setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                                console.log('[PASTE CODIGO] ? Código e produtor separados:', parsed);
                                return;
                              }
                              
                              // Se Não tiver padrão Código+produtor, deixa colar normalmente (já converte para upper)
                            }}
                            className={`${inputClass} font-bold text-center uppercase`}
                          />
                        </td>
                      );
                    }

                    // MOTIVO - Select editável
                    if (key === 'motivo') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <select
                            value={row.motivo}
                            onChange={(e) => {
                              const selectedMotivo = e.target.value;
                              const CulpabilidadeAuto = MOTIVOS_CulpabilidadeS[selectedMotivo];

                              // Calcula Ação automática baseado no motivo
                              const calcularAcaoAutomatica = (motivo: string, rota: string): string => {
                                const motivoLower = motivo.toLowerCase();
                                if (motivoLower === 'parou de fornecer') return 'Retirado da roteirização';
                                if (motivoLower === 'produtor suspenso') return 'Aguardando autorização';
                                if (motivoLower === 'alizarol positivo') return 'Leite Descartado';
                                if (rota) return `Será coletado na rota ${rota}`;
                                return '';
                              };

                              // Calcula Data Ação automática
                              const calcularDataAcaoAutomatica = (motivo: string, data: string, rota: string): string => {
                                const motivoLower = motivo.toLowerCase();
                                if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') return '-';
                                if (motivo && data && rota) {
                                  const match = data.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                                  if (match) {
                                    const [, day, month, year] = match;
                                    const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
                                    dateObj.setDate(dateObj.getDate() + 2);
                                    const d = String(dateObj.getDate()).padStart(2, '0');
                                    const m = String(dateObj.getMonth() + 1).padStart(2, '0');
                                    const y = dateObj.getFullYear();
                                    return `${d}/${m}/${y}`;
                                  }
                                }
                                return '';
                              };

                              const acaoAuto = calcularAcaoAutomatica(selectedMotivo, row.rota || '');
                              const dataAcaoAuto = calcularDataAcaoAutomatica(selectedMotivo, row.data || '', row.rota || '');

                              const updated = {
                                ...row,
                                motivo: selectedMotivo,
                                Culpabilidade: CulpabilidadeAuto || row.Culpabilidade || '',
                                acao: acaoAuto,
                                dataAcao: dataAcaoAuto
                              };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onBlur={() => {
                              if (row.id !== 'ghost' && !row.id.startsWith('temp')) {
                                void persistNonCollectionRow(row.id, 'Motivo');
                              }
                            }}
                            className={`w-full px-3 py-2 text-[11px] text-left truncate transition-all cursor-pointer outline-none border-none bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-slate-200 hover:bg-slate-200 dark:hover:bg-slate-700 focus:ring-2 focus:ring-blue-500 rounded ${
                              isDarkMode ? 'dark-mode-select' : ''
                            }`}
                          >
                            <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                            {Object.keys(MOTIVOS_CulpabilidadeS).map(label => (
                              <option
                                key={label}
                                value={label}
                                className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white"
                              >
                                {label}
                              </option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // OBSERVAÇÃO - Textarea editável
                    if (key === 'observacao') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <textarea
                            value={row.observacao}
                            onChange={(e) => {
                              const updated = { ...row, observacao: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onBlur={() => {
                              if (row.id !== 'ghost' && !row.id.startsWith('temp')) {
                                void persistNonCollectionRow(row.id, 'Observação');
                              }
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('observacao', val);
                              }
                            }}
                            rows={1}
                            className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                          />
                        </td>
                      );
                    }

                    // AÇÃO - Campo automático baseado no MOTIVO (editável)
                    if (key === 'acao') {
                      // Calcula o valor automático da Ação baseado no motivo
                      const getAcaoAutomatica = () => {
                        const motivo = (row.motivo || '').trim();
                        const rota = (row.rota || '').trim();

                        // Se Não tem motivo definido, retorna vazio
                        if (!motivo) return '';

                        // Regras específicas
                        if (motivo.toLowerCase() === 'parou de fornecer') {
                          return 'Retirado da roteirização';
                        }
                        if (motivo.toLowerCase() === 'produtor suspenso') {
                          return 'Aguardando autorização';
                        }
                        if (motivo.toLowerCase() === 'alizarol positivo') {
                          return 'Leite Descartado';
                        }

                        // Se tem rota, gera "Será coletado na rota X"
                        if (rota) {
                          return `Será coletado na rota ${rota}`;
                        }

                        return '';
                      };

                      const acaoValue = row.acao || getAcaoAutomatica();

                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <textarea
                            value={acaoValue}
                            onChange={(e) => {
                              const updated = { ...row, acao: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onBlur={() => {
                              if (row.id !== 'ghost' && !row.id.startsWith('temp')) {
                                void persistNonCollectionRow(row.id, 'Ação');
                              }
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('acao', val);
                              }
                            }}
                            rows={1}
                            className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                          />
                        </td>
                      );
                    }

                    // DATA AÇÃO - Input editável com máscara (preenchida automaticamente apenas para "Será coletado na rota...")
                    if (key === 'dataAcao') {
                      // Calcula data Ação automática: data + 2 dias APENAS para motivos que geram "Será coletado na rota..."
                      const getDataAcaoAutomatica = () => {
                        const motivo = (row.motivo || '').trim();
                        const dataNaoColeta = (row.data || '').trim();

                        // Se o usuário já preencheu dataAcao manualmente, respeita
                        if (row.dataAcao) return row.dataAcao;

                        // Para "parou de fornecer", "produtor suspenso" e "alizarol positivo", coloca hífen
                        const motivoLower = motivo.toLowerCase();
                        if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') {
                          return '-';
                        }

                        // Se tem motivo e data, calcula data + 2 dias
                        if (motivo && dataNaoColeta && row.rota) {
                          const match = dataNaoColeta.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                          if (match) {
                            const [, day, month, year] = match;
                            const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
                            dateObj.setDate(dateObj.getDate() + 2);
                            const d = String(dateObj.getDate()).padStart(2, '0');
                            const m = String(dateObj.getMonth() + 1).padStart(2, '0');
                            const y = dateObj.getFullYear();
                            return `${d}/${m}/${y}`;
                          }
                        }

                        return '';
                      };

                      const dataAcaoValue = getDataAcaoAutomatica();

                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={dataAcaoValue}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              const updated = { ...row, dataAcao: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onBlur={() => {
                              if (row.id !== 'ghost' && !row.id.startsWith('temp')) {
                                void persistNonCollectionRow(row.id, 'Data Ação');
                              }
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('dataAcao', val);
                              }
                            }}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // ÚLTIMA COLETA - Input editável com máscara
                    if (key === 'ultimaColeta') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <input
                            type="text"
                            value={row.ultimaColeta}
                            onChange={(e) => {
                              let val = e.target.value;
                              // Permite hífen como valor único (igual Data Ação)
                              if (val === '-') {
                                // já está ok
                              } else {
                                val = val.replace(/\D/g, '');
                                if (val.length > 8) val = val.slice(0, 8);
                                if (val.length >= 8) {
                                  val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                                }
                              }
                              const updated = { ...row, ultimaColeta: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onBlur={() => {
                              if (row.id !== 'ghost' && !row.id.startsWith('temp')) {
                                void persistNonCollectionRow(row.id, 'Última Coleta');
                              }
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('ultimaColeta', val);
                              }
                            }}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // OPERAÇÃO - Select editável
                    if (key === 'operacao') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <select
                            value={row.operacao}
                            onChange={(e) => {
                              const updated = { ...row, operacao: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            className={`${inputClass} text-center font-bold cursor-pointer bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200`}
                          >
                            <option value="" className="bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200">Selecione...</option>
                            {userConfigs.map(op => (
                              <option key={op.operacao} value={op.operacao} className="bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200">{op.operacao}</option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // Culpabilidade - Dropdown editável com opções fixas
                    if (key === 'Culpabilidade') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <select
                            value={row.Culpabilidade || ''}
                            onChange={(e) => {
                              const updated = { ...row, Culpabilidade: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onBlur={() => {
                              if (row.id !== 'ghost' && !row.id.startsWith('temp')) {
                                void persistNonCollectionRow(row.id, 'Culpabilidade');
                              }
                            }}
                            className={`${inputClass} text-center cursor-pointer`}
                          >
                            <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                            {Culpabilidade_OPCOES.map(culp => (
                              <option key={culp} value={culp} className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">{culp}</option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // SEMANA - Somente leitura
                    if (key === 'semana') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <span className={`${inputClass} text-center block py-2`}>
                            {row.semana || '-'}
                          </span>
                        </td>
                      );
                    }

                    // Default - Texto
                    return (
                      <td
                        key={key}
                        className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                          hiddenColumns.has(key) ? 'hidden' : ''
                        }`}
                        style={{ minWidth: colWidths[key] }}
                      >
                        <span className={`${inputClass} px-3 py-2 block`}>
                          {row[key as keyof NonCollection]}
                        </span>
                      </td>
                    );
                  })}
                  {/* Coluna de Ação */}
                  <td className="p-0 border border-slate-200/30 dark:border-slate-800/30 text-center">
                    <button
                      onClick={async () => {
                        const rowId = row.id;
                        // Remove localmente imediatamente para feedback visual
                        setNonCollections(prev => prev.filter(r => r.id !== rowId));
                        // Tenta excluir do SharePoint (apenas registros já persistidos)
                        if (rowId !== 'ghost' && !rowId.startsWith('temp')) {
                          try {
                            const token = await getValidToken() || currentUser.accessToken;
                            if (token) {
                              await SharePointService.deleteNonCollection(token, rowId);
                              console.log('[NC_DELETE] Não coleta excluída do SharePoint:', row.rota);
                            }
                          } catch (e: any) {
                            console.error('[NC_DELETE] Erro ao excluir do SharePoint:', e.message);
                            alert('Erro ao excluir do servidor. A linha foi removida apenas localmente.');
                            // Re-adiciona a linha se falhou ao excluir do SharePoint
                            loadData();
                          }
                        }
                      }}
                      className="p-2 text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors"
                      title="Excluir linha"
                    >
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              ))}

              {/* Ghost Row - Adição rápida */}
              <tr
                key="ghost"
                className="border-b-2 border-blue-200 dark:border-blue-800 transition-colors bg-slate-50 dark:bg-slate-800 italic text-slate-400 border-l-4 border-dashed border-slate-300 dark:border-slate-600"
              >
                {columns.map(({ key }) => {
                  const inputClass = `w-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold uppercase transition-all ${
                    isDarkMode ? 'text-slate-200 placeholder-slate-500' : 'text-slate-900 placeholder-slate-400'
                  }`;

                  if (key === 'rota') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.rota || ''}
                          onChange={(e) => updateGhostCell('rota', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('rota', val);
                            }
                          }}
                          placeholder="Cole rotas..."
                          className={`${inputClass} font-black text-center`}
                        />
                      </td>
                    );
                  }

                  if (key === 'data') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.data || ''}
                          onChange={(e) => {
                            let val = e.target.value.replace(/\D/g, '');
                            if (val.length > 8) val = val.slice(0, 8);
                            if (val.length >= 8) {
                              val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                            }
                            updateGhostCell('data', val);
                          }}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('data', val);
                            }
                          }}
                          placeholder="DD/MM/AAAA"
                          maxLength={10}
                          className={`${inputClass} text-center font-mono`}
                        />
                      </td>
                    );
                  }

                  if (key === 'operacao') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <select
                          value={ghostRow.operacao || ''}
                          onChange={(e) => updateGhostCell('operacao', e.target.value)}
                          className={`${inputClass} text-center font-bold cursor-pointer`}
                        >
                          <option value="">Selecione...</option>
                          {userConfigs.map(op => (
                            <option key={op.operacao} value={op.operacao}>{op.operacao}</option>
                          ))}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'produtor') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.produtor || ''}
                          onChange={(e) => updateGhostCell('produtor', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('produtor', val);
                            }
                          }}
                          placeholder="Cole produtores..."
                          className={`${inputClass} font-bold`}
                        />
                      </td>
                    );
                  }

                  if (key === 'codigo') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow.codigo || ''}
                          onChange={(e) => updateGhostCell('codigo', e.target.value.toUpperCase())}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('codigo', val);
                            }
                          }}
                          placeholder="Cole Códigos..."
                          className={`${inputClass} font-bold text-center uppercase`}
                        />
                      </td>
                    );
                  }

                  if (key === 'motivo') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <select
                          value={ghostRow.motivo || ''}
                          onChange={(e) => {
                            const selectedMotivo = e.target.value;
                            const CulpabilidadeAuto = MOTIVOS_CulpabilidadeS[selectedMotivo];

                            // Calcula Ação automática
                            const motivoLower = selectedMotivo.toLowerCase();
                            let acaoAuto = '';
                            if (motivoLower === 'parou de fornecer') acaoAuto = 'Retirado da roteirização';
                            else if (motivoLower === 'produtor suspenso') acaoAuto = 'Aguardando autorização';
                            else if (motivoLower === 'alizarol positivo') acaoAuto = 'Leite Descartado';
                            else if (ghostRow.rota) acaoAuto = `Será coletado na rota ${ghostRow.rota}`;

                            // Calcula Data Ação automática
                            let dataAcaoAuto = '';
                            if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') {
                              dataAcaoAuto = '-';
                            } else if (selectedMotivo && ghostRow.data && ghostRow.rota) {
                              const match = ghostRow.data.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                              if (match) {
                                const [, day, month, year] = match;
                                const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
                                dateObj.setDate(dateObj.getDate() + 2);
                                const d = String(dateObj.getDate()).padStart(2, '0');
                                const m = String(dateObj.getMonth() + 1).padStart(2, '0');
                                const y = dateObj.getFullYear();
                                dataAcaoAuto = `${d}/${m}/${y}`;
                              }
                            }

                            updateGhostCell('motivo', selectedMotivo);
                            if (CulpabilidadeAuto) updateGhostCell('Culpabilidade', CulpabilidadeAuto);
                            if (acaoAuto) updateGhostCell('acao', acaoAuto);
                            if (dataAcaoAuto) updateGhostCell('dataAcao', dataAcaoAuto);
                          }}
                          className={`${inputClass} text-left cursor-pointer`}
                        >
                          <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                          {Object.keys(MOTIVOS_CulpabilidadeS).map(label => (
                            <option key={label} value={label} className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">{label}</option>
                          ))}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'Culpabilidade') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <select
                          value={ghostRow.Culpabilidade || ''}
                          onChange={(e) => updateGhostCell('Culpabilidade', e.target.value)}
                          className={`${inputClass} text-center cursor-pointer`}
                        >
                          <option value="" className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">Selecione...</option>
                          {Culpabilidade_OPCOES.map(culp => (
                            <option key={culp} value={culp} className="bg-white dark:bg-slate-900 text-slate-900 dark:text-white">{culp}</option>
                          ))}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'observacao') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <textarea
                          value={ghostRow.observacao || ''}
                          onChange={(e) => updateGhostCell('observacao', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('observacao', val);
                            }
                          }}
                          rows={1}
                          className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                        />
                      </td>
                    );
                  }

                  if (key === 'acao') {
                    // Calcula o valor automático da Ação baseado no motivo (ghost row)
                    const getAcaoAutomatica = () => {
                      const motivo = (ghostRow.motivo || '').trim();
                      const rota = (ghostRow.rota || '').trim();

                      // Se Não tem motivo definido, retorna vazio
                      if (!motivo) return '';

                      // Regras específicas
                      if (motivo.toLowerCase() === 'parou de fornecer') {
                        return 'Retirado da roteirização';
                      }
                      if (motivo.toLowerCase() === 'produtor suspenso') {
                        return 'Aguardando autorização';
                      }
                      if (motivo.toLowerCase() === 'alizarol positivo') {
                        return 'Leite Descartado';
                      }

                      // Se tem rota, gera "Será coletado na rota X"
                      if (rota) {
                        return `Será coletado na rota ${rota}`;
                      }

                      return '';
                    };

                    const acaoValue = ghostRow.acao || getAcaoAutomatica();

                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <textarea
                          value={acaoValue}
                          onChange={(e) => updateGhostCell('acao', e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste('acao', val);
                            }
                          }}
                          rows={1}
                          className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                        />
                      </td>
                    );
                  }

                  if (key === 'semana') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <span className={`${inputClass} text-center block py-2`}>
                          {ghostRow.data ? getWeekString(ghostRow.data.split('/').reverse().join('-')) : '-'}
                        </span>
                      </td>
                    );
                  }

                  if (key === 'dataAcao' || key === 'ultimaColeta') {
                    // Para dataAcao, calcula automaticamente data + 2 dias APENAS para "Será coletado na rota..."
                    if (key === 'dataAcao') {
                      const getDataAcaoAutomatica = () => {
                        // Se o usuário já preencheu manualmente, respeita
                        if (ghostRow.dataAcao) return ghostRow.dataAcao;

                        const motivo = (ghostRow.motivo || '').trim();
                        const dataNaoColeta = (ghostRow.data || '').trim();

                        // Motivos que geram "Leite Descartado" ou "Aguardando autorização" ficam com hífen
                        const motivoLower = motivo.toLowerCase();
                        if (motivoLower === 'parou de fornecer' || motivoLower === 'produtor suspenso' || motivoLower === 'alizarol positivo') {
                          return '-';
                        }

                        // Se tem motivo, data e rota, calcula data + 2 dias
                        if (motivo && dataNaoColeta && ghostRow.rota) {
                          const match = dataNaoColeta.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                          if (match) {
                            const [, day, month, year] = match;
                            const dateObj = new Date(Number(year), Number(month) - 1, Number(day));
                            dateObj.setDate(dateObj.getDate() + 2);
                            const d = String(dateObj.getDate()).padStart(2, '0');
                            const m = String(dateObj.getMonth() + 1).padStart(2, '0');
                            const y = dateObj.getFullYear();
                            return `${d}/${m}/${y}`;
                          }
                        }

                        return '';
                      };

                      const dataAcaoValue = getDataAcaoAutomatica();

                      return (
                        <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                          <input
                            type="text"
                            value={dataAcaoValue}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              updateGhostCell(key, val);
                            }}
                            maxLength={10}
                            className={`${inputClass} text-center font-mono`}
                          />
                        </td>
                      );
                    }

                    // ultimaColeta mantém comportamento normal (vazio para usuário preencher, ou hífen)
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow[key] || ''}
                          onChange={(e) => {
                            let val = e.target.value;
                            // Permite hífen como valor único (igual Data Ação)
                            if (val === '-') {
                              // já está ok
                            } else {
                              val = val.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                            }
                            updateGhostCell(key, val);
                          }}
                          placeholder="DD/MM/AAAA"
                          maxLength={10}
                          className={`${inputClass} text-center font-mono`}
                        />
                      </td>
                    );
                  }

                  return (
                    <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                      <input
                        type="text"
                        value={ghostRow[key] || ''}
                        onChange={(e) => updateGhostCell(key, e.target.value)}
                        placeholder=""
                        className={`${inputClass}`}
                      />
                    </td>
                  );
                })}
                {/* Coluna de Ação */}
                <td className="p-0 border border-slate-200/30 dark:border-slate-800/30 text-center">
                  <button
                    onClick={handleAddFromGhost}
                    className="p-2 text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded-lg transition-colors"
                    title="Adicionar"
                  >
                    <Plus size={18} />
                  </button>
                </td>
              </tr>

              {filteredData.length === 0 && nonCollections.length === 0 && (
                <tr>
                  <td colSpan={columns.length + 1} className="px-4 py-16 text-center">
                    <div className="flex flex-col items-center gap-3">
                      <div className="w-16 h-16 rounded-full bg-slate-100 dark:bg-slate-800 flex items-center justify-center">
                        <Milk size={32} className="text-slate-400" />
                      </div>
                      <p className="text-sm font-bold text-slate-500 dark:text-slate-400 uppercase tracking-widest">
                        Nenhuma Não coleta registrada
                      </p>
                      <p className="text-xs font-medium text-slate-400 dark:text-slate-500">
                        Clique em "Adicionar Não Coleta" ou cole dados do Excel
                      </p>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Context Menu */}
      {contextMenu.visible && (
        <div
          ref={contextMenuRef}
          className="fixed z-[1000] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-2xl py-2 min-w-[200px] animate-in fade-in zoom-in-95 duration-150"
          style={{ top: contextMenu.y, left: contextMenu.x }}
        >
          <div className="px-3 py-2 border-b border-slate-200 dark:border-slate-700">
            <p className="text-[10px] font-black uppercase text-slate-400">Colunas</p>
          </div>
          {columns.map(({ key, label }) => (
            <button
              key={key}
              onClick={() => toggleColumn(key)}
              className="w-full px-3 py-2 text-left hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center gap-2"
            >
              {hiddenColumns.has(key) ? (
                <Square size={14} className="text-slate-400" />
              ) : (
                <CheckSquare size={14} className="text-blue-600" />
              )}
              <span className="text-[10px] font-bold uppercase text-slate-700 dark:text-slate-300">{label}</span>
            </button>
          ))}
        </div>
      )}

      {/* Modal de Adicionar Não Coleta */}
      {isAddModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[200] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-lg">
            <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white">
              <div className="flex items-center gap-3">
                <Milk size={24} />
                <h3 className="font-black uppercase tracking-widest text-base">Adicionar Não Coleta</h3>
              </div>
              <button
                onClick={() => setIsAddModalOpen(false)}
                className="p-2 hover:bg-slate-700 rounded-lg transition-colors"
              >
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
                  value={newNonCollectionData.operacao}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, operacao: e.target.value })}
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                >
                  <option value="">Selecione a Operação</option>
                  {userConfigs.map(config => (
                    <option key={config.operacao} value={config.operacao}>
                      {config.nomeExibicao || config.operacao}
                    </option>
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
                  value={newNonCollectionData.rota}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, rota: e.target.value })}
                  placeholder="Ex: ROTA 01"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                />
              </div>

              {/* Data */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Data *
                </label>
                <input
                  type="text"
                  value={newNonCollectionData.data}
                  onChange={e => {
                    let val = e.target.value.replace(/\D/g, '');
                    if (val.length > 8) val = val.slice(0, 8);
                    if (val.length >= 8) {
                      val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                    }
                    setNewNonCollectionData({ ...newNonCollectionData, data: val });
                  }}
                  placeholder="DD/MM/AAAA"
                  maxLength={10}
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors font-mono text-center"
                />
              </div>

              {/* Código */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Código
                </label>
                <input
                  type="text"
                  value={newNonCollectionData.codigo}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, codigo: e.target.value.toUpperCase() })}
                  placeholder="Ex: 123456 ou P001"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors uppercase"
                />
              </div>

              {/* Produtor */}
              <div>
                <label className="block text-[10px] font-black uppercase tracking-widest text-slate-500 dark:text-slate-400 mb-2">
                  Produtor *
                </label>
                <input
                  type="text"
                  value={newNonCollectionData.produtor}
                  onChange={e => setNewNonCollectionData({ ...newNonCollectionData, produtor: e.target.value })}
                  placeholder="Nome do produtor"
                  className="w-full p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm font-bold outline-none dark:text-white focus:border-primary-500 transition-colors"
                />
              </div>
            </div>
            <div className="p-6 pt-0 flex gap-3">
              <button
                onClick={() => setIsAddModalOpen(false)}
                className="flex-1 py-3 bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 rounded-xl font-black uppercase text-[10px] tracking-widest hover:bg-slate-200 dark:hover:bg-slate-700 transition-all"
              >
                Cancelar
              </button>
              <button
                onClick={handleAddNonCollection}
                className="flex-1 py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-black uppercase text-[10px] tracking-widest transition-all shadow-lg shadow-blue-500/20"
              >
                Adicionar
              </button>
            </div>
          </div>
        </div>
      )}


      {/* Modal de seleção de Operação (Bulk Paste) */}
      {isOperationModalOpen && (
        <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[300] flex items-center justify-center p-4 animate-in fade-in zoom-in-95 duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-blue-500/50 shadow-2xl animate-in zoom-in duration-300">
            <div className="flex items-center gap-3 text-blue-600 dark:text-blue-400 mb-6 font-black uppercase text-xs">
              <Milk size={24} />
              selecionar Operação
            </div>
            <p className="text-sm text-slate-600 dark:text-slate-400 mb-2 font-medium">
              {pendingBulkRoutes.length} rota(s) colada(s). Selecione a Operação:
            </p>
            <div className="mb-6 max-h-32 overflow-y-auto scrollbar-thin bg-slate-50 dark:bg-slate-800 rounded-xl p-3 border border-slate-200 dark:border-slate-700">
              {pendingBulkRoutes.slice(0, 10).map((rota, i) => (
                <div key={i} className="text-[10px] font-bold text-slate-600 dark:text-slate-400 py-1 truncate">
                  . {rota}
                </div>
              ))}
              {pendingBulkRoutes.length > 10 && (
                <div className="text-[10px] font-bold text-slate-400 italic mt-1">
                  ...e mais {pendingBulkRoutes.length - 10} rota(s)
                </div>
              )}
            </div>

            {userConfigs.length === 0 ? (
              <div className="text-center py-6 text-slate-400 dark:text-slate-500">
                <AlertTriangle size={32} className="mx-auto mb-3 opacity-50" />
                <p className="text-sm font-bold">Nenhuma Operação disponível</p>
                <p className="text-[10px] mt-1">Contate o administrador para configurar suas operações.</p>
              </div>
            ) : (
              <div className="grid grid-cols-2 gap-3">
                {userConfigs.map(op => (
                  <button
                    key={op.operacao}
                    onClick={() => createBulkRecordsWithOperation(op.operacao)}
                    disabled={isCreatingRecords}
                    className={`p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-2xl transition-all font-black text-xs uppercase relative
                      ${isCreatingRecords
                        ? 'opacity-40 cursor-not-allowed bg-slate-100 dark:bg-slate-800 text-slate-400 dark:text-slate-600'
                        : 'hover:bg-blue-600 hover:text-white hover:border-blue-700 dark:hover:bg-blue-600 dark:hover:border-blue-500 text-slate-700 dark:text-slate-300'
                      }`}
                  >
                    {isCreatingRecords ? (
                      <div className="flex items-center justify-center gap-2">
                        <Loader2 size={14} className="animate-spin" />
                        Criando...
                      </div>
                    ) : (
                      op.operacao
                    )}
                  </button>
                ))}
              </div>
            )}

            <button
              onClick={cancelBulkPaste}
              className="w-full mt-6 py-4 text-[10px] font-black uppercase text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 transition-colors"
            >
              Cancelar
            </button>
          </div>
        </div>
      )}

      {/* Modal de Atendimento Detalhado */}
      {isAttendanceModalOpen && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-[200] flex items-center justify-center p-4 animate-in fade-in zoom-in-95 duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-5xl overflow-hidden border border-blue-500/50 animate-in zoom-in duration-300 flex flex-col" style={{ maxHeight: '90vh' }}>
            {/* Header */}
            <div className="bg-gradient-to-r from-blue-600 via-blue-700 to-blue-600 text-white p-6 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 bg-white/20 rounded-2xl flex items-center justify-center">
                  <CheckCircle2 size={28} />
                </div>
                <div>
                  <h3 className="text-2xl font-black uppercase tracking-tight">Atendimento</h3>
                  <p className="text-xs text-blue-200 font-bold uppercase tracking-widest">Detalhamento por Operação</p>
                </div>
              </div>
              <button onClick={() => setIsAttendanceModalOpen(false)} className="hover:bg-white/10 p-2 rounded-full transition-all">
                <X size={24} />
              </button>
            </div>

            {/* Tabela de Atendimento */}
            <div className="flex-1 overflow-auto p-6 scrollbar-thin">
              <table className="w-full border-collapse">
                <thead>
                  <tr className="bg-slate-100 dark:bg-slate-800">
                    <th className="px-4 py-3 text-left text-[10px] font-black uppercase text-slate-600 dark:text-slate-400 tracking-wider border-b-2 border-slate-200 dark:border-slate-700 sticky left-0 bg-slate-100 dark:bg-slate-800 z-10">
                      Operação
                    </th>
                    {/* Atendimento Geral - Grupo */}
                    <th colSpan={3} className="px-4 py-2 text-center text-[10px] font-black uppercase text-blue-600 dark:text-blue-400 tracking-wider border-b-2 border-blue-300 dark:border-blue-800 bg-blue-50 dark:bg-blue-900/20">
                      Atendimento Geral
                    </th>
                    {/* Atendimento Interno - Grupo */}
                    <th colSpan={3} className="px-4 py-2 text-center text-[10px] font-black uppercase text-emerald-600 dark:text-emerald-400 tracking-wider border-b-2 border-emerald-300 dark:border-emerald-800 bg-emerald-50 dark:bg-emerald-900/20">
                      Atendimento Interno
                    </th>
                  </tr>
                  <tr className="bg-slate-50 dark:bg-slate-800/50">
                    <th className="px-4 py-2 text-left text-[9px] font-bold uppercase text-slate-500 dark:text-slate-400 border-b border-slate-200 dark:border-slate-700 sticky left-0 bg-slate-50 dark:bg-slate-800/50 z-10">
                    </th>
                    {/* Sub-colunas Geral */}
                    <th className="px-3 py-2 text-center text-[9px] font-bold uppercase text-slate-600 dark:text-slate-400 border-b border-slate-200 dark:border-slate-700">Coletas Prev.</th>
                    <th className="px-3 py-2 text-center text-[9px] font-bold uppercase text-slate-600 dark:text-slate-400 border-b border-slate-200 dark:border-slate-700">Não Coletas</th>
                    <th className="px-3 py-2 text-center text-[9px] font-bold uppercase text-blue-600 dark:text-blue-400 border-b border-slate-200 dark:border-slate-700">Atend. %</th>
                    {/* Sub-colunas Interno */}
                    <th className="px-3 py-2 text-center text-[9px] font-bold uppercase text-slate-600 dark:text-slate-400 border-b border-slate-200 dark:border-slate-700">Coletas Prev.</th>
                    <th className="px-3 py-2 text-center text-[9px] font-bold uppercase text-slate-600 dark:text-slate-400 border-b border-slate-200 dark:border-slate-700">NCs VIA</th>
                    <th className="px-3 py-2 text-center text-[9px] font-bold uppercase text-emerald-600 dark:text-emerald-400 border-b border-slate-200 dark:border-slate-700">Atend. %</th>
                  </tr>
                </thead>
                <tbody>
                  {attendanceDetails.map((item, idx) => {
                    const corGeral = item.pctGeral >= 90 ? 'text-green-600 dark:text-green-400' : item.pctGeral >= 70 ? 'text-amber-600 dark:text-amber-400' : 'text-red-600 dark:text-red-400';
                    const corInterno = item.pctInterno >= 90 ? 'text-green-600 dark:text-green-400' : item.pctInterno >= 70 ? 'text-amber-600 dark:text-amber-400' : 'text-red-600 dark:text-red-400';
                    const bgGeral = item.pctGeral >= 90 ? 'bg-green-50 dark:bg-green-900/20' : item.pctGeral >= 70 ? 'bg-amber-50 dark:bg-amber-900/20' : 'bg-red-50 dark:bg-red-900/20';
                    const bgInterno = item.pctInterno >= 90 ? 'bg-green-50 dark:bg-green-900/20' : item.pctInterno >= 70 ? 'bg-amber-50 dark:bg-amber-900/20' : 'bg-red-50 dark:bg-red-900/20';

                    return (
                      <tr key={item.operacao} className={`${idx % 2 === 0 ? 'bg-white dark:bg-slate-900' : 'bg-slate-50 dark:bg-slate-800/50'} hover:bg-blue-50 dark:hover:bg-slate-700/50 transition-colors`}>
                        <td className="px-4 py-3 font-black text-[11px] uppercase text-slate-800 dark:text-white border-b border-slate-200 dark:border-slate-700 sticky left-0 bg-inherit z-10">
                          {item.nomeExibicao}
                        </td>
                        {/* Colunas Geral */}
                        <td className="px-3 py-3 text-center text-[11px] font-bold text-slate-700 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700">{item.previstas}</td>
                        <td className="px-3 py-3 text-center text-[11px] font-bold text-slate-700 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700">{item.ncCount}</td>
                        <td className={`px-3 py-3 text-center text-[12px] font-black border-b border-slate-200 dark:border-slate-700 ${corGeral} ${bgGeral}`}>
                          {item.pctGeral.toFixed(1)}%
                        </td>
                        {/* Colunas Interno */}
                        <td className="px-3 py-3 text-center text-[11px] font-bold text-slate-700 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700">{item.previstas}</td>
                        <td className="px-3 py-3 text-center text-[11px] font-bold text-slate-700 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700">{item.ncVia}</td>
                        <td className={`px-3 py-3 text-center text-[12px] font-black border-b border-slate-200 dark:border-slate-700 ${corInterno} ${bgInterno}`}>
                          {item.pctInterno.toFixed(1)}%
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                {/* Rodapé com totais */}
                <tfoot>
                  <tr className="bg-slate-100 dark:bg-slate-800 font-black">
                    <td className="px-4 py-3 text-[11px] uppercase text-slate-800 dark:text-white border-t-2 border-slate-300 dark:border-slate-600 sticky left-0 bg-slate-100 dark:bg-slate-800 z-10">
                      TOTAL
                    </td>
                    <td className="px-3 py-3 text-center text-[11px] text-slate-800 dark:text-white border-t-2 border-slate-300 dark:border-slate-600">
                      {coletasPrevistas.reduce((sum, c) => sum + c.QntColeta, 0)}
                    </td>
                    <td className="px-3 py-3 text-center text-[11px] text-slate-800 dark:text-white border-t-2 border-slate-300 dark:border-slate-600">
                      {nonCollections.length}
                    </td>
                    <td className={`px-3 py-3 text-center text-[12px] border-t-2 border-slate-300 dark:border-slate-600 ${
                      (() => {
                        const tp = coletasPrevistas.reduce((s, c) => s + c.QntColeta, 0);
                        const tn = nonCollections.length;
                        const pct = tp > 0 ? ((tp - tn) / tp * 100) : 100;
                        return pct >= 90 ? 'text-green-600 dark:text-green-400' : pct >= 70 ? 'text-amber-600 dark:text-amber-400' : 'text-red-600 dark:text-red-400';
                      })()
                    }`}>
                      {(() => {
                        const tp = coletasPrevistas.reduce((s, c) => s + c.QntColeta, 0);
                        const tn = nonCollections.length;
                        return tp > 0 ? ((tp - tn) / tp * 100).toFixed(1) + '%' : '100.0%';
                      })()}
                    </td>
                    <td className="px-3 py-3 text-center text-[11px] text-slate-800 dark:text-white border-t-2 border-slate-300 dark:border-slate-600">
                      {coletasPrevistas.reduce((sum, c) => sum + c.QntColeta, 0)}
                    </td>
                    <td className="px-3 py-3 text-center text-[11px] text-slate-800 dark:text-white border-t-2 border-slate-300 dark:border-slate-600">
                      {nonCollections.filter(nc => nc.Culpabilidade === 'VIA').length}
                    </td>
                    <td className={`px-3 py-3 text-center text-[12px] border-t-2 border-slate-300 dark:border-slate-600 ${
                      (() => {
                        const tp = coletasPrevistas.reduce((s, c) => s + c.QntColeta, 0);
                        const nv = nonCollections.filter(nc => nc.Culpabilidade === 'VIA').length;
                        const pct = tp > 0 ? ((tp - nv) / tp * 100) : 100;
                        return pct >= 90 ? 'text-green-600 dark:text-green-400' : pct >= 70 ? 'text-amber-600 dark:text-amber-400' : 'text-red-600 dark:text-red-400';
                      })()
                    }`}>
                      {(() => {
                        const tp = coletasPrevistas.reduce((s, c) => s + c.QntColeta, 0);
                        const nv = nonCollections.filter(nc => nc.Culpabilidade === 'VIA').length;
                        return tp > 0 ? ((tp - nv) / tp * 100).toFixed(1) + '%' : '100.0%';
                      })()}
                    </td>
                  </tr>
                </tfoot>
              </table>

              {attendanceDetails.length === 0 && (
                <div className="py-16 text-center text-slate-400 dark:text-slate-500">
                  <AlertTriangle size={40} className="mx-auto mb-4 opacity-50" />
                  <p className="text-sm font-bold">Nenhuma operação disponível</p>
                </div>
              )}
            </div>

            {/* Footer */}
            <div className="p-4 bg-slate-50 dark:bg-slate-800/50 border-t border-slate-200 dark:border-slate-700 flex justify-center shrink-0">
              <button
                onClick={() => setIsAttendanceModalOpen(false)}
                className="px-8 py-3 bg-slate-200 dark:bg-slate-700 text-slate-600 dark:text-slate-300 rounded-xl font-black uppercase text-[10px] hover:bg-slate-300 dark:hover:bg-slate-600 transition-all tracking-widest border border-slate-300 dark:border-slate-600 shadow-sm"
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Histórico de Não Coletas */}
      {isHistoryModalOpen && (
        <div className={`fixed inset-0 bg-black/80 backdrop-blur-md z-[110] flex items-center justify-center animate-in zoom-in duration-300 ${isHistoryFullscreen ? 'p-0' : 'p-4'}`}>
          <div className={`rounded-[2.5rem] shadow-2xl w-full overflow-hidden border flex flex-col ${
            isHistoryFullscreen ? 'max-w-none w-full h-full rounded-none' : 'max-w-7xl max-h-[90vh]'
          } ${isDarkMode ? 'bg-slate-900 border-slate-700' : 'bg-white border-slate-200'}`}>
            {/* Header */}
            <div className={`p-6 flex justify-between items-center shrink-0 ${
              isDarkMode ? 'bg-slate-800' : 'bg-slate-100'
            }`}>
              <div className="flex items-center gap-4">
                <Database size={32} className={isDarkMode ? 'text-blue-400' : 'text-blue-600'} />
                <div>
                  <h3 className={`text-xl font-black uppercase tracking-tight ${isDarkMode ? 'text-white' : 'text-slate-800'}`}>Histórico de Não Coletas</h3>
                  <p className={`text-xs font-bold uppercase tracking-widest ${isDarkMode ? 'text-slate-400' : 'text-slate-500'}`}>Busca na lista nao_coletas_web_hist</p>
                </div>
                {archivedResults.length > 0 && (
                  <span className={`text-[10px] font-bold px-3 py-1 rounded-full ${
                    isDarkMode ? 'text-slate-400 bg-slate-700' : 'text-slate-500 bg-slate-200'
                  }`}>
                    {archivedResults.length} registro(s)
                  </span>
                )}
                {Object.keys(pendingHistoryEdits).length > 0 && (
                  <span className="text-[10px] font-black uppercase tracking-widest text-amber-500 bg-amber-100 dark:bg-amber-900/30 px-3 py-1 rounded-full border border-amber-300 dark:border-amber-700">
                    {Object.keys(pendingHistoryEdits).length} alteração(ões) pendente(s)
                  </span>
                )}
              </div>
              <div className="flex items-center gap-2">
                {Object.keys(pendingHistoryEdits).length > 0 && (
                  <button
                    onClick={savePendingHistoryEdits}
                    disabled={isSavingHistoryEdits}
                    className="flex items-center gap-2 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-black uppercase text-[10px] tracking-widest transition-all disabled:opacity-60"
                    title="Salvar alterações (Enter)"
                  >
                    {isSavingHistoryEdits ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                    Salvar
                  </button>
                )}
                <button
                  onClick={() => setIsHistoryFullscreen(!isHistoryFullscreen)}
                  className={`p-2 rounded-lg transition-all ${isDarkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-200'}`}
                  title={isHistoryFullscreen ? 'Sair da tela cheia' : 'Tela cheia'}
                >
                  {isHistoryFullscreen ? <Minimize2 size={20} className={isDarkMode ? 'text-slate-400' : 'text-slate-600'} /> : <Maximize2 size={20} className={isDarkMode ? 'text-slate-400' : 'text-slate-600'} />}
                </button>
                <button
                  onClick={closeHistoryModal}
                  className={`p-2 rounded-full transition-all ${isDarkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-200'}`}
                >
                  <X size={24} className={isDarkMode ? 'text-slate-400' : 'text-slate-600'} />
                </button>
              </div>
            </div>

            {/* Filtros */}
            <div className={`p-4 border-b shrink-0 ${isDarkMode ? 'bg-slate-900 border-slate-700' : 'bg-white border-slate-200'}`}>
              <div className="flex items-center gap-4 flex-wrap">
                <div className="flex items-center gap-2">
                  <label className={`text-[10px] font-black uppercase ${isDarkMode ? 'text-slate-400' : 'text-slate-500'}`}>Início:</label>
                  <input
                    type="date"
                    value={histStart}
                    onChange={(e) => setHistStart(e.target.value)}
                    className={`p-2 border rounded-xl text-sm font-bold outline-none focus:ring-2 focus:ring-blue-500 ${
                      isDarkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-300 text-slate-800'
                    }`}
                  />
                </div>
                <div className="flex items-center gap-2">
                  <label className={`text-[10px] font-black uppercase ${isDarkMode ? 'text-slate-400' : 'text-slate-500'}`}>Fim:</label>
                  <input
                    type="date"
                    value={histEnd}
                    onChange={(e) => setHistEnd(e.target.value)}
                    className={`p-2 border rounded-xl text-sm font-bold outline-none focus:ring-2 focus:ring-blue-500 ${
                      isDarkMode ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white border-slate-300 text-slate-800'
                    }`}
                  />
                </div>
                <button
                  onClick={handleSearchArchive}
                  disabled={isSearchingArchive}
                  className="flex items-center gap-2 px-4 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-black uppercase text-[10px] tracking-widest transition-all shadow-lg disabled:opacity-60"
                >
                  {isSearchingArchive ? <><Loader2 size={16} className="animate-spin" /> Buscando...</> : <><Search size={16} /> Buscar</>}
                </button>
                <span className={`text-[10px] font-bold ${isDarkMode ? 'text-slate-500' : 'text-slate-400'}`}>
                  {filteredArchivedResults.length} resultado(s)
                </span>
                {hasHistoryActiveFilters && (
                  <span className={`text-[10px] font-bold px-2 py-1 rounded-full ${
                    isDarkMode ? 'text-amber-300 bg-amber-900/30 border border-amber-700' : 'text-amber-700 bg-amber-100 border border-amber-300'
                  }`}>
                    FILTRADO ({filteredArchivedResults.length} de {archivedResults.length})
                  </span>
                )}
                {hasHistoryActiveFilters && (
                  <button
                    onClick={clearHistoryFilters}
                    className="px-3 py-2 rounded-lg bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-800 text-red-600 dark:text-red-300 text-[10px] font-black uppercase tracking-widest hover:bg-red-100 dark:hover:bg-red-900/50 transition-all"
                  >
                    Limpar Filtros
                  </button>
                )}
              </div>
            </div>

            {/* Tabela de Resultados */}
            <div className={`overflow-auto flex-1 ${isDarkMode ? 'bg-slate-950' : 'bg-slate-50'}`}>
              {filteredArchivedResults.length === 0 && !isSearchingArchive ? (
                <div className="flex items-center justify-center h-40">
                  <p className={`text-sm font-bold ${isDarkMode ? 'text-slate-500' : 'text-slate-400'}`}>
                    Nenhum resultado encontrado para o período/filtros selecionados.
                  </p>
                </div>
              ) : (
                <table className="w-full text-xs">
                  <thead className={`sticky top-0 z-10 ${isDarkMode ? 'bg-slate-900' : 'bg-slate-200'}`}>
                    <tr>
                      {historyColumns.map(({ key, label }) => {
                        const hasColumnFilter = (historySelectedFilters[key] || []).length > 0 || !!historyColFilters[key];

                        return (
                          <th
                            key={key}
                            className={`relative px-3 py-2 text-left font-black uppercase text-[9px] tracking-wider ${isDarkMode ? 'text-slate-400' : 'text-slate-600'}`}
                          >
                            <div className="flex items-center gap-2">
                              <span>{label}</span>
                              <button
                                onClick={() => setHistoryActiveFilterCol(historyActiveFilterCol === key ? null : key)}
                                className={`p-1 rounded transition-colors ${
                                  hasColumnFilter
                                    ? 'bg-blue-600 text-white'
                                    : isDarkMode
                                      ? 'hover:bg-slate-700 text-slate-400'
                                      : 'hover:bg-slate-300 text-slate-500'
                                }`}
                                title={`Filtrar ${label}`}
                              >
                                <Filter size={12} />
                              </button>
                            </div>

                            {historyActiveFilterCol === key && (
                              <div
                                ref={historyFilterDropdownRef}
                                className={`absolute top-full left-0 mt-2 z-50 w-56 border rounded-xl shadow-2xl p-3 animate-in fade-in zoom-in-95 duration-150 ${
                                  isDarkMode
                                    ? 'bg-slate-800 border-slate-700'
                                    : 'bg-white border-slate-200'
                                }`}
                              >
                                <div
                                  className={`flex items-center gap-2 mb-3 p-2 rounded-lg border ${
                                    isDarkMode ? 'bg-slate-900 border-slate-700' : 'bg-slate-50 border-slate-200'
                                  }`}
                                >
                                  <Search size={14} className="text-slate-400" />
                                  <input
                                    type="text"
                                    placeholder="Filtrar..."
                                    autoFocus
                                    value={historyColFilters[key] || ''}
                                    onChange={(e) => setHistoryColFilters({ ...historyColFilters, [key]: e.target.value })}
                                    className={`w-full bg-transparent outline-none text-[10px] font-bold ${
                                      isDarkMode ? 'text-white' : 'text-slate-800'
                                    }`}
                                  />
                                </div>

                                <div
                                  className={`max-h-48 overflow-y-auto space-y-1 py-2 border-t ${
                                    isDarkMode ? 'border-slate-700' : 'border-slate-100'
                                  }`}
                                >
                                  {getHistoryUniqueValues(key).map((value) => (
                                    <div
                                      key={value}
                                      onClick={() => toggleHistoryFilterValue(key, value)}
                                      className={`flex items-center gap-2 p-2 rounded-lg cursor-pointer transition-all ${
                                        isDarkMode ? 'hover:bg-slate-700' : 'hover:bg-slate-50'
                                      }`}
                                    >
                                      {(historySelectedFilters[key] || []).includes(value) ? (
                                        <CheckSquare size={14} className="text-blue-600" />
                                      ) : (
                                        <Square size={14} className="text-slate-300" />
                                      )}
                                      <span
                                        className={`text-[10px] font-bold uppercase truncate ${
                                          isDarkMode ? 'text-slate-300' : 'text-slate-700'
                                        }`}
                                      >
                                        {value}
                                      </span>
                                    </div>
                                  ))}
                                </div>

                                <button
                                  onClick={() => {
                                    setHistoryColFilters({ ...historyColFilters, [key]: '' });
                                    setHistorySelectedFilters({ ...historySelectedFilters, [key]: [] });
                                    setHistoryActiveFilterCol(null);
                                  }}
                                  className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-600 bg-red-50 dark:bg-red-900/30 hover:bg-red-100 dark:hover:bg-red-900/50 rounded-lg border border-red-100 dark:border-red-900/50 transition-colors"
                                >
                                  Limpar Filtro
                                </button>
                              </div>
                            )}
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredArchivedResults.map((nc) => {
                      const pending = pendingHistoryEdits[nc.id] || {};
                      const operacoes = Array.from(new Set([...userConfigs.map(c => c.operacao), nc.operacao].filter(Boolean)));
                      return (
                        <tr key={nc.id} className={`border-t ${isDarkMode ? 'border-slate-800 hover:bg-slate-900' : 'border-slate-200 hover:bg-slate-100'} transition-colors`}>
                          <td className={`px-3 py-2 font-bold ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {pending.semana || nc.semana}
                          </td>

                          <td className={`px-3 py-2 ${isDarkMode ? 'text-white' : 'text-slate-800'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'rota' ? (
                              <input
                                type="text"
                                value={pending.rota ?? nc.rota}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'rota', e.target.value)}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-bold outline-none"
                                autoFocus
                              />
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('rota'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 ${
                                  pending.rota ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                              >
                                {pending.rota || nc.rota || '---'}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'data' ? (
                              <input
                                type="text"
                                value={pending.data ?? formatDisplayDate(nc.data)}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'data', applyDateMask(e.target.value))}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-mono font-bold outline-none"
                                autoFocus
                              />
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('data'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 ${
                                  pending.data ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                              >
                                {pending.data || formatDisplayDate(nc.data)}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'codigo' ? (
                              <input
                                type="text"
                                value={pending.codigo ?? nc.codigo}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'codigo', e.target.value.toUpperCase())}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-bold outline-none uppercase"
                                autoFocus
                              />
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('codigo'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 ${
                                  pending.codigo ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                              >
                                {pending.codigo || nc.codigo || '---'}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'produtor' ? (
                              <input
                                type="text"
                                value={pending.produtor ?? nc.produtor}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'produtor', e.target.value)}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-bold outline-none"
                                autoFocus
                              />
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('produtor'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 ${
                                  pending.produtor ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                              >
                                {pending.produtor || nc.produtor || '---'}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 max-w-[260px] ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'motivo' ? (
                              <select
                                value={pending.motivo ?? nc.motivo}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'motivo', e.target.value)}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-bold outline-none"
                                autoFocus
                              >
                                <option value="">---</option>
                                {Object.keys(MOTIVOS_CulpabilidadeS).map(label => (
                                  <option key={label} value={label}>{label}</option>
                                ))}
                              </select>
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('motivo'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 truncate ${
                                  pending.motivo ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                                title={pending.motivo || nc.motivo}
                              >
                                {pending.motivo || nc.motivo || '---'}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 max-w-[260px] ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'acao' ? (
                              <textarea
                                value={pending.acao ?? nc.acao}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'acao', e.target.value)}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                rows={2}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-bold outline-none resize-none"
                                autoFocus
                              />
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('acao'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 truncate ${
                                  pending.acao ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                                title={pending.acao || nc.acao}
                              >
                                {pending.acao || nc.acao || '---'}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'dataAcao' ? (
                              <input
                                type="text"
                                value={pending.dataAcao ?? formatDisplayDate(nc.dataAcao)}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'dataAcao', applyDateMask(e.target.value))}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-mono font-bold outline-none"
                                autoFocus
                              />
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('dataAcao'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 ${
                                  pending.dataAcao ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                              >
                                {pending.dataAcao || formatDisplayDate(nc.dataAcao)}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 ${isDarkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'ultimaColeta' ? (
                              <input
                                type="text"
                                value={pending.ultimaColeta ?? formatDisplayDate(nc.ultimaColeta)}
                                onChange={(e) => {
                                  const raw = e.target.value;
                                  handleUpdateHistoryCell(
                                    nc.id,
                                    'ultimaColeta',
                                    raw === '-' ? raw : applyDateMask(raw)
                                  );
                                }}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-mono font-bold outline-none"
                                autoFocus
                              />
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('ultimaColeta'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 ${
                                  pending.ultimaColeta ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                              >
                                {pending.ultimaColeta || formatDisplayDate(nc.ultimaColeta)}
                              </div>
                            )}
                          </td>

                          <td className="px-3 py-2">
                            {editingHistoryId === nc.id && editingHistoryField === 'Culpabilidade' ? (
                              <select
                                value={pending.Culpabilidade ?? nc.Culpabilidade}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'Culpabilidade', e.target.value)}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-bold outline-none"
                                autoFocus
                              >
                                <option value="">---</option>
                                {Culpabilidade_OPCOES.map(culp => (
                                  <option key={culp} value={culp}>{culp}</option>
                                ))}
                              </select>
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('Culpabilidade'); }}
                                className={`inline-flex px-2 py-1 rounded-full text-[9px] font-black uppercase cursor-pointer ${
                                  (pending.Culpabilidade || nc.Culpabilidade) === 'VIA'
                                    ? 'bg-red-100 text-red-700 dark:bg-red-900/40 dark:text-red-300'
                                    : (pending.Culpabilidade || nc.Culpabilidade) === 'Cliente'
                                      ? 'bg-amber-100 text-amber-700 dark:bg-amber-900/40 dark:text-amber-300'
                                      : 'bg-slate-100 text-slate-600 dark:bg-slate-800 dark:text-slate-400'
                                }`}
                                title="Clique para editar"
                              >
                                {pending.Culpabilidade || nc.Culpabilidade || '-'}
                              </div>
                            )}
                          </td>

                          <td className={`px-3 py-2 ${isDarkMode ? 'text-blue-400' : 'text-blue-600'}`}>
                            {editingHistoryId === nc.id && editingHistoryField === 'operacao' ? (
                              <select
                                value={pending.operacao ?? nc.operacao}
                                onChange={(e) => handleUpdateHistoryCell(nc.id, 'operacao', e.target.value)}
                                onBlur={() => { setEditingHistoryId(null); setEditingHistoryField(null); }}
                                className="w-full bg-blue-100 dark:bg-blue-900/30 border-2 border-blue-500 px-2 py-1 rounded font-bold outline-none"
                                autoFocus
                              >
                                <option value="">---</option>
                                {operacoes.map(op => (
                                  <option key={op} value={op}>{op}</option>
                                ))}
                              </select>
                            ) : (
                              <div
                                onClick={() => { setEditingHistoryId(nc.id); setEditingHistoryField('operacao'); }}
                                className={`font-bold cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 rounded px-1 ${
                                  pending.operacao ? 'bg-amber-200 dark:bg-amber-800/50' : ''
                                }`}
                              >
                                {pending.operacao || nc.operacao || '---'}
                              </div>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            {/* Footer */}
            <div className={`p-4 border-t shrink-0 flex justify-end ${isDarkMode ? 'bg-slate-900 border-slate-700' : 'bg-white border-slate-200'}`}>
              <button
                onClick={closeHistoryModal}
                className={`px-6 py-3 rounded-xl font-black uppercase text-[10px] tracking-widest transition-all ${
                  isDarkMode
                    ? 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                }`}
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default NonCollectionsView;

