import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteConfig, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { getValidToken } from '../services/tokenService';
import * as XLSX from 'xlsx';
import { getBrazilDate, getWeekString } from '../utils/dateUtils';
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
  culpabilidade: string;
  operacao: string;
}

// Lista de MOTIVOS no padrão "MOTIVO - CULPABILIDADE"
const MOTIVOS_CULPABILIDADES = [
  { label: 'Rota Atrasada/Fábrica', culpabilidades: ['Cliente'] },
  { label: 'Rota Atrasada/Logística', culpabilidades: ['VIA'] },
  { label: 'Rota Atrasada/Manutenção', culpabilidades: ['VIA'] },
  { label: 'Rota Atrasada/Mão De Obra', culpabilidades: ['VIA'] },
  { label: 'Eqp. Cheio/Coleta Extra', culpabilidades: ['VIA', 'Cliente', 'Outros'], causaRaiz: true },
  { label: 'Eqp. Cheio/Eqp Menor', culpabilidades: ['VIA', 'Cliente'], causaRaiz: true },
  { label: 'Eqp. Cheio/Rota Atrasada', culpabilidades: ['VIA', 'Cliente', 'Outros'], causaRaiz: true },
  { label: 'Alizarol Positivo', culpabilidades: ['Outros'] },
  { label: 'Leite Com Antibiótico', culpabilidades: ['Outros'] },
  { label: 'Leite Congelado', culpabilidades: ['Outros'] },
  { label: 'Leite Descartado', culpabilidades: ['Outros'] },
  { label: 'Leite Quente', culpabilidades: ['Outros'] },
  { label: 'Não Coletado - Saúde Do Motorista', culpabilidades: ['VIA'] },
  { label: 'Não Coletado - Greve', culpabilidades: ['Outros'] },
  { label: 'Não Coletado - Solicitado Pelo Sdl', culpabilidades: ['Cliente'] },
  { label: 'Objeto Ou Sujeira No Leite', culpabilidades: ['Outros'] },
  { label: 'Parou De Fornecer', culpabilidades: ['Outros'] },
  { label: 'Produtor suspenso', culpabilidades: ['Cliente'] },
  { label: 'Passou Para 48Hrs', culpabilidades: ['Outros'] },
  { label: 'Problemas Mecânicos Equipamento', culpabilidades: ['VIA'] },
  { label: 'Produtor Solicitou A Não Coleta', culpabilidades: ['Outros'] },
  { label: 'Resfriador Vazio', culpabilidades: ['Outros'] },
  { label: 'Volume Insuficiente Para Medida', culpabilidades: ['Outros'] },
  { label: 'Coletado Por Outra Transportadora', culpabilidades: ['VIA', 'Cliente', 'Outros'], causaRaiz: true },
  { label: 'Descumprimento de roteirização', culpabilidades: ['VIA'] },
  { label: 'A rota não foi realizada', culpabilidades: ['VIA', 'Cliente'], causaRaiz: true },
  { label: 'Falta De Acesso', culpabilidades: ['Outros'] },
  { label: 'Jornada Excedida', culpabilidades: ['VIA'] },
  { label: 'Eqp. Cheio/Aumento de Vol.', culpabilidades: ['Outros'] },
  { label: 'Correção de Roteirzação', culpabilidades: ['VIA'] },
  { label: 'Crioscopia', culpabilidades: ['Outros'] },
  { label: 'Rota Atrasada/Infraestrutura', culpabilidades: ['Outros'] }
];

const CULPABILIDADE_OPCOES = ['VIA', 'Cliente', 'Outros'];

const NonCollectionsView: React.FC<{
  currentUser: User;
}> = ({ currentUser }) => {
  const [nonCollections, setNonCollections] = useState<NonCollection[]>([]);
  const [userOps, setUserOps] = useState<RouteConfig[]>([]);
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
    culpabilidade: 130,
    operacao: 120
  });

  const [hiddenColumns, setHiddenColumns] = useState<Set<string>>(() => {
    const saved = localStorage.getItem('non_collections_hidden_cols');
    if (saved) {
      return new Set(JSON.parse(saved));
    }
    return new Set();
  });

  const [contextMenu, setContextMenu] = useState<{ visible: boolean; x: number; y: number; col: string | null }>({ visible: false, x: 0, y: 0, col: null });
  const filterDropdownRef = useRef<HTMLDivElement>(null);
  const contextMenuRef = useRef<HTMLDivElement>(null);
  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

  // Estados para modal de adicionar não coleta
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [newNonCollectionData, setNewNonCollectionData] = useState<{
    rota: string;
    data: string;
    codigo: string;
    produtor: string;
    operacao: string;
  }>({
    rota: '',
    data: getBrazilDate().split('-').reverse().join('/'),
    codigo: '',
    produtor: '',
    operacao: ''
  });

  // Ghost Row para adição rápida via paste
  const [ghostRow, setGhostRow] = useState<Partial<NonCollection>>({
    id: 'ghost',
    semana: '',
    rota: '',
    data: getBrazilDate().split('-').reverse().join('/'),
    codigo: '',
    produtor: '',
    motivo: '',
    observacao: '',
    acao: '',
    dataAcao: getBrazilDate().split('-').reverse().join('/'),
    ultimaColeta: getBrazilDate().split('-').reverse().join('/'),
    culpabilidade: '',
    operacao: ''
  });

  const [showCausaRaiz, setShowCausaRaiz] = useState(false);

  // Estados para modal de seleção de operação (bulk paste)
  const [isOperationModalOpen, setIsOperationModalOpen] = useState(false);
  const [pendingBulkRoutes, setPendingBulkRoutes] = useState<string[]>([]);

  const updateGhostCell = (field: keyof NonCollection, value: string) => {
    const updatedGhost = { ...ghostRow, [field]: value };

    if (field === 'motivo') {
      const motivoData = MOTIVOS_CULPABILIDADES.find(m => m.label === value);
      if (motivoData) {
        if (motivoData.culpabilidades.length === 1) {
          updatedGhost.culpabilidade = motivoData.culpabilidades[0];
        }
        setShowCausaRaiz(!!motivoData.causaRaiz);
      }
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
        alert('Erro: Token não encontrado');
        return;
      }

      const dataParaSemana = ghostRow.data!.split('/').reverse().join('-');
      const semana = getWeekString(dataParaSemana);

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
        dataAcao: ghostRow.dataAcao || ghostRow.data!,
        ultimaColeta: ghostRow.ultimaColeta || ghostRow.data!,
        culpabilidade: ghostRow.culpabilidade || 'Não se aplica',
        operacao: ghostRow.operacao!
      };

      // Salva no SharePoint
      await SharePointService.saveNonCollection(token, newRecord);

      // Adiciona localmente
      setNonCollections(prev => [...prev, newRecord]);

      // Limpa ghost row
      setGhostRow({
        id: 'ghost',
        semana: '',
        rota: '',
        data: getBrazilDate().split('-').reverse().join('/'),
        codigo: '',
        produtor: '',
        motivo: '',
        observacao: '',
        acao: '',
        dataAcao: getBrazilDate().split('-').reverse().join('/'),
        ultimaColeta: getBrazilDate().split('-').reverse().join('/'),
        culpabilidade: '',
        operacao: ''
      });
      setShowCausaRaiz(false);
    } catch (e: any) {
      console.error('[ADD_NON_COLLECTION] Error:', e);
      alert(`Erro ao adicionar não coleta: ${e.message}`);
    }
  };

  /**
   * Cria novas linhas de não coleta após o usuário selecionar a operação no modal.
   */
  const createBulkRecordsWithOperation = async (operacao: string) => {
    const dataFormatada = getBrazilDate().split('-').reverse().join('/');
    const dataParaSemana = dataFormatada.split('/').reverse().join('-');
    const semana = getWeekString(dataParaSemana);

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
      dataAcao: dataFormatada,
      ultimaColeta: dataFormatada,
      culpabilidade: 'Não se aplica',
      operacao
    }));

    setNonCollections(prev => [...prev, ...newRecords]);
    console.log('[BULK_PASTE] ✅', newRecords.length, 'linhas criadas com operação:', operacao);

    // Limpa estados do modal
    setIsOperationModalOpen(false);
    setPendingBulkRoutes([]);
  };

  /**
   * Cancela o modal de seleção de operação sem criar linhas.
   */
  const cancelBulkPaste = () => {
    setIsOperationModalOpen(false);
    setPendingBulkRoutes([]);
  };

  // Fecha dropdown ao clicar fora
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (filterDropdownRef.current && !filterDropdownRef.current.contains(e.target as Node)) {
        setActiveFilterCol(null);
      }
      if (contextMenuRef.current && !contextMenuRef.current.contains(e.target as Node)) {
        setContextMenu(prev => ({ ...prev, visible: false }));
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
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

  // Carrega operações do usuário ao montar
  useEffect(() => {
    loadUserOperations();
  }, [currentUser]);

  const loadUserOperations = async () => {
    try {
      const token = await getValidToken() || currentUser.accessToken;
      if (!token) {
        console.error('[NonCollections] Token não encontrado');
        return;
      }

      const configs = await SharePointService.getRouteConfigs(token, currentUser.email, true);
      setUserOps(configs || []);
      console.log('[NonCollections] Operações carregadas:', configs?.map(c => c.operacao));

      // Carrega não coletas do SharePoint
      const spNonCollections = await SharePointService.getNonCollections(token, currentUser.email);
      setNonCollections(spNonCollections);
      console.log('[NonCollections] ✅ Não coletas carregadas:', spNonCollections.length);
    } catch (e) {
      console.error('[NonCollections] Erro ao carregar operações:', e);
    } finally {
      setIsLoading(false);
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
        dataAcao: newNonCollectionData.data,
        ultimaColeta: newNonCollectionData.data,
        culpabilidade: 'Não se aplica',
        operacao: newNonCollectionData.operacao
      };

      setNonCollections(prev => [...prev, newRecord]);
      setIsAddModalOpen(false);
      setNewNonCollectionData({
        rota: '',
        data: getBrazilDate().split('-').reverse().join('/'),
        codigo: '',
        produtor: '',
        operacao: ''
      });
    } catch (e) {
      console.error('[ADD_NON_COLLECTION] Error:', e);
      alert('Erro ao adicionar não coleta');
    }
  };

  /**
   * Tenta separar código e produtor de uma string colada.
   * Padrão esperado: código alfanumérico no início seguido de texto (ex: "202520769MARCELO DINIZ COUTO" ou "P0274001RODRIGO ALVES PEREIRA")
   * Retorna null se não conseguir identificar o padrão.
   */
  const parseCodigoProdutor = (line: string): { codigo: string; produtor: string } | null => {
    const trimmed = line.trim();
    
    // Regex: captura caracteres alfanuméricos no início (código) seguidos de texto (produtor)
    // O código pode ter letras e números (ex: "P0274001", "202520769")
    // O produtor é todo o restante da string
    // Ex: "P0274001RODRIGO ALVES PEREIRA (IN )" → codigo: "P0274001", produtor: "RODRIGO ALVES PEREIRA (IN )"
    const match = trimmed.match(/^([A-Za-z0-9]+)(.+)$/);
    
    if (match) {
      return {
        codigo: match[1].toUpperCase(),
        produtor: match[2].trim()
      };
    }
    
    return null;
  };

  const handleBulkPaste = (field: keyof NonCollection, value: string) => {
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
      // Abre modal para selecionar operação antes de criar as linhas
      setPendingBulkRoutes(lines);
      setIsOperationModalOpen(true);
    } else {
      // COMPORTAMENTO 2: Atualizar linhas existentes (para CÓDIGO, PRODUTOR, MOTIVO, OBSERVAÇÃO, etc.)
      console.log('[BULK_PASTE] Atualizando linhas existentes...');
      const updatedRecords: NonCollection[] = [];

      for (let i = 0; i < Math.min(lines.length, nonCollections.length); i++) {
        const record = nonCollections[i];
        let finalValue = lines[i];
        let codigo = '';
        let produtor = '';

        // Tenta separar código e produtor se o valor tiver o padrão
        if (field === 'produtor' || field === 'codigo') {
          const parsed = parseCodigoProdutor(lines[i]);
          if (parsed) {
            codigo = parsed.codigo;
            produtor = parsed.produtor;
            console.log('[BULK_PASTE] ✅ Código e produtor separados:', parsed);
          }
        }

        // Formata se for data
        if (field === 'data' || field === 'dataAcao' || field === 'ultimaColeta') {
          if (finalValue.includes('-')) {
            const [year, month, day] = finalValue.split('-');
            finalValue = `${day}/${month}/${year}`;
          }
        }

        // CÓDIGO: converte para maiúsculo (se estiver atualizando)
        if (field === 'codigo' && !codigo) {
          finalValue = finalValue.toUpperCase();
        }

        updatedRecords.push({
          ...record,
          ...(codigo && produtor ? { codigo, produtor } : { [field]: finalValue })
        });
        console.log('[BULK_PASTE] Atualizando linha', i, 'campo', field);
      }

      // Mantém as linhas que não foram atualizadas
      const remainingRecords = nonCollections.slice(updatedRecords.length);

      setNonCollections([...updatedRecords, ...remainingRecords]);
      console.log('[BULK_PASTE] ✅', updatedRecords.length, 'linhas atualizadas', remainingRecords.length, 'mantidas');
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
    { key: 'culpabilidade', label: 'CULPABILIDADE' },
    { key: 'operacao', label: 'OPERAÇÃO' }
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
        <div>
          <h1 className="text-2xl font-black uppercase text-slate-800 dark:text-white tracking-tight">
            Não Coletas
          </h1>
          <p className="text-xs font-bold text-slate-500 dark:text-slate-400 mt-1 uppercase tracking-widest">
            Acompanhamento de ocorrências
          </p>
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
            onClick={loadUserOperations}
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

      {/* Filters Bar */}
      <div className="flex items-center gap-3 px-6 py-3">
        <div className="flex items-center gap-2 px-4 py-2.5 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl shadow-sm flex-1 max-w-md">
          <Search size={16} className="text-slate-400" />
          <input
            type="text"
            placeholder="Buscar por produtor..."
            className="flex-1 bg-transparent outline-none text-sm font-medium text-slate-700 dark:text-slate-300 placeholder-slate-400"
            value={colFilters['produtor'] || ''}
            onChange={(e) => setColFilters({ ...colFilters, produtor: e.target.value })}
          />
        </div>

        <div className="flex items-center gap-2 px-4 py-2.5 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-xl shadow-sm">
          <CheckCircle2 size={16} className="text-blue-600 dark:text-blue-400" />
          <span className="text-xs font-black text-blue-700 dark:text-blue-300 uppercase">
            TOTAL: {nonCollections.length} não coleta(s)
          </span>
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
                              
                              // Linha única com código + produtor: separa automaticamente
                              const parsed = parseCodigoProdutor(val);
                              if (parsed) {
                                e.preventDefault();
                                const updated = { ...row, codigo: parsed.codigo, produtor: parsed.produtor };
                                setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                                console.log('[PASTE PRODUTOR] ✅ Código e produtor separados:', parsed);
                                return;
                              }
                              
                              // Se não tiver padrão código+produtor, deixa colar normalmente
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
                              
                              // Linha única com código + produtor: separa automaticamente
                              const parsed = parseCodigoProdutor(val);
                              if (parsed) {
                                e.preventDefault();
                                const updated = { ...row, codigo: parsed.codigo, produtor: parsed.produtor };
                                setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                                console.log('[PASTE CODIGO] ✅ Código e produtor separados:', parsed);
                                return;
                              }
                              
                              // Se não tiver padrão código+produtor, deixa colar normalmente (já converte para upper)
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
                              const updated = { ...row, motivo: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            className={`w-full px-3 py-2 text-[11px] text-left truncate transition-all cursor-pointer outline-none border-none bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-slate-200 hover:bg-slate-200 dark:hover:bg-slate-700 focus:ring-2 focus:ring-blue-500 rounded ${
                              isDarkMode ? 'dark-mode-select' : ''
                            }`}
                          >
                            <option value="" className="bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200">Selecione...</option>
                            {MOTIVOS_CULPABILIDADES.map(m => (
                              <option 
                                key={m.label} 
                                value={m.label}
                                className="bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200"
                              >
                                {m.label}
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
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('observacao', val);
                              }
                            }}
                            placeholder="Detalhes + causa raiz..."
                            rows={1}
                            className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                          />
                        </td>
                      );
                    }

                    // AÇÃO - Textarea editável
                    if (key === 'acao') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <textarea
                            value={row.acao}
                            onChange={(e) => {
                              const updated = { ...row, acao: e.target.value };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
                            }}
                            onPaste={(e) => {
                              const val = e.clipboardData.getData('text');
                              if (val.includes('\n')) {
                                e.preventDefault();
                                handleBulkPaste('acao', val);
                              }
                            }}
                            placeholder="Ação..."
                            rows={1}
                            className={`${inputClass} resize-none whitespace-pre-wrap break-words min-h-[48px]`}
                          />
                        </td>
                      );
                    }

                    // DATA AÇÃO - Input editável com máscara
                    if (key === 'dataAcao') {
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
                            value={row.dataAcao}
                            onChange={(e) => {
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              const updated = { ...row, dataAcao: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
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
                              let val = e.target.value.replace(/\D/g, '');
                              if (val.length > 8) val = val.slice(0, 8);
                              if (val.length >= 8) {
                                val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
                              }
                              const updated = { ...row, ultimaColeta: val };
                              setNonCollections(prev => prev.map(r => r.id === row.id ? updated : r));
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
                            {userOps.map(op => (
                              <option key={op.operacao} value={op.operacao} className="bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-200">{op.operacao}</option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // CULPABILIDADE - Badge (somente leitura)
                    if (key === 'culpabilidade') {
                      return (
                        <td
                          key={key}
                          className={`p-0 border border-slate-200/30 dark:border-slate-800/30 ${
                            hiddenColumns.has(key) ? 'hidden' : ''
                          }`}
                          style={{ minWidth: colWidths[key] }}
                        >
                          <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-[9px] font-black uppercase tracking-wider ${
                            row.culpabilidade?.toLowerCase() === 'produtor'
                              ? 'bg-orange-100 dark:bg-orange-900/40 text-orange-700 dark:text-orange-300'
                              : row.culpabilidade?.toLowerCase() === 'logística' || row.culpabilidade?.toLowerCase() === 'logistica'
                              ? 'bg-yellow-100 dark:bg-yellow-900/40 text-yellow-700 dark:text-yellow-300'
                              : row.culpabilidade?.toLowerCase() === 'clima'
                              ? 'bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300'
                              : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
                          }`}>
                            {row.culpabilidade}
                          </span>
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
                  {/* Coluna de ação */}
                  <td className="p-0 border border-slate-200/30 dark:border-slate-800/30 text-center">
                    <button
                      onClick={() => {
                        setNonCollections(prev => prev.filter(r => r.id !== row.id));
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
                          {userOps.map(op => (
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
                          placeholder="Cole códigos..."
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
                          onChange={(e) => updateGhostCell('motivo', e.target.value)}
                          className={`${inputClass} text-left cursor-pointer`}
                        >
                          <option value="">Selecione...</option>
                          {MOTIVOS_CULPABILIDADES.map(m => (
                            <option key={m.label} value={m.label}>{m.label}</option>
                          ))}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'culpabilidade') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <select
                          value={ghostRow.culpabilidade || ''}
                          onChange={(e) => updateGhostCell('culpabilidade', e.target.value)}
                          className={`${inputClass} text-center cursor-pointer`}
                          disabled={!ghostRow.motivo}
                        >
                          <option value="">Selecione...</option>
                          {(() => {
                            const motivoData = MOTIVOS_CULPABILIDADES.find(m => m.label === ghostRow.motivo);
                            const opcoes = motivoData?.culpabilidades || CULPABILIDADE_OPCOES;
                            return opcoes.map(culp => (
                              <option key={culp} value={culp}>{culp}</option>
                            ));
                          })()}
                        </select>
                      </td>
                    );
                  }

                  if (key === 'observacao' || key === 'acao') {
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <textarea
                          value={ghostRow[key] || ''}
                          onChange={(e) => updateGhostCell(key, e.target.value)}
                          onPaste={(e) => {
                            const val = e.clipboardData.getData('text');
                            if (val.includes('\n')) {
                              e.preventDefault();
                              handleBulkPaste(key, val);
                            }
                          }}
                          placeholder={key === 'observacao' ? 'Detalhes + causa raiz...' : 'Ação...'}
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
                    return (
                      <td key={`ghost-${key}`} className="p-0 border border-slate-200/30 dark:border-slate-800/30" style={{ verticalAlign: 'middle' }}>
                        <input
                          type="text"
                          value={ghostRow[key] || ''}
                          onChange={(e) => {
                            let val = e.target.value.replace(/\D/g, '');
                            if (val.length > 8) val = val.slice(0, 8);
                            if (val.length >= 8) {
                              val = `${val.slice(0, 2)}/${val.slice(2, 4)}/${val.slice(4, 8)}`;
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
                {/* Coluna de ação */}
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
                        Nenhuma não coleta registrada
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
                  <option value="">Selecione a operação</option>
                  {userOps.map(config => (
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

      {/* Footer Info */}
      <div className="px-6 py-3 text-center">
        <p className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-widest">
          Visualização apenas • Dados não persistidos
        </p>
      </div>

      {/* Modal de Seleção de Operação (Bulk Paste) */}
      {isOperationModalOpen && (
        <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[300] flex items-center justify-center p-4 animate-in fade-in zoom-in-95 duration-200">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-blue-500/50 shadow-2xl animate-in zoom-in duration-300">
            <div className="flex items-center gap-3 text-blue-600 dark:text-blue-400 mb-6 font-black uppercase text-xs">
              <Milk size={24} />
              Selecionar Operação
            </div>
            <p className="text-sm text-slate-600 dark:text-slate-400 mb-2 font-medium">
              {pendingBulkRoutes.length} rota(s) colada(s). Selecione a operação:
            </p>
            <div className="mb-6 max-h-32 overflow-y-auto scrollbar-thin bg-slate-50 dark:bg-slate-800 rounded-xl p-3 border border-slate-200 dark:border-slate-700">
              {pendingBulkRoutes.slice(0, 10).map((rota, i) => (
                <div key={i} className="text-[10px] font-bold text-slate-600 dark:text-slate-400 py-1 truncate">
                  • {rota}
                </div>
              ))}
              {pendingBulkRoutes.length > 10 && (
                <div className="text-[10px] font-bold text-slate-400 italic mt-1">
                  ...e mais {pendingBulkRoutes.length - 10} rota(s)
                </div>
              )}
            </div>

            {userOps.length === 0 ? (
              <div className="text-center py-6 text-slate-400 dark:text-slate-500">
                <AlertTriangle size={32} className="mx-auto mb-3 opacity-50" />
                <p className="text-sm font-bold">Nenhuma operação disponível</p>
                <p className="text-[10px] mt-1">Contate o administrador para configurar suas operações.</p>
              </div>
            ) : (
              <div className="grid grid-cols-2 gap-3">
                {userOps.map(op => (
                  <button
                    key={op.operacao}
                    onClick={() => createBulkRecordsWithOperation(op.operacao)}
                    className="p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-2xl hover:bg-blue-600 hover:text-white hover:border-blue-700 dark:hover:bg-blue-600 dark:hover:border-blue-500 transition-all font-black text-xs uppercase text-slate-700 dark:text-slate-300"
                  >
                    {op.operacao}
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
    </div>
  );
};

export default NonCollectionsView;
