
import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Task, OperationStatus, User } from '../types';
import { SharePointService, DailyWarning } from '../services/sharepointService';
import { 
  Maximize2, Minimize2, Loader2, Database, 
  ShieldCheck, AlertCircle, RefreshCw, CheckCircle,
  Activity, Lock, CheckCircle2, PaintBucket,
  HelpCircle, X, LogOut, ChevronDown, ChevronRight,
  RotateCcw, Save, UserCheck, Bell, MessageSquarePlus, Megaphone,
  Clock, Check
} from 'lucide-react';

const STATUS_CONFIG: Record<string, { label: string, color: string, next: OperationStatus, shortcut: string, desc: string }> = {
  'OK': { label: 'OK', color: 'bg-green-200 text-green-800 border-green-300 dark:bg-green-900/60 dark:text-green-300 dark:border-green-800', next: 'EA', shortcut: '1', desc: 'Concluído' },
  'EA': { label: 'EA', color: 'bg-yellow-200 text-yellow-800 border-yellow-300 dark:bg-yellow-900/60 dark:text-yellow-300 dark:border-yellow-800', next: 'AR', shortcut: '2', desc: 'Em Andamento' },
  'ATT': { label: 'ATT', color: 'bg-blue-200 text-blue-800 border-blue-300 dark:bg-blue-900/60 dark:text-blue-300 dark:border-blue-800', next: 'AT', shortcut: '3', desc: 'Atualizar' },
  'AR': { label: 'AR', color: 'bg-orange-200 text-orange-800 border-orange-300 dark:bg-orange-900/60 dark:text-orange-300 dark:border-orange-800', next: 'ATT', shortcut: '4', desc: 'Aguardando Retorno' },
  'AT': { label: 'AT', color: 'bg-red-500 text-white border-red-600 dark:bg-red-800 dark:text-white dark:border-red-700', next: 'PR', shortcut: '5', desc: 'Atrasado' },
  'PR': { label: 'PR', color: 'bg-slate-200 text-slate-600 border-slate-300 dark:bg-slate-700 dark:text-slate-300 dark:border-slate-600', next: 'OK', shortcut: '6', desc: 'Pendente' },
};

// Added teamMembers to fix the TS error in App.tsx
interface TaskManagerProps {
  tasks: Task[];
  setTasks: React.Dispatch<React.SetStateAction<Task[]>>;
  locations: string[];
  setLocations: any;
  onUserSwitch: any;
  collapsedCategories: string[];
  setCollapsedCategories: any;
  currentUser: User;
  onLogout: () => void;
  teamMembers: string[];
}

const TaskManager: React.FC<TaskManagerProps> = ({ 
  tasks, 
  setTasks, 
  locations, 
  collapsedCategories,
  setCollapsedCategories,
  onUserSwitch, 
  currentUser,
  onLogout,
  teamMembers
}) => {
  const [activeTool, setActiveTool] = useState<OperationStatus | null>(null);
  const [isUpdating, setIsUpdating] = useState(false);
  const [compact, setCompact] = useState(true);
  
  // Estados para Avisos Diários
  const [dailyWarnings, setDailyWarnings] = useState<DailyWarning[]>([]);
  const [isCreateWarningModalOpen, setIsCreateWarningModalOpen] = useState(false);
  const [isViewWarningsModalOpen, setIsViewWarningsModalOpen] = useState(false);
  const [newWarning, setNewWarning] = useState<Omit<DailyWarning, 'id' | 'visualizado'>>({
      operacao: '', celula: currentUser.email, rota: '', descricao: '', dataOcorrencia: new Date().toISOString().split('T')[0]
  });

  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [resetResponsible, setResetResponsible] = useState('');
  const [registeredUsers, setRegisteredUsers] = useState<string[]>([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState(false);

  const [isDragging, setIsDragging] = useState(false);
  const paintedThisDrag = useRef<Set<string>>(new Set());
  
  const autoCollapsedSessionRef = useRef<Set<string>>(new Set());
  const manuallyOpenedRef = useRef<Set<string>>(new Set());

  const fetchWarnings = async () => {
    const token = currentUser.accessToken || (window as any).__access_token;
    if (token) {
        const warnings = await SharePointService.getDailyWarnings(token, currentUser.email);
        setDailyWarnings(warnings);
        // FIX: Auto-open modal if there are pending warnings
        if (warnings.length > 0) {
            setIsViewWarningsModalOpen(true);
        }
    }
  };

  useEffect(() => {
      fetchWarnings();
      // Re-busca a cada 5 minutos para ver se há novos avisos de outros usuários
      const interval = setInterval(fetchWarnings, 300000); 
      return () => clearInterval(interval);
  }, [currentUser]);

  const handleAddWarning = async (e: React.FormEvent) => {
    e.preventDefault();
    const token = currentUser.accessToken || (window as any).__access_token;
    if (!token) return;
    setIsUpdating(true);
    try {
        await SharePointService.addDailyWarning(token, newWarning);
        setIsCreateWarningModalOpen(false);
        setNewWarning({ operacao: '', celula: currentUser.email, rota: '', descricao: '', dataOcorrencia: new Date().toISOString().split('T')[0] });
        alert("Aviso registrado com sucesso!");
    } catch (err: any) {
        alert("Erro ao salvar aviso: " + err.message);
    } finally {
        setIsUpdating(false);
    }
  };

  const handleMarkAsViewed = async (id: string) => {
      const token = currentUser.accessToken || (window as any).__access_token;
      if (!token) return;
      try {
          await SharePointService.markWarningAsViewed(token, id);
          // Remove localmente para feedback imediato
          setDailyWarnings(prev => prev.filter(w => w.id !== id));
          // Se não houver mais avisos, fecha o modal sozinho
          if (dailyWarnings.length <= 1) {
              setIsViewWarningsModalOpen(false);
          }
      } catch (err: any) {
          console.error("Erro ao marcar como lido:", err);
          alert("Não foi possível marcar como lido no servidor.");
      }
  };

  const getCategoryStats = (category: string) => {
    const catTasks = tasks.filter(t => (t.category || 'Geral') === category);
    if (catTasks.length === 0) return { percent: 0, isComplete: false };
    let totalCells = 0, okCells = 0;
    catTasks.forEach(t => {
      locations.forEach(loc => {
        totalCells++;
        if (t.operations[loc] === 'OK') okCells++;
      });
    });
    const percent = totalCells === 0 ? 0 : Math.round((okCells / totalCells) * 100);
    return { percent, isComplete: percent === 100 };
  };

  useEffect(() => {
    const categories = Array.from(new Set<string>(tasks.map(t => t.category || 'Geral')));
    categories.forEach((cat: string) => {
        const stats = getCategoryStats(cat);
        
        if (stats.isComplete && 
            !collapsedCategories.includes(cat) && 
            !autoCollapsedSessionRef.current.has(cat) && 
            !manuallyOpenedRef.current.has(cat)) {
            
            setCollapsedCategories((prev: string[]) => prev.includes(cat) ? prev : [...prev, cat]);
            autoCollapsedSessionRef.current.add(cat);
        } else if (!stats.isComplete) {
            autoCollapsedSessionRef.current.delete(cat);
            manuallyOpenedRef.current.delete(cat);
        }
    });
  }, [tasks]);

  useEffect(() => {
    const handleMouseUp = () => {
      setIsDragging(false);
      paintedThisDrag.current.clear();
    };
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement;
      if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT') return;
      switch (e.key) {
        case '1': setActiveTool('OK'); break;
        case '2': setActiveTool('EA'); break;
        case '3': setActiveTool('ATT'); break;
        case '4': setActiveTool('AR'); break;
        case '5': setActiveTool('AT'); break;
        case '6': setActiveTool('PR'); break;
        case 'Escape': setActiveTool(null); break;
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, []);

  const handleOpenResetModal = async () => {
    setIsResetModalOpen(true);
    setIsLoadingUsers(true);
    try {
        const token = currentUser.accessToken || (window as any).__access_token;
        if (token) {
            const users = await SharePointService.getRegisteredUsers(token, currentUser.email);
            setRegisteredUsers(users);
            if (users.length === 1) {
                setResetResponsible(users[0]);
            } else {
                setResetResponsible('');
            }
        }
    } catch (e) {
        console.error("Erro ao carregar usuários:", e);
    } finally {
        setIsLoadingUsers(false);
    }
  };

  const handleUpdateStatus = async (taskId: string, location: string, status: OperationStatus) => {
    if (!currentUser.accessToken) return;
    
    const originalTasks = [...tasks];
    setTasks(prev => prev.map(t => t.id === taskId ? { 
      ...t, 
      operations: { ...t.operations, [location]: status } 
    } : t));
    
    setIsUpdating(true);
    try {
      const today = new Date().toISOString().split('T')[0];
      const todayKey = today.replace(/-/g, '');
      const uniqueKey = `${todayKey}_${taskId}_${location}`;
      
      await SharePointService.updateStatus(currentUser.accessToken, {
        DataReferencia: today,
        TarefaID: taskId,
        OperacaoSigla: location,
        Status: status,
        Usuario: currentUser.name,
        Title: uniqueKey
      });
    } catch (err: any) {
      console.error(`Erro ao sincronizar:`, err);
      alert(`Falha ao salvar no SharePoint: ${err.message}`);
      setTasks(originalTasks);
    } finally {
      setIsUpdating(false);
    }
  };

  const handlePaintRow = async (taskId: string) => {
    if (!activeTool || !currentUser.accessToken) return;
    
    const originalTasks = [...tasks];
    setTasks(prev => prev.map(t => t.id === taskId ? { 
      ...t, 
      operations: locations.reduce((acc, loc) => ({...acc, [loc]: activeTool}), {})
    } : t));

    setIsUpdating(true);
    try {
      const today = new Date().toISOString().split('T')[0];
      const todayKey = today.replace(/-/g, '');
      
      await Promise.all(locations.map(loc => {
        const uniqueKey = `${todayKey}_${taskId}_${loc}`;
        return SharePointService.updateStatus(currentUser.accessToken!, {
            DataReferencia: today,
            TarefaID: taskId,
            OperacaoSigla: loc,
            Status: activeTool,
            Usuario: currentUser.name,
            Title: uniqueKey
        });
      }));
    } catch (err: any) {
      alert(`Erro na sincronização em lote: ${err.message}`);
      setTasks(originalTasks);
    } finally {
      setIsUpdating(false);
    }
  };

  const handleResetChecklist = async () => {
    if (!resetResponsible.trim() || !currentUser.accessToken) return;
    
    setIsUpdating(true);
    try {
        await SharePointService.saveHistory(currentUser.accessToken, {
            id: Date.now().toString(),
            timestamp: new Date().toISOString(),
            tasks: tasks,
            resetBy: resetResponsible,
            email: currentUser.email
        });
        
        const today = new Date().toISOString().split('T')[0];
        const todayKey = today.replace(/-/g, '');
        
        const resetPromises: Promise<any>[] = [];
        tasks.forEach(task => {
            locations.forEach(loc => {
                const uniqueKey = `${todayKey}_${task.id}_${loc}`;
                resetPromises.push(SharePointService.updateStatus(currentUser.accessToken!, {
                    DataReferencia: today,
                    TarefaID: task.id,
                    OperacaoSigla: loc,
                    Status: 'PR',
                    Usuario: resetResponsible,
                    Title: uniqueKey
                }));
            });
        });

        await Promise.all(resetPromises);

        setTasks(prev => prev.map(t => ({
            ...t,
            operations: locations.reduce((acc, loc) => ({ ...acc, [loc]: 'PR' }), {})
        })));

        autoCollapsedSessionRef.current.clear();
        manuallyOpenedRef.current.clear();
        setIsResetModalOpen(false);
        alert("Checklist resetado e salvo com sucesso no SharePoint!");
    } catch (error: any) {
        console.error("Erro no Reset:", error);
        alert(`ERRO CRÍTICO: Não foi possível resetar. Detalhe: ${error.message}`);
    } finally {
        setIsUpdating(false);
    }
  };

  const onCellInteraction = (taskId: string, loc: string, forcedStatus?: OperationStatus) => {
    const task = tasks.find(t => t.id === taskId);
    if (!task) return;
    
    const currentStatus = task.operations[loc] || 'PR';
    const nextStatus = forcedStatus || (activeTool || STATUS_CONFIG[currentStatus].next);
    
    if (currentStatus !== nextStatus) {
      handleUpdateStatus(taskId, loc, nextStatus);
    }
  };

  const toggleCategory = (cat: string) => {
    const { isComplete } = getCategoryStats(cat);
    const isCurrentlyCollapsed = collapsedCategories.includes(cat);

    if (isCurrentlyCollapsed) {
      setCollapsedCategories((prev: string[]) => prev.filter(c => c !== cat));
      if (isComplete) {
        manuallyOpenedRef.current.add(cat);
      }
    } else {
      if (isComplete) {
        setCollapsedCategories((prev: string[]) => [...prev, cat]);
        manuallyOpenedRef.current.delete(cat);
      }
    }
  };

  const groupedTasks = useMemo(() => tasks.reduce((acc, t) => {
    const cat = t.category || 'Geral';
    if (!acc[cat]) acc[cat] = [];
    acc[cat].push(t);
    return acc;
  }, {} as Record<string, Task[]>), [tasks]);

  return (
    <div className="flex flex-col h-full bg-white dark:bg-slate-900 rounded-2xl border dark:border-slate-800 shadow-sm overflow-hidden relative font-sans transition-colors duration-500">
      
      {/* BADGE DE AVISOS PENDENTES */}
      {dailyWarnings.length > 0 && (
          <div onClick={() => setIsViewWarningsModalOpen(true)} className="bg-orange-500 hover:bg-orange-600 transition-colors text-white px-4 py-2 flex items-center justify-center gap-3 cursor-pointer shadow-lg z-[60]">
              <Megaphone size={16} className="animate-bounce" />
              <span className="text-[10px] font-black uppercase tracking-widest">Atenção: Você possui {dailyWarnings.length} aviso(s) operacional(is) pendente(s)</span>
              <ChevronRight size={14} />
          </div>
      )}

      {/* HEADER / TOOLBAR */}
      <div className="px-4 py-3 border-b dark:border-slate-800 flex flex-col xl:flex-row justify-between items-center bg-gray-50/80 dark:bg-slate-800/80 backdrop-blur-md gap-3 shrink-0 z-50">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <div className="p-2 bg-blue-600 rounded-lg text-white shadow-lg shadow-blue-500/20">
              <Activity size={20} />
            </div>
            <h2 className="text-lg font-bold text-gray-800 dark:text-white whitespace-nowrap">
              Checklist CCO
            </h2>
          </div>
          <div className="h-6 w-px bg-gray-300 dark:bg-slate-700 hidden md:block" />
          {isUpdating ? (
            <div className="flex items-center gap-2 text-[10px] text-blue-500 animate-pulse font-black uppercase tracking-widest">
              <Loader2 size={12} className="animate-spin"/> Gravando
            </div>
          ) : (
            <div className="flex items-center gap-2 text-[10px] text-green-500 font-bold uppercase tracking-widest">
              <ShieldCheck size={12}/> Protegido
            </div>
          )}
        </div>

        <div className="flex flex-wrap items-center justify-center gap-3">
          <div className="flex items-center gap-2 bg-white dark:bg-slate-950 px-3 py-1.5 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm">
            <div className="flex items-center gap-2 text-[10px] font-bold text-slate-400 uppercase tracking-tight mr-1">
              <PaintBucket size={14} className={activeTool ? 'text-blue-500' : 'text-slate-400'} />
              {activeTool || 'Pincel'}
            </div>
            <div className="flex items-center gap-1">
              {(Object.entries(STATUS_CONFIG) as [string, any][]).map(([key, cfg]) => (
                <button 
                  key={key} 
                  onClick={() => setActiveTool(activeTool === key ? null : key as OperationStatus)} 
                  className={`w-7 h-7 rounded-lg font-black text-[9px] transition-all duration-200 border flex items-center justify-center relative group ${cfg.color} ${activeTool === key ? 'ring-2 ring-offset-2 ring-blue-500 scale-110 z-10' : 'opacity-80 hover:opacity-100 hover:scale-105'}`}
                  title={`${cfg.desc} [${cfg.shortcut}]`}
                >
                  {cfg.label}
                  <span className="absolute -bottom-4 text-[8px] text-slate-400 opacity-0 group-hover:opacity-100 font-mono">{cfg.shortcut}</span>
                </button>
              ))}
            </div>
          </div>

          <div className="flex items-center gap-1">
            <button 
              onClick={() => setIsCreateWarningModalOpen(true)}
              className="flex items-center gap-2 px-3 py-2 bg-blue-50 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 rounded-xl hover:bg-blue-100 dark:hover:bg-blue-900/50 transition-all border border-blue-100 dark:border-blue-800 shadow-sm"
              title="Adicionar Aviso Operacional"
            >
              <MessageSquarePlus size={18} />
              <span className="text-xs font-bold hidden sm:inline">Aviso</span>
            </button>

            <button 
              onClick={handleOpenResetModal}
              className="flex items-center gap-2 px-3 py-2 bg-amber-50 dark:bg-amber-900/30 text-amber-600 dark:text-amber-400 rounded-xl hover:bg-amber-100 dark:hover:bg-amber-900/50 transition-all border border-amber-100 dark:border-amber-800 shadow-sm"
              title="Resetar Checklist e Salvar no SharePoint"
            >
              <RotateCcw size={18} />
              <span className="text-xs font-bold hidden sm:inline">Resetar</span>
            </button>
            <button onClick={() => setCompact(!compact)} className={`p-2 rounded-xl transition-all ${!compact ? 'bg-blue-100 text-blue-600 dark:bg-blue-900' : 'text-slate-400 hover:bg-slate-100'}`} title="Modo Visualização">
              {compact ? <Maximize2 size={18}/> : <Minimize2 size={18}/>}
            </button>
            <div className="w-px h-6 bg-slate-200 dark:bg-slate-700 mx-2" />
            <button onClick={onLogout} className="flex items-center gap-2 px-3 py-1.5 bg-red-50 dark:bg-red-900/30 text-red-600 dark:text-red-400 rounded-xl hover:bg-red-100 transition-all font-bold text-xs">
                <LogOut size={16}/> Sair
            </button>
          </div>
        </div>
      </div>

      {/* MODAL DE LEITURA DE AVISOS */}
      {isViewWarningsModalOpen && (
          <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-[110] flex items-center justify-center p-4 animate-in zoom-in duration-300">
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden border border-orange-500/30">
                  <div className="bg-orange-600 text-white p-6 flex justify-between items-center">
                      <div className="flex items-center gap-4">
                          <Bell size={32} className="animate-pulse" />
                          <div>
                              <h3 className="text-xl font-black uppercase tracking-tight">Avisos Operacionais do Dia</h3>
                              <p className="text-xs text-orange-200 font-bold uppercase tracking-widest">Favor revisar e acusar recebimento</p>
                          </div>
                      </div>
                      <button onClick={() => setIsViewWarningsModalOpen(false)} className="hover:bg-white/10 p-2 rounded-full transition-all">
                          <X size={24} />
                      </button>
                  </div>
                  <div className="p-6 max-h-[60vh] overflow-y-auto space-y-4 bg-slate-50 dark:bg-slate-950 scrollbar-thin">
                      {dailyWarnings.length === 0 ? (
                          <div className="py-12 text-center text-slate-400 italic font-bold">Todos os avisos foram lidos. Bom trabalho!</div>
                      ) : dailyWarnings.map(warning => (
                          <div key={warning.id} className="bg-white dark:bg-slate-900 p-6 rounded-[2rem] border-2 border-slate-200 dark:border-slate-800 shadow-sm group hover:border-orange-500 transition-all relative">
                              <div className="flex justify-between items-start mb-4">
                                  <div>
                                      <span className="px-3 py-1 bg-orange-100 dark:bg-orange-900/30 text-orange-600 dark:text-orange-400 text-[10px] font-black rounded-full uppercase mb-2 inline-block">
                                          Operação: {warning.operacao}
                                      </span>
                                      <h4 className="text-lg font-black text-slate-800 dark:text-white uppercase">Rota: {warning.rota}</h4>
                                  </div>
                                  
                                  {/* FIX: Botão de marcação como lido mais explícito */}
                                  <button 
                                    onClick={() => handleMarkAsViewed(warning.id)}
                                    className="p-3 px-5 bg-green-500 hover:bg-green-600 text-white rounded-2xl shadow-lg transition-all flex items-center gap-2 group-hover:scale-105 active:scale-95"
                                    title="Marcar como lido e não mostrar novamente"
                                  >
                                      <Check size={18} />
                                      <span className="text-[10px] font-black uppercase whitespace-nowrap">Marcar como lido</span>
                                  </button>
                              </div>
                              <p className="text-sm text-slate-600 dark:text-slate-400 leading-relaxed font-medium bg-slate-50 dark:bg-slate-800/50 p-4 rounded-2xl border border-slate-100 dark:border-slate-700">
                                  {warning.descricao}
                              </p>
                              <div className="mt-4 flex items-center justify-between">
                                  <div className="flex items-center gap-2 text-[9px] text-slate-400 font-bold uppercase tracking-widest">
                                      <Clock size={12} /> Data Ocorrência: {new Date(warning.dataOcorrencia).toLocaleDateString('pt-BR')}
                                  </div>
                                  <span className="text-[8px] font-bold text-slate-300 dark:text-slate-600">ID: {warning.id}</span>
                              </div>
                          </div>
                      ))}
                  </div>
                  <div className="p-6 bg-white dark:bg-slate-900 border-t dark:border-slate-800 flex justify-center gap-4">
                      <button onClick={() => setIsViewWarningsModalOpen(false)} className="px-8 py-3 bg-slate-100 dark:bg-slate-800 text-slate-500 dark:text-slate-300 rounded-xl font-black uppercase text-[10px] hover:bg-slate-200 transition-all tracking-widest border border-slate-200 dark:border-slate-700 shadow-sm">Dispensar Painel</button>
                  </div>
              </div>
          </div>
      )}

      {/* MODAL DE CRIAÇÃO DE AVISO */}
      {isCreateWarningModalOpen && (
          <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-[110] flex items-center justify-center p-4 animate-in fade-in duration-200">
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden border dark:border-slate-700">
                  <div className="bg-blue-600 text-white p-6 flex justify-between items-center">
                      <div className="flex items-center gap-3">
                          <MessageSquarePlus size={24} />
                          <h3 className="text-lg font-black uppercase tracking-tight">Novo Aviso Operacional</h3>
                      </div>
                      <button onClick={() => setIsCreateWarningModalOpen(false)}><X size={24}/></button>
                  </div>
                  <form onSubmit={handleAddWarning} className="p-8 space-y-4">
                      <div className="grid grid-cols-2 gap-4">
                          <div className="space-y-1">
                              <label className="text-[10px] font-black text-slate-400 uppercase">Operação</label>
                              <select 
                                required
                                value={newWarning.operacao}
                                onChange={e => setNewWarning({...newWarning, operacao: e.target.value})}
                                className="w-full p-3 border dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 dark:text-white text-[11px] font-bold outline-none"
                              >
                                  <option value="">Selecione...</option>
                                  {locations.map(loc => (
                                      <option key={loc} value={loc}>{loc}</option>
                                  ))}
                              </select>
                          </div>
                          <div className="space-y-1">
                              <label className="text-[10px] font-black text-slate-400 uppercase">Rota</label>
                              <input 
                                required
                                type="text"
                                placeholder="Ex: 001 ou ROTA X"
                                value={newWarning.rota}
                                onChange={e => setNewWarning({...newWarning, rota: e.target.value})}
                                className="w-full p-3 border dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 dark:text-white text-[11px] font-bold outline-none"
                              />
                          </div>
                      </div>
                      <div className="space-y-1">
                          <label className="text-[10px] font-black text-slate-400 uppercase">Data da Ocorrência</label>
                          <input 
                            required
                            type="date"
                            value={newWarning.dataOcorrencia}
                            onChange={e => setNewWarning({...newWarning, dataOcorrencia: e.target.value})}
                            className="w-full p-3 border dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 dark:text-white text-[11px] font-bold outline-none"
                          />
                      </div>
                      <div className="space-y-1">
                          <label className="text-[10px] font-black text-slate-400 uppercase">Descrição da Ocorrência</label>
                          <textarea 
                            required
                            rows={4}
                            placeholder="Descreva detalhadamente o ocorrido..."
                            value={newWarning.descricao}
                            onChange={e => setNewWarning({...newWarning, descricao: e.target.value})}
                            className="w-full p-4 border dark:border-slate-700 rounded-xl bg-slate-50 dark:bg-slate-800 dark:text-white text-[11px] font-medium outline-none resize-none"
                          />
                      </div>
                      <div className="pt-4 flex gap-3">
                          <button type="button" onClick={() => setIsCreateWarningModalOpen(false)} className="flex-1 py-3 bg-slate-100 dark:bg-slate-800 text-slate-500 rounded-xl font-black uppercase text-[10px] tracking-widest">Cancelar</button>
                          <button 
                            type="submit" 
                            disabled={isUpdating}
                            className="flex-[2] py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-black uppercase text-[10px] tracking-widest shadow-xl shadow-blue-500/30 flex items-center justify-center gap-2 transition-all active:scale-95 disabled:opacity-50"
                          >
                              {isUpdating ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                              Salvar no SharePoint
                          </button>
                      </div>
                  </form>
              </div>
          </div>
      )}

      {/* RESET MODAL */}
      {isResetModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4 animate-in fade-in duration-200">
             <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-md overflow-hidden border dark:border-slate-700">
                <div className="bg-amber-700 dark:bg-slate-800 text-white p-4 flex justify-between items-center">
                    <div className="flex items-center gap-2">
                        <div className="p-2 bg-amber-600 rounded-lg">
                            <RotateCcw size={20} />
                        </div>
                        <div>
                            <h3 className="font-bold text-lg">Resetar Checklist</h3>
                            <p className="text-[10px] text-amber-200 uppercase tracking-tighter">Snapshot será salvo no SharePoint</p>
                        </div>
                    </div>
                    <button onClick={() => setIsResetModalOpen(false)} className="hover:bg-white/10 p-1 rounded-full transition-colors">
                        <X size={24} />
                    </button>
                </div>
                <div className="p-6 bg-gray-50 dark:bg-slate-900">
                    <div className="mb-6">
                        <label className="block text-xs font-bold text-gray-500 dark:text-gray-400 uppercase mb-2">Quem está realizando o reset?</label>
                        
                        <div className="relative">
                            {isLoadingUsers ? (
                                <div className="flex items-center gap-2 p-3 bg-white dark:bg-slate-800 border dark:border-slate-700 rounded-xl text-slate-400 text-sm italic">
                                    <Loader2 size={16} className="animate-spin" /> Buscando nomes cadastrados...
                                </div>
                            ) : registeredUsers.length > 0 ? (
                                <select 
                                    value={resetResponsible}
                                    onChange={(e) => setResetResponsible(e.target.value)}
                                    className="w-full p-3 pr-10 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-sm dark:text-white focus:ring-2 focus:ring-amber-500 outline-none appearance-none font-bold shadow-sm"
                                    autoFocus
                                >
                                    <option value="">Selecione seu nome...</option>
                                    {registeredUsers.map(name => (
                                        <option key={name} value={name}>{name}</option>
                                    ))}
                                </select>
                            ) : (
                                <div className="p-4 bg-red-50 dark:bg-red-900/20 border border-red-100 dark:border-red-900 rounded-xl text-red-600 dark:text-red-400 text-xs flex flex-col gap-2">
                                    <div className="flex items-center gap-2 font-bold uppercase">
                                        <AlertCircle size={16} /> Dados não encontrados
                                    </div>
                                    <p className="opacity-80">Não foi possível carregar a lista de usuários da tabela 'Usuarios_cco'. Verifique se existem registros no SharePoint.</p>
                                </div>
                            )}
                            {!isLoadingUsers && registeredUsers.length > 0 && (
                                <div className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none">
                                    <ChevronDown size={18} />
                                </div>
                            )}
                        </div>
                    </div>

                    <div className="flex gap-3 mt-8">
                        <button onClick={() => setIsResetModalOpen(false)} className="flex-1 py-3 bg-gray-200 dark:bg-slate-700 text-gray-700 dark:text-gray-200 font-bold rounded-xl hover:bg-gray-300 dark:hover:bg-slate-600 transition-colors">Cancelar</button>
                        <button 
                            onClick={handleResetChecklist} 
                            disabled={!resetResponsible.trim() || isUpdating || isLoadingUsers} 
                            className="flex-[2] py-3 bg-amber-600 text-white font-bold rounded-xl shadow-lg flex items-center justify-center gap-2 disabled:opacity-50 transition-all hover:bg-amber-700 active:scale-95"
                        >
                            {isUpdating ? <Loader2 size={20} className="animate-spin" /> : <Save size={20} />}
                            Confirmar Reset
                        </button>
                    </div>
                    
                    <div className="mt-4 text-center">
                        <p className="text-[10px] text-slate-400 font-medium italic">
                            O histórico será vinculado ao acesso corporativo logado.
                        </p>
                    </div>
                </div>
             </div>
        </div>
      )}

      {/* MAIN TABLE AREA */}
      <div className="flex-1 overflow-auto bg-slate-100 dark:bg-slate-950 transition-colors duration-500 scrollbar-thin">
        {tasks.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center p-12 text-center">
                <Database size={48} className="text-blue-600 mb-6 opacity-20"/>
                <h3 className="text-lg font-black dark:text-white mb-2">Nenhuma tarefa encontrada</h3>
            </div>
        ) : (
            <table className={`min-w-full border-separate border-spacing-0 select-none ${compact ? 'text-[10px]' : 'text-[11px]'}`}>
              <thead className="sticky top-0 z-[40]">
                <tr className="bg-blue-900 dark:bg-blue-950 text-white shadow-xl">
                  <th className="p-3 border-r border-blue-800 dark:border-blue-900 text-left sticky left-0 bg-blue-900 dark:bg-blue-950 z-[45] min-w-[350px] shadow-[4px_0_12px_-4px_rgba(0,0,0,0.4)] font-black uppercase tracking-widest text-[9px]">
                    Ação / Descrição da Tarefa
                  </th>
                  {locations.map(loc => (
                    <th key={loc} className="p-3 border-r border-blue-800 dark:border-blue-900 w-24 text-center font-bold">
                      {loc.replace('LAT-', '').replace('ITA-', '')}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(Object.entries(groupedTasks) as [string, Task[]][]).map(([cat, catTasks]) => {
                  const isCollapsed = collapsedCategories.includes(cat);
                  const { percent, isComplete } = getCategoryStats(cat);
                  
                  const canBeMinimized = isComplete || isCollapsed;
                  
                  return (
                    <React.Fragment key={cat}>
                      <tr 
                        className={`bg-blue-600 dark:bg-blue-900 text-white transition-colors h-10 group relative overflow-hidden cursor-pointer ${!canBeMinimized ? 'opacity-90' : 'hover:bg-blue-700'}`} 
                        onClick={() => toggleCategory(cat)}
                        title={!canBeMinimized ? "Finalize as tarefas para poder minimizar esta categoria" : "Clique para expandir/colapsar"}
                      >
                        <td colSpan={locations.length + 1} className="p-0 border-y border-blue-700 sticky left-0 z-30 overflow-hidden">
                          <div className={`absolute inset-y-0 left-0 transition-all duration-1000 pointer-events-none ${isComplete ? 'bg-green-500' : 'bg-blue-400'}`} style={{ width: `${percent}%` }} />
                          <div className="absolute inset-0 px-4 flex items-center justify-between z-10 pointer-events-auto">
                            <div className="flex items-center gap-3">
                              {isCollapsed ? <ChevronRight size={14}/> : <ChevronDown size={14}/>}
                              <span className="text-[10px] font-black uppercase tracking-widest">{cat}</span>
                            </div>
                            <span className="text-[9px] font-black bg-black/20 px-2 py-0.5 rounded-lg flex items-center gap-1">
                                {isComplete && <CheckCircle size={10} />}
                                {percent}% {isComplete ? '(FINALIZADO)' : ''}
                            </span>
                          </div>
                        </td>
                      </tr>
                      {!isCollapsed && catTasks.map(task => (
                        <tr key={task.id} className="bg-white dark:bg-slate-900 border-b border-slate-100 dark:border-slate-800/50 hover:bg-blue-50/30 dark:hover:bg-slate-800/50 transition-colors group">
                          <td 
                            className={`p-4 border-r border-slate-100 dark:border-slate-800 sticky left-0 bg-inherit z-30 shadow-[4px_0_8px_-4px_rgba(0,0,0,0.1)] transition-all ${activeTool ? 'cursor-crosshair hover:bg-blue-50 dark:hover:bg-blue-900/20' : ''}`}
                            onClick={() => handlePaintRow(task.id)}
                          >
                            <div className="flex flex-col gap-1.5">
                                <div className="font-bold text-slate-800 dark:text-slate-100 text-[13px] leading-tight">
                                    {task.title}
                                </div>
                                {task.description && (
                                  <div className="text-[11px] font-normal text-slate-500 dark:text-slate-400 leading-snug whitespace-pre-wrap opacity-90">
                                      {task.description}
                                  </div>
                                )}
                            </div>
                            {activeTool && <div className="absolute right-2 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100"><PaintBucket size={14} className="text-blue-500" /></div>}
                          </td>
                          {locations.map(loc => {
                            const status = task.operations[loc] || 'PR';
                            const cfg = STATUS_CONFIG[status];
                            return (
                              <td 
                                key={loc} 
                                className="p-0 border-r border-slate-100 dark:border-slate-800 h-12 relative" 
                                onMouseDown={() => {
                                    setIsDragging(true);
                                    onCellInteraction(task.id, loc);
                                    paintedThisDrag.current.add(`${task.id}-${loc}`);
                                }}
                                onMouseEnter={() => {
                                    if (isDragging && !paintedThisDrag.current.has(`${task.id}-${loc}`)) {
                                        onCellInteraction(task.id, loc, activeTool || undefined);
                                        paintedThisDrag.current.add(`${task.id}-${loc}`);
                                    }
                                }}
                              >
                                <div className={`absolute inset-[3px] rounded-lg flex items-center justify-center transition-all duration-200 font-black text-[10px] ${cfg.color} hover:brightness-95 active:scale-90 shadow-sm cursor-pointer`}>
                                  {cfg.label}
                                </div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
        )}
      </div>

      <div className="px-4 py-2 bg-white dark:bg-slate-900 border-t dark:border-slate-800 flex justify-between items-center text-[9px] text-slate-400 font-bold uppercase tracking-widest shrink-0">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
                <div className="w-1.5 h-1.5 rounded-full bg-green-500"></div>
                Operador: {currentUser.name}
            </div>
          </div>
          <div className="flex items-center gap-2">
              <HelpCircle size={10} className="text-blue-500"/> Atalhos: (1-6) Pintar, (Arraste) Pintura Contínua, (Título da Ação) Pintar Linha Toda
          </div>
      </div>
    </div>
  );
};

export default TaskManager;
