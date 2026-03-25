import React, { useState, useEffect, useRef } from 'react';
import { HashRouter as Router, Routes, Route, useNavigate } from 'react-router-dom';
import { CheckSquare, History, Truck, LogOut, ChevronLeft, ChevronRight, Loader2, Search, LayoutDashboard, TowerControl, RefreshCw, AlertTriangle } from 'lucide-react';
import TaskManager from './components/TaskManager';
import HistoryViewer from './components/HistoryViewer';
import RouteDepartureView from './components/RouteDeparture';
import SharePointExplorer from './components/SharePointExplorer';
import SendReportView from './components/SendReportView';
import Login from './components/Login';
import PWAInstallPrompt from './components/PWAInstallPrompt';
import { SharePointService } from './services/sharepointService';
import { logout as msalLogout } from './services/authService';
import { msalInstance } from './services/authService';
import { startTokenRefreshLoop, stopTokenRefreshLoop, clearTokenState, getValidToken } from './services/tokenService';
import { Task, User } from './types';
import { setCurrentUser as setStorageUser } from './services/storageService';
import { getBrazilDate, getBrazilHours, getBrazilISOString, isAfter10amBrazil, getBrazilMinutes } from './utils/dateUtils';
import { InteractionRequiredAuthError } from '@azure/msal-browser';

const SCOPES = ["User.Read", "Sites.ReadWrite.All"];

const SidebarLink = ({ to, icon: Icon, label, active, collapsed }: any) => (
  <a href={`#${to}`} className={`flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${active ? 'bg-blue-600 text-white' : 'text-slate-500 hover:bg-slate-100'} ${collapsed ? 'justify-center' : ''}`}>
    <Icon size={20} />
    {!collapsed && <span className="font-medium whitespace-nowrap">{label}</span>}
  </a>
);

// Modal exibido quando a sessão expira e não é possível renovar silenciosamente
const SessionExpiredModal: React.FC<{ onRenew: () => void; isRenewing: boolean }> = ({ onRenew, isRenewing }) => (
  <div className="fixed inset-0 z-[9999] bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4 animate-in fade-in duration-300">
    <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-10 w-full max-w-sm border border-amber-500/50 shadow-2xl flex flex-col items-center gap-6 text-center animate-in zoom-in duration-300">
      <div className="w-16 h-16 rounded-full bg-amber-100 dark:bg-amber-900/30 flex items-center justify-center">
        <AlertTriangle size={32} className="text-amber-500" />
      </div>
      <div>
        <h3 className="text-lg font-black uppercase text-slate-800 dark:text-white tracking-tight">Sessão Expirada</h3>
        <p className="text-sm text-slate-500 dark:text-slate-400 mt-2 font-medium">
          Sua sessão Microsoft expirou. Clique abaixo para renovar sem perder seu trabalho.
        </p>
      </div>
      <button
        onClick={onRenew}
        disabled={isRenewing}
        className="w-full py-4 bg-blue-600 hover:bg-blue-700 text-white rounded-2xl font-black uppercase text-[11px] tracking-widest flex items-center justify-center gap-3 transition-all disabled:opacity-60 shadow-lg shadow-blue-500/20"
      >
        {isRenewing
          ? <><Loader2 size={18} className="animate-spin" /> Renovando...</>
          : <><RefreshCw size={18} /> Renovar Sessão</>
        }
      </button>
      <p className="text-[9px] text-slate-400 font-bold uppercase tracking-widest">
        Seus dados não serão perdidos
      </p>
    </div>
  </div>
);

const AppContent = () => {
  const [currentUser, setUser] = useState<User | null>(null);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [locations, setLocations] = useState<string[]>([]);
  const [teamMembers, setTeamMembers] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [syncMessage, setSyncMessage] = useState("Iniciando...");
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [collapsed, setCollapsed] = useState(true);
  const [collapsedCategories, setCollapsedCategories] = useState<string[]>([]);

  // Estado do modal de sessão expirada
  const [sessionExpired, setSessionExpired] = useState(false);
  const [isRenewing, setIsRenewing] = useState(false);
  const lastTokenErrorRef = useRef<number>(0);

  // Ref para a função de cleanup do loop de refresh
  const stopRefreshLoopRef = useRef<(() => void) | null>(null);

  const navigate = useNavigate();

  // Chamado quando o loop detecta que a sessão não pode ser renovada silenciosamente
  const handleSessionExpired = () => {
    // Debounce: só exibe modal se não houve erro nos últimos 10 segundos
    const now = Date.now();
    if (now - lastTokenErrorRef.current < 10000) {
      console.warn('[APP] Sessão expirada ignorada (debounce)');
      return;
    }
    
    // Só exibe se já não estiver exibindo
    if (sessionExpired) {
      console.warn('[APP] Sessão expirada ignorada (modal já aberto)');
      return;
    }

    lastTokenErrorRef.current = now;
    console.warn('[APP] Sessão expirada — exibindo modal de renovação');
    setSessionExpired(true);
  };

  // Renova a sessão via popup (sem perder estado)
  const handleRenewSession = async () => {
    setIsRenewing(true);
    try {
      const accounts = msalInstance.getAllAccounts();
      const account = accounts[0];

      let response;
      try {
        // Tenta primeiro sem UI (pode funcionar se o cookie de sessão ainda é válido)
        response = await msalInstance.acquireTokenSilent({
          scopes: SCOPES,
          account,
          forceRefresh: true,
        });
      } catch {
        // Fallback: popup de login (prompt: none = sem tela de seleção se já logado)
        response = await msalInstance.acquireTokenPopup({
          scopes: SCOPES,
          account,
          prompt: 'none',
        }).catch(() =>
          // Último recurso: popup com seleção de conta
          msalInstance.acquireTokenPopup({ scopes: SCOPES })
        );
      }

      if (response?.accessToken) {
        // Atualiza window E estado (renovação manual requer atualização do estado)
        (window as any).__access_token = response.accessToken;
        setUser(prev => prev ? { ...prev, accessToken: response.accessToken } : prev);
        setSessionExpired(false);
        lastTokenErrorRef.current = 0; // Reset debounce após renovação bem-sucedida
        console.log('[APP] ✅ Sessão renovada com sucesso');
      }
    } catch (err: any) {
      console.error('[APP] Falha ao renovar sessão:', err.message);
      // Se não conseguiu renovar de jeito nenhum, faz logout limpo
      await handleLogout();
    } finally {
      setIsRenewing(false);
    }
  };

  const handleLogin = (user: User) => {
    setUser(user);
    (window as any).__access_token = user.accessToken;
    setSessionExpired(false);
    lastTokenErrorRef.current = Date.now(); // Previne modal imediato após login

    // Inicia o loop de refresh proativo (background — sem re-renderização)
    if (stopRefreshLoopRef.current) stopRefreshLoopRef.current();
    stopRefreshLoopRef.current = startTokenRefreshLoop(handleSessionExpired);

    loadDataFromSharePoint(user);
  };

  const loadDataFromSharePoint = async (user: User) => {
    // Sempre pega o token mais fresco disponível
    const token = await getValidToken() || user.accessToken;
    if (!token) {
      console.error('[APP] Token não encontrado');
      return;
    }

    (window as any).__access_token = token;
    setIsLoading(true);

    try {
      setSyncMessage("Carregando Definições...");
      const spTasks = await SharePointService.getTasks(token);
      const spOps = await SharePointService.getOperations(token, user.email);
      const spMembers = await SharePointService.getTeamMembers(token);
      setTeamMembers(spMembers);

      setSyncMessage("Sincronizando Matriz 1:1...");
      await SharePointService.ensureMatrix(token, spTasks, spOps);

      setSyncMessage("Recuperando Status...");
      const today = getBrazilDate();
      const spStatus = await SharePointService.getStatusByDate(token, today);

      const opSiglas = spOps.map(o => o.Title);
      setLocations(opSiglas);

      const matrixTasks: Task[] = spTasks.map(t => {
        const ops: Record<string, any> = {};
        opSiglas.forEach(sigla => {
          const statusMatch = spStatus.find(s => s.TarefaID === t.id && s.OperacaoSigla === sigla);
          ops[sigla] = statusMatch ? statusMatch.Status : 'PR';
        });

        return {
          id: t.id,
          title: t.Title,
          description: t.Descricao,
          category: t.Categoria,
          timeRange: t.Horario,
          operations: ops,
          createdAt: new Date().toISOString(),
          isDaily: true,
          active: t.Ativa
        };
      });

      setTasks(matrixTasks.filter(t => t.active !== false));
    } catch (err: any) {
      console.error("[APP] Erro ao carregar SharePoint:", err.message);
      setSyncMessage("Erro na sincronização");
    } finally {
      setIsLoading(false);
    }
  };

  // Auto-save às 10:00h (Brasília)
  useEffect(() => {
    if (!currentUser || tasks.length === 0) return;

    const checkAutoSaveTrigger = async () => {
      if (isAfter10amBrazil()) {
        const todayBrazil = getBrazilDate();
        const safeEmail = currentUser.email.replace(/[^a-zA-Z0-9]/g, '_');
        const autoSaveFlag = `auto_save_done_${safeEmail}_${todayBrazil}`;

        if (localStorage.getItem(autoSaveFlag) !== 'true') {
          console.log(`[AUTO_SAVE] Executando às ${getBrazilHours()}:${String(getBrazilMinutes()).padStart(2, '0')} (Brasília)`);
          try {
            // Usa sempre o token mais fresco
            const token = await getValidToken() || currentUser.accessToken!;
            await SharePointService.saveHistory(token, {
              id: Date.now().toString(),
              timestamp: getBrazilISOString(),
              tasks: tasks,
              resetBy: 'Salvamento automático (10:00h)',
              email: currentUser.email
            });
            localStorage.setItem(autoSaveFlag, 'true');
            console.log('[AUTO_SAVE] Concluído com sucesso');
          } catch (e) {
            console.error("[AUTO_SAVE] Falha:", e);
          }
        }
      }
    };

    checkAutoSaveTrigger();
    const interval = setInterval(checkAutoSaveTrigger, 60000);
    return () => clearInterval(interval);
  }, [currentUser, tasks]);

  // Escuta o evento token-expired disparado pelo sharepointService (com debounce)
  useEffect(() => {
    const onTokenExpired = () => {
      // Debounce: ignora eventos repetidos dentro de 10 segundos
      const now = Date.now();
      if (now - lastTokenErrorRef.current < 10000) {
        console.warn('[EVENT_LISTENER] token-expired ignorado (debounce)');
        return;
      }
      handleSessionExpired();
    };
    window.addEventListener('token-expired', onTokenExpired);
    return () => window.removeEventListener('token-expired', onTokenExpired);
  }, [sessionExpired]);

  const handleLogout = async () => {
    // Para o loop de refresh antes de deslogar
    if (stopRefreshLoopRef.current) {
      stopRefreshLoopRef.current();
      stopRefreshLoopRef.current = null;
    }
    clearTokenState();

    await msalLogout();
    setUser(null);
    setStorageUser(null);
    delete (window as any).__access_token;
    navigate('/');
  };

  useEffect(() => {
    if (isDarkMode) document.documentElement.classList.add('dark');
    else document.documentElement.classList.remove('dark');
  }, [isDarkMode]);

  // Cleanup do loop ao desmontar o componente
  useEffect(() => {
    return () => {
      if (stopRefreshLoopRef.current) stopRefreshLoopRef.current();
    };
  }, []);

  if (!currentUser) return <Login onLogin={handleLogin} />;

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 overflow-hidden">

      {/* Modal de sessão expirada — sobrepõe tudo */}
      {sessionExpired && (
        <SessionExpiredModal onRenew={handleRenewSession} isRenewing={isRenewing} />
      )}

      <aside className={`bg-white dark:bg-slate-900 border-r dark:border-slate-800 transition-all ${collapsed ? 'w-20' : 'w-64'} p-4 flex flex-col`}>
        <div className="mb-10 flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold">V</div>
          {!collapsed && <h1 className="font-bold dark:text-white text-sm">CCO Digital</h1>}
        </div>
        <nav className="flex-1 space-y-2">
          <SidebarLink to="/" icon={CheckSquare} label="Checklist" active={window.location.hash === '#/'} collapsed={collapsed} />
          <SidebarLink to="/departures" icon={Truck} label="Saídas" active={window.location.hash === '#/departures'} collapsed={collapsed} />
          <SidebarLink to="/resumo" icon={TowerControl} label="Resumo" active={window.location.hash === '#/resumo'} collapsed={collapsed} />
          <SidebarLink to="/history" icon={History} label="Histórico" active={window.location.hash === '#/history'} collapsed={collapsed} />
          <SidebarLink to="/explorer" icon={Search} label="Explorador" active={window.location.hash === '#/explorer'} collapsed={collapsed} />
        </nav>
        <div className="mt-auto space-y-2 border-t pt-4 dark:border-slate-800">
           <button onClick={() => setCollapsed(!collapsed)} className="p-2 w-full flex justify-center text-slate-500 hover:bg-slate-100 rounded-lg">
             {collapsed ? <ChevronRight size={20}/> : <ChevronLeft size={20}/>}
           </button>
        </div>
      </aside>

      <main className="flex-1 overflow-hidden p-4">
        {isLoading ? (
          <div className="h-full flex items-center justify-center flex-col gap-4 text-blue-600">
             <Loader2 size={40} className="animate-spin" />
             <p className="font-bold animate-pulse text-sm uppercase tracking-widest">{syncMessage}</p>
          </div>
        ) : (
          <Routes>
            <Route path="/" element={
              <TaskManager
                tasks={tasks}
                setTasks={setTasks}
                locations={locations}
                setLocations={setLocations}
                onUserSwitch={() => loadDataFromSharePoint(currentUser)}
                collapsedCategories={collapsedCategories}
                setCollapsedCategories={setCollapsedCategories}
                currentUser={currentUser}
                onLogout={handleLogout}
                teamMembers={teamMembers}
              />
            } />
            <Route path="/departures" element={<RouteDepartureView currentUser={currentUser} />} />
            <Route path="/resumo" element={<SendReportView currentUser={currentUser} />} />
            <Route path="/history" element={<HistoryViewer currentUser={currentUser} />} />
            <Route path="/explorer" element={<SharePointExplorer currentUser={currentUser} />} />
          </Routes>
        )}
      </main>
      <PWAInstallPrompt />
    </div>
  );
};

const App = () => (<Router><AppContent /></Router>);
export default App;
