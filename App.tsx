
import React, { useState, useEffect } from 'react';
import { HashRouter as Router, Routes, Route, useNavigate } from 'react-router-dom';
import { CheckSquare, History, Truck, Moon, Sun, LogOut, ChevronLeft, ChevronRight, Loader2, Search, LayoutDashboard, TowerControl } from 'lucide-react';
import TaskManager from './components/TaskManager';
import HistoryViewer from './components/HistoryViewer';
import RouteDepartureView from './components/RouteDeparture';
import SharePointExplorer from './components/SharePointExplorer';
import SendReportView from './components/SendReportView';
import Login from './components/Login';
import { SharePointService } from './services/sharepointService';
import { logout as msalLogout } from './services/authService';
import { Task, User } from './types';
import { setCurrentUser as setStorageUser } from './services/storageService';
import { getBrazilDate, getBrazilHours, getBrazilISOString, isAfter10amBrazil, getBrazilMinutes } from './utils/dateUtils';

const SidebarLink = ({ to, icon: Icon, label, active, collapsed }: any) => (
  <a href={`#${to}`} className={`flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${active ? 'bg-blue-600 text-white' : 'text-slate-500 hover:bg-slate-100'} ${collapsed ? 'justify-center' : ''}`}>
    <Icon size={20} />
    {!collapsed && <span className="font-medium whitespace-nowrap">{label}</span>}
  </a>
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

  const navigate = useNavigate();

  const handleLogin = (user: User) => {
    setUser(user);
    (window as any).__access_token = user.accessToken;
    loadDataFromSharePoint(user);
  };

  const loadDataFromSharePoint = async (user: User) => {
    if (!user.accessToken) {
      console.error('[APP] Token não encontrado');
      return;
    }

    (window as any).__access_token = user.accessToken;
    setIsLoading(true);

    try {
      setSyncMessage("Carregando Definições...");
      const spTasks = await SharePointService.getTasks(user.accessToken);
      const spOps = await SharePointService.getOperations(user.accessToken, user.email);
      const spMembers = await SharePointService.getTeamMembers(user.accessToken);
      setTeamMembers(spMembers);

      setSyncMessage("Sincronizando Matriz 1:1...");
      await SharePointService.ensureMatrix(user.accessToken, spTasks, spOps);

      setSyncMessage("Recuperando Status...");
      const today = getBrazilDate();
      const spStatus = await SharePointService.getStatusByDate(user.accessToken, today);

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

  useEffect(() => {
    if (!currentUser || tasks.length === 0) return;

    const checkAutoSaveTrigger = async () => {
      // Verifica se já passou das 10:00h no fuso de Brasília
      if (isAfter10amBrazil()) {
        const todayBrazil = getBrazilDate();
        const safeEmail = currentUser.email.replace(/[^a-zA-Z0-9]/g, '_');
        // Flag específica por dia no fuso brasileiro
        const autoSaveFlag = `auto_save_done_${safeEmail}_${todayBrazil}`;

        if (localStorage.getItem(autoSaveFlag) !== 'true') {
          console.log(`[AUTO_SAVE] Executando salvamento automático às ${getBrazilHours()}:${String(getBrazilMinutes()).padStart(2, '0')} (Brasília)`);
          try {
            await SharePointService.saveHistory(currentUser.accessToken!, {
              id: Date.now().toString(),
              timestamp: getBrazilISOString(),
              tasks: tasks,
              resetBy: 'Salvamento automático (10:00h)',
              email: currentUser.email
            });
            localStorage.setItem(autoSaveFlag, 'true');
            console.log('[AUTO_SAVE] Salvamento concluído com sucesso');
          } catch (e) {
            console.error("[AUTO_SAVE] Falha no backup automático:", e);
          }
        }
      }
    };

    // Verifica imediatamente e depois a cada minuto
    checkAutoSaveTrigger();
    const interval = setInterval(checkAutoSaveTrigger, 60000);
    return () => clearInterval(interval);
  }, [currentUser, tasks]);

  const handleLogout = async () => {
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

  if (!currentUser) return <Login onLogin={handleLogin} />;

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 overflow-hidden">
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
           <button onClick={() => setIsDarkMode(!isDarkMode)} className="p-2 w-full flex justify-center text-slate-500 hover:bg-slate-100 rounded-lg">
             {isDarkMode ? <Sun size={20}/> : <Moon size={20}/>}
           </button>
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
    </div>
  );
};

const App = () => (<Router><AppContent /></Router>);
export default App;
