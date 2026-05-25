# Plano: Histórico para usuário ALL + Filtro de Célula nos cards

## Contexto

Usuários do tipo "ALL" (viewer com acesso a todas as filiais) atualmente só veem "Saídas" e "Não Coletas". Precisam ter acesso ao Histórico, puxando dados de todas as filiais. Além disso, tanto nos painéis de saídas quanto no histórico, deve haver um filtro de célula (operação) próximo aos cards KPI.

---

## Etapa 1: `sharepointService.ts` — Retornar `isAllViewer` e adicionar `fetchAll` no `getHistory`

**Arquivo**: `services/sharepointService.ts`

### 1a. `getRouteConfigsByAccess` (linha 884)
- Adicionar `isAllViewer: boolean` ao retorno
- Linha 900 (editor): `return { configs: editableConfigs, canEdit: true, isAllViewer: false }`
- Linha 917 (ALL viewer): `return { configs: allConfigs, canEdit: false, isAllViewer: true }`
- Linha 939 (viewer normal): `return { configs: readableConfigs, canEdit: false, isAllViewer: false }`
- Linha 942 (catch): `return { configs: [], canEdit: false, isAllViewer: false }`

### 1b. `getHistory` (linha 601)
- Adicionar parâmetro `fetchAll?: boolean`
- Na linha 629, quando `fetchAll === true`, pular o filtro por email:
  ```
  .filter(record => fetchAll ? true : record.email?.toLowerCase() === userEmail.toLowerCase().trim())
  ```

---

## Etapa 2: `App.tsx` — Expor "Histórico" no sidebar + rota para ALL viewers

**Arquivo**: `App.tsx`

### 2a. Novo state (após os states existentes ~linha 86)
```tsx
const [isAllViewer, setIsAllViewer] = useState(false);
```

### 2b. Capturar flag (linha ~223, dentro de `loadDataFromSharePoint`)
```tsx
setIsAllViewer(routeAccess.isAllViewer);
```

### 2c. Sidebar (linha ~423) — Adicionar link Histórico para ALL viewers
Após o link "Não Coletas" (linha 423), antes do `{!isViewerOnly && (`:
```tsx
{isAllViewer && (
  <SidebarLink to="/history" icon={History} label="Histórico" active={window.location.hash === '#/history'} collapsed={collapsed} />
)}
```

### 2d. Rotas (linha ~471) — Adicionar rota `/history` para ALL viewers
Dentro do bloco `isViewerOnly`, antes do catch-all:
```tsx
{isAllViewer && (
  <Route path="/history" element={<HistoryViewer currentUser={currentUser} viewerMode="saidas" />} />
)}
```

---

## Etapa 3: `HistoryViewer.tsx` — Suportar modo ALL viewer + filtro de célula

**Arquivo**: `components/HistoryViewer.tsx`

### 3a. Props — Adicionar `viewerMode`
```tsx
interface HistoryViewerProps {
    currentUser: AppUser;
    viewerMode?: 'saidas';
}
```

### 3b. `fetchHistory` (~linha 54)
- Trocar `getRouteConfigs` por `getRouteConfigsByAccess` para obter configs corretas (ALL viewer não tem email em nenhuma config):
  ```tsx
  const routeAccess = await SharePointService.getRouteConfigsByAccess(token, currentUser.email);
  setUserConfigs(routeAccess.configs);
  ```
- Passar `fetchAll` no `getHistory`:
  ```tsx
  const data = await SharePointService.getHistory(token, currentUser.email, viewerMode === 'saidas');
  ```

### 3c. Filtro de célula (novo state + UI)
- Novo state: `const [celulaFilter, setCelulaFilter] = useState<string>('todas');`
- Derivar opções: `celulaOptions` memoizado de `userConfigs` (`['todas', ...unique ops]`)
- Aplicar filtro nos dados exibidos (no `processedHistory` ou `filteredAllowedLocations`)
- UI: `<select>` dropdown posicionado no header (após o título "Histórico"), seguindo estilo existente com classes Tailwind (bg-white/dark:bg-slate-900, border, rounded, text-xs, font-bold)

---

## Etapa 4: `RouteDeparture.tsx` — Filtro de célula nos cards KPI

**Arquivo**: `components/RouteDeparture.tsx`

### 4a. Novo state (~linha 620)
```tsx
const [celulaFilter, setCelulaFilter] = useState<string>('todas');
```

### 4b. Derivar opções
```tsx
const celulaOptions = useMemo(() => {
  const ops = Array.from(new Set(userConfigs.map(c => c.operacao).filter(Boolean))).sort();
  return ['todas', ...ops];
}, [userConfigs]);
```

### 4c. Aplicar filtro em `filteredRoutes` (~linha 3902)
Após filtro existente de `myOps`, adicionar:
```tsx
if (celulaFilter !== 'todas') {
  result = result.filter(r => r.operacao === celulaFilter);
}
```
Adicionar `celulaFilter` ao array de dependências do useMemo.

### 4d. Aplicar filtro nos indicadores de performance (~linha 3969)
No `useEffect`, filtrar `allUserRoutes` por `celulaFilter`:
```tsx
const allUserRoutes = routes.filter(r => {
  if (myOps.size === 0) return false;
  if (!myOps.has(r.operacao)) return false;
  if (celulaFilter !== 'todas' && r.operacao !== celulaFilter) return false;
  return true;
});
```
Adicionar `celulaFilter` ao array de dependências.

### 4e. UI — Dropdown após os cards KPI (~linha 4070)
Dentro do div `flex items-center gap-3 ml-8` (linha 4041), após o último card (linha 4069), adicionar:
```tsx
<select
  value={celulaFilter}
  onChange={(e) => setCelulaFilter(e.target.value)}
  className={`ml-3 px-3 py-2 rounded-xl text-[10px] font-bold uppercase outline-none cursor-pointer ${isDarkMode ? 'bg-slate-800 text-slate-200 border border-slate-700' : 'bg-white text-slate-700 border border-slate-300'}`}
>
  {celulaOptions.map(op => (
    <option key={op} value={op}>
      {op === 'todas' ? 'Todas as Células' : op}
    </option>
  ))}
</select>
```

### 4f. Filtro de célula no modal de histórico (History Modal)
- Novo state: `historyCelulaFilter` (resetar ao abrir o modal)
- Aplicar no `filteredArchivedResults`
- Dropdown no header do modal, próximo aos cards "Desempenho do Período"

---

## Verificação

1. Login com usuário editor → sidebar mostra Histórico normalmente (sem mudança)
2. Login com usuário ALL viewer → sidebar mostra Saídas, Não Coletas e **Histórico**
3. ALL viewer acessa Histórico → vê dados de todas as filiais (não filtrado por email)
4. ALL viewer NÃO pode editar nada no histórico (canEdit = false)
5. Filtro de célula na tela de Saídas → filtra rotas e recalcula KPIs
6. Filtro de célula no modal de histórico → filtra resultados arquivados
7. Filtro de célula no HistoryViewer → filtra os dados exibidos
8. Selecionar "Todas as Células" volta ao comportamento original
