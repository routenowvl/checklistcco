# ✅ Correção - Coluna SAÍDA (Controle de Saídas)

## Problema

Na guia **CONTROLE DE SAÍDAS**, ao preencher a coluna **SAÍDA** com data e horário completos (ex: `02/04/2026 14:30:00`), o sistema estava descartando a parte da data quando o usuário saía da célula, mantendo apenas o horário (`14:30:00`).

### Causa Raiz

O código tinha dois problemas:

1. **`onChange` chamava `formatTimeInput()`** - Esta função sempre retorna apenas `HH:MM:SS`, descartando qualquer data
2. **`displayValue` sempre mostrava apenas horário** - Mesmo quando o valor salvo continha data completa, a exibição mostrava apenas o horário, impedindo o usuário de ver/editar a data

## Solução Implementada

### 1. Adicionado estado para controlar edição
```typescript
const [editingSaidaCell, setEditingSaidaCell] = useState<string | null>(null);
```

### 2. Lógica de exibição inteligente
- **Quando NÃO está editando**: Se o valor tem data completa (`DD/MM/AAAA HH:MM:SS`), mostra apenas o horário para limpeza visual
- **Quando está editando**: Mostra o valor completo (com data) para permitir edição

### 3. Fluxo de interação
```
1. Usuário clica na célula → onFocus → mostra valor completo (com data)
2. Usuário digita → onChange → atualiza sem formatar (preserva data enquanto digita)
3. Usuário sai → onBlur → valida e formata:
   - Se digitou data completa (DD/MM/AAAA HH:MM:SS) → salva completo
   - Se digitou apenas horário → formata como HH:MM:SS
   - Se digitou "-" → salva "-"
   - Se vazio → limpa
```

### 4. Remoção do `formatTimeInput` no `onChange`
Agora o `onChange` apenas atualiza o valor bruto sendo digitado, sem formatação. A formatação ocorre apenas no `onBlur`, após o usuário terminar de digitar.

## Comportamento Atual

| Entrada do Usuário | Resultado Armazenado | Exibição (fora da edição) |
|-------------------|---------------------|--------------------------|
| `02/04/2026 14:30:00` | `02/04/2026 14:30:00` | `14:30:00` |
| `14:30:00` | `14:30:00` | `14:30:00` |
| `143000` | `14:30:00` | `14:30:00` |
| `14:30` | `14:30:00` | `14:30:00` |
| `-` | `-` | `-` |

## Arquivos Modificados

- `components/RouteDeparture.tsx`
  - Adicionado estado `editingSaidaCell`
  - Atualizada lógica de renderização da coluna `saida`
  - Adicionado `onFocus` para mostrar valor completo durante edição
  - Removida chamada do `formatTimeInput` no `onChange`

## Testes Recomendados

1. **Data completa**: Digite `02/04/2026 14:30:00` → Tab → Verifique se manteve a data
2. **Apenas horário**: Digite `14:30:00` → Tab → Verifique se formatou corretamente
3. **Atalho numérico**: Digite `143000` → Tab → Verifique se formatou para `14:30:00`
4. **Valor "-"**: Digite `-` → Tab → Verifique se manteve "-"
5. **Edição subsequente**: Clique em uma célula com data completa → Verifique se mostra a data completa para edição

---

**Corrigido em:** 02 de abril de 2026  
**Autor:** Qwen Code  
**Build:** ✅ Aprovado (vite build OK)
