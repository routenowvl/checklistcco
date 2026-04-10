# TODO: Ajuste Tema Branco - CONTROLE DE SAÍDAS

## ✅ Planejado e Aprovado
- [x] Analisar RouteDeparture.tsx e App.tsx
- [x] Confirmar plano com usuário

## 🔄 Implementação (RouteDeparture.tsx)
- [ ] 1. Criar TODO.md
- [ ] 2. Atualizar container principal (bg-slate-100 → gradient white/slate-50)
- [ ] 3. Table container (bg-white → white/95 + backdrop-blur + shadow)
- [ ] 4. Header thead (bg-[#1e293b] → gradient slate-900/950)
- [ ] 5. Rows/inputs (slate-50/100 → white/80 + slate-200/50 borders + text-slate-900)
- [ ] 6. Modals/popups (containers + text)
- [ ] 7. Status badges/accents (crisp whites)
- [ ] 8. Toggle button + small polishes
- [ ] 9. Test white mode visibility
- [ ] 10. Marcar como completo

## ✅ Teste & Validação
- [ ] White mode: alta contraste/visibilidade
- [ ] Dark mode: inalterado
- [ ] UX fluida, hover states suaves

---

# ✅ Data Automática - NÃO COLETAS

## Implementado
- [x] Criar função `getNonCollectionDateForCurrentTime()` em dateUtils.ts
- [x] Atualizar ghost row (estado inicial + reset após adicionar)
- [x] Atualizar modal de adicionar (estado inicial + reset após adicionar)
- [x] Atualizar `createBulkRecordsWithOperation` (bulk paste)
- [x] Atualizar `handleBulkPaste` (novas linhas criadas)
- [x] Atualizar `handleAddNonCollection` (reset do modal)
- [x] Build aprovado

## Regra de Negócio
| Horário (Brasília) | Data Atribuída |
|-------------------|----------------|
| 00:00 - 11:59 | Dia anterior (D-1) |
| 12:00 - 23:59 | Dia atual (D) |

## Validação
- [ ] Testar inserção pela manhã (antes das 12h) — data deve ser do dia anterior
- [ ] Testar inserção à tarde/noite (após 12h) — data deve ser do dia atual
- [ ] Testar bulk paste com múltiplas linhas — todas com mesma data
- [ ] Testar modal de adicionar — data correta ao abrir
