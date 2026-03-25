# 🔒 Sistema de Trava para Envio de Saídas e Não Coletas

## Problema Resolvido

Quando duas ou mais pessoas estavam com o sistema aberto simultaneamente, o envio **automático** de e-mail (quando todas as rotas de uma operação estão OK) era disparado múltiplas vezes, causando:
- Envio duplicado de e-mails
- Popup de "enviando" travado na tela
- Necessidade de F5 para fechar o popup

**Importante:** O envio de **Resumo Geral** é MANUAL e NÃO possui trava, pois é acionado explicitamente pelo usuário.

## Solução Implementada

Sistema de **trava distribuída** usando o próprio SharePoint como fonte da verdade, garantindo que apenas um usuário possa enviar o e-mail automático de cada operação por vez.

---

## ⚠️ Configuração Necessária no SharePoint

### Adicionar Campos na Lista `CONFIG_OPERACAO_SAIDA_DE_ROTAS`

Adicione os seguintes campos na lista do SharePoint:

| Nome Interno | Nome Exibido | Tipo | Descrição |
|--------------|--------------|------|-----------|
| `LockEnvio` | Lock Envio | Texto (única linha) | Indica se há trava ativa (`true`/`false`) |
| `LockUser` | Lock Usuário | Texto (única linha) | E-mail do usuário que adquiriu a trava |
| `LockTimestamp` | Lock Timestamp | Texto (única linha) | Data/hora ISO quando a trava foi adquirida |

### Como Adicionar os Campos

1. Acesse a lista `CONFIG_OPERACAO_SAIDA_DE_ROTAS` no SharePoint
2. Clique em **"Adicionar coluna"** → **"Texto"**
3. Crie as 3 colunas conforme tabela acima
4. **Importante:** Não torne os campos obrigatórios

---

## 🔄 Como Funciona

### Fluxo de Envio Automático (Saídas/Não Coletas)

1. **Sistema detecta que todas as rotas de uma operação estão OK** (polling automático)
2. **Verifica trava** para aquela operação específica
3. Se houver trava ativa (não expirada):
   - ❌ Envio é **bloqueado**
   - Mensagem alerta: "Outro usuário está enviando..."
4. Se não houver trava:
   - ✅ Sistema **adquire trava** para aquela operação
   - Envia e-mail automático via webhook
   - Atualiza SharePoint com último envio
   - **Libera trava** automaticamente

### Envio de Resumo Geral (MANUAL)

- **NÃO possui trava** - usuário clica explicitamente para enviar
- Envia todas as operações de uma vez
- Cada usuário pode enviar seu resumo quando quiser

### Timeout Automático

- Trava expira após **2 minutos** se não for liberada
- Previne travamento permanente em caso de erro/falha
- Timeout é verificado antes de bloquear novo envio

---

## 📊 Mensagens para o Usuário

### Sucesso (Envio Automático)
- ✅ "Enviado com sucesso!"

### Trava Ativa (Outro Usuário) - Envio Automático
- ⚠️ "⚠️ Já existe envio em andamento para [operacao] (envio em andamento por [usuario]). Aguarde alguns segundos e tente novamente."

### Falha ao Adquirir Trava - Envio Automático
- ⚠️ "⚠️ Não foi possível adquirir trava para [operacao]. Tente novamente em alguns segundos."

### Resumo Geral (Manual)
- ✅ "Enviado com sucesso!" (sem trava)

---

## 🛡️ Recursos de Segurança

1. **Validação de Propriedade**: Só adquire trava se operação pertencer ao usuário
2. **Identificação do Usuário**: Trava inclui e-mail de quem adquiriu
3. **Timestamp**: Permite detectar travas abandonadas
4. **Liberação Automática**: Trava é liberada mesmo em caso de erro
5. **Timeout de 2 Minutos**: Previne bloqueio permanente

---

## 🧪 Testes Recomendados

### Cenário 1: Envio Automático Normal
1. Todas as rotas de uma operação ficam OK
2. Sistema dispara envio automático
3. ✅ Envio ocorre normalmente
4. ✅ Trava é liberada após envio

### Cenário 2: Conflito Simultâneo (Envio Automático)
1. Todas as rotas ficam OK para dois usuários
2. **Simultaneamente**, ambos os sistemas tentam enviar
3. ✅ Usuário A adquire trava e envia
4. ✅ Usuário B recebe mensagem de bloqueio
5. Após liberação, Usuário B pode enviar (se ainda necessário)

### Cenário 3: Trava Abandonada (Erro)
1. Simule falha de rede durante envio automático
2. Aguarde 2 minutos
3. ✅ Trava expira automaticamente
4. ✅ Novo envio é permitido

### Cenário 4: Resumo Geral (Manual)
1. Dois usuários clicam em "Enviar Resumo" simultaneamente
2. ✅ Ambos os envios ocorrem normalmente (sem trava)
3. ✅ Cada usuário recebe confirmação individual

---

## 📝 Logs de Depuração

O sistema gera logs detalhados no console:

```
[LOCK_CHECK] Trava ativa para LAT-CWB por usuario@viagroup.com.br em 2026-03-24T14:30:00.000Z
[LOCK_ACQUIRE] Trava adquirida por usuario@viagroup.com.br para LAT-CWB em 2026-03-24T14:30:00.000Z
[LOCK_RELEASE] Trava liberada para LAT-CWB
[LOCK_CHECK] Trava expirada para LAT-CWB (usuário: usuario@viagroup.com.br, tempo: 125s)
```

---

## 🔧 Comandos do SharePointService

### Verificar Trava
```typescript
const lock = await SharePointService.checkSendLock(token, operacao);
// Retorna: { locked: boolean, user?: string, timestamp?: string, expired?: boolean }
```

### Adquirir Trava
```typescript
const result = await SharePointService.acquireSendLock(token, operacao, userEmail);
// Retorna: { success: boolean, message?: string }
```

### Liberar Trava
```typescript
await SharePointService.releaseSendLock(token, operacao);
```

---

## 🚨 Importante

- **NÃO remova as colunas de trava** da lista SharePoint
- **Mantenha o polling ativo** (5 segundos) para atualização em tempo real
- **Timeout de 2 minutos** pode ser ajustado em `sharepointService.ts` (linha ~1047)

---

**Desenvolvido com ❤️ para VIA Group**
