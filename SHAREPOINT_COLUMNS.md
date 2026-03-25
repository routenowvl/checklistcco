# 📋 Script de Criação de Colunas - SharePoint

## Lista: CONFIG_OPERACAO_SAIDA_DE_ROTAS

### Colunas Necessárias para Sistema de Trava

---

## Método 1: Interface Web (Manual)

1. Acesse o site do SharePoint onde está a lista `CONFIG_OPERACAO_SAIDA_DE_ROTAS`
2. Clique em **"Configurações"** (engrenagem) → **"Adicionar coluna"**
3. Para cada coluna abaixo:

### Coluna 1: LockEnvio
- **Nome:** LockEnvio
- **Nome Exibido:** Lock Envio
- **Tipo:** Texto (única linha)
- **Obrigatório:** Não
- **Descrição:** Indica se há trava ativa para envio (true/false)

### Coluna 2: LockUser
- **Nome:** LockUser
- **Nome Exibido:** Lock Usuário
- **Tipo:** Texto (única linha)
- **Obrigatório:** Não
- **Descrição:** E-mail do usuário que adquiriu a trava

### Coluna 3: LockTimestamp
- **Nome:** LockTimestamp
- **Nome Exibido:** Lock Timestamp
- **Tipo:** Texto (única linha)
- **Obrigatório:** Não
- **Descrição:** Data/hora ISO quando a trava foi adquirida

---

## Método 2: PowerShell PnP (Automatizado)

```powershell
# Conectar ao SharePoint
Connect-PnPOnline -Url "https://vialacteoscombr.sharepoint.com/sites/CCO" -Interactive

# Adicionar colunas na lista CONFIG_OPERACAO_SAIDA_DE_ROTAS
$list = Get-PnPList -Identity "CONFIG_OPERACAO_SAIDA_DE_ROTAS"

# Coluna LockEnvio
Add-PnPField -List $list -DisplayName "Lock Envio" -InternalName "LockEnvio" -Type Text -AddToDefaultView

# Coluna LockUser
Add-PnPField -List $list -DisplayName "Lock Usuário" -InternalName "LockUser" -Type Text -AddToDefaultView

# Coluna LockTimestamp
Add-PnPField -List $list -DisplayName "Lock Timestamp" -InternalName "LockTimestamp" -Type Text -AddToDefaultView

Write-Host "Colunas criadas com sucesso!" -ForegroundColor Green
```

---

## Método 3: Microsoft Graph API (Programático)

### Requisição HTTP

```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-id}/columns
Authorization: Bearer {token}
Content-Type: application/json

{
  "name": "LockEnvio",
  "displayName": "Lock Envio",
  "description": "Indica se há trava ativa para envio (true/false)",
  "columnGroup": "Colunas de Trava",
  "text": {
    "allowMultipleLines": false,
    "isPlainText": true
  }
}
```

Repita para `LockUser` e `LockTimestamp`.

---

## Validação Após Criação

1. Acesse a lista `CONFIG_OPERACAO_SAIDA_DE_ROTAS`
2. Clique em **"Configurações"** → **"Configurações da lista"**
3. Verifique se as 3 colunas estão listadas
4. Adicione as colunas à visualização padrão para facilitar depuração

---

## Visualização Recomendada

Crie uma nova visualização chamada **"Com Trava"** incluindo:

- OPERACAO
- EMAIL
- UltimoEnvioSaida
- Status
- **LockEnvio** ← Nova
- **LockUser** ← Nova
- **LockTimestamp** ← Nova
- UltimoEnvioResumoSaida
- StatusResumoSaida

---

## ⚠️ Importante

- **NÃO** torne as colunas obrigatórias
- **NÃO** remova as colunas após criação
- Mantenha os **nomos internos** exatamente como especificado
- Teste em ambiente de homologação antes de produção

---

**Desenvolvido com ❤️ para VIA Group**
**Data:** 24 de março de 2026
