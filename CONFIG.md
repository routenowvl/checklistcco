# Configuração do Ambiente

## Variáveis de Ambiente

Este projeto utiliza variáveis de ambiente para configurar credenciais e endpoints. Siga os passos abaixo:

### 1. Copie o arquivo de exemplo

```bash
cp .env.example .env
```

### 2. Preencha as variáveis no arquivo `.env`

#### Google Gemini (IA)
```
GEMINI_API_KEY=sua_chave_aqui
```
Obtenha sua chave em: https://aistudio.google.com/apikey

#### Microsoft Azure AD (Autenticação)
```
VITE_AZURE_CLIENT_ID=seu_client_id_aqui
VITE_AZURE_TENANT_ID=seu_tenant_id_aqui
```
Estes valores são obtidos no Azure Portal > Azure Active Directory > App registrations

#### SharePoint Online
```
VITE_SHAREPOINT_SITE_PATH=seu_site_sharepoint_aqui
```
Exemplo: `vialacteoscombr.sharepoint.com:/sites/CCO`

#### Webhooks (n8n)
```
VITE_WEBHOOK_SAIDAS_URL=https://n8n.datastack.viagroup.com.br/webhook/seu_webhook_aqui
VITE_WEBHOOK_RESUMO_URL=https://n8n.datastack.viagroup.com.br/webhook/seu_webhook_aqui
```

#### Servidor de Desenvolvimento
```
VITE_SERVER_PORT=3000
```

### 3. Instale as dependências

```bash
npm install
```

### 4. Execute o projeto

```bash
npm run dev
```

## ⚠️ Importante

- **NUNCA** commit o arquivo `.env` no repositório
- O arquivo `.env` já está listado no `.gitignore`
- Use o `.env.example` como template para outros desenvolvedores
- As variáveis prefixadas com `VITE_` são expostas no código frontend

## Scripts Disponíveis

| Comando | Descrição |
|---------|-----------|
| `npm run dev` | Inicia o servidor de desenvolvimento |
| `npm run build` | Compila o projeto para produção |
| `npm run preview` | Visualiza a build de produção |
