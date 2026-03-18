<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# Checklist CCO - VIA Group 📱

Sistema de gestão de operações logísticas em tempo real para o CCO (Centro de Controle Operacional).

> **Agora disponível como PWA!** Instale diretamente do navegador e tenha acesso rápido como um aplicativo nativo.

## 🚀 Executar Localmente

**Pré-requisitos:** Node.js 18+

1. Instale as dependências:
   ```bash
   npm install
   ```

2. Configure as variáveis de ambiente:
   - Copie `.env.example` para `.env`
   - Preencha com suas chaves (Gemini API, Azure AD, SharePoint)

3. Execute o app em modo de desenvolvimento:
   ```bash
   npm run dev
   ```

4. Acesse em `http://localhost:3000`

## 📱 Instalar como PWA

Após abrir o aplicativo no navegador:

1. **Aguarde alguns segundos** - Um popup aparecerá no canto inferior direito
2. **Clique em "Instalar"**
3. O app será instalado como um aplicativo nativo!

**Ou manualmente:**
- **Chrome/Edge:** Clique no ícone de instalar (⬇️) na barra de endereço
- **Menu:** `⋯` → `Aplicativos` → `Instalar este aplicativo`

Veja mais detalhes em [PWA.md](PWA.md)

## 🛠️ Scripts

| Comando | Descrição |
|---------|-----------|
| `npm run dev` | Inicia servidor de desenvolvimento |
| `npm run build` | Compila para produção (com PWA) |
| `npm run preview` | Visualiza build de produção |

## 📁 Estrutura do Projeto

```
checklist-cco/
├── components/       # Componentes React
├── services/         # Integrações (SharePoint, Auth, IA)
├── utils/            # Utilitários (datas, etc.)
├── public/           # Ícones PWA e assets estáticos
├── dist/             # Build de produção (gerado)
└── PWA.md            # Documentação completa do PWA
```

## 📄 Documentação

- [CONFIG.md](CONFIG.md) - Configuração de variáveis de ambiente
- [PWA.md](PWA.md) - Guia completo de instalação do PWA

---

**Desenvolvido com ❤️ para VIA Group**
