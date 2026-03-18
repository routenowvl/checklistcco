# 📱 Instalação do PWA - Checklist CCO

O **Checklist CCO** agora é um **PWA (Progressive Web App)**! Isso significa que você pode instalá-lo no seu computador como um aplicativo nativo, sem precisar de loja de aplicativos.

---

## ✨ Benefícios

- ✅ **Acesso rápido** diretamente do desktop/menu iniciar
- ✅ **Funciona offline** (parcialmente - cache de recursos)
- ✅ **Sem barra de navegador** - experiência de app nativo
- ✅ **Atualizações automáticas** em segundo plano
- ✅ **Ícone próprio** na área de trabalho

---

## 🚀 Como Instalar

### **Google Chrome / Edge (Windows/Mac/Linux)**

1. Abra o Checklist CCO no navegador
2. Aguarde alguns segundos - um **popup aparecerá no canto inferior direito**
3. Clique em **"Instalar"**
4. O aplicativo será instalado automaticamente

**Ou manualmente:**

- **Chrome:** Clique no ícone de **instalar** (⬇️) na barra de endereço
- **Edge:** Clique em **⋯** → **Aplicativos** → **Instalar este aplicativo**

### **Após a instalação:**

- O app abrirá em uma janela separada (sem barra de navegador)
- O ícone aparecerá na sua área de trabalho
- No Windows: Disponível no Menu Iniciar
- No Mac: Disponível na pasta Aplicativos e Dock

---

## 📋 Pré-requisitos

- **HTTPS** (ou localhost para desenvolvimento)
- Navegador compatível:
  - ✅ Chrome 67+
  - ✅ Edge 79+
  - ✅ Safari 11.1+ (iOS 11.3+)
  - ✅ Firefox 68+

---

## 🔄 Atualizações

O PWA se atualiza **automaticamente** em segundo plano. Sempre que você abrir o aplicativo, ele verificará se há novas versões.

---

## 🗑️ Como Desinstalar

### Windows:
1. Abra **Configurações** → **Aplicativos**
2. Encontre "Checklist CCO - VIA Group"
3. Clique em **Desinstalar**

### Mac:
1. Abra a pasta **Aplicativos**
2. Arraste "Checklist CCO" para a Lixeira

### Navegador:
1. Clique nos **⋯** (três pontos) do navegador
2. Vá em **Aplicativos** ou **Mais ferramentas**
3. Remova o aplicativo

---

## 🛠️ Desenvolvimento

### Comandos úteis:

```bash
# Desenvolvimento (com PWA habilitado)
npm run dev

# Build de produção
npm run build

# Preview da build
npm run preview
```

### Estrutura de arquivos PWA:

```
public/
├── pwa-192x192.png       # Ícone 192x192
├── pwa-512x512.png       # Ícone 512x512
├── apple-touch-icon.png  # Ícone Apple (180x180)
└── favicon.ico           # Favicon

dist/
├── manifest.webmanifest  # Manifesto PWA (gerado automaticamente)
├── sw.js                 # Service Worker (gerado automaticamente)
└── workbox-*.js          # Workbox (gerenciamento de cache)
```

### Personalização:

Para alterar configurações do PWA, edite `vite.config.ts`:

```typescript
VitePWA({
  manifest: {
    name: 'Checklist CCO - VIA Group',
    short_name: 'CCO',
    // ... outras configurações
  }
})
```

---

## ⚠️ Limitações

- **Funcionalidades offline:** Apenas recursos estáticos são cacheados. Dados do SharePoint requerem conexão.
- **Primeiro acesso:** Requer internet para carregar pela primeira vez.
- **Token de autenticação:** Sessão do Azure MSAL expira normalmente - requer login periódico.

---

## 📞 Suporte

Dúvidas ou problemas? Entre em contato com a equipe de desenvolvimento.

---

**Desenvolvido com ❤️ para VIA Group**
