# Gerador de Reports — a/b lab

Aplicação web para geração automática de reports mensais de mídia paga no template da a/b lab.

## Como fazer deploy no Vercel (5 minutos)

### 1. Suba o projeto no GitHub
- Crie um repositório no GitHub
- Faça upload de todos os arquivos desta pasta

### 2. Conecte ao Vercel
- Acesse [vercel.com](https://vercel.com) e faça login com o GitHub
- Clique em **Add New Project**
- Selecione o repositório criado
- Clique em **Deploy**

### 3. Configure a variável de ambiente
- No painel do Vercel, vá em **Settings → Environment Variables**
- Adicione:
  - **Name:** `ANTHROPIC_API_KEY`
  - **Value:** sua chave da API Anthropic (encontre em [console.anthropic.com](https://console.anthropic.com))
- Clique em **Save** e depois **Redeploy**

### 4. Compartilhe o link
- O Vercel vai gerar um link tipo `ablab-reports.vercel.app`
- Compartilhe com qualquer pessoa da agência — sem instalação necessária

---

## Como usar

1. Preencha nome da marca, período e meta
2. Informe os botões de agendamento do mês anterior
3. Faça upload dos 3 arquivos:
   - **Google Ads CSV** — exportar em: Campanhas → exportar → CSV
   - **Meta Ads XLSX** — exportar em: Gerenciador de Anúncios → Exportar → Excel
   - **GA4 CSV** — exportar o relatório de Conversões x Campanhas
4. Adicione contexto do mês (opcional)
5. Clique em **Gerar Report**
6. Baixe o PowerPoint pronto no template da a/b lab

---

## Desenvolvimento local

```bash
npm install
cp .env.example .env.local
# Preencha ANTHROPIC_API_KEY no .env.local
npm run dev
```

Acesse `http://localhost:3000`
