# Dashboard S&OP — Popular Pet

## Estrutura do repositório

```
sopp-dashboard/
├── index.html                        ← Dashboard (não editar)
├── data.json                         ← Dados (gerado automaticamente)
├── update.py                         ← Script de geração
├── planilha/
│   └── BaseHtmlv1.xlsx               ← Substituir aqui toda semana
└── .github/
    └── workflows/
        └── update.yml                ← Automação (GitHub Actions)
```

---

## Como atualizar o dashboard (toda semana)

### Opção A — Pelo GitHub (sem instalar nada)

1. Acesse o repositório no GitHub
2. Clique na pasta `planilha/`
3. Clique em **Add file → Upload files**
4. Arraste o novo Excel (substitui o anterior)
5. Clique em **Commit changes**
6. O GitHub Actions roda automaticamente e atualiza o `data.json`
7. Em ~60 segundos o dashboard está atualizado

### Opção B — Pelo terminal

```bash
# Substituir o Excel
cp novo_relatorio.xlsx planilha/BaseHtmlv1.xlsx

# Publicar
git add planilha/
git commit -m "S&OP semana 22"
git push
```
O GitHub Actions cuida do resto automaticamente.

---

## Configuração inicial (1x só)

### 1. Criar repositório no GitHub
1. Vá em github.com → **New repository**
2. Nome: `sopp-dashboard`
3. Visibilidade: **Public** (necessário para GitHub Pages gratuito)
4. Clique em **Create repository**

### 2. Fazer upload dos arquivos
Upload de todos os arquivos desta pasta mantendo a estrutura.

### 3. Ativar GitHub Pages
1. **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: **main** | Folder: **/ (root)**
4. **Save**

### 4. Ativar GitHub Actions
1. **Settings → Actions → General**
2. **Allow all actions** → **Save**

### 5. Acessar o dashboard
```
https://SEU_USUARIO.github.io/sopp-dashboard
```

---

## Como funciona o botão Atualizar

Ao clicar **Atualizar**, o dashboard busca o `data.json` mais recente do próprio GitHub Pages — mesmo domínio, sem CORS, sem erros.
