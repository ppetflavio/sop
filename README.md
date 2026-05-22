# Dashboard S&OP — Popular Pet

## Estrutura do repositório

```
/
├── index.html    ← Dashboard (não editar)
├── data.json     ← Dados processados (atualizar semanalmente)
├── update.py     ← Script de atualização
└── README.md
```

## Como publicar no GitHub Pages

### 1ª vez (configuração)
1. Crie um repositório no GitHub (ex: `sopp-dashboard`)
2. Faça upload dos 3 arquivos: `index.html`, `data.json`, `update.py`
3. Vá em **Settings → Pages → Source: main branch → / (root)**
4. Acesse: `https://SEU_USUARIO.github.io/sopp-dashboard`

---

## Como atualizar os dados

### Requisitos
```bash
pip install pandas openpyxl numpy
```

### Passo a passo (toda semana)
```bash
# 1. Baixe a planilha do Google Sheets como .xlsx
# 2. Rode o script
python update.py BaseHtmlv1.xlsx

# 3. Publique no GitHub
git add data.json
git commit -m "S&OP semana $(date +%V/%Y)"
git push
```

O dashboard atualiza automaticamente em ~30 segundos após o push.

---

## Como funciona o botão Atualizar

- Clica em **Atualizar** → busca `data.json` do GitHub Pages
- Sem CORS, sem proxy, sem erros
- Todas as seções são atualizadas: semáforos, KPIs, charts, tabelas, mês seguinte, compradores

## Estrutura do data.json

O `data.json` contém os dados agregados por canal (Marketplace, Shopee, Amazon, Meli) com:
- KPIs mês atual e mês seguinte
- Top 100 SKUs por receita
- SKUs com necessidade de compra
- Compradores consolidados
- Críticos (zerado + acima da meta)
- Distribuição por categoria

**Gerado automaticamente pelo `update.py` a partir do Excel exportado do Google Sheets.**
