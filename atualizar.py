#!/usr/bin/env python3
"""
atualizar.py — Popular Pet S&OP
Lê o CSV exportado pelo XLSM e gera data.json + produtos.json para o GitHub Pages.

Uso:
    python atualizar.py planilha/BaseGeral_20260618_132740.csv
    python atualizar.py              ← busca automaticamente o .csv em planilha/
"""

import sys, os, json, glob
import pandas as pd
import numpy as np
from datetime import datetime

# ─────────────────────────────────────────────
# 1. LOCALIZAR O CSV
# ─────────────────────────────────────────────
def find_csv(arg=None):
    if arg and os.path.isfile(arg):
        return arg
    # busca automática dentro da pasta planilha/
    candidates = sorted(glob.glob("planilha/*.csv"), key=os.path.getmtime, reverse=True)
    if candidates:
        print(f"[INFO] CSV encontrado: {candidates[0]}")
        return candidates[0]
    raise FileNotFoundError("Nenhum CSV encontrado em planilha/. Forneça o caminho como argumento.")

csv_path = find_csv(sys.argv[1] if len(sys.argv) > 1 else None)

# ─────────────────────────────────────────────
# 2. LER E LIMPAR O CSV
# ─────────────────────────────────────────────
print(f"[INFO] Lendo: {csv_path}")
df = pd.read_csv(csv_path, sep=None, engine='python', dtype=str)

# Remove BOM da primeira coluna se presente
df.columns = [c.replace('\ufeff', '').strip() for c in df.columns]

# Função para converter string BR → float  (ex: "66,03 " → 66.03)
def br_float(series):
    return (
        series.astype(str)
              .str.strip()
              .str.replace(r'[%\s]', '', regex=True)   # remove % e espaços
              .str.replace('.', '', regex=False)         # remove sep milhar BR
              .str.replace(',', '.', regex=False)        # vírgula → ponto decimal
              .pipe(pd.to_numeric, errors='coerce')
              .fillna(0)
    )

# Colunas numéricas diretas (pandas já leu como número ou precisam de conversão)
NUM_COLS = [
    'Meta Atual (Q)', 'Meta Atual (C$)', 'Meta Atual ($)',
    'Acum Meta (Q)', 'Acum Realizado (Q)', 'Acum Meta ($)', 'Acum Realizado ($)', 'Acum Realizado (C$)',
    'Meta - Realizado (Q)',
    'Reprojeção (Q)', 'Reprojeção (C$)', 'Reprojeção ($)',
    'Necess. (Q) Atual.', 'Necess. (C$) Atual.', 'Necess. ($) Atual.',
    'Necess. (Q) Atual Rep', 'Necess. (C$) Atual Rep', 'Necess. ($) Atual Rep',
    'Reprojeção x Meta (Q)', 'Desvio Abaixo ($)', 'Desvio Acima ($)',
    'Reprojeção - Realizado (Q)',
    'Estoque L1 (Q)', 'Estoque L1 (C$)', 'Estoque L1 ($)',
    'Estoque Bling (Q)', 'Estoque Blin Ajust. (Q)', 'Estoque Bling (C$)', 'Estoque Bling ($)',
    'Estoque Full (Q)', 'Estoque Full (C$)', 'Estoque Full ($)',
    'Estoque Total (Q)', 'Estoque Total (C$)', 'Estoque Total ($)',
    'Estoque Mínimo (Q)', 'Est. Total x Est. Mínimo (Q)',
    'Estoque Meta (Q)', 'Estoque Meta (C$)', 'Estoque Meta ($)',
    'Pedidos Pendentes (Q)', 'Pedidos Pendentes (C$)', 'Pedidos Pendentes ($)',
    'Estoque Total + Ped. Pendente (Q)', 'Estoque Total + Ped. Pendente (C$)', 'Estoque Total + Ped. Pendente ($)',
    'Estoque Futuro 1 (Q)', 'Estoque Futuro 1  (C$)', 'Estoque Futuro 1  ($)',
    'Estoque Meta Seg (Q)', 'Estoque Meta Seg (C$)', 'Estoque Meta Seg ($)',
    'Meta Seguinte (Q)', 'Meta Seguinte (C$)', 'Meta Seguinte ($)',
    'Meta Seg. Nec (Q)', 'Meta Seg. Nec (C$)', 'Meta Seg. Nec ($)',
    'Estoque Futuro 2 (Q)', 'Estoque Futuro 2 (C$)', 'Estoque Futuro 2 ($)',
    'Reprojeção x Estoque (Q)', 'Reprojeção x Estoque (C$)', 'Reprojeção x Estoque ($)',
    'Reprojeção x Estoque Nec. (Q)', 'Reprojeção x Estoque Nec. (C$)', 'Reprojeção x Estoque Nec. ($)',
    'Ped. Pendente [Dig] (Q)', 'Ped. Pendente [Dig] ($)',
    'Ped. Pendente [999] (Q)', 'Ped. Pendente [999] (C$)',
    'Leadtime Médio (Dias)',
    'Calc. Venda diária', 'Calc. Dias Estoque',
    'Faturado XML Central (Q)',
    '999 Estoque (Q)', '999 Ped. Pendente (Q)', '999 Estoque + Ped Pendente (Q)',
    '999 Demanda Lojas 30d (Q)', '999 Disponível (Q)',
    'Nivel Estoque %', 'Nivel Estoque Rep %',
    'Empresa',
]

# Colunas monetárias em string BR (vírgula decimal)
STR_MONEY_COLS = ['Vlr. Custo ($)', 'Vlr. Venda ($)', 'Vlr Ticket ($)']

for c in NUM_COLS:
    if c in df.columns:
        df[c] = br_float(df[c])

for c in STR_MONEY_COLS:
    if c in df.columns:
        df[c] = br_float(df[c])

# Percentuais como string (ex: "4%", "26,1%") → float puro
for c in ['Margem  (%)', 'Impostos (%)', 'Meta %']:
    if c in df.columns:
        df[c] = (
            df[c].astype(str)
                 .str.strip()
                 .str.replace('%', '', regex=False)
                 .str.replace(',', '.', regex=False)
                 .pipe(pd.to_numeric, errors='coerce')
                 .fillna(0)
        )

# Empresa como int
df['Empresa'] = df['Empresa'].astype(int)

# Produto/SKU como string limpa
df['Produto'] = df['Produto'].astype(str).str.strip()

print(f"[INFO] {len(df)} linhas | {len(df.columns)} colunas | Empresas: {sorted(df['Empresa'].unique())}")

# ─────────────────────────────────────────────
# 3. HELPER: converter linha para dict JSON-safe
# ─────────────────────────────────────────────
def safe_val(v):
    if isinstance(v, float):
        if np.isnan(v) or np.isinf(v):
            return 0
        return round(v, 4)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return round(float(v), 4)
    if pd.isna(v) if not isinstance(v, (list, dict)) else False:
        return ""
    return v

# ─────────────────────────────────────────────
# 4. LÓGICA DE NEGÓCIO — Nec. Líquida Global
#    Pedidos pendentes são globais (entram na Empresa 1 e cobrem todos os canais)
#    Nec. Líquida = MAX(0, Nec.Compra − PP_global_por_SKU)
# ─────────────────────────────────────────────
# PP global por SKU = soma dos Pedidos Pendentes (Q) de TODAS as empresas daquele SKU
# (na prática só Empresa 1 tem PP, mas a regra é global)
pp_global = df.groupby('Produto')['Pedidos Pendentes (Q)'].sum().rename('PP_Global')
df = df.join(pp_global, on='Produto')

df['Nec. Liquida (Q)'] = df.apply(
    lambda r: max(0, r['Est. Total x Est. Mínimo (Q)'] - r['PP_Global']),
    axis=1
)

# ─────────────────────────────────────────────
# 5. RUPTURA — SKUs com estoque=0 / SKUs com venda>0
# ─────────────────────────────────────────────
def ruptura_stats(sub):
    """Retorna dict com métricas de ruptura para um subconjunto do df."""
    total_skus     = len(sub)
    skus_zero      = int((sub['Estoque Total (Q)'] == 0).sum())
    skus_com_venda = int((sub['Acum Realizado (Q)'] > 0).sum())
    skus_ruptura   = int(((sub['Estoque Total (Q)'] == 0) & (sub['Acum Realizado (Q)'] > 0)).sum())
    receita_risco  = round(float(
        sub.loc[(sub['Estoque Total (Q)'] == 0) & (sub['Acum Realizado (Q)'] > 0), 'Meta Atual ($)'].sum()
    ), 2)
    taxa = round(skus_ruptura / skus_com_venda * 100, 1) if skus_com_venda > 0 else 0.0
    return {
        "total_skus":     total_skus,
        "skus_zero":      skus_zero,
        "skus_com_venda": skus_com_venda,
        "skus_ruptura":   skus_ruptura,
        "taxa_ruptura":   taxa,
        "receita_risco":  receita_risco,
    }

# ─────────────────────────────────────────────
# 6. MAPEAMENTO DE CANAIS
# ─────────────────────────────────────────────
CANAIS = {
    "empresa1":   1,
    "shopee":   995,
    "amazon":   996,
    "meli":     997,
}

def build_canal(empresa_id):
    sub = df[df['Empresa'] == empresa_id].copy()
    rows = []
    for _, r in sub.iterrows():
        rows.append({
            # Identificação
            "sku":          safe_val(r['Produto']),
            "descricao":    safe_val(r.get('Descrição', '')),
            "empresa_sku":  safe_val(r.get('Empresa|Sku', '')),
            "categoria":    safe_val(r.get('Categoria', '')),
            "marca":        safe_val(r.get('Marca', '')),
            "comprador":    safe_val(r.get('Comprador', '')),
            "status":       safe_val(r.get('Status Compras', '')),
            "pareto":       safe_val(r.get('Pareto', '')),
            "meta_flag":    safe_val(r.get('Meta', '')),
            "full_flag":    safe_val(r.get('Full', '')),
            # Preços
            "vlr_custo":    safe_val(r['Vlr. Custo ($)']),
            "vlr_venda":    safe_val(r['Vlr. Venda ($)']),
            "vlr_ticket":   safe_val(r['Vlr Ticket ($)']),
            # Mês Atual — Meta vs Realizado
            "meta_q":       safe_val(r['Meta Atual (Q)']),
            "meta_c":       safe_val(r['Meta Atual (C$)']),
            "meta_v":       safe_val(r['Meta Atual ($)']),
            "real_q":       safe_val(r['Acum Realizado (Q)']),
            "real_v":       safe_val(r['Acum Realizado ($)']),
            "gap_meta_q":   safe_val(r['Meta - Realizado (Q)']),
            "meta_pct":     safe_val(r.get('Meta %', 0)),
            "meta_st":      safe_val(r.get('Meta St.', '')),
            # Acumulado
            "acum_meta_q":  safe_val(r['Acum Meta (Q)']),
            "acum_real_q":  safe_val(r['Acum Realizado (Q)']),
            "acum_meta_v":  safe_val(r['Acum Meta ($)']),
            "acum_real_v":  safe_val(r['Acum Realizado ($)']),
            # Reprojeção
            "repr_q":       safe_val(r['Reprojeção (Q)']),
            "repr_v":       safe_val(r['Reprojeção ($)']),
            "repr_x_meta":  safe_val(r['Reprojeção x Meta (Q)']),
            "desvio_abaixo":safe_val(r['Desvio Abaixo ($)']),
            "desvio_acima": safe_val(r['Desvio Acima ($)']),
            # Necessidade Atual
            "nec_q":        safe_val(r['Necess. (Q) Atual.']),
            "nec_c":        safe_val(r['Necess. (C$) Atual.']),
            "nec_v":        safe_val(r['Necess. ($) Atual.']),
            # Estoque
            "est_total_q":  safe_val(r['Estoque Total (Q)']),
            "est_total_c":  safe_val(r['Estoque Total (C$)']),
            "est_total_v":  safe_val(r['Estoque Total ($)']),
            "est_min_q":    safe_val(r['Estoque Mínimo (Q)']),
            "est_bling_q":  safe_val(r['Estoque Bling (Q)']),
            "est_full_q":   safe_val(r['Estoque Full (Q)']),
            "est_l1_q":     safe_val(r['Estoque L1 (Q)']),
            "nivel_est_pct":safe_val(r['Nivel Estoque %']),
            "nivel_est_st": safe_val(r.get('Nivel Estoque St.', '')),
            # Pedidos Pendentes
            "ped_pend_q":   safe_val(r['Pedidos Pendentes (Q)']),
            "ped_pend_v":   safe_val(r['Pedidos Pendentes ($)']),
            "ped_dig_q":    safe_val(r['Ped. Pendente [Dig] (Q)']),
            "ped_dig_v":    safe_val(r['Ped. Pendente [Dig] ($)']),
            "data_ped_dig": safe_val(r.get('Data Pedido [Dig]', '')),
            "prev_ent_dig": safe_val(r.get('Previsão Entrega [Dig]', '')),
            # Nec. Compra (BI) — Empresa 1: compra ao fornecedor; Full: transferência
            "nec_compra_q": safe_val(r['Est. Total x Est. Mínimo (Q)']),
            # Nec. Líquida global (descontado PP compartilhado)
            "nec_liq_q":    safe_val(r['Nec. Liquida (Q)']),
            # Estoque Futuro / Mês Seguinte
            "est_fut1_q":   safe_val(r['Estoque Futuro 1 (Q)']),
            "est_fut1_v":   safe_val(r['Estoque Futuro 1  ($)']),
            "est_fut2_q":   safe_val(r['Estoque Futuro 2 (Q)']),
            "meta_seg_q":   safe_val(r['Meta Seguinte (Q)']),
            "meta_seg_v":   safe_val(r['Meta Seguinte ($)']),
            "meta_seg_nec_q": safe_val(r['Meta Seg. Nec (Q)']),
            "est_meta_seg_q": safe_val(r['Estoque Meta Seg (Q)']),
            # Cobertura
            "dias_estoque": safe_val(r.get('Calc. Dias Estoque', 0)),
            "venda_diaria": safe_val(r.get('Calc. Venda diária', 0)),
            "leadtime":     safe_val(r.get('Leadtime Médio (Dias)', 0)),
            # Saúde / Cálculos
            "calc_saude":   safe_val(r.get('Calc. Validar Saúde', '')),
            "calc_formatado": safe_val(r.get('Calc. Formatado', '')),
            "st_oport":     safe_val(r.get('St.Oportunidades', '')),
        })
    return rows

# ─────────────────────────────────────────────
# 7. GERAR TOTAIS POR CANAL (KPIs)
# ─────────────────────────────────────────────
def kpis_canal(empresa_id):
    sub = df[df['Empresa'] == empresa_id]
    rup = ruptura_stats(sub)
    
    # Nec. Líquida: só somar da Empresa 1 (evita dupla contagem com Fulls)
    nec_liq_v = round(float(
        sub.loc[sub['Nec. Liquida (Q)'] > 0, 'Nec. Liquida (Q)'].multiply(
            sub.loc[sub['Nec. Liquida (Q)'] > 0, 'Vlr. Custo ($)']
        ).sum()
    ), 2)
    
    return {
        "meta_q":        round(float(sub['Meta Atual (Q)'].sum()), 0),
        "meta_v":        round(float(sub['Meta Atual ($)'].sum()), 2),
        "real_q":        round(float(sub['Acum Realizado (Q)'].sum()), 0),
        "real_v":        round(float(sub['Acum Realizado ($)'].sum()), 2),
        "repr_q":        round(float(sub['Reprojeção (Q)'].sum()), 0),
        "repr_v":        round(float(sub['Reprojeção ($)'].sum()), 2),
        "nec_compra_q":  round(float(sub['Est. Total x Est. Mínimo (Q)'].sum()), 0),
        "nec_compra_c":  round(float((sub['Est. Total x Est. Mínimo (Q)'] * sub['Vlr. Custo ($)']).sum()), 2),
        "nec_liq_q":     round(float(sub['Nec. Liquida (Q)'].sum()), 0),
        "nec_liq_v":     nec_liq_v,
        "ped_pend_q":    round(float(sub['Pedidos Pendentes (Q)'].sum()), 0),
        "ped_pend_v":    round(float(sub['Pedidos Pendentes ($)'].sum()), 2),
        "est_total_q":   round(float(sub['Estoque Total (Q)'].sum()), 0),
        "est_total_v":   round(float(sub['Estoque Total ($)'].sum()), 2),
        "meta_seg_q":    round(float(sub['Meta Seguinte (Q)'].sum()), 0),
        "meta_seg_v":    round(float(sub['Meta Seguinte ($)'].sum()), 2),
        "meta_seg_nec_q":round(float(sub['Meta Seg. Nec (Q)'].sum()), 0),
        "desvio_abaixo": round(float(sub['Desvio Abaixo ($)'].sum()), 2),
        "desvio_acima":  round(float(sub['Desvio Acima ($)'].sum()), 2),
        **rup,
    }

# ─────────────────────────────────────────────
# 8. TOP-50 MARCA
# ─────────────────────────────────────────────
def top_marcas(empresa_id, n=50):
    sub = df[df['Empresa'] == empresa_id].copy()
    grp = sub.groupby('Marca').agg(
        meta_v    =('Meta Atual ($)', 'sum'),
        real_v    =('Acum Realizado ($)', 'sum'),
        repr_v    =('Reprojeção ($)', 'sum'),
        meta_seg_v=('Meta Seguinte ($)', 'sum'),
        nec_liq_q =('Nec. Liquida (Q)', 'sum'),
        skus      =('Produto', 'count'),
    ).reset_index()
    grp = grp.sort_values('meta_v', ascending=False).head(n)
    result = []
    for _, r in grp.iterrows():
        result.append({
            "marca":       str(r['Marca']),
            "meta_v":      round(float(r['meta_v']), 2),
            "real_v":      round(float(r['real_v']), 2),
            "repr_v":      round(float(r['repr_v']), 2),
            "meta_seg_v":  round(float(r['meta_seg_v']), 2),
            "nec_liq_q":   round(float(r['nec_liq_q']), 0),
            "skus":        int(r['skus']),
        })
    return result

# ─────────────────────────────────────────────
# 9. MONTAR O data.json
# ─────────────────────────────────────────────
print("[INFO] Montando data.json...")

data = {
    "meta": {
        "gerado_em":   datetime.now().strftime("%d/%m/%Y %H:%M"),
        "arquivo_csv": os.path.basename(csv_path),
        "total_linhas": len(df),
        "empresas":    sorted(df['Empresa'].unique().tolist()),
    },
    "kpis": {canal: kpis_canal(emp_id) for canal, emp_id in CANAIS.items()},
    "ruptura": {canal: ruptura_stats(df[df['Empresa'] == emp_id]) for canal, emp_id in CANAIS.items()},
    "top_marcas_atual":    {canal: top_marcas(emp_id) for canal, emp_id in CANAIS.items()},
    "top_marcas_seguinte": {canal: top_marcas(emp_id) for canal, emp_id in CANAIS.items()},
    "rows": {canal: build_canal(emp_id) for canal, emp_id in CANAIS.items()},
}

# ─────────────────────────────────────────────
# 10. MONTAR O produtos.json (lookup rápido)
# ─────────────────────────────────────────────
print("[INFO] Montando produtos.json...")

produtos = {}
for _, r in df[df['Empresa'] == 1].iterrows():
    sku = str(r['Produto']).strip()
    produtos[sku] = {
        "sku":       sku,
        "descricao": str(r.get('Descrição', '')).strip(),
        "categoria": str(r.get('Categoria', '')).strip(),
        "marca":     str(r.get('Marca', '')).strip(),
        "comprador": str(r.get('Comprador', '')).strip(),
    }

# ─────────────────────────────────────────────
# 11. SALVAR
# ─────────────────────────────────────────────
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

with open("produtos.json", "w", encoding="utf-8") as f:
    json.dump(produtos, f, ensure_ascii=False, separators=(',', ':'))

size_data    = os.path.getsize("data.json") / 1024
size_prod    = os.path.getsize("produtos.json") / 1024
print(f"[OK] data.json gerado: {size_data:.1f} KB")
print(f"[OK] produtos.json gerado: {size_prod:.1f} KB")
print(f"[OK] Gerado em: {data['meta']['gerado_em']}")
