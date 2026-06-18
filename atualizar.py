#!/usr/bin/env python3
"""
atualizar.py — Popular Pet S&OP
Lê o CSV exportado pelo XLSM e gera data.json + produtos.json para o GitHub Pages.
Lógica consolidada idêntica à aba Compras do XLSM v5.
"""

import sys, os, json, glob
import pandas as pd
import numpy as np
from datetime import datetime

# ─── 1. LOCALIZAR O CSV ──────────────────────────────────────
def find_csv(arg=None):
    if arg and os.path.isfile(arg):
        return arg
    candidates = sorted(glob.glob("planilha/*.csv"), key=os.path.getmtime, reverse=True)
    if candidates:
        print(f"[INFO] CSV encontrado: {candidates[0]}")
        return candidates[0]
    raise FileNotFoundError("Nenhum CSV encontrado em planilha/.")

csv_path = find_csv(sys.argv[1] if len(sys.argv) > 1 else None)

# ─── 2. LER E LIMPAR ─────────────────────────────────────────
print(f"[INFO] Lendo: {csv_path}")
df = pd.read_csv(csv_path, sep=None, engine='python', dtype=str)
df.columns = [c.replace('\ufeff', '').strip() for c in df.columns]

def br_float(s):
    return (s.astype(str).str.strip()
             .str.replace(r'[%\s]', '', regex=True)
             .str.replace('.', '', regex=False)
             .str.replace(',', '.', regex=False)
             .pipe(pd.to_numeric, errors='coerce')
             .fillna(0))

df['Empresa'] = pd.to_numeric(df['Empresa'].astype(str).str.strip(), errors='coerce').fillna(0).astype(int)
df['Produto'] = df['Produto'].astype(str).str.strip()

NUM_COLS = [
    'Meta Atual (Q)', 'Meta Atual (C$)', 'Meta Atual ($)',
    'Acum Meta (Q)', 'Acum Realizado (Q)', 'Acum Meta ($)', 'Acum Realizado ($)', 'Acum Realizado (C$)',
    'Meta - Realizado (Q)', 'Reprojeção (Q)', 'Reprojeção (C$)', 'Reprojeção ($)',
    'Necess. (Q) Atual.', 'Necess. (C$) Atual.', 'Necess. ($) Atual.',
    'Estoque Total (Q)', 'Estoque Total (C$)', 'Estoque Total ($)',
    'Estoque Mínimo (Q)', 'Est. Total x Est. Mínimo (Q)',
    'Estoque Meta (Q)', 'Estoque Meta (C$)', 'Estoque Meta ($)',
    'Pedidos Pendentes (Q)', 'Pedidos Pendentes (C$)', 'Pedidos Pendentes ($)',
    'Estoque Total + Ped. Pendente (Q)',
    'Estoque Futuro 1 (Q)', 'Estoque Futuro 1  (C$)', 'Estoque Futuro 1  ($)',
    'Estoque Meta Seg (Q)', 'Estoque Meta Seg (C$)', 'Estoque Meta Seg ($)',
    'Meta Seguinte (Q)', 'Meta Seguinte (C$)', 'Meta Seguinte ($)',
    'Meta Seg. Nec (Q)', 'Meta Seg. Nec (C$)', 'Meta Seg. Nec ($)',
    'Estoque Futuro 2 (Q)',
    'Ped. Pendente [Dig] (Q)', 'Ped. Pendente [Dig] ($)',
    'Ped. Pendente [999] (Q)', 'Ped. Pendente [999] (C$)',
    'Leadtime Médio (Dias)', 'Calc. Venda diária', 'Calc. Dias Estoque',
    'Faturado XML Central (Q)',
    'Nivel Estoque %', 'Nivel Estoque Rep %',
    'Desvio Abaixo ($)', 'Desvio Acima ($)',
]
for c in NUM_COLS:
    if c in df.columns:
        df[c] = br_float(df[c])

for c in ['Vlr. Custo ($)', 'Vlr. Venda ($)', 'Vlr Ticket ($)']:
    if c in df.columns:
        df[c] = br_float(df[c])

for c in ['Margem  (%)', 'Impostos (%)']:
    if c in df.columns:
        df[c] = (df[c].astype(str).str.strip()
                  .str.replace('%', '', regex=False)
                  .str.replace(',', '.', regex=False)
                  .pipe(pd.to_numeric, errors='coerce').fillna(0))

print(f"[INFO] {len(df)} linhas | {df['Empresa'].unique().tolist()}")

# ─── 3. LÓGICA CONSOLIDADA (igual XLSM aba Compras v5) ───────
# Agrupa todas empresas por SKU para calcular a necessidade real
meta_cons   = df.groupby('Produto')['Meta Atual (Q)'].sum()
real_cons   = df.groupby('Produto')['Acum Realizado (Q)'].sum()
est_cons    = df.groupby('Produto')['Estoque Total (Q)'].sum()
pp_cons     = df.groupby('Produto')['Pedidos Pendentes (Q)'].sum()
estmin_cons = df.groupby('Produto')['Estoque Mínimo (Q)'].sum()
metaseg_cons= df.groupby('Produto')['Meta Seguinte (Q)'].sum()

e1 = df[df['Empresa'] == 1].copy()
e1['meta_cons']    = e1['Produto'].map(meta_cons)
e1['real_cons']    = e1['Produto'].map(real_cons)
e1['est_cons']     = e1['Produto'].map(est_cons)
e1['pp_cons']      = e1['Produto'].map(pp_cons)
e1['estmin_cons']  = e1['Produto'].map(estmin_cons)
e1['metaseg_cons'] = e1['Produto'].map(metaseg_cons)

# necLiq consolidada = MAX(0, EstMin.total - Estoque.total - PP.total)
e1['nec_liq_cons'] = (e1['estmin_cons'] - e1['est_cons'] - e1['pp_cons']).clip(lower=0)
e1['_invest']      = e1['nec_liq_cons'] * e1['Vlr. Custo ($)']
e1['_fatur']       = e1['nec_liq_cons'] * e1['Vlr Ticket ($)']

# ─── 4. FILTROS DE COERÊNCIA (XLSM aba Compras) ──────────────
# 1) ATIVO  2) meta_cons > 0  3) real_cons < meta_cons  4) nec_liq > 0
filtro_compras = (
    (e1['Status Compras'] == 'ATIVO') &
    (e1['meta_cons'] > 0) &
    (e1['real_cons'] < e1['meta_cons']) &
    (e1['nec_liq_cons'] > 0)
)
compras = e1[filtro_compras].copy().sort_values('_invest', ascending=False)

invest_total = float(compras['_invest'].sum())
fatur_total  = float(compras['_fatur'].sum())
lucro_total  = fatur_total - invest_total
margem_total = lucro_total / fatur_total * 100 if fatur_total > 0 else 0

print(f"[INFO] Compras: {len(compras)} SKUs | Invest R${invest_total:,.0f} | Fat R${fatur_total:,.0f} | Margem {margem_total:.1f}%")

# ─── 5. BREAKDOWN MÊS ATUAL × MÊS SEGUINTE ──────────────────
# Peso Mês Atual por SKU = estmin_parte_atual / estmin_total
# (simplificado: usar proporção do estoque mínimo próprio da emp1)
e1_comp = compras.copy()
e1_comp['est_min_e1'] = e1_comp['Estoque Mínimo (Q)']
total_min = e1_comp['estmin_cons'].replace(0, np.nan)
e1_comp['peso_atual'] = (e1_comp['Estoque Mínimo (Q)'] / total_min).fillna(0.5)
invest_atual = float((e1_comp['_invest'] * e1_comp['peso_atual']).sum())
invest_seg   = invest_total - invest_atual
fatur_atual  = float((e1_comp['_fatur'] * e1_comp['peso_atual']).sum())
fatur_seg    = fatur_total - fatur_atual

# ─── 6. HELPER JSON-SAFE ─────────────────────────────────────
def safe(v):
    if isinstance(v, (float, np.floating)):
        return 0 if (np.isnan(v) or np.isinf(v)) else round(float(v), 4)
    if isinstance(v, (np.integer,)):
        return int(v)
    if pd.isna(v) if not isinstance(v, (list, dict)) else False:
        return ''
    return v

# ─── 7. ROWS PARA O DASHBOARD ────────────────────────────────
def row_to_dict(r):
    invest = safe(r['nec_liq_cons'] * r['Vlr. Custo ($)'])
    fatur  = safe(r['nec_liq_cons'] * r['Vlr Ticket ($)'])
    lucro  = round(float(fatur) - float(invest), 2) if invest and fatur else 0
    margem = round(float(lucro) / float(fatur) * 100, 1) if fatur else 0
    return {
        'sku':          safe(r['Produto']),
        'descricao':    safe(r.get('Descrição', '')),
        'categoria':    safe(r.get('Categoria', '')),
        'marca':        safe(r.get('Marca', '')),
        'comprador':    safe(r.get('Comprador', '')),
        'status':       safe(r.get('Status Compras', '')),
        'pareto':       safe(r.get('Pareto', '')),
        'leadtime':     safe(r.get('Leadtime Médio (Dias)', 0)),
        # Estoque Empresa 1
        'est_e1':       safe(r['Estoque Total (Q)']),
        'est_min_e1':   safe(r['Estoque Mínimo (Q)']),
        # Consolidado rede
        'est_cons':     safe(r['est_cons']),
        'estmin_cons':  safe(r['estmin_cons']),
        'pp_cons':      safe(r['pp_cons']),
        # Necessidade
        'nec_compra_q': safe(r['Est. Total x Est. Mínimo (Q)']),  # bruta emp1
        'nec_liq_q':    safe(r['nec_liq_cons']),                  # líquida consolidada
        'ped_pend_q':   safe(r['Pedidos Pendentes (Q)']),
        'ped_dig_q':    safe(r['Ped. Pendente [Dig] (Q)']),
        'data_ped_dig': safe(r.get('Data Pedido [Dig]', '')),
        'prev_ent_dig': safe(r.get('Previsão Entrega [Dig]', '')),
        # Preços
        'vlr_custo':    safe(r['Vlr. Custo ($)']),
        'vlr_venda':    safe(r['Vlr. Venda ($)']),
        'vlr_ticket':   safe(r['Vlr Ticket ($)']),
        # Financeiro da compra
        'invest':       invest,
        'fatur_esp':    fatur,
        'lucro_bruto':  lucro,
        'margem_pct':   margem,
        # Meta
        'meta_q':       safe(r['Meta Atual (Q)']),
        'meta_v':       safe(r['Meta Atual ($)']),
        'real_q':       safe(r['Acum Realizado (Q)']),
        'real_v':       safe(r['Acum Realizado ($)']),
        'meta_seg_q':   safe(r['Meta Seguinte (Q)']),
        'meta_seg_cons':safe(r['metaseg_cons']),
        'falta_meta_q': safe(r['meta_cons'] - r['real_cons']),
        # Estoque futuro
        'est_fut1_q':   safe(r.get('Estoque Futuro 1 (Q)', 0)),
    }

rows_compras = [row_to_dict(r) for _, r in compras.iterrows()]

# ─── 8. COMPRADORES ──────────────────────────────────────────
compradores = {}
for row in rows_compras:
    c = row['comprador'] or 'N/A'
    if c not in compradores:
        compradores[c] = {'nec_liq': 0, 'invest': 0, 'fatur': 0, 'lucro': 0, 'skus': 0, 'ped_pend': 0}
    compradores[c]['nec_liq']  += row['nec_liq_q']
    compradores[c]['invest']   += row['invest']
    compradores[c]['fatur']    += row['fatur_esp']
    compradores[c]['lucro']    += row['lucro_bruto']
    compradores[c]['skus']     += 1
    compradores[c]['ped_pend'] += row['ped_pend_q']
for c in compradores:
    compradores[c] = {k: round(float(v), 2) if isinstance(v, float) else v
                      for k, v in compradores[c].items()}
    f = compradores[c]['fatur']
    compradores[c]['margem'] = round(compradores[c]['lucro'] / f * 100, 1) if f else 0

# ─── 9. RUPTURA ──────────────────────────────────────────────
e1_all = df[df['Empresa'] == 1].copy()
for c in ['Estoque Total (Q)', 'Acum Realizado (Q)', 'Meta Atual (Q)', 'Meta Atual ($)']:
    if c in e1_all.columns:
        e1_all[c] = br_float(e1_all[c]) if e1_all[c].dtype == object else e1_all[c]

ruptura = e1_all[(e1_all['Estoque Total (Q)'] == 0) & (e1_all['Acum Realizado (Q)'] > 0)].copy()
skus_com_venda = int((e1_all['Acum Realizado (Q)'] > 0).sum())
taxa_rup = round(len(ruptura) / skus_com_venda * 100, 1) if skus_com_venda else 0

rows_ruptura = []
for _, r in ruptura.sort_values('Meta Atual ($)', ascending=False).iterrows():
    rows_ruptura.append({
        'sku': safe(r['Produto']), 'descricao': safe(r.get('Descrição', '')),
        'categoria': safe(r.get('Categoria', '')), 'marca': safe(r.get('Marca', '')),
        'comprador': safe(r.get('Comprador', '')),
        'real_q': safe(r['Acum Realizado (Q)']), 'meta_q': safe(r['Meta Atual (Q)']),
        'meta_v': safe(r['Meta Atual ($)']), 'nec_compra_q': safe(r['Est. Total x Est. Mínimo (Q)']),
        'leadtime': safe(r.get('Leadtime Médio (Dias)', 0)),
        'ped_pend_q': safe(r['Pedidos Pendentes (Q)']),
    })

# ─── 10. TOP MARCAS ──────────────────────────────────────────
marcas = (e1_all.groupby('Marca').agg(
    meta_v=('Meta Atual ($)', 'sum'), real_v=('Acum Realizado ($)', 'sum'),
    repr_v=('Reprojeção ($)', 'sum'), meta_seg_v=('Meta Seguinte ($)', 'sum'), skus=('Produto', 'count')
).reset_index().sort_values('meta_v', ascending=False).head(50))
rows_marcas = [{'marca': str(r.Marca), 'meta_v': round(float(r.meta_v), 2),
                'real_v': round(float(r.real_v), 2), 'repr_v': round(float(r.repr_v), 2),
                'meta_seg_v': round(float(r.meta_seg_v), 2), 'skus': int(r.skus)}
               for r in marcas.itertuples()]

# ─── 11. MONTAR data.json ────────────────────────────────────
data = {
    'meta': {
        'gerado_em': datetime.now().strftime('%d/%m/%Y %H:%M'),
        'arquivo_csv': os.path.basename(csv_path),
        'total_linhas': len(df),
    },
    'compras': {
        'total': {
            'skus': len(compras),
            'invest': round(invest_total, 2),
            'fatur':  round(fatur_total, 2),
            'lucro':  round(lucro_total, 2),
            'margem': round(margem_total, 1),
            'invest_atual': round(invest_atual, 2),
            'invest_seg':   round(invest_seg, 2),
            'fatur_atual':  round(fatur_atual, 2),
            'fatur_seg':    round(fatur_seg, 2),
            'pct_atual':    round(invest_atual / invest_total * 100, 1) if invest_total else 0,
            'pct_seg':      round(invest_seg   / invest_total * 100, 1) if invest_total else 0,
        },
        'compradores': compradores,
        'rows': rows_compras,
    },
    'ruptura': {
        'total_skus': int(len(e1_all)),
        'skus_zero': int((e1_all['Estoque Total (Q)'] == 0).sum()),
        'skus_com_venda': skus_com_venda,
        'skus_ruptura': len(ruptura),
        'taxa_ruptura': taxa_rup,
        'receita_risco': round(float(ruptura['Meta Atual ($)'].sum()), 2),
        'rows': rows_ruptura,
    },
    'marcas': rows_marcas,
}

with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, separators=(',', ':'))

with open('produtos.json', 'w', encoding='utf-8') as f:
    produtos = {str(r['Produto']): {'sku': str(r['Produto']), 'descricao': str(r.get('Descrição', '')),
                'categoria': str(r.get('Categoria', '')), 'marca': str(r.get('Marca', '')),
                'comprador': str(r.get('Comprador', ''))}
                for _, r in df[df['Empresa'] == 1].iterrows()}
    json.dump(produtos, f, ensure_ascii=False, separators=(',', ':'))

print(f"[OK] data.json: {os.path.getsize('data.json')//1024} KB")
print(f"[OK] produtos.json: {os.path.getsize('produtos.json')//1024} KB")
