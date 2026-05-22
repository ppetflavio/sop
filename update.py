"""
update.py — Atualiza o data.json com os dados da planilha e faz push para GitHub
Uso: python update.py planilha.xlsx
"""
import sys, json, numpy as np, pandas as pd
from pathlib import Path

class NpEnc(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, np.integer): return int(o)
        if isinstance(o, np.floating): return float(o)
        return super().default(o)

def ms_ste(row):
    if row['Estoque Mês Seguinte (Q)'] <= 0: return 'Estoque Zerado'
    if row['Estoque Mês Seguinte (Q)'] < row['Estoque Segurança Mês Seguinte (Q)']: return 'Estoque Falta'
    return 'Estoque Sobra'

def ms_stv(row):
    if row['Meta Mês Seguinte (Q)'] == 0: return 'Sem Meta'
    if row['Necessidade Compra Mês Seguinte (Q)'] > 0: return 'Abaixo Meta'
    return 'Acima Meta'

def build(xlsx_path):
    df = pd.read_excel(xlsx_path)
    df['Comprador'] = df['Comprador'].str.strip()
    EMPS = {1:'Marketplace', 995:'Shopee Full', 996:'Amazon Full', 997:'Meli Full'}
    result = {}
    for emp_num, emp_name in EMPS.items():
        d = df[df['Empresa'] == emp_num].copy()
        if len(d) == 0: continue
        f  = lambda col: float(d[col].sum())
        fp = lambda c1,c2: float((d[c1]*d[c2]).sum())
        mq=f('Meta Mês Total (Q)'); rq=f('Realizado Acumulado (Q)'); pq=f('Projeção Venda (Q)')
        ms=f('Meta Mês Seguinte (Q)'); es=f('Estoque Mês Seguinte (Q)')
        ns=f('Necessidade Compra Mês Seguinte (Q)')
        ps=f('Peso Total Kg (Necessidde compra Mês seguinte)')
        nv=fp('Necessidade Compra Mês Seguinte (Q)','Preço Custo ($)')
        cats=[]
        for cat,cd in d.groupby('Categoria'):
            cats.append({'Categoria':str(cat),'real':float(cd['Realizado Acumulado (Q)'].sum()),
                'meta':float(cd['Meta Mês Total (Q)'].sum()),'proj':float(cd['Projeção Venda (Q)'].sum()),
                'receita':float((cd['Realizado Acumulado (Q)']*cd['Preço Praticado ($)']).sum()),
                'meta_seg':float(cd['Meta Mês Seguinte (Q)'].sum()),
                'estq_seg':float(cd['Estoque Mês Seguinte (Q)'].sum()),
                'estqss_seg':float(cd['Estoque Segurança Mês Seguinte (Q)'].sum()),
                'nec_seg':float(cd['Necessidade Compra Mês Seguinte (Q)'].sum()),
                'nec_peso_seg':float(cd['Peso Total Kg (Necessidde compra Mês seguinte)'].sum()),
                'nec_val_seg':float((cd['Necessidade Compra Mês Seguinte (Q)']*cd['Preço Custo ($)']).sum()),
                'rec_meta_seg':float((cd['Meta Mês Seguinte (Q)']*cd['Preço Praticado ($)']).sum())})
        cats.sort(key=lambda x:-x['receita'])
        top100=d.assign(_r=d['Realizado Acumulado (Q)']*d['Preço Praticado ($)']).nlargest(100,'_r')
        t100=[{'Produto':str(r['Produto']),'Descrição':str(r['Descrição']),'Marca':str(r['Marca']),
            'Categoria':str(r['Categoria']),'Comprador':str(r['Comprador']),'Pareto':str(r['Pareto']),
            'Peso Kg':float(r['Peso Kg']),'Preço Custo ($)':float(r['Preço Custo ($)']),
            'Preço Praticado ($)':float(r['Preço Praticado ($)']),'Meta Mês Total (Q)':float(r['Meta Mês Total (Q)']),
            'Realizado Acumulado (Q)':float(r['Realizado Acumulado (Q)']),'Projeção Venda (Q)':float(r['Projeção Venda (Q)']),
            'Receita_Real':float(r['Realizado Acumulado (Q)']*r['Preço Praticado ($)']),'Status Venda':str(r['Status Venda']),
            'Status Estoque':str(r['Status Estoque']),'Estoque Total (Q)':float(r['Estoque Total (Q)']),
            'Necessidade Compra (Q)':float(r['Necessidade Compra (Q)']),'Leadtime Médio (Dias)':float(r['Leadtime Médio (Dias)'])}
            for _,r in top100.iterrows()]
        dseg=d[d['Necessidade Compra Mês Seguinte (Q)']>0].assign(_i=d['Necessidade Compra Mês Seguinte (Q)']*d['Preço Custo ($)']).sort_values('_i',ascending=False)
        tseg=[{'Produto':str(r['Produto']),'Descrição':str(r['Descrição']),'Marca':str(r['Marca']),
            'Categoria':str(r['Categoria']),'Comprador':str(r['Comprador']),'Pareto':str(r['Pareto']),
            'Peso Kg':float(r['Peso Kg']),'Preço Custo ($)':float(r['Preço Custo ($)']),
            'Preço Praticado ($)':float(r['Preço Praticado ($)']),'Projeção Venda (Q)':float(r['Projeção Venda (Q)']),
            'Meta Mês Seguinte (Q)':float(r['Meta Mês Seguinte (Q)']),'Estoque Mês Seguinte (Q)':float(r['Estoque Mês Seguinte (Q)']),
            'Estoque Segurança Mês Seguinte (Q)':float(r['Estoque Segurança Mês Seguinte (Q)']),'Ped. Pendente [Dig] (Q)':float(r['Ped. Pendente [Dig] (Q)']),
            'Leadtime Médio (Dias)':float(r['Leadtime Médio (Dias)']),'Necessidade Compra Mês Seguinte (Q)':float(r['Necessidade Compra Mês Seguinte (Q)']),
            'Nec_Peso_Seg':float(r['Peso Total Kg (Necessidde compra Mês seguinte)']),'Nec_Valor_Seg':float(r['Necessidade Compra Mês Seguinte (Q)']*r['Preço Custo ($)']),
            'Status_Estq_Seg':ms_ste(r),'Status_Venda_Seg':ms_stv(r)}
            for _,r in dseg.iterrows() if r['Necessidade Compra Mês Seguinte (Q)']>0]
        comps=[]
        for comp,cd in d.groupby('Comprador'):
            comp=str(comp).strip()
            if not comp or comp in ('0','nan',''): continue
            comps.append({'Comprador':comp,'skus':int(len(cd)),
                'nec':float(cd['Necessidade Compra Mês Seguinte (Q)'].sum()),
                'nec_seg':float(cd['Necessidade Compra Mês Seguinte (Q)'].sum()),
                'peso_seg':float(cd['Peso Total Kg (Necessidde compra Mês seguinte)'].sum()),
                'nec_val':float((cd['Necessidade Compra Mês Seguinte (Q)']*cd['Preço Custo ($)']).sum()),
                'receita':float((cd['Realizado Acumulado (Q)']*cd['Preço Praticado ($)']).sum()),
                'meta':float(cd['Meta Mês Total (Q)'].sum())})
        crit=d[(d['Status Estoque']=='Estoque Zerado')&(d['Status Venda']=='Acima Meta')].assign(_r=d['Projeção Venda (Q)']*d['Preço Praticado ($)']).nlargest(50,'_r')
        criticos=[{'Produto':str(r['Produto']),'Descrição':str(r['Descrição']),'Marca':str(r['Marca']),
            'Categoria':str(r['Categoria']),'Comprador':str(r['Comprador']),'Projeção Venda (Q)':float(r['Projeção Venda (Q)']),
            'Necessidade Compra (Q)':float(r['Necessidade Compra (Q)']),'Receita_Risco':float(r['Projeção Venda (Q)']*r['Preço Praticado ($)']),
            'Necessidade Compra Mês Seguinte (Q)':float(r['Necessidade Compra Mês Seguinte (Q)']),'Leadtime Médio (Dias)':float(r['Leadtime Médio (Dias)'])}
            for _,r in crit.iterrows()]
        result[str(emp_num)]={
            'nome':emp_name,'skus':int(len(d)),
            'meta_qtd':mq,'real_qtd':rq,'proj_qtd':pq,
            'meta_receita':fp('Meta Mês Total (Q)','Preço Praticado ($)'),
            'real_receita':fp('Realizado Acumulado (Q)','Preço Praticado ($)'),
            'proj_receita':fp('Projeção Venda (Q)','Preço Praticado ($)'),
            'estq_qtd':f('Estoque Total (Q)'),
            'estq_valor':fp('Estoque Total (Q)','Preço Custo ($)'),
            'nec_qtd':f('Necessidade Compra (Q)'),
            'aderencia_pct':round(rq/max(mq,1)*100,1),
            'proj_vs_meta_pct':round(pq/max(mq,1)*100,1),
            'status_estoque':d['Status Estoque'].value_counts().to_dict(),
            'status_venda':d['Status Venda'].value_counts().to_dict(),
            'meta_seg':ms,'estq_seg':es,
            'estqss_seg':f('Estoque Segurança Mês Seguinte (Q)'),
            'nec_seg':ns,'peso_seg':ps,'nec_val':nv,
            'rec_meta_seg':fp('Meta Mês Seguinte (Q)','Preço Praticado ($)'),
            'cobertura_seg':round(es/max(ms,1)*100,1),
            'skus_ok_seg':int((d['Estoque Mês Seguinte (Q)']>=d['Estoque Segurança Mês Seguinte (Q)']).sum()),
            'skus_nec_seg':int((d['Necessidade Compra Mês Seguinte (Q)']>0).sum()),
            'status_estq_seg':d.apply(ms_ste,axis=1).value_counts().to_dict(),
            'status_venda_seg':d.apply(ms_stv,axis=1).value_counts().to_dict(),
            'categoria':cats,'top100':t100,'top_seg':tseg,
            'criticos':criticos,'compradores':comps}
    return result

if __name__ == '__main__':
    xlsx = sys.argv[1] if len(sys.argv)>1 else 'planilha.xlsx'
    print(f"Processando {xlsx}...")
    D = build(xlsx)
    total = sum(v['nec_seg'] for v in D.values())
    print(f"Total nec_seg: {total:.0f} un | {sum(v['skus'] for v in D.values())} SKUs")
    out = Path('data.json')
    out.write_text(json.dumps(D, cls=NpEnc, ensure_ascii=False, separators=(',',':')))
    print(f"data.json gerado: {out.stat().st_size//1024} KB")
    print("\nPara publicar no GitHub Pages:")
    print("  git add data.json")
    print("  git commit -m 'Atualiza dados S&OP'")
    print("  git push")
    print("\nDashboard atualizado em ~30 segundos!")
