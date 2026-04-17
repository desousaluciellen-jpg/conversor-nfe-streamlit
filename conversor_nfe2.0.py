import streamlit as st
import openpyxl
from collections import defaultdict
import io, zipfile, pandas as pd, plotly.express as px
from datetime import datetime

st.set_page_config(page_title="Conversor NFe 2.0", layout="wide")
st.title("📄 Conversor NFe 2.0 Pro")

def extrair(arquivo):
    wb=openpyxl.load_workbook(arquivo,data_only=True); ws=wb.active
    headers=[ws.cell(1,c).value for c in range(1,ws.max_column+1)]
    notas=defaultdict(lambda:{'serie':None,'numero':None,'data':None,'chave':None,'itens':[]})
    emit={}
    for r in range(2,ws.max_row+1):
        row={h:ws.cell(r,c).value for c,h in enumerate(headers,1)}
        chave=row.get('ChaveAcesso')
        if chave and row.get('DescricaoProduto'):
            if not emit:
                emit={'cnpj':str(row.get('CnpjEmitente') or ''),'razao':str(row.get('RazaoSocialEmitente') or '')}
            notas[chave]['chave']=chave
            notas[chave]['numero']=row.get('NumeroDocumento')
            notas[chave]['serie']=row.get('SerieDocumento')
            notas[chave]['data']=row.get('DataEmissaoNfe')
            notas[chave]['itens'].append({
                'cfop':str(row.get('CfopProduto') or ''),
                'desc':row.get('DescricaoProduto'),
                'ncm':row.get('NcmProduto'),
                'qtd':float(row.get('QuantidadeUnidadeComercial') or 0),
                'vTotal':float(row.get('ValorTotalProduto') or 0),
                'vDesc':float(row.get('ValorDesconto') or 0),
                'icmsTag':row.get('TipoIcmsTag') or '',
                'icmsTrib':row.get('IcmsTributacao') or '',
            })
    return notas,emit

up=st.file_uploader("Planilha",type="xlsx")
if up:
    notas,emit=extrair(up)
    if not notas:
        st.error("Nenhuma nota encontrada. Verifique colunas ChaveAcesso e DescricaoProduto.")
        st.stop()
    
    notas_list=[]; itens_list=[]
    for ch,n in notas.items():
        bruto=sum(i['vTotal'] for i in n['itens']); desc=sum(i['vDesc'] for i in n['itens'])
        notas_list.append({'Chave':ch,'Numero':n['numero'],'Serie':n['serie'],'Data':n['data'],
                           'CFOPs':', '.join(sorted(set(i['cfop'] for i in n['itens']))),
                           'Bruto':bruto,'Desc':desc,'Liquido':bruto-desc,
                           'ICMS_Tags':', '.join(sorted(set(i['icmsTag'] for i in n['itens'] if i['icmsTag']))),
                           'Itens':len(n['itens'])})
        for i in n['itens']:
            itens_list.append({'Chave':ch,'Numero':n['numero'],'CFOP':i['cfop'],'Descricao':i['desc'],
                               'NCM':i['ncm'],'Qtd':i['qtd'],'V_Total':i['vTotal'],'V_Desc':i['vDesc'],
                               'V_Liq':i['vTotal']-i['vDesc'],'ICMS_Tag':i['icmsTag'],'ICMS_Trib':i['icmsTrib']})
    
    df_n=pd.DataFrame(notas_list); df_i=pd.DataFrame(itens_list)
    st.success(f"{len(df_n)} notas | Emitente: {emit.get('razao')} - {emit.get('cnpj')}")
    st.dataframe(df_n.head(10), use_container_width=True)
    
    # Excel com xlsxwriter (evita bug openpyxl)
    excel=io.BytesIO()
    try:
        with pd.ExcelWriter(excel, engine='xlsxwriter') as w:
            pd.DataFrame([{'Emitente':emit.get('razao'),'CNPJ':emit.get('cnpj')}]).to_excel(w,'Resumo',index=False)
            df_n.to_excel(w,'Notas_Geradas',index=False)
            df_i.to_excel(w,'Itens_Detalhados',index=False)
            df_i.groupby('CFOP').agg(Bruto=('V_Total','sum'),Desc=('V_Desc','sum')).reset_index().to_excel(w,'Por_CFOP',index=False)
    except Exception as e:
        st.error(f"Erro ao gerar Excel: {e}")
        st.stop()
    excel.seek(0)
    
    # ZIP
    zipb=io.BytesIO()
    with zipfile.ZipFile(zipb,'w') as z:
        for _,r in df_n.iterrows():
            z.writestr(f"xmls/NFe_{r['Numero']}.xml", f"<NFe>{r['Chave']}</NFe>")
        z.writestr("Relatorio_Completo.xlsx", excel.getvalue())
    zipb.seek(0)
    st.download_button("⬇️ Baixar", zipb, "pacote.zip")
