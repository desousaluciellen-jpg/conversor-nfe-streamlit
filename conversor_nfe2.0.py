import streamlit as st
import openpyxl
from collections import defaultdict
from datetime import datetime
import io
import zipfile
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Conversor NFe 2.0 Pro", layout="wide")
st.markdown('<style>.hero{background:linear-gradient(135deg,#1e3a8a,#3b82f6);padding:1.8rem;border-radius:14px;color:white}</style>', unsafe_allow_html=True)

def extrair(arquivo):
    wb=openpyxl.load_workbook(arquivo,data_only=True); ws=wb.active
    headers=[ws.cell(1,c).value for c in range(1,ws.max_column+1)]
    notas=defaultdict(lambda:{'serie':None,'numero':None,'dataEmissao':None,'chave':None,'itens':[]})
    emitente={}
    for r in range(2,ws.max_row+1):
        row={h:ws.cell(r,c).value for c,h in enumerate(headers,1)}
        chave=row.get('ChaveAcesso')
        if chave and row.get('DescricaoProduto'):
            if not emitente:
                emitente={'cnpj':str(row.get('CnpjEmitente') or ''),'razao':str(row.get('RazaoSocialEmitente') or row.get('NomeEmitente') or '')}
            notas[chave]['chave']=chave
            notas[chave]['serie']=row.get('SerieDocumento')
            notas[chave]['numero']=row.get('NumeroDocumento')
            notas[chave]['dataEmissao']=row.get('DataEmissaoNfe')
            notas[chave]['itens'].append({
                'cfop':str(row.get('CfopProduto') or ''),
                'descricao':row.get('DescricaoProduto'),
                'ncm':row.get('NcmProduto'),
                'qtd':float(row.get('QuantidadeUnidadeComercial') or 0),
                'vUnit':float(row.get('ValorUnitarioComercial') or 0),
                'vTotal':float(row.get('ValorTotalProduto') or 0),
                'vDesc':float(row.get('ValorDesconto') or 0),
                'icmsTag':row.get('TipoIcmsTag') or '',
                'icmsTrib':row.get('IcmsTributacao') or '',
            })
    return notas, emitente

st.markdown('<div class="hero"><h2>📄 Conversor NFe 2.0 Pro</h2></div>', unsafe_allow_html=True)
up=st.file_uploader("Planilha", type="xlsx")

if up:
    notas,emit=extrair(up)
    # preparar dataframes
    lista_notas=[]
    lista_itens=[]
    for chave,n in notas.items():
        total_bruto=sum(i['vTotal'] for i in n['itens'])
        total_desc=sum(i['vDesc'] for i in n['itens'])
        total_liq=total_bruto-total_desc
        cfops=', '.join(sorted(set(i['cfop'] for i in n['itens'])))
        icms_tags=', '.join(sorted(set(i['icmsTag'] for i in n['itens'] if i['icmsTag'])))
        lista_notas.append({
            'ChaveAcesso':chave,'Numero':n['numero'],'Serie':n['serie'],
            'DataEmissao':n['dataEmissao'],'CFOPs':cfops,
            'Valor_Bruto':total_bruto,'Descontos':total_desc,'Valor_Liquido':total_liq,
            'ICMS_Tags':icms_tags,'Qtd_Itens':len(n['itens'])
        })
        for i in n['itens']:
            lista_itens.append({
                'ChaveAcesso':chave,'Numero':n['numero'],'Serie':n['serie'],
                'DataEmissao':n['dataEmissao'],'CFOP':i['cfop'],
                'Descricao':i['descricao'],'NCM':i['ncm'],'Qtd':i['qtd'],
                'V_Unit':i['vUnit'],'V_Total':i['vTotal'],'V_Desc':i['vDesc'],
                'V_Liquido':i['vTotal']-i['vDesc'],
                'ICMS_Tag':i['icmsTag'],'ICMS_Trib':i['icmsTrib']
            })
    df_notas=pd.DataFrame(lista_notas)
    df_itens=pd.DataFrame(lista_itens)
    
    # resumo CFOP
    resumo_cfop=df_itens.groupby('CFOP').agg(Qtde=('V_Total','count'),Bruto=('V_Total','sum'),Desc=('V_Desc','sum')).reset_index()
    resumo_cfop['Liquido']=resumo_cfop['Bruto']-resumo_cfop['Desc']
    
    st.success(f"Emitente: {emit.get('razao','')} | CNPJ: {emit.get('cnpj','')} | {len(df_notas)} notas processadas")
    
    col1,col2=st.columns(2)
    with col1:
        st.dataframe(df_notas.head(20), use_container_width=True)
    with col2:
        fig=px.pie(resumo_cfop, names='CFOP', values='Liquido', hole=0.4, title="Valor Líquido por CFOP")
        st.plotly_chart(fig, use_container_width=True)
    
    # Excel completo
    excel=io.BytesIO()
    with pd.ExcelWriter(excel, engine='openpyxl') as w:
        # Resumo
        pd.DataFrame([{'Emitente':emit.get('razao'),'CNPJ':emit.get('cnpj'),'Total_Notas':len(df_notas),'Total_Bruto':df_notas['Valor_Bruto'].sum()}]).to_excel(w, 'Resumo', index=False)
        resumo_cfop.to_excel(w, 'Por_CFOP', index=False)
        df_notas.to_excel(w, 'Notas_Geradas', index=False)
        df_itens.to_excel(w, 'Itens_Detalhados', index=False)
    excel.seek(0)
    
    # ZIP
    zip_buf=io.BytesIO()
    with zipfile.ZipFile(zip_buf,'w') as z:
        for _,n in df_notas.iterrows():
            z.writestr(f"xmls/NFe_{int(n['Numero'])}.xml", f"<NFe>{n['ChaveAcesso']}</NFe>")
        z.writestr("relatorios/Relatorio_Completo.xlsx", excel.getvalue())
    zip_buf.seek(0)
    
    st.download_button("⬇️ Baixar ZIP (XMLs + Relatório Excel Completo)", zip_buf, "Pacote_NFe_Completo.zip")
