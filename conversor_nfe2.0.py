import streamlit as st
import openpyxl
from collections import defaultdict
from datetime import datetime
import io
import zipfile
import pandas as pd
import plotly.express as px
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Conversor NFe 2.0 Pro", page_icon="📄", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
html,body{font-family:'Inter',sans-serif}
.hero{background:linear-gradient(135deg,#1e3a8a,#3b82f6);padding:2rem;border-radius:16px;color:white;margin-bottom:1.5rem}
.metric-card{background:white;padding:1rem;border-radius:12px;border:1px solid #e5e7eb;text-align:center}
.metric-value{font-size:1.6rem;font-weight:700;color:#1e3a8a}
div.stDownloadButton>button{background:#1e3a8a;color:white;border-radius:10px;font-weight:600;padding:0.8rem;width:100%}
</style>
""", unsafe_allow_html=True)

def extrair(arquivo):
    wb=openpyxl.load_workbook(arquivo,data_only=True); ws=wb.active
    headers=[ws.cell(1,c).value for c in range(1,ws.max_column+1)]
    notas=defaultdict(lambda:{'serie':None,'numero':None,'dataEmissao':None,'itens':[]})
    emitente_info = {}
    for r in range(2,ws.max_row+1):
        row={h:ws.cell(r,c).value for c,h in enumerate(headers,1)}
        chave=row.get('ChaveAcesso')
        if chave and row.get('DescricaoProduto'):
            # Captura dados do emitente da primeira nota válida
            if not emitente_info:
                emitente_info = {
                    'cnpj': str(row.get('CnpjEmitente') or ''),
                    'razao': str(row.get('RazaoSocialEmitente') or row.get('NomeEmitente') or 'Não informado'),
                    'fantasia': str(row.get('NomeFantasiaEmitente') or row.get('xFant') or ''),
                    'ie': str(row.get('InscricaoEstadualEmitente') or row.get('IE') or ''),
                    'cidade': str(row.get('CidadeEmitente') or row.get('MunicipioEmitente') or ''),
                    'uf': str(row.get('UFEmitente') or row.get('UF') or ''),
                }
            notas[chave]['serie']=row.get('SerieDocumento')
            notas[chave]['numero']=row.get('NumeroDocumento')
            notas[chave]['dataEmissao']=row.get('DataEmissaoNfe')
            notas[chave]['itens'].append({
                'cfop':str(row.get('CfopProduto') or ''),
                'valorTotal':float(row.get('ValorTotalProduto') or 0),
                'valorDesconto':float(row.get('ValorDesconto') or 0),
            })
    return notas, emitente_info

st.markdown('<div class="hero"><h1>📄 Conversor NFe 2.0 Pro</h1><p>XML + Excel + PDF + Gráfico — Emitente automático do Excel</p></div>', unsafe_allow_html=True)

up=st.file_uploader("Planilha .xlsx", type="xlsx")
if up:
    notas, emitente = extrair(up)
    if not notas:
        st.error("Nenhuma nota válida encontrada na planilha.")
        st.stop()
    
    itens=[i for n in notas.values() for i in n['itens']]
    df=pd.DataFrame(itens)
    total_bruto=df['valorTotal'].sum(); total_desc=df['valorDesconto'].sum(); total_liq=total_bruto-total_desc
    datas=[n['dataEmissao'] for n in notas.values() if n['dataEmissao']]
    datas_dt=[d if isinstance(d,datetime) else pd.to_datetime(d,errors='coerce') for d in datas]
    periodo=f"{min(datas_dt).strftime('%d/%m/%Y')} a {max(datas_dt).strftime('%d/%m/%Y')}" if datas_dt else "-"
    
    resumo=df.groupby('cfop').agg(Qtde_Itens=('valorTotal','count'),Valor_Bruto=('valorTotal','sum'),Descontos=('valorDesconto','sum')).reset_index()
    resumo['Valor_Liquido']=resumo['Valor_Bruto']-resumo['Descontos']
    
    # Info do emitente lida do Excel
    st.info(f"**Emitente detectado:** {emitente['razao']} • **CNPJ:** {emitente['cnpj']} • **IE:** {emitente['ie']} • **Período:** {periodo}", icon="🏢")
    
    c1,c2,c3,c4=st.columns(4)
    vals=[len(notas), total_bruto, total_desc, total_liq]
    labels=["Notas Emitidas","Total Bruto","Descontos","Total Líquido"]
    for col,lab,val in zip([c1,c2,c3,c4],labels,vals):
        with col:
            val_str = f"R$ {val:,.2f}" if lab!="Notas Emitidas" else str(int(val))
            st.markdown(f'<div class="metric-card"><div>{lab}</div><div class="metric-value">{val_str}</div></div>', unsafe_allow_html=True)
    
    colA,colB=st.columns([1.2,1])
    with colA:
        st.subheader("Resumo por CFOP")
        st.dataframe(resumo.style.format({'Valor_Bruto':'R$ {:,.2f}','Descontos':'R$ {:,.2f}','Valor_Liquido':'R$ {:,.2f}'}), use_container_width=True)
    with colB:
        st.subheader("Gráfico por CFOP")
        fig=px.pie(resumo, names='cfop', values='Valor_Liquido', hole=0.4, title="Distribuição Líquida")
        fig.update_traces(textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
    
    # Download
    excel_buf=io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
        pd.DataFrame([{'Emitente':emitente['razao'],'CNPJ':emitente['cnpj'],'Período':periodo,'Bruto':total_bruto,'Descontos':total_desc,'Líquido':total_liq}]).to_excel(writer, sheet_name='Resumo', index=False)
        resumo.to_excel(writer, sheet_name='Por CFOP', index=False)
    excel_buf.seek(0)
    
    pdf_buf=io.BytesIO()
    doc=SimpleDocTemplate(pdf_buf, pagesize=A4)
    styles=getSampleStyleSheet(); elems=[Paragraph(f"<b>Resumo Fiscal</b><br/>{emitente['razao']} - {emitente['cnpj']}", styles['Title'])]
    data=[['CFOP','Itens','Bruto','Desc.','Líquido']]
    for _,r in resumo.iterrows(): data.append([r['cfop'],int(r['Qtde_Itens']),f"R$ {r['Valor_Bruto']:,.2f}",f"R$ {r['Descontos']:,.2f}",f"R$ {r['Valor_Liquido']:,.2f}"])
    t=Table(data); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1e3a8a')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.5,colors.grey)]))
    elems.append(t); doc.build(elems); pdf_buf.seek(0)
    
    zip_buf=io.BytesIO()
    with zipfile.ZipFile(zip_buf,'w',zipfile.ZIP_DEFLATED) as z:
        for chave,nota in notas.items():
            z.writestr(f"xmls/NFe_{nota['numero']}.xml", f'<NFe><emit><CNPJ>{emitente["cnpj"]}</CNPJ></emit></NFe>')
        z.writestr("resumo/Resumo_NFe.xlsx", excel_buf.getvalue())
        z.writestr("resumo/Resumo_NFe.pdf", pdf_buf.getvalue())
    zip_buf.seek(0)
    
    st.download_button("⬇️ BAIXAR PACOTE COMPLETO", zip_buf, f"Pacote_NFe_{datetime.now().strftime('%Y%m%d')}.zip", "application/zip")
else:
    st.info("Envie a planilha. O emitente será lido automaticamente do arquivo.")
