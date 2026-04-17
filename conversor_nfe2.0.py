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
    for r in range(2,ws.max_row+1):
        row={h:ws.cell(r,c).value for c,h in enumerate(headers,1)}
        chave=row.get('ChaveAcesso')
        if chave and row.get('DescricaoProduto'):
            notas[chave]['serie']=row.get('SerieDocumento')
            notas[chave]['numero']=row.get('NumeroDocumento')
            notas[chave]['dataEmissao']=row.get('DataEmissaoNfe')
            notas[chave]['itens'].append({
                'cfop':str(row.get('CfopProduto') or ''),
                'valorTotal':float(row.get('ValorTotalProduto') or 0),
                'valorDesconto':float(row.get('ValorDesconto') or 0),
            })
    return notas

def gerar_xml(chave,nota,emit):
    return f'<?xml version="1.0"?><NFe><infNFe Id="{chave}"><ide><nNF>{nota["numero"]}</nNF></ide></infNFe></NFe>'

def criar_excel(resumo_cfop, total_bruto,total_desc,total_liq, emitente, periodo):
    output=io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # aba resumo
        df_resumo=pd.DataFrame([{
            'Emitente':emitente['razao'],'CNPJ':emitente['cnpj'],'Período':periodo,
            'Total Bruto':total_bruto,'Descontos':total_desc,'Total Líquido':total_liq
        }])
        df_resumo.to_excel(writer, sheet_name='Resumo Geral', index=False)
        resumo_cfop.to_excel(writer, sheet_name='Por CFOP', index=False)
    output.seek(0); return output

def criar_pdf(resumo_cfop, total_bruto,total_desc,total_liq, emitente, periodo):
    buffer=io.BytesIO()
    doc=SimpleDocTemplate(buffer, pagesize=A4)
    styles=getSampleStyleSheet(); elems=[]
    elems.append(Paragraph(f"<b>Resumo Fiscal NFe</b>", styles['Title']))
    elems.append(Paragraph(f"Emitente: {emitente['razao']} - CNPJ: {emitente['cnpj']}<br/>Período: {periodo}", styles['Normal']))
    elems.append(Spacer(1,12))
    data=[['CFOP','Itens','Bruto','Desc.','Líquido']]
    for _,r in resumo_cfop.iterrows():
        data.append([r['cfop'], int(r['Qtde_Itens']), f"R$ {r['Valor_Bruto']:,.2f}", f"R$ {r['Descontos']:,.2f}", f"R$ {r['Valor_Liquido']:,.2f}"])
    data.append(['TOTAL','',f"R$ {total_bruto:,.2f}",f"R$ {total_desc:,.2f}",f"R$ {total_liq:,.2f}"])
    t=Table(data, colWidths=[60,50,90,70,90])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1e3a8a')),('TEXTCOLOR',(0,0),(-1,0),colors.white),
                           ('ALIGN',(1,0),(-1,-1),'RIGHT'),('GRID',(0,0),(-1,-1),0.5,colors.grey)]))
    elems.append(t)
    doc.build(elems); buffer.seek(0); return buffer

with st.sidebar:
    st.markdown("### Emitente")
    emitente={"cnpj":st.text_input("CNPJ","00.000.000/0001-91"),"razao":st.text_input("Razão","DNA RESTAURANTE LTDA")}

st.markdown('<div class="hero"><h1>📄 Conversor NFe 2.0 Pro</h1><p>XML + Excel + PDF + Gráfico por CFOP</p></div>', unsafe_allow_html=True)

up=st.file_uploader("Planilha .xlsx", type="xlsx")
if up:
    notas=extrair(up)
    itens=[i for n in notas.values() for i in n['itens']]
    df=pd.DataFrame(itens)
    total_bruto=df['valorTotal'].sum(); total_desc=df['valorDesconto'].sum(); total_liq=total_bruto-total_desc
    datas=[n['dataEmissao'] for n in notas.values() if n['dataEmissao']]
    periodo=f"{min(datas).strftime('%d/%m/%Y')} a {max(datas).strftime('%d/%m/%Y')}" if datas else "-"
    
    resumo=df.groupby('cfop').agg(Qtde_Itens=('valorTotal','count'),Valor_Bruto=('valorTotal','sum'),Descontos=('valorDesconto','sum')).reset_index()
    resumo['Valor_Liquido']=resumo['Valor_Bruto']-resumo['Descontos']
    resumo=resumo.rename(columns={'cfop':'cfop'})
    
    c1,c2,c3,c4=st.columns(4)
    for col,lab,val in zip([c1,c2,c3,c4],["Notas","Bruto","Descontos","Líquido"],[len(notas),total_bruto,total_desc,total_liq]):
        with col: st.markdown(f'<div class="metric-card"><div>{lab}</div><div class="metric-value">{"R$ {:,.2f}".format(val) if "R" not in lab and lab!="Notas" else (str(val) if lab=="Notas" else "R$ {:,.2f}".format(val))}</div></div>', unsafe_allow_html=True)
    
    colA,colB=st.columns([1.2,1])
    with colA:
        st.subheader("Resumo por CFOP")
        st.dataframe(resumo.style.format({'Valor_Bruto':'R$ {:,.2f}','Descontos':'R$ {:,.2f}','Valor_Liquido':'R$ {:,.2f}'}), use_container_width=True)
    with colB:
        st.subheader("Gráfico por CFOP")
        fig=px.pie(resumo, names='cfop', values='Valor_Liquido', hole=0.4, title="Distribuição do Valor Líquido")
        fig.update_traces(textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
    
    # Gerar arquivos
    excel_bytes=criar_excel(resumo,total_bruto,total_desc,total_liq,emitente,periodo)
    pdf_bytes=criar_pdf(resumo,total_bruto,total_desc,total_liq,emitente,periodo)
    
    zip_buf=io.BytesIO()
    with zipfile.ZipFile(zip_buf,'w',zipfile.ZIP_DEFLATED) as z:
        for chave,nota in notas.items():
            z.writestr(f"xmls/NFe_{nota['numero']}.xml", gerar_xml(chave,nota,emitente))
        z.writestr("resumo/Resumo_NFe.xlsx", excel_bytes.getvalue())
        z.writestr("resumo/Resumo_NFe.pdf", pdf_bytes.getvalue())
    zip_buf.seek(0)
    
    st.download_button("⬇️ BAIXAR PACOTE COMPLETO (XMLs + Excel + PDF)", zip_buf, f"Pacote_NFe_{datetime.now().strftime('%Y%m%d')}.zip", "application/zip")
else:
    st.info("Faça upload para gerar XMLs, Excel, PDF e gráfico.")
