import streamlit as st
import openpyxl
from datetime import datetime
import io
import zipfile
import pandas as pd
import plotly.express as px
import base64
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Central XML Fiscal", page_icon="📄", layout="wide")

# --- HERO ---
try:
    with open("xml_icon.png", "rb") as f:
        icon_b64 = base64.b64encode(f.read()).decode()
    icon_html = f'<img src="data:image/png;base64,{icon_b64}" width="72" height="72">'
except:
    icon_html = "📄"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
.hero{{background:linear-gradient(135deg,#1e3a8a,#3b82f6);padding:1.6rem 2rem;border-radius:16px;color:white;display:flex;align-items:center;gap:1rem;margin-bottom:1.5rem}}
.hero-text h1{{margin:0;font-size:2rem;font-weight:700}}
.hero-text p{{margin:.3rem 0 0;opacity:.92}}
.metric-card{{background:#f8fafc;padding:1rem;border-radius:12px;border:1px solid #e5e7eb;text-align:center}}
.metric-value{{font-size:1.6rem;font-weight:700;color:#1e3a8a}}
div.stDownloadButton>button{{background:#1e3a8a;color:white;border-radius:10px;font-weight:600;padding:0.8rem;width:100%}}
</style>
<div class="hero">{icon_html}<div class="hero-text"><h1>Central XML Fiscal</h1><p>Importe sua planilha de itens SAT e converta seus dados do Excel em arquivos XML prontos para o seu sistema contábil.</p></div></div>
""", unsafe_allow_html=True)

def extrair(arquivo):
    arquivo.seek(0)
    wb = openpyxl.load_workbook(arquivo, data_only=True)
    ws = wb.active
    headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]

    notas = {} # <-- agora é dict normal, não defaultdict
    emitente_info = {'cnpj':'','razao':'Não informado','fantasia':'','ie':'','cidade':'','uf':''}

    for r in range(2, ws.max_row+1):
        row = {h: ws.cell(r,c).value for c,h in enumerate(headers,1) if h}
        chave = row.get('ChaveAcesso')
        if chave and row.get('DescricaoProduto'):
            if chave not in notas:
                notas[chave] = {'serie':None,'numero':None,'dataEmissao':None,'itens':[]}
            if emitente_info['cnpj'] == '':
                emitente_info.update({
                    'cnpj': str(row.get('CnpjEmitente') or ''),
                    'razao': str(row.get('RazaoSocialEmitente') or row.get('NomeEmitente') or 'Não informado'),
                })
            notas[chave]['serie'] = row.get('SerieDocumento')
            notas[chave]['numero'] = row.get('NumeroDocumento')
            notas[chave]['dataEmissao'] = row.get('DataEmissaoNfe')
            notas[chave]['itens'].append({
                'cfop': str(row.get('CfopProduto') or ''),
                'valorTotal': float(row.get('ValorTotalProduto') or 0),
                'valorDesconto': float(row.get('ValorDesconto') or 0),
            })
    return notas, emitente_info

up = st.file_uploader("Planilha.xlsx", type=["xlsx"])
if up:
    notas, emitente = extrair(up)
    if not notas:
        st.error("Nenhuma nota válida encontrada.")
        st.stop()

    itens = [i for n in notas.values() for i in n['itens']]
    df = pd.DataFrame(itens)
    total_bruto = df['valorTotal'].sum()
    total_desc = df['valorDesconto'].sum()
    total_liq = total_bruto - total_desc

    datas = [n['dataEmissao'] for n in notas.values() if n['dataEmissao']]
    datas_dt = pd.to_datetime(datas, errors='coerce').dropna()
    periodo = f"{datas_dt.min():%d/%m/%Y} a {datas_dt.max():%d/%m/%Y}" if not datas_dt.empty else "-"

    resumo = df.groupby('cfop', as_index=False).agg(
        Qtde_Itens=('valorTotal','count'),
        Valor_Bruto=('valorTotal','sum'),
        Descontos=('valorDesconto','sum')
    )
    resumo['Valor_Liquido'] = resumo['Valor_Bruto'] - resumo['Descontos']

    st.info(f"**Emitente:** {emitente['razao']} • **CNPJ:** {emitente['cnpj']} • **Período:** {periodo}")

    c1,c2,c3,c4 = st.columns(4)
    for col,lab,val in zip([c1,c2,c3,c4], ["Notas","Total Bruto","Descontos","Total Líquido"], [len(notas), total_bruto, total_desc, total_liq]):
        val_str = f"R$ {val:,.2f}" if lab!="Notas" else str(int(val))
        col.markdown(f'<div class="metric-card"><div>{lab}</div><div class="metric-value">{val_str}</div></div>', unsafe_allow_html=True)

    colA,colB = st.columns([1.2,1])
    with colA:
        st.subheader("Resumo por CFOP")
        st.dataframe(resumo.style.format({'Valor_Bruto':'R$ {:,.2f}','Descontos':'R$ {:,.2f}','Valor_Liquido':'R$ {:,.2f}'}), use_container_width=True)
    with colB:
        fig = px.pie(resumo, names='cfop', values='Valor_Liquido', hole=0.4)
        st.plotly_chart(fig, use_container_width=True)

    # downloads
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine='openpyxl') as w:
        pd.DataFrame([{'Emitente':emitente['razao'],'CNPJ':emitente['cnpj'],'Período':periodo}]).to_excel(w, index=False, sheet_name='Resumo')
        resumo.to_excel(w, index=False, sheet_name='CFOP')

    pdf_buf = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buf, pagesize=A4)
    styles = getSampleStyleSheet()
    elems = [Paragraph(f"Central XML Fiscal - {emitente['razao']}", styles['Title'])]
    data = [['CFOP','Itens','Bruto','Desc','Líquido']] + [[r.cfop, int(r.Qtde_Itens), f"R$ {r.Valor_Bruto:,.2f}", f"R$ {r.Descontos:,.2f}", f"R$ {r.Valor_Liquido:,.2f}"] for r in resumo.itertuples()]
    t = Table(data); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1e3a8a')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.5,colors.grey)]))
    elems.append(t); doc.build(elems)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf,'w') as z:
        for n in notas.values():
            z.writestr(f"xmls/NFe_{n['numero'] or 'SN'}.xml", f"<NFe><emit><CNPJ>{emitente['cnpj']}</CNPJ></emit></NFe>")
        z.writestr("Resumo.xlsx", excel_buf.getvalue())
        z.writestr("Resumo.pdf", pdf_buf.getvalue())

    st.download_button("⬇️ BAIXAR PACOTE", zip_buf.getvalue(), f"XML_{datetime.now():%Y%m%d}.zip")
else:
    st.info("Envie a planilha.xlsx")
