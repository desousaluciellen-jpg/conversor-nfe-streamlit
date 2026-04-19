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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import base64

# ===== CÓDIGO ORIGINAL DE CONVERSÃO =====
def to_float(valor):
    if valor is None or valor == '': return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    try: return float(str(valor).strip().replace('.', '').replace(',', '.'))
    except: return 0.0

def extrair_dados_planilha(arquivo_xlsx):
    wb = openpyxl.load_workbook(arquivo_xlsx, data_only=True)
    ws = wb.active
    headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
    notas = defaultdict(lambda: {'serie':None,'numero':None,'dataEmissao':None,'cnpjEmitente':None,'itens':[]})
    emitente_info = {}
    for r in range(2, ws.max_row+1):
        row = {h: ws.cell(r,c).value for c,h in enumerate(headers,1)}
        chave = row.get('ChaveAcesso')
        if chave and row.get('DescricaoProduto'):
            if not emitente_info:
                emitente_info = {
                    'cnpj': str(row.get('CnpjEmitente') or ''),
                    'razao': str(row.get('RazaoSocialEmitente') or 'Não informado'),
                    'ie': str(row.get('InscricaoEstadualEmitente') or ''),
                }
            notas[chave]['serie'] = row.get('SerieDocumento')
            notas[chave]['numero'] = row.get('NumeroDocumento')
            notas[chave]['dataEmissao'] = row.get('DataEmissaoNfe')
            notas[chave]['cnpjEmitente'] = row.get('CnpjEmitente')
            notas[chave]['itens'].append({
                'codigo': row.get('Produto'),
                'descricao': row.get('DescricaoProduto'),
                'ncm': row.get('NcmProduto'),
                'cfop': row.get('CfopProduto'),
                'quantidade': to_float(row.get('QuantidadeUnidadeComercial')),
                'valorUnitario': to_float(row.get('ValorUnitarioComercial')),
                'valorTotal': to_float(row.get('ValorTotalProduto')),
                'valorFrete': to_float(row.get('ValorFrete')),
                'valorSeguro': to_float(row.get('ValorSeguro')),
                'valorDesconto': to_float(row.get('ValorDesconto')),
                'valorOutras': to_float(row.get('ValorOutrasDespesas')),
                'valorIcmsBc': to_float(row.get('ValorIcmsBc')), # <— NOVO
                'valorIcms': to_float(row.get('ValorIcms')), # <— NOVO
                'icmsTributacao': row.get('IcmsTributacao'), # <— CST
                'icmsTag': row.get('TipoIcmsTag'),
            })
    return notas, emitente_info

def gerar_xml(chave, nota):
    #... (sua função gerar_xml permanece exatamente igual)
    data = nota['dataEmissao']
    data_str = data.strftime('%Y-%m-%d') if isinstance(data, datetime) else str(data)[:10]
    xml_str = f'''<?xml version="1.0" encoding="utf-8"?><nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00"><NFe xmlns="http://www.portalfiscal.inf.br/nfe"><infNFe Id="NFe{chave}" versao="4.00"><ide><cUF>42</cUF><cNF>{str(nota['numero']).zfill(8)}</cNF><natOp>VENDA</natOp><mod>65</mod><serie>{nota['serie']}</serie><nNF>{nota['numero']}</nNF><dhEmi>{data_str}T20:10:12-03:00</dhEmi><tpNF>1</tpNF><idDest>1</idDest><cMunFG>4209102</cMunFG><tpImp>4</tpImp><tpEmis>1</tpEmis><cDV>6</cDV><tpAmb>1</tpAmb><finNFe>1</finNFe><indFinal>1</indFinal><indPres>1</indPres><procEmi>0</procEmi><verProc>1.00</verProc></ide><emit><CNPJ>{nota['cnpjEmitente']}</CNPJ><xNome>DNA RESTAURANTE LTDA</xNome></emit>'''
    total_prod = total_frete = total_seg = total_desc = total_outro = 0
    for idx, item in enumerate(nota['itens'],1):
        total_prod+=item['valorTotal']; total_frete+=item['valorFrete']; total_seg+=item['valorSeguro']; total_desc+=item['valorDesconto']; total_outro+=item['valorOutras']
        icms_xml = '<ICMS><ICMSSN500><orig>0</orig><CSOSN>500</CSOSN></ICMSSN500></ICMS>'
        pis_xml = '<PIS><PISNT><CST>04</CST></PISNT></PIS>'; cofins_xml = '<COFINS><COFINSNT><CST>04</CST></COFINSNT></COFINS>'
        prod_xml = f'''<prod><cProd>{item['codigo']}</cProd><xProd>{item['descricao']}</xProd><NCM>{item['ncm']}</NCM><CFOP>{item['cfop']}</CFOP><qCom>{item['quantidade']:.2f}</qCom><vUnCom>{item['valorUnitario']:.2f}</vUnCom><vProd>{item['valorTotal']:.2f}</vProd>'''
        if item['valorFrete']>0: prod_xml+=f'<vFrete>{item["valorFrete"]:.2f}</vFrete>'
        if item['valorSeguro']>0: prod_xml+=f'<vSeg>{item["valorSeguro"]:.2f}</vSeg>'
        if item['valorDesconto']>0: prod_xml+=f'<vDesc>{item["valorDesconto"]:.2f}</vDesc>'
        if item['valorOutras']>0: prod_xml+=f'<vOutro>{item["valorOutras"]:.2f}</vOutro>'
        prod_xml+='</prod>'; xml_str+=f'<det nItem="{idx}">{prod_xml}<imposto>{icms_xml}{pis_xml}{cofins_xml}</imposto></det>'
    vNF = total_prod+total_frete+total_seg+total_outro-total_desc
    xml_str+=f'''<total><ICMSTot><vProd>{total_prod:.2f}</vProd><vFrete>{total_frete:.2f}</vFrete><vSeg>{total_seg:.2f}</vSeg><vDesc>{total_desc:.2f}</vDesc><vOutro>{total_outro:.2f}</vOutro><vNF>{vNF:.2f}</vNF></ICMSTot></total></infNFe></NFe></nfeProc>'''
    return xml_str

# ===== INTERFACE =====
st.set_page_config(page_title="Central XML Fiscal", page_icon="📄", layout="wide")
try:
    with open("xml_icon.png","rb") as f: icon = f'<img src="data:image/png;base64,{base64.b64encode(f.read()).decode()}" width="64">'
except: icon = "📄"
st.markdown(f'<div style="background:linear-gradient(135deg,#1e3a8a,#3b82f6);padding:1.2rem;border-radius:12px;color:white;display:flex;gap:1rem;align-items:center">{icon}<div><h2 style="margin:0">Central XML Fiscal</h2><p style="margin:0">Importe sua planilha SAT e gere XMLs + relatórios</p></div></div>', unsafe_allow_html=True)

up = st.file_uploader("Planilha.xlsx", type="xlsx")
if up:
    notas, emitente = extrair_dados_planilha(up)
    itens = [i for n in notas.values() for i in n['itens']]
    df = pd.DataFrame(itens)
    total_bruto = df['valorTotal'].sum(); total_desc = df['valorDesconto'].sum()

    # --- Relatório Completo ---
    completo = []
    for chave, nota in sorted(notas.items()):
        for it in nota['itens']:
            completo.append({
                'Chave': chave,
                'Nº': nota['numero'],
                'Série': nota['serie'],
                'Data': nota['dataEmissao'],
                'CFOP': it['cfop'],
                'Produto': it['descricao'],
                'Qtd': it['quantidade'],
                'V.Un': it['valorUnitario'],
                'Bruto': it['valorTotal'],
                'Desc': it['valorDesconto'],
                'Frete': it['valorFrete'],
                'Seg': it['valorSeguro'],
                'Outras': it['valorOutras'],
                'Líquido': it['valorTotal'] + it['valorFrete'] + it['valorSeguro'] + it['valorOutras'] - it['valorDesconto'],
                'BC ICMS': it['valorIcmsBc'],
                'ICMS': it['valorIcms'],
                'CST': it['icmsTributacao'],
            })
    df_completo = pd.DataFrame(completo).sort_values(['CFOP','Nº'])

    # Métricas e gráficos (iguais)
    st.info(f"Emitente: {emitente['razao']} | Notas: {len(notas)}")
    resumo = df.groupby('cfop').agg(Bruto=('valorTotal','sum'),Desc=('valorDesconto','sum')).reset_index()
    resumo['Líquido'] = resumo['Bruto']-resumo['Desc']
    st.dataframe(resumo, use_container_width=True)

    # Excel com 3 abas
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine='openpyxl') as w:
        pd.DataFrame([emitente]).to_excel(w, sheet_name='Resumo', index=False)
        resumo.to_excel(w, sheet_name='Por CFOP', index=False)
        df_completo.to_excel(w, sheet_name='Completo', index=False) # <— NOVA ABA

    # ZIP com XMLs + Excel
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf,'w') as z:
        for chave, nota in notas.items():
            z.writestr(f"NFe_{nota['numero']}.xml", gerar_xml(chave, nota))
        z.writestr("Relatorio_NFe.xlsx", excel_buf.getvalue())

    st.download_button("⬇️ BAIXAR PACOTE", zip_buf.getvalue(), f"XMLs_{datetime.now():%Y%m%d}.zip")
