#!/usr/bin/env python3
import streamlit as st
import openpyxl
from collections import defaultdict
from datetime import datetime
import io
import zipfile
import base64

# ======= SEU CÓDIGO ORIGINAL (INTACTO) =======
def to_float(valor):
    if valor is None or valor == '':
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        s = str(valor).strip().replace('.', '').replace(',', '.')
        return float(s)
    except:
        return 0.0

def extrair_dados_planilha(arquivo_xlsx):
    wb = openpyxl.load_workbook(arquivo_xlsx)
    ws = wb.active
    headers = []
    for col in range(1, ws.max_column + 1):
        headers.append(ws.cell(row=1, column=col).value)
    notas = defaultdict(lambda: {
        'serie': None,
        'numero': None,
        'dataEmissao': None,
        'cnpjEmitente': None,
        'itens': []
    })
    for row in range(2, ws.max_row + 1):
        row_data = {}
        for col, header in enumerate(headers, 1):
            row_data[header] = ws.cell(row=row, column=col).value
        chave = row_data.get('ChaveAcesso')
        if chave and row_data.get('DescricaoProduto'):
            notas[chave]['serie'] = row_data.get('SerieDocumento')
            notas[chave]['numero'] = row_data.get('NumeroDocumento')
            notas[chave]['dataEmissao'] = row_data.get('DataEmissaoNfe')
            notas[chave]['cnpjEmitente'] = row_data.get('CnpjEmitente')
            item = {
                'numero': row_data.get('Item'),
                'codigo': row_data.get('Produto'),
                'descricao': row_data.get('DescricaoProduto'),
                'ncm': row_data.get('NcmProduto'),
                'cfop': row_data.get('CfopProduto'),
                'quantidade': to_float(row_data.get('QuantidadeUnidadeComercial')),
                'valorUnitario': to_float(row_data.get('ValorUnitarioComercial')),
                'valorTotal': to_float(row_data.get('ValorTotalProduto')),
                'valorFrete': to_float(row_data.get('ValorFrete')),
                'valorSeguro': to_float(row_data.get('ValorSeguro')),
                'valorDesconto': to_float(row_data.get('ValorDesconto')),
                'valorOutras': to_float(row_data.get('ValorOutrasDespesas')),
                'icmsTag': row_data.get('TipoIcmsTag'),
                'icmsTributacao': row_data.get('IcmsTributacao'),
            }
            notas[chave]['itens'].append(item)
    return notas

def gerar_xml(chave, nota):
    data = nota['dataEmissao']
    if isinstance(data, datetime):
        data_str = data.strftime('%Y-%m-%d')
    else:
        data_str = str(data)[:10]
    xml_str = f'''<?xml version="1.0" encoding="utf-8"?><nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00"><NFe xmlns="http://www.portalfiscal.inf.br/nfe"><infNFe Id="NFe{chave}" versao="4.00"><ide><cUF>42</cUF><cNF>{str(nota['numero']).zfill(8)}</cNF><natOp>VENDA DE MERCADORIAS</natOp><mod>65</mod><serie>{nota['serie']}</serie><nNF>{nota['numero']}</nNF><dhEmi>{data_str}T20:10:12-03:00</dhEmi><tpNF>1</tpNF><idDest>1</idDest><cMunFG>4209102</cMunFG><tpImp>4</tpImp><tpEmis>1</tpEmis><cDV>6</cDV><tpAmb>1</tpAmb><finNFe>1</finNFe><indFinal>1</indFinal><indPres>1</indPres><procEmi>0</procEmi><verProc>1.00</verProc></ide><emit><CNPJ>{nota['cnpjEmitente']}</CNPJ><xNome>DNA RESTAURANTE LTDA</xNome><xFant>DNA SUSHI</xFant><enderEmit><xLgr>Rua Monsenhor Gercino</xLgr><nro>1285</nro><xBairro>Itaum</xBairro><cMun>4209102</cMun><xMun>Joinville</xMun><UF>SC</UF><CEP>89210009</CEP><cPais>1058</cPais><xPais>BRASIL</xPais><fone>47991986628</fone></enderEmit><IE>262605775</IE><CRT>1</CRT></emit>'''
    total_prod = 0; total_frete = 0; total_seg = 0; total_desc = 0; total_outro = 0
    for idx, item in enumerate(nota['itens'], 1):
        total_prod += item['valorTotal']; total_frete += item.get('valorFrete', 0); total_seg += item.get('valorSeguro', 0); total_desc += item.get('valorDesconto', 0); total_outro += item.get('valorOutras', 0)
        icms_tag = item.get('icmsTag', 'ICMSSN500')
        if icms_tag == 'ICMSSN102':
            icms_xml = f'''<ICMS><ICMSSN102><orig>0</orig><CSOSN>102</CSOSN></ICMSSN102></ICMS>'''
            pis_xml = f'''<PIS><PISOutr><CST>49</CST><qBCProd>0.0000</qBCProd><vAliqProd>0.0000</vAliqProd><vPIS>0.00</vPIS></PISOutr></PIS>'''
            cofins_xml = f'''<COFINS><COFINSOutr><CST>49</CST><vBC>0.00</vBC><pCOFINS>0.00</pCOFINS><vCOFINS>0.00</vCOFINS></COFINSOutr></COFINS>'''
        else:
            icms_xml = f'''<ICMS><ICMSSN500><orig>0</orig><CSOSN>500</CSOSN><vBCSTRet>0.00</vBCSTRet><pST>0.00</pST><vICMSSTRet>0.00</vICMSSTRet></ICMSSN500></ICMS>'''
            pis_xml = f'''<PIS><PISNT><CST>04</CST></PISNT></PIS>'''
            cofins_xml = f'''<COFINS><COFINSNT><CST>04</CST></COFINSNT></COFINS>'''
        prod_xml = f'''<prod><cProd>{item['codigo']}</cProd><cEAN>SEM GTIN</cEAN><xProd>{item['descricao']}</xProd><NCM>{item['ncm']}</NCM><CFOP>{item['cfop']}</CFOP><uCom>UN</uCom><qCom>{item['quantidade']:.2f}</qCom><vUnCom>{item['valorUnitario']:.2f}</vUnCom><vProd>{item['valorTotal']:.2f}</vProd><cEANTrib>SEM GTIN</cEANTrib><uTrib>UN</uTrib><qTrib>{item['quantidade']:.2f}</qTrib><vUnTrib>{item['valorUnitario']:.2f}</vUnTrib>'''
        if item.get('valorFrete', 0) > 0: prod_xml += f'''<vFrete>{item['valorFrete']:.2f}</vFrete>'''
        if item.get('valorSeguro', 0) > 0: prod_xml += f'''<vSeg>{item['valorSeguro']:.2f}</vSeg>'''
        if item.get('valorDesconto', 0) > 0: prod_xml += f'''<vDesc>{item['valorDesconto']:.2f}</vDesc>'''
        if item.get('valorOutras', 0) > 0: prod_xml += f'''<vOutro>{item['valorOutras']:.2f}</vOutro>'''
        prod_xml += f'''<indTot>1</indTot></prod>'''
        xml_str += f'''<det nItem="{idx}">{prod_xml}<imposto>{icms_xml}{pis_xml}{cofins_xml}</imposto></det>'''
    vNF = total_prod + total_frete + total_seg + total_outro - total_desc
    xml_str += f'''<total><ICMSTot><vBC>0.00</vBC><vICMS>0.00</vICMS><vICMSDeson>0.00</vICMSDeson><vFCP>0.00</vFCP><vBCST>0.00</vBCST><vST>0.00</vST><vFCPST>0.00</vFCPST><vFCPSTRet>0.00</vFCPSTRet><vProd>{total_prod:.2f}</vProd><vFrete>{total_frete:.2f}</vFrete><vSeg>{total_seg:.2f}</vSeg><vDesc>{total_desc:.2f}</vDesc><vII>0.00</vII><vIPI>0.00</vIPI><vIPIDevol>0.00</vIPIDevol><vPIS>0.00</vPIS><vCOFINS>0.00</vCOFINS><vOutro>{total_outro:.2f}</vOutro><vNF>{vNF:.2f}</vNF></ICMSTot></total><transp><modFrete>9</modFrete></transp><pag><detPag><tPag>04</tPag><vPag>{vNF:.2f}</vPag><card><tpIntegra>2</tpIntegra></card></detPag></pag></infNFe></NFe></nfeProc>'''
    return xml_str
# ======= FIM DO SEU CÓDIGO ORIGINAL =======

# --- INTERFACE NOVA ---
st.set_page_config(page_title="Central XML Fiscal", page_icon="📄", layout="centered")

try:
    with open("xml_icon.png","rb") as f:
        icon_b64 = base64.b64encode(f.read()).decode()
    icon = f'<img src="data:image/png;base64,{icon_b64}" width="64" height="64">'
except:
    icon = "📄"

st.markdown(f"""
<style>
.hero{{background:linear-gradient(135deg,#1e3a8a,#3b82f6);padding:1.5rem 2rem;border-radius:16px;color:white;display:flex;align-items:center;gap:1rem;margin-bottom:1.5rem}}
.hero h1{{margin:0;font-size:1.8rem}}.hero p{{margin:.2rem 0 0;opacity:.9}}
</style>
<div class="hero">{icon}<div><h1>Central XML Fiscal</h1><p>Importe sua planilha de itens SAT e converta seus dados do Excel em arquivos XML prontos para o seu sistema contábil.</p></div></div>
""", unsafe_allow_html=True)

up = st.file_uploader("Selecione a planilha.xlsx", type=["xlsx"])

if up:
    with st.spinner("Processando planilha..."):
        notas = extrair_dados_planilha(up)

    if not notas:
        st.error("Nenhuma nota encontrada na planilha.")
    else:
        st.success(f"✓ {len(notas)} notas encontradas")

        # Gera ZIP em memória usando SUA função gerar_xml
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for chave, nota in sorted(notas.items()):
                xml_content = gerar_xml(chave, nota)
                filename = f"NFe_{nota['numero']}_Seq{nota['serie']}.xml"
                zipf.writestr(filename, xml_content)

        st.download_button(
            "⬇️ BAIXAR XMLs (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"XMLs_NFe_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            use_container_width=True
        )
else:
    st.info("Envie o arquivo Excel exportado do seu sistema.")
