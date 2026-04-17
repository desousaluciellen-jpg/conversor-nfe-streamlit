import streamlit as st
import openpyxl
from collections import defaultdict
from datetime import datetime
import io
import zipfile

st.set_page_config(page_title="Conversor NFe 2.0", page_icon="📄", layout="centered")

def extrair_dados_planilha(arquivo_xlsx):
    wb = openpyxl.load_workbook(arquivo_xlsx)
    ws = wb.active
    
    headers = [ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)]
    
    notas = defaultdict(lambda: {
        'serie': None, 'numero': None, 'dataEmissao': None,
        'cnpjEmitente': None, 'itens': []
    })
    
    for row in range(2, ws.max_row + 1):
        row_data = {header: ws.cell(row=row, column=col).value for col, header in enumerate(headers, 1)}
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
                'quantidade': float(row_data.get('QuantidadeUnidadeComercial') or 0),
                'valorUnitario': float(row_data.get('ValorUnitarioComercial') or 0),
                'valorTotal': float(row_data.get('ValorTotalProduto') or 0),
                'valorDesconto': float(row_data.get('ValorDesconto') or 0),
                'icmsTag': row_data.get('TipoIcmsTag'),
                'icmsTributacao': row_data.get('IcmsTributacao'),
            }
            notas[chave]['itens'].append(item)
    return notas

def gerar_xml(chave, nota):
    data = nota['dataEmissao']
    data_str = data.strftime('%Y-%m-%d') if isinstance(data, datetime) else str(data)[:10]
    
    xml_str = f'<?xml version="1.0" encoding="utf-8"?><nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00"><NFe xmlns="http://www.portalfiscal.inf.br/nfe"><infNFe Id="NFe{chave}" versao="4.00"><ide><cUF>42</cUF><cNF>{str(nota["numero"]).zfill(8)}</cNF><natOp>VENDA DE MERCADORIAS</natOp><mod>65</mod><serie>{nota["serie"]}</serie><nNF>{nota["numero"]}</nNF><dhEmi>{data_str}T20:10:12-03:00</dhEmi><tpNF>1</tpNF><idDest>1</idDest><cMunFG>4209102</cMunFG><tpImp>4</tpImp><tpEmis>1</tpEmis><cDV>6</cDV><tpAmb>1</tpAmb><finNFe>1</finNFe><indFinal>1</indFinal><indPres>1</indPres><procEmi>0</procEmi><verProc>1.00</verProc></ide><emit><CNPJ>{nota["cnpjEmitente"]}</CNPJ><xNome>DNA RESTAURANTE LTDA</xNome><xFant>DNA SUSHI</xFant><enderEmit><xLgr>Rua Monsenhor Gercino</xLgr><nro>1285</nro><xBairro>Itaum</xBairro><cMun>4209102</cMun><xMun>Joinville</xMun><UF>SC</UF><CEP>89210009</CEP><cPais>1058</cPais><xPais>BRASIL</xPais><fone>47991986628</fone></enderEmit><IE>262605775</IE><CRT>1</CRT></emit>'
    
    total_valor = 0
    total_desconto = 0
    for idx, item in enumerate(nota['itens'], 1):
        total_valor += item['valorTotal']
        total_desconto += item['valorDesconto']
        icms_tag = item.get('icmsTag', 'ICMSSN500')
        
        if icms_tag == 'ICMSSN102':
            icms_xml = '<ICMS><ICMSSN102><orig>0</orig><CSOSN>102</CSOSN></ICMSSN102></ICMS>'
            pis_xml = '<PIS><PISOutr><CST>49</CST><qBCProd>0.0000</qBCProd><vAliqProd>0.0000</vAliqProd><vPIS>0.00</vPIS></PISOutr></PIS>'
            cofins_xml = '<COFINS><COFINSOutr><CST>49</CST><vBC>0.00</vBC><pCOFINS>0.00</pCOFINS><vCOFINS>0.00</vCOFINS></COFINSOutr></COFINS>'
        else:
            icms_xml = '<ICMS><ICMSSN500><orig>0</orig><CSOSN>500</CSOSN><vBCSTRet>0.00</vBCSTRet><pST>0.00</pST><vICMSSTRet>0.00</vICMSSTRet></ICMSSN500></ICMS>'
            pis_xml = '<PIS><PISNT><CST>04</CST></PISNT></PIS>'
            cofins_xml = '<COFINS><COFINSNT><CST>04</CST></COFINSNT></COFINS>'
        
        desconto_xml = f'<vDesc>{item["valorDesconto"]:.2f}</vDesc>' if item['valorDesconto'] > 0 else ''
        xml_str += f'<det nItem="{idx}"><prod><cProd>{item["codigo"]}</cProd><cEAN>SEM GTIN</cEAN><xProd>{item["descricao"]}</xProd><NCM>{item["ncm"]}</NCM><CFOP>{item["cfop"]}</CFOP><uCom>UN</uCom><qCom>{item["quantidade"]:.2f}</qCom><vUnCom>{item["valorUnitario"]:.2f}</vUnCom><vProd>{item["valorTotal"]:.2f}</vProd><cEANTrib>SEM GTIN</cEANTrib><uTrib>UN</uTrib><qTrib>{item["quantidade"]:.2f}</qTrib><vUnTrib>{item["valorUnitario"]:.2f}</vUnTrib>{desconto_xml}<indTot>1</indTot></prod><imposto>{icms_xml}{pis_xml}{cofins_xml}</imposto></det>'

    total_liquido = total_valor - total_desconto
    xml_str += f'<total><ICMSTot><vBC>0.00</vBC><vICMS>0.00</vICMS><vICMSDeson>0.00</vICMSDeson><vFCP>0.00</vFCP><vBCST>0.00</vBCST><vST>0.00</vST><vFCPST>0.00</vFCPST><vFCPSTRet>0.00</vFCPSTRet><vProd>{total_valor:.2f}</vProd><vFrete>0.00</vFrete><vSeg>0.00</vSeg><vDesc>{total_desconto:.2f}</vDesc><vII>0.00</vII><vIPI>0.00</vIPI><vIPIDevol>0.00</vIPIDevol><vPIS>0.00</vPIS><vCOFINS>0.00</vCOFINS><vOutro>0.00</vOutro><vNF>{total_liquido:.2f}</vNF></ICMSTot></total><transp><modFrete>9</modFrete></transp><pag><detPag><tPag>04</tPag><vPag>{total_liquido:.2f}</vPag><card><tpIntegra>2</tpIntegra></card></detPag></pag></infNFe></NFe></nfeProc>'
    return xml_str

st.title("📄 Conversor NFe 2.0")
st.write("Envie sua planilha .xlsx e receba os XMLs da NFe prontos para download.")

uploaded = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])

if uploaded:
    with st.spinner("Processando..."):
        notas = extrair_dados_planilha(uploaded)
        st.success(f"{len(notas)} notas encontradas")
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for chave, nota in sorted(notas.items()):
                xml = gerar_xml(chave, nota)
                filename = f"NFe_{nota['numero']}_Seq{nota['serie']}.xml"
                zipf.writestr(filename, xml)
        
        zip_buffer.seek(0)
        st.download_button(
            label="⬇️ Baixar ZIP com XMLs",
            data=zip_buffer,
            file_name="nfe_xmls.zip",
            mime="application/zip"
        )
        
        with st.expander("Ver detalhes"):
            for chave, nota in list(notas.items())[:5]:
                st.write(f"**NFe {nota['numero']}** - Série {nota['serie']} - {len(nota['itens'])} itens")
else:
    st.info("Faça upload da planilha para começar.")
