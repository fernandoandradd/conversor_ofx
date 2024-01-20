import xml.etree.ElementTree as ET
import openpyxl
import streamlit as st
from tkinter import Tk
from tkinter import filedialog
from datetime import datetime
import os
from lxml import etree
from bs4 import BeautifulSoup
from ofxparse import OfxParser
import io

def analisar_ofx(conteudo_ofx):
    indice_inicio = conteudo_ofx.find(b'<OFX>')

    if indice_inicio == -1:
        st.error("Erro: Não foi possível encontrar o início dos dados XML no arquivo OFX.")
        return None

    conteudo_xml = conteudo_ofx[indice_inicio:]

    try:
        raiz = ET.fromstring(conteudo_xml)
    except ET.ParseError as e:
        st.error(f"Erro ao analisar o conteúdo OFX: {e}")
        return None

    dados = []

    for stmttrn in raiz.findall(".//STMTTRN"):
        trntype = stmttrn.find("TRNTYPE").text.replace("CREDIT", "CREDITO").replace("DEBIT", "DEBITO")
        dtposted = stmttrn.find("DTPOSTED").text
        dtposted_formatada = datetime.strptime(dtposted[:8], "%Y%m%d").strftime("%d/%m/%Y")
        trnamt = stmttrn.find("TRNAMT").text.replace(".", ",").replace("-", "")
        memo = stmttrn.find("MEMO").text

        dados.append([trntype, dtposted_formatada, trnamt, memo])

    return dados

def converter_ofx_para_excel(conteudo_ofx):
    dados = analisar_ofx(conteudo_ofx)

    if dados is None:
        return None

    planilha_excel = openpyxl.Workbook()
    planilha_ativa = planilha_excel.active

    cabecalho = ["TIPO TRANSAÇÃO", "DATA TRANSAÇÃO", "VALOR", "HISTÓRICO"]
    planilha_ativa.append(cabecalho)

    for linha in dados:
        planilha_ativa.append(linha)

    return planilha_excel

def save_excel(planilha_excel, file_name='output.xlsx'):
    buffer = io.BytesIO()
    planilha_excel.save(buffer)
    buffer.seek(0)
    return buffer
# -------------------------
def preprocessar_ofx_bb(conteudo_ofx_bb):
    # Substituir quebras de linha nas tags de data específicas do BB
    conteudo_ofx_bb_corrigido = (
        conteudo_ofx_bb.replace(b'<DTSERVER>', b'</DTSERVER>')

    )
    return conteudo_ofx_bb_corrigido


def analisar_ofx_bb(conteudo_ofx_bb):
    try:
        # Cria um objeto BytesIO para simular um arquivo
        ofx_file = io.BytesIO(conteudo_ofx_bb)

        # Parse do arquivo OFX
        ofx_obj = OfxParser.parse(ofx_file)

    except Exception as e:
        print(f"Erro ao analisar o conteúdo OFX: {e}")
        return None

    dados_bb = []



    for stmttrn in ofx_obj.account.statement.transactions:
        print(stmttrn[1])
        trntype_bb = stmttrn.type.replace("CREDIT", "CREDITO").replace("DEBIT", "DEBITO")
        dtposted_formatada_bb = stmttrn.date.strftime("%d/%m/%Y")
        trnamt_bb = str(stmttrn.amount)
        memo_bb = stmttrn.payee


        dados_bb.append([trntype_bb, dtposted_formatada_bb, trnamt_bb, memo_bb])

    return dados_bb



def converter_ofx_para_excel_bb(conteudo_ofx_bb):
    dados_bb = analisar_ofx_bb(conteudo_ofx_bb)


    if dados_bb is None:
        return None

    planilha_excel_bb = openpyxl.Workbook()
    planilha_ativa_bb = planilha_excel_bb.active


    cabecalho_bb = ["TIPO TRANSAÇÃO", "DATA TRANSAÇÃO", "VALOR", "HISTÓRICO"]
    planilha_ativa_bb.append(cabecalho_bb)

    for linha in dados_bb:
        planilha_ativa_bb.append(linha)

    return planilha_excel_bb

def obter_caminho_salvar_bb():

    caminho_salvar_bb = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    root = Tk()  # Criar uma instância de Tk para poder chamar destroy()
    root.destroy()
    return caminho_salvar_bb

# ------------------------------------------

def generate_unique_key(base_key, suffix):
    return f"{base_key}_{suffix}"

def main():
    st.set_page_config(page_title="Conversor OFX", layout="centered")
    st.image("LOGO.jpg", use_column_width=True)

    st.title("Conversor OFX para Excel")

    st.subheader(":blue[Converter OFX Banco do Nordeste - BNB]")

    caminho_arquivo = st.file_uploader("Arraste o arquivo para baixo ou clique para Selecionar o arquivo OFX - BANCO DO NORDESTE", type=["ofx"], help="Clique para Selecionar o arquivo OFX")

    if caminho_arquivo:
        conteudo_ofx = caminho_arquivo.read()

        planilha_excel = converter_ofx_para_excel(conteudo_ofx)


        if planilha_excel is not None:
            st.success("Conversão concluída. Agora você pode escolher onde salvar o arquivo Excel.")

            planilha_excel = save_excel(planilha_excel)

            st.download_button(label='📥 Download Current Result',
                               data=planilha_excel,
                               file_name='df_test.xlsx')



    st.subheader(":blue[Converter OFX Banco do Brasil - BB]")

    caminho_arquivo_bb = st.file_uploader("Arraste o arquivo para baixo ou clique para Selecionar o arquivo OFX - BANCO DO BRASIL", type=["ofx"], help="Clique para Selecionar o arquivo OFX")

    if caminho_arquivo_bb:
        conteudo_ofx_bb = caminho_arquivo_bb.read()

        planilha_excel_bb = converter_ofx_para_excel_bb(conteudo_ofx_bb)

        if planilha_excel_bb is not None:
            st.success("Conversão concluída. Agora você pode escolher onde salvar o arquivo Excel.")

            if st.button("Exportar para Excel - BB"):
                caminho_salvar_bb = obter_caminho_salvar_bb()

                if caminho_salvar_bb:
                    planilha_excel_bb.save(caminho_salvar_bb)
                    st.success(f"Arquivo Excel salvo em {caminho_salvar_bb}")

if __name__ == "__main__":
    main()
