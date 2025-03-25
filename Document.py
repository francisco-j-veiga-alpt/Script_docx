import os
from docx import Document
import re
import openpyxl

def substituir_texto(arquivo, texto_antigo, texto_novo):
    doc = Document(arquivo)

    def substituir_em_runs(runs):
        texto_completo = "".join(run.text for run in runs)
        novo_texto = re.sub(re.escape(texto_antigo), texto_novo, texto_completo, flags=re.IGNORECASE)
        
        if texto_completo != novo_texto:
            for i, run in enumerate(runs):
                if i == 0:
                    run.text = novo_texto
                else:
                    run.text = ""

    for paragrafo in doc.paragraphs:
        substituir_em_runs(paragrafo.runs)

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    substituir_em_runs(paragrafo.runs)
    
    doc.save(arquivo)


def processar_texto(arquivo_xlsx, caminho_arquivos=None):
    workbook = openpyxl.load_workbook(arquivo_xlsx)
    sheet = workbook.active

    substituicoes = []
    for row in sheet.iter_rows(min_row=1, values_only=True):
        if len(row) >= 2:
            texto_antigo, texto_novo = row[0], row[1]
            substituicoes.append((texto_antigo, texto_novo))

    if caminho_arquivos:
        for arquivo in os.listdir(caminho_arquivos):
            if arquivo.endswith(".docx"):
                caminho_completo = os.path.join(caminho_arquivos, arquivo)
                for texto_antigo, texto_novo in substituicoes:
                    substituir_texto(caminho_completo, texto_antigo, texto_novo)
    else:
        print("Nenhum caminho de arquivos especificado.")

    workbook.close()



if __name__ == "__main__":
    
    print("Bem-vindo ao programa de substituição de texto em documentos Word!")
    
    # Perguntar ao utilizador pelo caminho do arquivo Excel
    while True:
        arquivo_xlsx = input("Digite o caminho completo do arquivo Excel (.xlsx): ").strip()
        if arquivo_xlsx.lower().endswith('.xlsx') and os.path.isfile(arquivo_xlsx):
            break
        else:
            print("Erro: O arquivo deve ter a extensão .xlsx e existir. Tente novamente.")

    # Perguntar ao utilizador pelo caminho da pasta com os documentos Word
    while True:
        pasta_docx = input("Digite o caminho da pasta contendo os documentos Word: ").strip()
        if os.path.isdir(pasta_docx):
            break
        else:
            print("Erro: O caminho especificado não é uma pasta válida. Tente novamente.")

    print(f"\nProcessando arquivos na pasta: {pasta_docx}")
    print(f"Usando substituições do arquivo: {arquivo_xlsx}\n")

    arquivos_processados = processar_texto(arquivo_xlsx, pasta_docx)

    print(f"\nProcessamento concluído!")
    print(f"Total de arquivos processados: {arquivos_processados}")