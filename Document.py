import os
from docx import Document
import re

def substituir_texto(arquivo, texto_antigo, texto_novo, pasta_saida):
    doc = Document(arquivo)
    texto_antigo_lower = texto_antigo.lower()

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

    nome_relativo = os.path.relpath(arquivo, start=pasta_base)
    caminho_saida = os.path.join(pasta_saida, nome_relativo)
    pasta_destino = os.path.dirname(caminho_saida)
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    caminho_saida = caminho_saida.replace(".docx", "_modificado.docx")
    
    doc.save(caminho_saida)
    print(f"Substituição concluída: {caminho_saida}")

if __name__ == "__main__":
    # Diretório base
    pasta_base = os.path.join(os.path.expanduser("~"), "Documents", "change")

    if not os.path.exists(pasta_base):
        print(f"Erro: A pasta '{pasta_base}' não existe.")
        exit(1)

    # Perguntar o nome da pasta de saída
    nome_pasta_saida = input("Digite o nome da pasta de saída: ").strip()
    pasta_saida = os.path.join(pasta_base, nome_pasta_saida)

    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)
        print(f"Pasta de saída criada: {pasta_saida}")

    texto_antigo = input("Digite a palavra ou frase que deseja substituir: ").strip()
    texto_novo = input("Digite a nova palavra ou frase: ").strip()

    arquivos_encontrados = []

    # Busca recursiva por .docx
    for raiz, _, arquivos in os.walk(pasta_base):
        for arquivo in arquivos:
            if arquivo.endswith(".docx") and not arquivo.endswith("_modificado.docx"):
                caminho_completo = os.path.join(raiz, arquivo)
                arquivos_encontrados.append(caminho_completo)

    if not arquivos_encontrados:
        print("Nenhum arquivo .docx encontrado na pasta 'Documents/change'.")
        exit(1)

    for caminho_arquivo in arquivos_encontrados:
        substituir_texto(caminho_arquivo, texto_antigo, texto_novo, pasta_saida)

    print("Processamento concluído para todos os arquivos.")
