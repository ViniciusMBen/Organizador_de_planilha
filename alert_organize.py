import pandas as pd
import os
from datetime import datetime

def encontrar_arquivo_recente(extensao=".xlsx"):
    """
    Encontra o arquivo mais recente com a extensão especificada na pasta atual.

    :param extensao: Extensão dos arquivos a serem considerados (default: ".xlsx").
    :return: Nome do arquivo mais recente ou None se não houver arquivos.
    """
    arquivos = [f for f in os.listdir() if f.endswith(extensao)]
    if not arquivos:
        return None
    arquivos.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return arquivos[0]

def processar_excel(entrada, saida, colunas):
    """
    Processa o arquivo Excel para dividir os dados por `application_name`,
    contar as mensagens (`message`) e salvar o resultado em abas separadas.

    :param entrada: Caminho do arquivo de entrada (.xlsx).
    :param saida: Caminho do arquivo de saída (.xlsx).
    :param colunas: Lista das colunas a serem analisadas.
    """
    try:
        df = pd.read_excel(entrada)
        
        for coluna in colunas:
            if coluna not in df.columns:
                print(f"A coluna '{coluna}' não foi encontrada no arquivo.")
                return

        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            grupos = df.groupby('application_name')
            for app, grupo in grupos:
                contagem = grupo['message'].value_counts().reset_index()
                contagem.columns = ['message', 'Contagem']

                contagem.to_excel(writer, index=False, sheet_name=str(app)[:31])

        print(f"Arquivo processado com sucesso! Saída salva em '{saida}'.")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

colunas_analisar = ['application_name', 'message']  

arquivo_entrada = encontrar_arquivo_recente()

if arquivo_entrada:
    print(f"Arquivo mais recente encontrado: {arquivo_entrada}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = f"relatorio_processado_{timestamp}.xlsx"

    processar_excel(arquivo_entrada, arquivo_saida, colunas_analisar)
else:
    print("Nenhum arquivo .xlsx encontrado na pasta.")
