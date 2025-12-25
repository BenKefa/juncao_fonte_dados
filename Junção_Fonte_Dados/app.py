## ===== Importar bibliotecas ===== ##
import pandas as pd
import locale
import nbformat
import os
import shutil

## ===== Juntar todas as planilhas de pedidos em uma única base para importação do Dash ===== #
diretorio_origem = r'C:/Projetos/Junção_Fonte_Dados/dados_dispersos'  # Pasta de origem, onde as planilhas estão em pleno uso
diretorio_destino = r'C:/Projetos/Junção_Fonte_Dados/backup'  # Pasta de destino, de onde o app vai buscar os dados para unificar em um único arquivo


# Excluir tudo dentro da pasta de backup "diretorio_destino" para popular com os novos arquivos atualizados  
def excluir_tudo_em_diretorio(diretorio_destino):
    for arquivo in os.listdir(diretorio_destino):
        arquivo_completo = os.path.join(diretorio_destino, arquivo)
        if os.path.isfile(arquivo_completo):
            os.remove(arquivo_completo)
        elif os.path.isdir(arquivo_completo):
            excluir_tudo_em_diretorio(arquivo_completo)

excluir_tudo_em_diretorio(diretorio_destino)

# Identificar todos os arquivos excel no formato .xlsx dentro da pasta "diretorio_origem" e copiar para a pasta "diretorio_destino"
for filename in os.listdir(diretorio_origem):
    if filename.endswith('.xlsx'):
        arquivo_excel = os.path.join(diretorio_origem, filename)

        try:
            destino_arquivo = os.path.join(diretorio_destino, filename)
            shutil.copy(arquivo_excel, destino_arquivo)
        except PermissionError:
            print(f'O arquivo {filename} está aberto no Excel. Não foi copiado.')
           

 
# Função para importar os dados de pedidos
def importar_base_dados():
    todos_arquivos = []
    for planilha in os.listdir(diretorio_destino):
        if planilha.endswith('.xlsx'):
            arquivo_excel = os.path.join(diretorio_destino, planilha)
            resumo_arquivos = pd.read_excel(arquivo_excel, header=0)
            todos_arquivos.append(resumo_arquivos)

    # Unificar todas as planilhas
    arquivo_unificado = pd.concat(todos_arquivos, ignore_index=True)
    
    # Salvar arquivo unificado   
    arquivo_unificado.to_excel('C:/Projetos/Junção_Fonte_Dados/dados_agrupados/dados_agrupados.xlsx', index=False)

    return arquivo_unificado
arquivo_unificado = importar_base_dados()


