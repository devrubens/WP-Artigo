import pandas as pd
import matplotlib.pyplot as plt

# Caminho do arquivo
file_path = "AddSeuArquivo.xlsx"

def gerar_graficos(sheet_name, colunas_selecionadas, linha_inicio, linha_fim):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    print(f"DataFrame carregado da aba '{sheet_name}':")
    print(df.head())  
    
    total_linhas, total_colunas = df.shape
    print(f"Número de linhas: {total_linhas}, Número de colunas: {total_colunas}")
    
    print("\nÍndice e nomes das colunas disponíveis:")
    for i, col_name in enumerate(df.columns):
        print(f"{i}: {col_name}")
    
    if linha_inicio >= total_linhas or linha_fim > total_linhas:
        raise IndexError("Os índices de linhas estão fora dos limites.")
    if any(col >= total_colunas for col in colunas_selecionadas):
        raise IndexError("Os índices de colunas estão fora dos limites.")
    
    # seleção de linhas e colunas
    df_selecionado = df.iloc[linha_inicio:linha_fim, colunas_selecionadas]
    
    df_filtrado = df_selecionado[df_selecionado.apply(lambda row: row.astype(str).str.contains('X').any(), axis=1)]
    
    contagem = df_filtrado.apply(lambda col: col.astype(str).str.contains('X').sum())

    # Grafico de barra
    plt.figure(figsize=(10, 6))
    contagem.plot(kind='bar', color='skyblue')
    plt.show()

    # Pizza
    plt.figure(figsize=(8, 8))
    contagem.plot(kind='pie', autopct='%1.1f%%', startangle=140)
    plt.ylabel('')
    plt.show()

    # Linhas
    plt.figure(figsize=(10, 6))
    contagem.plot(kind='line', marker='o', linestyle='-', color='green')
    plt.grid(True)
    plt.show()

# Função para uso

gerar_graficos(
    sheet_name='Nome da pagina para vc analisar',  
    colunas_selecionadas=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],   # Colunas 
    linha_inicio=2,                          #linha de inicio
    linha_fim=20                             # Linha de fim final
)
