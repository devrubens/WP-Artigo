import pandas as pd
import matplotlib.pyplot as plt

file_path = "seuarquivo.xlsx"

def gerar_graficos(sheet_name, colunas_selecionadas, linha_inicio, linha_fim):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    print(f"DataFrame carregado da aba '{sheet_name}':")
    print(df.head())  
    
    total_linhas, total_colunas = df.shape
    print(f"Número de linhas: {total_linhas}, Número de colunas: {total_colunas}")
    
    if linha_inicio >= total_linhas or linha_fim > total_linhas:
        print(f"Erro: Tentativa de acessar linhas fora dos limites. Total de linhas disponíveis: {total_linhas}.")
        raise IndexError("Os índices de linhas estão fora dos limites.")
    
    if any(col >= total_colunas for col in colunas_selecionadas):
        raise IndexError("Os índices de colunas estão fora dos limites.")
    
    # Seleção de linhas e colunas
    df_selecionado = df.iloc[linha_inicio:linha_fim, colunas_selecionadas]
    
    df_filtrado = df_selecionado[df_selecionado.apply(lambda row: row.astype(str).str.contains('X').any(), axis=1)]
    
    contagem = df_filtrado.apply(lambda col: col.astype(str).str.contains('X').sum())

    # Tamanho da fonte
    font_size = 14

    # Gráfico de barra
    plt.figure(figsize=(10, 6))
    contagem.plot(kind='bar', color='skyblue')
    plt.xticks(fontsize=font_size)
    plt.yticks(fontsize=font_size)
    plt.show()

    # Gráfico de pizza
    plt.figure(figsize=(8, 8))
    contagem.plot(kind='pie', autopct='%1.1f%%', startangle=140, fontsize=font_size)
    plt.ylabel('')
    plt.show()

    # Gráfico de linhas
    plt.figure(figsize=(10, 6))
    contagem.plot(kind='line', marker='o', linestyle='-', color='green')
    plt.xlabel('Colunas', fontsize=font_size)
    plt.ylabel('Contagem', fontsize=font_size)
    plt.xticks(fontsize=font_size)
    plt.yticks(fontsize=font_size)
    plt.grid(True)
    plt.show()

# Função para uso

gerar_graficos(
    sheet_name='Dificuldades (Abraão)',  
    colunas_selecionadas=[2, 3, 4, 5, 6, 7, 8, 9, 10, 11],   # Colunas 
    linha_inicio=2,                          #linha de inicio
    linha_fim=19                             # Linha de fim final
)
