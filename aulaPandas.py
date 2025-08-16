import pandas as pd

def escolher_aba_excel(caminho_arquivo):
    xls = pd.ExcelFile(caminho_arquivo)
    abas = xls.sheet_names
    
    print("Selecione a aba que deseja carregar:")
    for i, aba in enumerate(abas, start=1):
        print(f"{i} - {aba}")
    
    escolha = int(input("Digite o número da aba: "))
    
    if 1 <= escolha <= len(abas):
        aba_escolhida = abas[escolha - 1]
        df = pd.read_excel(caminho_arquivo, sheet_name=aba_escolhida)
        print(f"Aba '{aba_escolhida}' carregada!")
        return df, aba_escolhida
    else:
        print("Número inválido.")
        return None, None


# Exemplo de uso
url3 = "/content/Exemplo 3  - Importação em Excel.xlsx"
df, aba = escolher_aba_excel(url3)

# Agora df é o DataFrame da aba escolhida
if df is not None:
    print("Colunas disponíveis:", df.columns.tolist())
    
    # Você pode ATUAR diretamente nele:
    print("\nExemplo: média numérica")
    print(df.mean(numeric_only=True))
    
    print("\nExemplo: quantidade de linhas")
    print(len(df))
