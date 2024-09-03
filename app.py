from flask import Flask, request, render_template, send_file
import pandas as pd
import threading
import itertools
import time
import sys
from scipy.stats import zscore
import webbrowser
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Variável global para controle da animação de carregamento
carregando = False

# Função para exibir a introdução no console
def exibir_informacoes_do_desenvolvedor():
    """
    Exibe informações sobre o desenvolvedor Anderson Marinho no console.
    """
    # Códigos de escape ANSI para colorir o texto no console
    cores = {
        'azul': '\033[94m',  # Texto azul
        'negrito': '\033[1m', # Texto negrito
        'reset': '\033[0m'    # Reset para a cor padrão
    }

    # Informação para ser exibida no console
    informacao = f"""
{cores['azul']}{cores['negrito']}Bem-vindo ao Analisador de Folha de Pagamento! - Gross to Gross Dashboard e Análise{cores['reset']}
Desenvolvido por: Anderson Marinho - Especialista em RH, DP, Advocacia e ADS
Conectando a excelência em gestão de pessoas, conformidade legal e automação de processos.

Portfólio: {cores['azul']}https://advmarinho.github.io/PortifolioRH/#servicos{cores['reset']}
GitHub: {cores['azul']}https://github.com/advmarinho/ {cores['reset']}
"""

    # Exibe a informação no console
    print(informacao)

# Chama a função para exibir as informações do desenvolvedor ao iniciar o script
exibir_informacoes_do_desenvolvedor()

# Função para a animação de carregamento
def animacao_carregamento():
    """
    Exibe uma animação de carregamento no console durante a execução do processo de análise de folha de pagamento.
    A animação usa caracteres especiais para criar um efeito de rotação e exibe a mensagem 'Carregando' em azul.
    """
    cores = {
        'azul': '\033[94m',   # ANSI escape code para texto azul
        'reset': '\033[0m'    # Reset para retornar à cor padrão
    }

    print(f"{cores['azul']}Ação: Análise de Folha de Pagamento em Progresso...{cores['reset']}")

    for c in itertools.cycle(['|', '/', '-', '\\']):
        if not carregando:
            break
        sys.stdout.write(f"\r{cores['azul']}Carregando {c}{cores['reset']}")
        sys.stdout.flush()
        time.sleep(0.1)

    # Mensagem final quando o carregamento for concluído
    sys.stdout.write('\rConcluído!                \n')


# Inicializar tkinter e abrir diálogo para selecionar o arquivo CSV
Tk().withdraw()  # Ocultar a janela principal do tkinter
file_path = askopenfilename(
    filetypes=[("CSV files", "*.csv")],
    title="Selecione o arquivo CSV"
)

# Verifica se um arquivo foi selecionado
if not file_path:
    print("Nenhum arquivo selecionado. O programa será encerrado.")
    exit()
else:
    try:
        # Iniciar a animação de carregamento em um thread separado
        carregando = True
        t = threading.Thread(target=animacao_carregamento)
        t.start()

        # Iniciar o contador de tempo para ler o arquivo CSV
        start_time = time.time()

        # Ler o arquivo CSV usando pandas, especificando o delimitador e ignorando linhas mal formatadas
        df = pd.read_csv(file_path, sep=';', low_memory=False, on_bad_lines='skip')

        # Parar a animação de carregamento
        carregando = False
        t.join()

        # Calcular o tempo de execução para ler o arquivo CSV
        end_time = time.time()
        elapsed_time = end_time - start_time

        print(f"Tempo para carregar o arquivo CSV: {elapsed_time:.6f} segundos")

        # Filtrar apenas funcionários ativos e verbas de tipo "Provento" e "Desconto"
        df = df[(df['Sit Folha-*'] != 'DEMITIDO') & (df['Tipo Verba'].isin(['Provento', 'Desconto']))]

        # Inicializar listas de funcionários e verbas com base nos dados do DataFrame
        lista_funcionarios = df['Nome'].unique().tolist() if 'Nome' in df.columns else []
        lista_verbas = df['DescVerba'].unique().tolist() if 'DescVerba' in df.columns else []

        # Converter colunas para o tipo numérico, se aplicável
        for col in ['Salario', 'Horas Lanc', 'Vlr Lancam']:
            if col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = pd.to_numeric(df[col].str.replace(',', '.'), errors='coerce')

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
        carregando = False
        t.join()
        exit()

def filtrar_ultimos_12_meses(df):
    # Certificar-se de que Cod Periodo é tratado como string para facilitar a ordenação
    df['Cod Periodo'] = df['Cod Periodo'].astype(str)
    
    # Ordenar por 'Cod Periodo' de forma regressiva
    df = df.sort_values(by='Cod Periodo', ascending=False)

    # Filtrar os últimos 12 meses
    ultimos_12_periodos = df['Cod Periodo'].unique()[:12]
    df_ultimos_12_meses = df[df['Cod Periodo'].isin(ultimos_12_periodos)]

    return df_ultimos_12_meses

def calcular_tendencias(df):
    # Calcular a média e o desvio padrão para cada verba nos últimos 12 meses
    df['Cod Periodo'] = df['Cod Periodo'].astype(str)
    df_pivot = df.pivot_table(
        index=['Matricula**', 'Nome', 'Filial', 'VB TMF**', 'DescVerba', 'Xdeb e Xcred'],  # Inclua 'Xdeb e Xcred' aqui
        columns='Cod Periodo',
        values='Vlr Lancam',
        aggfunc='sum'
    ).reset_index()

    # Calcular a média dos últimos meses para comparar com cada mês
    periodos_anteriores = df_pivot.columns[-12:]
    df_pivot['Media_meses'] = df_pivot[periodos_anteriores].mean(axis=1)
    df_pivot['Desvio_meses'] = df_pivot[periodos_anteriores].std(axis=1)

    # Inicializar colunas para armazenar informações de meses fora do padrão e a explicação
    df_pivot['Mes_Fora_Padrao'] = None
    df_pivot['Explicacao'] = None

    # Loop para calcular a diferença relativa para cada mês
    for periodo in periodos_anteriores:
        df_pivot['Diferenca_Relativa_' + periodo] = (df_pivot[periodo] - df_pivot['Media_meses']) / df_pivot['Desvio_meses']

        # Identificar outliers para cada mês
        outliers = df_pivot[(df_pivot['Diferenca_Relativa_' + periodo] > 2) | (df_pivot['Diferenca_Relativa_' + periodo] < -2)]
        
        # Adicionar informações ao DataFrame
        df_pivot.loc[outliers.index, 'Mes_Fora_Padrao'] = periodo
        df_pivot.loc[outliers.index, 'Explicacao'] = outliers.apply(
            lambda row: f"Valor significativamente {'acima' if row['Diferenca_Relativa_' + periodo] > 2 else 'abaixo'} da média para o mês {periodo}.",
            axis=1
        )

    # Identificar verbas únicas que aparecem apenas uma vez
    df_verbas_unicas = df_pivot[periodos_anteriores].notna().sum(axis=1) == 1
    df_pivot.loc[df_verbas_unicas, 'Mes_Fora_Padrao'] = df_pivot.loc[df_verbas_unicas, periodos_anteriores].idxmax(axis=1)
    df_pivot.loc[df_verbas_unicas, 'Explicacao'] = "Verba que apareceu apenas uma vez no período analisado."

    # Retornar apenas os outliers
    df_outliers = df_pivot[df_pivot['Mes_Fora_Padrao'].notna()]

    return df_outliers


def eda_basica(df):
    # Mostrar informações básicas sobre o DataFrame
    info_df = df.info()

    # Mostrar estatísticas descritivas
    estatisticas_df = df.describe(include='all')

    # Identificar valores ausentes
    valores_ausentes = df.isnull().sum()

    # Identificar colunas numéricas e categóricas
    colunas_numericas = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
    colunas_categoricas = df.select_dtypes(include=['object']).columns.tolist()

    # Verificar correlação entre colunas numéricas
    correlacao_df = df[colunas_numericas].corr()

    return {
        'info': info_df,
        'estatisticas': estatisticas_df,
        'valores_ausentes': valores_ausentes,
        'colunas_numericas': colunas_numericas,
        'colunas_categoricas': colunas_categoricas,
        'correlacao': correlacao_df
    }


# Função para preparar os dados dos outliers para exibição
def prepare_outliers_for_display(df_outliers):
    if not df_outliers.empty:
        # Definir as colunas para exibição
        display_columns = ['Filial', 'Nome', 'Matricula**', 'VB TMF**', 'DescVerba', 'Xdeb e Xcred', 'Media_meses', 'Desvio_meses', 'Mes_Fora_Padrao', 'Explicacao']
        
        # Arredondar colunas numéricas para 2 casas decimais
        df_outliers['Media_meses'] = df_outliers['Media_meses'].round(2)
        df_outliers['Desvio_meses'] = df_outliers['Desvio_meses'].round(2)
        
        # Contar o número de linhas do DataFrame
        row_count = len(df_outliers)
        
        # Converter o DataFrame para HTML
        outliers_display = df_outliers[display_columns].to_html(classes='table table-striped table-bordered', index=False)
        
        # Retornar o HTML da tabela e o número de linhas
        return outliers_display, row_count
    
    return "Nenhum outlier encontrado nos últimos 12 meses.", 0


# Aplicar filtros e cálculos
df_ultimos_12_meses = filtrar_ultimos_12_meses(df)
df_outliers = calcular_tendencias(df_ultimos_12_meses)
resultados_eda = eda_basica(df)
outliers_display = prepare_outliers_for_display(df_outliers)

# **Adicionar a funcionalidade para salvar o resultado em um arquivo Excel**
try:
    df_outliers.to_excel("outliers_result.xlsx", index=False)
    print("Os resultados foram salvos no arquivo 'outliers_result.xlsx'.")
except Exception as e:
    print(f"Erro ao salvar o arquivo Excel: {e}")



# Configuração do Flask
app = Flask(__name__, template_folder=os.path.join(os.path.dirname(__file__), 'templates'))

@app.route('/download_excel')
def download_excel():
    try:
        # Defina o caminho absoluto para salvar o arquivo
        excel_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "outliers_result.xlsx")
        
        # Gerar o arquivo Excel
        df_outliers.to_excel(excel_path, index=False)
        
        # Enviar o arquivo para download
        return send_file(excel_path, as_attachment=True)
    except Exception as e:
        print(f"Erro ao gerar o arquivo Excel: {e}")
        return "Erro ao gerar o arquivo Excel.", 500


@app.route('/')
def home():
    try:
        # Receber os parâmetros de filtro do formulário
        nome_filtro = request.args.get('nome', '')
        cod_periodo_filtro = request.args.get('cod_periodo', '')
        desc_verba_filtro = request.args.get('desc_verba', '')
        xdeb_xcred_filtro = request.args.get('xdeb_xcred', '')

        # Filtrar os dados com base nos parâmetros selecionados
        df_filtrado = df_outliers.copy()
        
        if nome_filtro:
            df_filtrado = df_filtrado[df_filtrado['Nome'] == nome_filtro]
        if cod_periodo_filtro:
            df_filtrado = df_filtrado[df_filtrado['Mes_Fora_Padrao'] == cod_periodo_filtro]
        if desc_verba_filtro:
            df_filtrado = df_filtrado[df_filtrado['DescVerba'] == desc_verba_filtro]
        if xdeb_xcred_filtro:
            df_filtrado = df_filtrado[df_filtrado['Xdeb e Xcred'] == xdeb_xcred_filtro]

        # Ordenar por nome para exibição
        df_filtrado = df_filtrado.sort_values(by='Nome')

        # Preparar os dados para exibição
        outliers_display, row_count = prepare_outliers_for_display(df_filtrado)

        # Listas para os filtros
        funcionarios = sorted(lista_funcionarios) if lista_funcionarios else []
        verbas = lista_verbas if lista_verbas else []
        cod_periodos = df_outliers['Mes_Fora_Padrao'].unique().tolist() if 'Mes_Fora_Padrao' in df_outliers.columns else []
        xdeb_xcred_vals = df_outliers['Xdeb e Xcred'].unique().tolist() if 'Xdeb e Xcred' in df_outliers.columns else []

        return render_template('index.html', 
                               funcionarios=funcionarios, 
                               verbas=verbas, 
                               cod_periodos=cod_periodos,
                               xdeb_xcred_vals=xdeb_xcred_vals,
                               selected_nome=nome_filtro,
                               selected_cod_periodo=cod_periodo_filtro,
                               selected_desc_verba=desc_verba_filtro,
                               selected_xdeb_xcred=xdeb_xcred_filtro,
                               outliers_display=outliers_display,
                               row_count=row_count)
    except Exception as e:
        print(f"Erro ao renderizar a página inicial: {e}")
        return "Ocorreu um erro ao processar sua solicitação.", 500


def open_browser():
    # Abre o navegador na URL localhost na porta 5000
    print("\033[93m\nCopie e Cole o endereço no navegador - http://127.0.0.1:5000  \n\033[0m")
    
    webbrowser.open_new("http://127.0.0.1:5000")


if __name__ == '__main__':
    # Inicia um thread para abrir o navegador após 1 segundo
    threading.Timer(1, open_browser).start()
    # Inicia o servidor Flask na porta 5000
    app.run(debug=False, port=5000)
