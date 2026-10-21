import pandas as pd
import plotly.express as px
from pathlib import Path
import re
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# REGRAS DE CLASSIFICAÇÃO
# ==========================================
def classificar_formulario(nome_original):
    """
    Classifica o formulário aplicando normalização de texto e busca por raízes de palavras.
    """
    if pd.isna(nome_original):
        return 'Ignorado (Vazio)'
    
    # Normaliza o texto: minúsculas, sem espaços extras nas pontas
    nome_norm = str(nome_original).strip().lower()
    
    # Trata acentuação básica para a regra
    nome_norm = nome_norm.replace('ç', 'c').replace('õ', 'o').replace('ã', 'a')
    
    # REGRA: Deve conter 'inspecao' ou 'inspecoes' E estar no contexto de 'csc' ou 'regional'
    if 'inspec' in nome_norm and ('csc' in nome_norm or 'regional' in nome_norm or 'regionais' in nome_norm):
        return 'Inspeções CSC (Válido)'
    
    # Se não atendeu à regra específica, vai para a fila de validação
    return 'Pendente de Validação'

# ==========================================
# FUNÇÕES DE EXTRAÇÃO E VALIDAÇÃO
# ==========================================
def extrair_competencia_nome(nome_arquivo):
    match = re.search(r'(20\d{2})\.(\d{2})', nome_arquivo)
    if match:
        return pd.Period(f"{match.group(1)}-{match.group(2)}", freq='M')
    return None

def encontrar_coluna_data(df):
    possiveis_nomes = ['Aberto(a)', 'Aberto', 'Data de Abertura', 'Created', 'Data']
    for col in df.columns:
        if str(col).strip() in possiveis_nomes:
            return str(col).strip()
    return None

# ==========================================
# PIPELINE DE PROCESSAMENTO
# ==========================================
def processar_arquivos(diretorio_base):
    base_path = Path(diretorio_base)
    lista_registros = []
    
    nomes_excluidos = {'ativo_x_mrsls_csc_sla', '2025_x_mrsls_csc_sla'}
    extensoes_validas = {'.xlsx', '.xls', '.csv'}

    if not base_path.exists():
        print(f"Erro: Diretório '{base_path}' não encontrado.")
        return pd.DataFrame()

    for arquivo in base_path.iterdir():
        if not arquivo.is_file() or arquivo.name.startswith('~$') or arquivo.suffix.lower() not in extensoes_validas or arquivo.stem in nomes_excluidos:
            continue

        try:
            if arquivo.suffix.lower() == '.csv':
                df = pd.read_csv(arquivo, low_memory=False)
            else:
                df = pd.read_excel(arquivo)
                
            df.columns = [str(c).strip() for c in df.columns]
            
            if 'Formulário' not in df.columns:
                continue
                
            col_data = encontrar_coluna_data(df)
            
            # Aplica a regra de classificação em uma nova coluna
            df['Categoria_Classificada'] = df['Formulário'].apply(classificar_formulario)
            
            # Identificação do Mês
            if col_data and not df[col_data].isna().all():
                df['Data_Real'] = pd.to_datetime(df[col_data], errors='coerce')
                df['Mes_Apuracao'] = df['Data_Real'].dt.to_period('M')
            else:
                df['Mes_Apuracao'] = extrair_competencia_nome(arquivo.name)
            
            # Extrai apenas as colunas necessárias para a consolidação
            df_extract = df[['Mes_Apuracao', 'Formulário', 'Categoria_Classificada']].copy()
            df_extract['Arquivo_Origem'] = arquivo.name
            df_extract['Quantidade'] = 1
            
            lista_registros.append(df_extract)
            
        except Exception as e:
            print(f"[Erro] Falha ao processar '{arquivo.name}': {e}")

    if not lista_registros:
        return pd.DataFrame()
        
    df_consolidado = pd.concat(lista_registros, ignore_index=True)
    df_consolidado = df_consolidado.dropna(subset=['Mes_Apuracao'])
    
    # =================================================================
    # NOVA REGRA: Filtra e mantém exclusivamente os registros do ano 2026
    # =================================================================
    df_consolidado = df_consolidado[df_consolidado['Mes_Apuracao'].dt.year == 2026]
    
    return df_consolidado

def gerar_grafico_interativo(df_valido, diretorio_saida):
    # Agrupa por Mês e Nome Original do Formulário para detalhar no tooltip
    df_plot = df_valido.groupby(['Mes_Apuracao', 'Formulário']).size().reset_index(name='Chamados')
    df_plot['Mes_Str'] = df_plot['Mes_Apuracao'].dt.strftime('%m/%Y')
    df_plot = df_plot.sort_values('Mes_Apuracao')

    # Calcula os totais absolutos por mês para exibir no topo das barras
    totais_mensais = df_plot.groupby('Mes_Str', sort=False)['Chamados'].sum().reset_index(name='Total_Mes')

    # Criação do gráfico mantendo a separação por 'Formulário' para preservar o tooltip
    fig = px.bar(
        df_plot, 
        x='Mes_Str', 
        y='Chamados', 
        color='Formulário', 
        title='Quantidade Mensal de Chamados (Inspeções CSC e Regionais)',
        labels={'Mes_Str': 'Competência', 'Chamados': 'Qtd de Chamados', 'Formulário': 'Nome Original'},
        # REMOVIDO: text='Chamados' (para não mostrar os números fracionados nos bloquinhos)
        color_discrete_sequence=['#1f497d'] # Azul escuro uniforme
    )

    fig.update_layout(
        barmode='stack',
        hovermode='x unified',
        xaxis_title="Mês/Ano",
        yaxis_title="Total de Chamados",
        showlegend=False, 
        plot_bgcolor='white'
    )
    
    # Remove as linhas divisórias e oculta qualquer texto remanescente dentro da barra
    fig.update_traces(marker_line_width=0, textposition='none')
    
    # Adiciona o número total cravado no topo de cada barra azul
    for _, row in totais_mensais.iterrows():
        fig.add_annotation(
            x=row['Mes_Str'], 
            y=row['Total_Mes'], 
            text=str(row['Total_Mes']),
            showarrow=False,
            yshift=10, 
            font=dict(size=14, color="black", family="Arial Black")
        )
    
    caminho_html = diretorio_saida / 'Dashboard_Interativo_Inspecoes.html'
    fig.write_html(caminho_html)
    print(f"\n[Visualização] Gráfico interativo gerado com total único nas barras: {caminho_html}")
    fig.show()

# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
def main():
    caminho_entrada = r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla"
    pasta_resultados = Path(caminho_entrada) / "resultados_analise"
    pasta_resultados.mkdir(exist_ok=True)
    
    print("Iniciando varredura e classificação...")
    df_geral = processar_arquivos(caminho_entrada)
    
    if df_geral.empty:
        print("Nenhum dado válido de 2026 encontrado para processamento.")
        return

    # Separa os dados em Válidos e Pendentes
    df_inspecoes = df_geral[df_geral['Categoria_Classificada'] == 'Inspeções CSC (Válido)']
    df_pendentes = df_geral[df_geral['Categoria_Classificada'] == 'Pendente de Validação']

    # 1. Tabela Detalhada (Por Mês e Nome do Formulário)
    tabela_detalhada = df_inspecoes.groupby(['Mes_Apuracao', 'Formulário']).size().reset_index(name='Quantidade')
    tabela_detalhada = tabela_detalhada.sort_values(['Mes_Apuracao', 'Quantidade'], ascending=[True, False])
    
    # 2. Resumo de Auditoria de Nomes (Quais formulários o código encontrou e como classificou)
    auditoria_nomes = df_geral.groupby(['Formulário', 'Categoria_Classificada']).size().reset_index(name='Total_Chamados')
    auditoria_nomes = auditoria_nomes.sort_values('Total_Chamados', ascending=False)

    # Exportações
    caminho_detalhado = pasta_resultados / 'Detalhamento_Inspecoes_Mensal_2026.xlsx'
    caminho_auditoria = pasta_resultados / 'Auditoria_Classificacao_Formularios_2026.xlsx'
    
    tabela_detalhada['Mes_Apuracao'] = tabela_detalhada['Mes_Apuracao'].astype(str)
    tabela_detalhada.to_excel(caminho_detalhado, index=False)
    auditoria_nomes.to_excel(caminho_auditoria, index=False)
    
    print("\n--- RESUMO DE CLASSIFICAÇÃO (Apenas 2026) ---")
    print(f"Total de chamados de Inspeção válidos: {len(df_inspecoes)}")
    print(f"Total de chamados ignorados/pendentes: {len(df_pendentes)}")
    print(f"\nArquivos exportados para a pasta: {pasta_resultados}")
    
    # Gera o gráfico interativo apenas com os dados validados
    if not df_inspecoes.empty:
        gerar_grafico_interativo(df_inspecoes, pasta_resultados)

if __name__ == "__main__":
    main()