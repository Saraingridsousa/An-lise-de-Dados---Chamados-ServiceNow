import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import re
import warnings

# Suprime avisos desnecessários do openpyxl ao ler arquivos
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def extrair_competencia_nome(nome_arquivo):
    """Tenta extrair o padrão de ano e mês (YYYY.MM) do nome do arquivo."""
    match = re.search(r'(20\d{2})\.(\d{2})', nome_arquivo)
    if match:
        return pd.Period(f"{match.group(1)}-{match.group(2)}", freq='M')
    return None

def encontrar_coluna_data(df):
    """Busca dinamicamente por uma coluna que represente a data de abertura."""
    possiveis_nomes = ['Aberto(a)', 'Aberto', 'Data de Abertura', 'Created', 'Data']
    for col in df.columns:
        if str(col).strip() in possiveis_nomes:
            return col
    return None

def processar_arquivos(diretorio_base):
    base_path = Path(diretorio_base)
    resultados_arquivos = []
    
    resumo = {
        'encontrados': 0, 'processados': 0, 'ignorados': 0, 'erros': 0,
        'total_registros_lidos': 0, 'total_regional': 0, 'total_csc': 0
    }
    
    nomes_excluidos = {'ativo_x_mrsls_csc_sla', '2025_x_mrsls_csc_sla'}
    extensoes_validas = {'.xlsx', '.xls', '.csv'}

    if not base_path.exists():
        print(f"Erro: O diretório '{base_path}' não foi encontrado.")
        return pd.DataFrame(), resumo, []

    for arquivo in base_path.iterdir():
        if not arquivo.is_file():
            continue
        
        resumo['encontrados'] += 1
        
        # Ignora arquivos temporários e extensões inválidas
        if arquivo.name.startswith('~$') or arquivo.suffix.lower() not in extensoes_validas:
            resumo['ignorados'] += 1
            continue
            
        # Ignora arquivos pelos nomes-base específicos
        if arquivo.stem in nomes_excluidos:
            resumo['ignorados'] += 1
            continue

        try:
            # Leitura do arquivo
            if arquivo.suffix.lower() == '.csv':
                df = pd.read_csv(arquivo, low_memory=False)
            else:
                df = pd.read_excel(arquivo)
                
            resumo['total_registros_lidos'] += len(df)
            
            # Normalizar nomes das colunas (remove espaços no início/fim)
            df.columns = [str(c).strip() for c in df.columns]
            
            if 'Formulário' not in df.columns:
                print(f"[Aviso] Arquivo ignorado (Sem coluna 'Formulário'): {arquivo.name}")
                resumo['ignorados'] += 1
                continue
                
            # Identificação da competência (Prioridade: Dados reais -> Nome do arquivo)
            col_data = encontrar_coluna_data(df)
            
            # Filtro das categorias
            filtro_regional = df['Formulário'] == 'Inspeções CSC Regional'
            filtro_csc = df['Formulário'] == 'Inspeção CSC'
            
            qtd_regional = filtro_regional.sum()
            qtd_csc = filtro_csc.sum()
            
            df_filtrado = df[filtro_regional | filtro_csc].copy()
            
            resumo['total_regional'] += qtd_regional
            resumo['total_csc'] += qtd_csc
            
            if len(df_filtrado) > 0 and col_data:
                # Agrupa pela data real do dado se houver chamados válidos
                df_filtrado[col_data] = pd.to_datetime(df_filtrado[col_data], errors='coerce')
                df_filtrado['Mes_Real'] = df_filtrado[col_data].dt.to_period('M')
                
                # Para cada mês encontrado no arquivo, registrar o volume
                agrupado = df_filtrado.groupby('Mes_Real')['Formulário'].value_counts().unstack(fill_value=0)
                
                for mes, row in agrupado.iterrows():
                    reg = row.get('Inspeções CSC Regional', 0)
                    csc = row.get('Inspeção CSC', 0)
                    resultados_arquivos.append({
                        'Arquivo': arquivo.name,
                        'Mes': mes,
                        'Inspecoes_CSC_Regional': reg,
                        'Inspecao_CSC': csc
                    })
            else:
                # Fallback: Se não encontrou chamados das categorias ou não tem coluna de data,
                # usa o mês de referência do nome do arquivo (se conseguiu ler alguma coisa)
                mes_ref = extrair_competencia_nome(arquivo.name)
                resultados_arquivos.append({
                    'Arquivo': arquivo.name,
                    'Mes': mes_ref if mes_ref else pd.Period('NaT', freq='M'),
                    'Inspecoes_CSC_Regional': qtd_regional,
                    'Inspecao_CSC': qtd_csc
                })
                
            resumo['processados'] += 1
            print(f"[Sucesso] Lido: {arquivo.name} | Regional: {qtd_regional} | CSC: {qtd_csc}")
            
        except Exception as e:
            print(f"[Erro] Falha ao processar '{arquivo.name}': {e}")
            resumo['erros'] += 1

    return pd.DataFrame(resultados_arquivos), resumo

def consolidar_dados(df_bruto):
    if df_bruto.empty:
        return pd.DataFrame(), []
    
    # Remove registros onde não foi possível identificar o mês
    df_valido = df_bruto.dropna(subset=['Mes']).copy()
    
    # Identificar sobreposições (mesmo mês vindo de arquivos diferentes)
    contagem_arquivos_por_mes = df_valido.groupby('Mes')['Arquivo'].nunique()
    sobreposicoes = contagem_arquivos_por_mes[contagem_arquivos_por_mes > 1].index.tolist()
    
    # Agrupar por Mês
    df_consolidado = df_valido.groupby('Mes').agg({
        'Inspecoes_CSC_Regional': 'sum',
        'Inspecao_CSC': 'sum'
    }).reset_index()
    
    df_consolidado['Total_Chamados'] = df_consolidado['Inspecoes_CSC_Regional'] + df_consolidado['Inspecao_CSC']
    
    # Ordenação cronológica
    df_consolidado = df_consolidado.sort_values('Mes').reset_index(drop=True)
    return df_consolidado, sobreposicoes

def gerar_grafico(df_consolidado, diretorio_saida):
    if df_consolidado.empty:
        print("Sem dados para gerar gráfico.")
        return

    # Preparar eixo X como string legível
    meses_str = df_consolidado['Mes'].dt.strftime('%m/%Y')
    valores = df_consolidado['Total_Chamados']

    plt.figure(figsize=(10, 6))
    barras = plt.bar(meses_str, valores, color='#2c7bb6', edgecolor='black')

    plt.title('Quantidade Mensal de Chamados de Inspeção CSC', fontsize=14, fontweight='bold', pad=15)
    plt.xlabel('Competência (Mês/Ano)', fontsize=12)
    plt.ylabel('Quantidade de Chamados', fontsize=12)
    plt.xticks(rotation=45)
    
    # Adicionar rótulos nas barras
    for barra in barras:
        altura = barra.get_height()
        plt.text(barra.get_x() + barra.get_width()/2., altura + 0.5,
                 f'{int(altura)}', ha='center', va='bottom', fontsize=11)
                 
    # Remove a borda superior e direita (Spines) para visual mais limpo
    plt.gca().spines['top'].set_visible(False)
    plt.gca().spines['right'].set_visible(False)

    plt.tight_layout()
    
    caminho_grafico = diretorio_saida / 'grafico_inspecoes_mensal.png'
    plt.savefig(caminho_grafico, dpi=300)
    plt.show()

def executar_pipeline():
    caminho_entrada = r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla"
    pasta_resultados = Path(caminho_entrada) / "resultados_analise"
    pasta_resultados.mkdir(exist_ok=True)
    
    print("--- INICIANDO LEITURA DOS ARQUIVOS ---")
    df_bruto, resumo = processar_arquivos(caminho_entrada)
    
    print("\n--- CONSOLIDANDO DADOS ---")
    df_consolidado, sobreposicoes = consolidar_dados(df_bruto)
    
    print("\n================ RESUMO DA EXECUÇÃO ================")
    print(f"Arquivos Encontrados: {resumo['encontrados']}")
    print(f"Arquivos Processados: {resumo['processados']}")
    print(f"Arquivos Ignorados: {resumo['ignorados']}")
    print(f"Arquivos com Erro: {resumo['erros']}")
    print(f"Total de Linhas Analisadas (todas categorias): {resumo['total_registros_lidos']}")
    print(f"Total 'Inspeções CSC Regional': {resumo['total_regional']}")
    print(f"Total 'Inspeção CSC': {resumo['total_csc']}")
    print(f"Total Combinado Geral: {resumo['total_regional'] + resumo['total_csc']}")
    
    if sobreposicoes:
        print("\n[ALERTA DE SOBREPOSIÇÃO]")
        meses_alerta = [m.strftime('%m/%Y') for m in sobreposicoes]
        print(f"Atenção! Os seguintes meses tiveram dados extraídos de mais de um arquivo distinto: {', '.join(meses_alerta)}.")
        print("Existe risco de chamados duplicados nesses meses caso as planilhas parciais (ex: intervalos de dias) possuam os mesmos chamados da planilha mensal. Recomenda-se investigar a origem nesses arquivos.")
        
    print("\n================ RESULTADO TABULAR ================")
    if not df_consolidado.empty:
        # Exibe no console formatando a data
        df_exibicao = df_consolidado.copy()
        df_exibicao['Mes'] = df_exibicao['Mes'].dt.strftime('%m/%Y')
        print(df_exibicao.to_string(index=False))
        
        # Salvar CSV
        caminho_csv = pasta_resultados / 'tabela_consolidada_inspecoes.csv'
        df_consolidado.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig')
        print(f"\nTabela salva em: {caminho_csv}")
        
        # Gerar e salvar Gráfico
        print("Gerando gráfico...")
        gerar_grafico(df_consolidado, pasta_resultados)
    else:
        print("Nenhum dado válido para consolidar.")

if __name__ == "__main__":
    executar_pipeline()