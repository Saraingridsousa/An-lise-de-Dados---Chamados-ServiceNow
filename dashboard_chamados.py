import pandas as pd
import plotly.express as px
from pathlib import Path
import re
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# UTILITÁRIO PARA TRADUZIR COLUNA
# ==========================================
def indice_para_coluna_excel(idx):
    col = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        col = chr(65 + remainder) + col
    return col

# ==========================================
# REGRAS DE CLASSIFICAÇÃO (BUSCA EXATA)
# ==========================================
def classificar_formulario(nome_original):
    if pd.isna(nome_original):
        return 'Ignorado (Vazio)'
    
    nome_limpo = str(nome_original).strip()
    
    formularios_validos = [
        "Inspeções CSC Regional",
        "Inspeções CSC Regional 1.0",
        "Inspeção CSC"
    ]
    
    if nome_limpo in formularios_validos:
        return 'Inspeções CSC (Válido)'
    
    return 'Pendente de Validação'

# ==========================================
# EXTRAÇÃO DE MÊS DO NOME DO ARQUIVO
# ==========================================
def extrair_competencia_nome(nome_arquivo):
    match = re.search(r'(20\d{2})\.(\d{2})', nome_arquivo)
    if match:
        return pd.Period(f"{match.group(1)}-{match.group(2)}", freq='M')
    return None

def encontrar_coluna(df, possiveis_nomes):
    colunas_df = {str(c).strip().lower(): c for c in df.columns}
    for nome in possiveis_nomes:
        if nome.lower() in colunas_df:
            return colunas_df[nome.lower()]
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

    print("\n" + "="*70)
    print("INÍCIO DA VARREDURA DE ARQUIVOS")
    print("="*70)

    for arquivo in base_path.iterdir():
        if not arquivo.is_file() or arquivo.name.startswith('~$') or arquivo.suffix.lower() not in extensoes_validas or arquivo.stem in nomes_excluidos:
            continue

        try:
            mes_arquivo = extrair_competencia_nome(arquivo.name)
            
            if not mes_arquivo:
                print(f"[!] IGNORADO: {arquivo.name} (Sem mês identificado no nome)")
                continue
                
            if arquivo.suffix.lower() == '.csv':
                df = pd.read_csv(arquivo, low_memory=False)
            else:
                df = pd.read_excel(arquivo)
                
            df.columns = [str(c).strip() for c in df.columns]
            
            col_form = encontrar_coluna(df, ['Formulário', 'Formulario', 'Item'])
            col_id = encontrar_coluna(df, ['Número', 'Numero', 'ID', 'Ticket'])
            
            if not col_form:
                print(f"[X] IGNORADO: {arquivo.name}")
                print(f"    Motivo: Coluna 'Formulário' não encontrada.\n")
                continue
                
            idx_col = df.columns.get_loc(col_form)
            letra_col = indice_para_coluna_excel(idx_col)
            
            print(f"[+] LENDO: {arquivo.name}")
            print(f"    -> Coluna 'Formulário' encontrada: Letra {letra_col}")
            
            df['Categoria_Classificada'] = df[col_form].apply(classificar_formulario)
            df['Mes_Apuracao'] = mes_arquivo
            
            if col_id:
                df['ID_Chamado'] = df[col_id].astype(str)
            else:
                df['ID_Chamado'] = arquivo.name + "_" + df.index.astype(str)
                
            df['Formulário_Original'] = df[col_form].astype(str).str.strip()
            
            df_extract = df[['ID_Chamado', 'Mes_Apuracao', 'Formulário_Original', 'Categoria_Classificada']].copy()
            df_extract['Arquivo_Origem'] = arquivo.name
            
            # Limpa chamados duplicados APENAS dentro desta planilha
            df_extract = df_extract.drop_duplicates(subset=['ID_Chamado'])
            
            validos_aqui = df_extract[df_extract['Categoria_Classificada'] == 'Inspeções CSC (Válido)']
            if validos_aqui.empty:
                print("    -> Nenhum chamado das categorias exatas foi encontrado nesta planilha.\n")
            else:
                contagem_por_nome = validos_aqui['Formulário_Original'].value_counts()
                # O NÚMERO IMPRESSO AQUI AGORA VAI DIRETO PARA O GRÁFICO:
                print(f"    -> Total de formulários válidos contabilizados: {len(validos_aqui)}")
                for nome, qtd in contagem_por_nome.items():
                    print(f"       - {qtd} chamados de: '{nome}'")
                print("") 
            
            lista_registros.append(df_extract)
            
        except Exception as e:
            print(f"[!] ERRO ao processar '{arquivo.name}': {e}\n")

    print("="*70)

    if not lista_registros:
        return pd.DataFrame()
        
    df_consolidado = pd.concat(lista_registros, ignore_index=True)
    df_consolidado = df_consolidado.dropna(subset=['Mes_Apuracao'])
    
    # Mantém apenas os meses de 2026
    df_consolidado = df_consolidado[df_consolidado['Mes_Apuracao'].dt.year == 2026]
    
    return df_consolidado

# ==========================================
# GERAÇÃO DO GRÁFICO INTERATIVO
# ==========================================
def gerar_grafico_interativo(df_valido, diretorio_saida):
    df_plot = df_valido.groupby(['Mes_Apuracao', 'Formulário_Original']).size().reset_index(name='Chamados')
    df_plot['Mes_Str'] = df_plot['Mes_Apuracao'].dt.strftime('%m/%Y')
    
    # Ordena rigidamente pela linha do tempo
    df_plot = df_plot.sort_values('Mes_Apuracao')

    totais_mensais = df_plot.groupby('Mes_Str', sort=False)['Chamados'].sum().reset_index(name='Total_Mes')

    fig = px.bar(
        df_plot, 
        x='Mes_Str', 
        y='Chamados', 
        color='Formulário_Original', 
        title='Quantidade Mensal de Chamados (Inspeções CSC)',
        labels={'Mes_Str': 'Competência', 'Chamados': 'Qtd de Chamados', 'Formulário_Original': 'Nome Original'},
        color_discrete_sequence=['#1f497d'] 
    )

    fig.update_layout(
        barmode='stack',
        hovermode='x unified',
        xaxis_title="Mês/Ano",
        yaxis_title="Total de Chamados",
        showlegend=False, 
        plot_bgcolor='white'
    )
    
    # Trava do eixo X: Garante que os meses não saiam de ordem
    meses_em_ordem = totais_mensais['Mes_Str'].tolist()
    fig.update_xaxes(categoryorder='array', categoryarray=meses_em_ordem)
    
    fig.update_traces(marker_line_width=0, textposition='none')
    
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
    print(f"\n[Visualização] Gráfico gerado com sucesso: {caminho_html}")
    fig.show()

# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
def main():
    caminho_entrada = r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla"
    pasta_resultados = Path(caminho_entrada) / "resultados_analise"
    pasta_resultados.mkdir(exist_ok=True)
    
    df_geral = processar_arquivos(caminho_entrada)
    
    if df_geral.empty:
        print("Nenhum dado válido de 2026 encontrado para processamento.")
        return

    df_inspecoes = df_geral[df_geral['Categoria_Classificada'] == 'Inspeções CSC (Válido)']
    
    tabela_detalhada = df_inspecoes.groupby(['Mes_Apuracao', 'Formulário_Original']).size().reset_index(name='Quantidade')
    tabela_detalhada = tabela_detalhada.sort_values(['Mes_Apuracao', 'Quantidade'], ascending=[True, False])
    
    auditoria_nomes = df_geral.groupby(['Formulário_Original', 'Categoria_Classificada']).size().reset_index(name='Total_Chamados')
    auditoria_nomes = auditoria_nomes.sort_values('Total_Chamados', ascending=False)

    caminho_detalhado = pasta_resultados / 'Detalhamento_Inspecoes_Mensal_2026.xlsx'
    caminho_auditoria = pasta_resultados / 'Auditoria_Classificacao_Formularios_2026.xlsx'
    
    tabela_detalhada['Mes_Apuracao'] = tabela_detalhada['Mes_Apuracao'].astype(str)
    tabela_detalhada.to_excel(caminho_detalhado, index=False)
    auditoria_nomes.to_excel(caminho_auditoria, index=False)
    
    print("\n--- RESUMO FINAL DA EXECUÇÃO ---")
    print(f"Total de chamados validados (Ano 2026): {len(df_inspecoes)}")
    
    if not df_inspecoes.empty:
        gerar_grafico_interativo(df_inspecoes, pasta_resultados)

if __name__ == "__main__":
    main()