import pandas as pd
import plotly.express as px
from pathlib import Path

def gerar_grafico_gestor():
    # Caminho do arquivo detalhado que geramos na Etapa 2
    caminho_base = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla")
    arquivo_dados = caminho_base / "resultados_correlacao_gerencias" / "1_Chamados_Classificados_por_Gerencia.xlsx"
    
    if not arquivo_dados.exists():
        print("Erro: O arquivo 1_Chamados_Classificados_por_Gerencia.xlsx não foi encontrado.")
        return

    # Lê os dados garantindo que o mês seja tratado como texto para o gráfico
    df = pd.read_excel(arquivo_dados, dtype=str)
    
    # Filtra apenas os chamados onde o e-mail ou nome do gestor contém "fabiana"
    # O preenchimento com str(x).lower() evita erros com células vazias
    filtro_fabiana = df['Email_Gestor'].apply(lambda x: 'fabiana' in str(x).lower()) | \
                     df['Nome_Gestor'].apply(lambda x: 'fabiana' in str(x).lower())
    
    df_fabiana = df[filtro_fabiana].copy()
    
    if df_fabiana.empty:
        print("Nenhum chamado encontrado para a gestora Fabiana Mara.")
        return

    # Agrupa a quantidade de chamados por Mês
    df_plot = df_fabiana.groupby('Competência_Mensal').size().reset_index(name='Total_Chamados')
    df_plot = df_plot.sort_values('Competência_Mensal')

    # Criação do Gráfico Interativo no Plotly
    fig = px.bar(
        df_plot, 
        x='Competência_Mensal', 
        y='Total_Chamados',
        title='Evolução Mensal de Chamados - Gestão: Fabiana Mara (Exclui Inspeções CSC)',
        labels={'Competência_Mensal': 'Mês de Apuração', 'Total_Chamados': 'Quantidade de Chamados'},
        color_discrete_sequence=['#1f497d'] # Azul padrão
    )

    fig.update_layout(
        xaxis_title="Mês e Ano",
        yaxis_title="Total de Chamados",
        plot_bgcolor='white'
    )
    
    # Adiciona o número total cravado no topo de cada barra
    for _, row in df_plot.iterrows():
        fig.add_annotation(
            x=row['Competência_Mensal'], 
            y=row['Total_Chamados'], 
            text=str(row['Total_Chamados']),
            showarrow=False,
            yshift=10, 
            font=dict(size=14, color="black", family="Arial Black")
        )

    # Remove contornos para visual limpo
    fig.update_traces(marker_line_width=0)
    
    caminho_html = arquivo_dados.parent / 'Dashboard_Fabiana_Mara.html'
    fig.write_html(caminho_html)
    print(f"\n[Sucesso] Gráfico gerado e salvo em: {caminho_html}")
    fig.show()

if __name__ == "__main__":
    gerar_grafico_gestor()