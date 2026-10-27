import pandas as pd
import plotly.express as px
from pathlib import Path


CAMINHO_BASE = Path(
    r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla"
)
ARQUIVO_ENTRADA = (
    CAMINHO_BASE
    / "resultados_correlacao_gerencias"
    / "1_Chamados_Classificados_por_Gerencia.xlsx"
)
PASTA_RESULTADOS = CAMINHO_BASE / "resultados_servicos_regionais_mg"

FORMULARIOS_EXCLUIDOS = {
    "INSPEÇÕES CSC REGIONAL",
    "INSPEÇÕES CSC REGIONAL 1.0",
    "INSPEÇÃO CSC",
}
GERENCIA_ALVO = "SERVIÇOS REGIONAIS - MG"
MESES_2026 = [f"2026-{mes:02d}" for mes in range(1, 11)]
NOMES_MESES = {
    "01": "Janeiro",
    "02": "Fevereiro",
    "03": "Março",
    "04": "Abril",
    "05": "Maio",
    "06": "Junho",
    "07": "Julho",
    "08": "Agosto",
    "09": "Setembro",
    "10": "Outubro",
}


def normalizar_texto(serie):
    return serie.fillna("").astype(str).str.strip().str.upper()


def carregar_e_filtrar():
    if not ARQUIVO_ENTRADA.exists():
        raise FileNotFoundError(
            f"Planilha consolidada não encontrada: {ARQUIVO_ENTRADA}"
        )

    try:
        df = pd.read_excel(ARQUIVO_ENTRADA, dtype=str)
    except PermissionError as exc:
        raise PermissionError(
            "Não foi possível ler a planilha consolidada. "
            "Feche o arquivo no Excel e execute o script novamente: "
            f"{ARQUIVO_ENTRADA}"
        ) from exc
    colunas_obrigatorias = {
        "Competência_Mensal",
        "Formulário",
        "Gerência Geral",
    }
    colunas_ausentes = colunas_obrigatorias.difference(df.columns)
    if colunas_ausentes:
        raise ValueError(
            "Colunas obrigatórias ausentes na planilha consolidada: "
            + ", ".join(sorted(colunas_ausentes))
        )

    formulario_normalizado = normalizar_texto(df["Formulário"])
    gerencia_normalizada = normalizar_texto(df["Gerência Geral"])

    filtro = (
        gerencia_normalizada.eq(GERENCIA_ALVO)
        & ~formulario_normalizado.isin(FORMULARIOS_EXCLUIDOS)
    )
    return df.loc[filtro].copy()


def gerar_grafico(df_filtrado):
    contagem_mensal = (
        df_filtrado.groupby("Competência_Mensal", as_index=False)
        .size()
        .rename(columns={"size": "Chamados_Abertos"})
    )
    contagem_mensal = (
        contagem_mensal.set_index("Competência_Mensal")
        .reindex(MESES_2026, fill_value=0)
        .rename_axis("Competência_Mensal")
        .reset_index()
    )

    detalhes_por_mes = (
        df_filtrado.groupby(["Competência_Mensal", "Formulário"])
        .size()
        .reset_index(name="Quantidade")
        .sort_values(["Competência_Mensal", "Quantidade"], ascending=[True, False])
    )
    detalhes_hover = (
        detalhes_por_mes.groupby("Competência_Mensal")
        .apply(
            lambda grupo: "<br>".join(
                f"{formulario}: {quantidade}"
                for formulario, quantidade in zip(
                    grupo["Formulário"], grupo["Quantidade"]
                )
            ),
            include_groups=False,
        )
        .rename("Detalhamento")
    )
    contagem_mensal["Detalhamento"] = (
        contagem_mensal["Competência_Mensal"]
        .map(detalhes_hover)
        .fillna("")
    )
    contagem_mensal["Mês"] = contagem_mensal["Competência_Mensal"].str[-2:].map(
        NOMES_MESES
    )

    fig = px.bar(
        contagem_mensal,
        x="Mês",
        y="Chamados_Abertos",
        text="Chamados_Abertos",
        title=(
            "Formulários abertos por mês - Serviços Regionais - MG"
        ),
        labels={
            "Mês": "Mês",
            "Chamados_Abertos": "Chamados abertos",
        },
    )
    fig.update_traces(
        marker_color="#1f497d",
        textposition="outside",
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Total de chamados: %{y}<br><br>"
            "%{customdata[0]}<extra></extra>"
        ),
        customdata=contagem_mensal[["Detalhamento", "Competência_Mensal"]],
    )
    fig.update_layout(
        xaxis_title="Mês",
        yaxis_title="Quantidade de chamados",
        plot_bgcolor="white",
        showlegend=False,
    )
    return contagem_mensal, fig


def main():
    PASTA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    df_filtrado = carregar_e_filtrar()

    arquivo_filtrado = (
        PASTA_RESULTADOS
        / "Chamados_Servicos_Regionais_MG_Sem_Inspecoes.xlsx"
    )
    df_filtrado.to_excel(arquivo_filtrado, index=False)

    if df_filtrado.empty:
        print("Nenhum chamado encontrado para os filtros informados.")
        return

    contagem_mensal, figura = gerar_grafico(df_filtrado)
    detalhamento_formularios = (
        df_filtrado.groupby(
            ["Competência_Mensal", "Formulário"], as_index=False
        )
        .size()
        .rename(columns={"size": "Chamados_Abertos"})
        .sort_values(
            ["Competência_Mensal", "Chamados_Abertos"],
            ascending=[True, False],
        )
    )
    arquivo_contagem = (
        PASTA_RESULTADOS
        / "Chamados_Servicos_Regionais_MG_Contagem_Mensal.xlsx"
    )
    arquivo_grafico = (
        PASTA_RESULTADOS
        / "Dashboard_Servicos_Regionais_MG.html"
    )
    contagem_mensal.astype({"Competência_Mensal": str}).to_excel(
        arquivo_contagem, index=False
    )
    arquivo_detalhamento = (
        PASTA_RESULTADOS
        / "Chamados_Servicos_Regionais_MG_Por_Formulario.xlsx"
    )
    detalhamento_formularios.to_excel(arquivo_detalhamento, index=False)
    figura.write_html(arquivo_grafico)

    print(f"Chamados após os filtros: {len(df_filtrado)}")
    print(f"Planilha filtrada: {arquivo_filtrado}")
    print(f"Contagem mensal: {arquivo_contagem}")
    print(f"Detalhamento por formulário: {arquivo_detalhamento}")
    print(f"Gráfico: {arquivo_grafico}")


if __name__ == "__main__":
    main()
