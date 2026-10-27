import pandas as pd
import plotly.graph_objects as go
from pathlib import Path


CAMINHO_BASE = Path(
    r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla"
)
ARQUIVO_ENTRADA = (
    CAMINHO_BASE
    / "resultados_correlacao_gerencias"
    / "1_Chamados_Classificados_por_Gerencia.xlsx"
)
PASTA_RESULTADOS = CAMINHO_BASE / "resultados_servicos_regionais_rj"

FORMULARIOS_EXCLUIDOS = {
    "INSPEÇÕES CSC REGIONAL",
    "INSPEÇÕES CSC REGIONAL 1.0",
    "INSPEÇÃO CSC",
}
GERENCIA_ALVO = "GER SERVICOS REGIONAIS CSC - RJ"
GESTOR_ALVO = "CARLA.SALOME@MRS.COM.BR"
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
        "Grupo_Atribuicao",
        "Gerência Geral",
    }
    colunas_ausentes = colunas_obrigatorias.difference(df.columns)
    if colunas_ausentes:
        raise ValueError(
            "Colunas obrigatórias ausentes na planilha consolidada: "
            + ", ".join(sorted(colunas_ausentes))
        )

    gerencia_normalizada = normalizar_texto(df["Gerência Geral"])
    gestor_normalizado = normalizar_texto(df["Email_Gestor_RH"])
    filtro = (
        gerencia_normalizada.eq(GERENCIA_ALVO)
        | gestor_normalizado.eq(GESTOR_ALVO)
    )
    return df.loc[filtro].copy()


def classificar_formularios(df_gerencia):
    df = df_gerencia.copy()
    formulario_normalizado = normalizar_texto(df["Formulário"])
    df["Tipo_Contagem"] = formulario_normalizado.map(
        lambda formulario: (
            "Inspeções CSC"
            if formulario in FORMULARIOS_EXCLUIDOS
            else "Demais formulários"
        )
    )
    df["Detalhamento"] = df["Formulário"]
    df.loc[
        df["Tipo_Contagem"].eq("Inspeções CSC"),
        "Detalhamento",
    ] = df.loc[
        df["Tipo_Contagem"].eq("Inspeções CSC"),
        "Grupo_Atribuicao",
    ]
    return df


def preparar_contagens(df_gerencia):
    df = classificar_formularios(df_gerencia)
    tipos = ["Demais formulários", "Inspeções CSC"]
    contagem = (
        df.groupby(
            ["Competência_Mensal", "Tipo_Contagem"], as_index=False
        )
        .size()
        .rename(columns={"size": "Chamados_Abertos"})
    )
    grade = pd.MultiIndex.from_product(
        [MESES_2026, tipos],
        names=["Competência_Mensal", "Tipo_Contagem"],
    )
    contagem = (
        contagem.set_index(["Competência_Mensal", "Tipo_Contagem"])
        .reindex(grade, fill_value=0)
        .rename_axis(["Competência_Mensal", "Tipo_Contagem"])
        .reset_index()
    )

    detalhes = (
        df.groupby(
            ["Competência_Mensal", "Tipo_Contagem", "Detalhamento"],
            as_index=False,
        )
        .size()
        .rename(columns={"size": "Quantidade"})
        .sort_values(
            ["Competência_Mensal", "Tipo_Contagem", "Quantidade"],
            ascending=[True, True, False],
        )
    )
    detalhes_hover = (
        detalhes.groupby(["Competência_Mensal", "Tipo_Contagem"])
        .apply(
            lambda grupo: "<br>".join(
                f"{formulario}: {quantidade}"
                for formulario, quantidade in zip(
                    grupo["Detalhamento"], grupo["Quantidade"]
                )
            ),
            include_groups=False,
        )
        .to_dict()
    )
    contagem["Detalhamento"] = [
        detalhes_hover.get((mes, tipo), "")
        for mes, tipo in zip(
            contagem["Competência_Mensal"], contagem["Tipo_Contagem"]
        )
    ]
    contagem["Mês"] = contagem["Competência_Mensal"].str[-2:].map(NOMES_MESES)
    return df, contagem, detalhes


def gerar_grafico(contagem):
    figura = go.Figure()
    cores = {
        "Demais formulários": "#1f497d",
        "Inspeções CSC": "#f2c94c",
    }
    for tipo in ["Demais formulários", "Inspeções CSC"]:
        dados_tipo = contagem[contagem["Tipo_Contagem"] == tipo]
        figura.add_trace(
            go.Bar(
                name=tipo,
                x=dados_tipo["Mês"],
                y=dados_tipo["Chamados_Abertos"],
                text=dados_tipo["Chamados_Abertos"],
                marker_color=cores[tipo],
                customdata=dados_tipo[["Detalhamento"]].to_numpy(),
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    f"<b>{tipo}</b><br>"
                    "Total: %{y}<br><br>"
                    "%{customdata[0]}<extra></extra>"
                ),
                textposition="outside",
            )
        )
    figura.update_layout(
        title="Chamados abertos por mês - Serviços Regionais - RJ",
        xaxis_title="Mês",
        yaxis_title="Quantidade de chamados",
        plot_bgcolor="white",
        barmode="group",
        legend_title="Tipo de formulário",
    )
    return figura


def main():
    PASTA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    df_gerencia = carregar_e_filtrar()

    if df_gerencia.empty:
        print("Nenhum chamado encontrado para a gerência informada.")
        return

    df_classificado, contagem_mensal, detalhamento = preparar_contagens(
        df_gerencia
    )
    df_sem_inspecoes = df_classificado[
        df_classificado["Tipo_Contagem"] == "Demais formulários"
    ].drop(columns=["Tipo_Contagem"])
    df_inspecoes = df_classificado[
        df_classificado["Tipo_Contagem"] == "Inspeções CSC"
    ].drop(columns=["Tipo_Contagem"])

    arquivo_filtrado = (
        PASTA_RESULTADOS
        / "Chamados_Servicos_Regionais_RJ_Sem_Inspecoes.xlsx"
    )
    arquivo_inspecoes = (
        PASTA_RESULTADOS
        / "Chamados_Servicos_Regionais_RJ_Inspecoes_CSC.xlsx"
    )
    arquivo_contagem = (
        PASTA_RESULTADOS
        / "Chamados_Servicos_Regionais_RJ_Contagem_Mensal.xlsx"
    )
    arquivo_detalhamento = (
        PASTA_RESULTADOS
        / "Chamados_Servicos_Regionais_RJ_Por_Formulario.xlsx"
    )
    arquivo_grafico = (
        PASTA_RESULTADOS
        / "Dashboard_Servicos_Regionais_RJ.html"
    )

    df_sem_inspecoes.to_excel(arquivo_filtrado, index=False)
    df_inspecoes.to_excel(arquivo_inspecoes, index=False)
    contagem_mensal.to_excel(arquivo_contagem, index=False)
    detalhamento.to_excel(arquivo_detalhamento, index=False)
    gerar_grafico(contagem_mensal).write_html(arquivo_grafico)

    print(f"Chamados da gerência: {len(df_gerencia)}")
    print(f"Demais formulários (barra azul): {len(df_sem_inspecoes)}")
    print(f"Inspeções CSC (barra amarela): {len(df_inspecoes)}")
    print(f"Gráfico: {arquivo_grafico}")


if __name__ == "__main__":
    main()
