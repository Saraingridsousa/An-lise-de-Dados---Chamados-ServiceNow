import pandas as pd
from pathlib import Path
import re

def main():
    print("[1] Iniciando cruzamento e purificação de dados: ServiceNow + RH")
    
    # Define caminhos
    caminho_base = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla")
    arquivo_canais = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - Chamados Service Now\Canal por Coordenação.xlsx")
    arquivo_rh = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - sys_user\Ativos RH 3.xlsx")
    arquivo_desligados = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - sys_user\RH- Relatório de desligados.xlsx")
    pasta_resultados = caminho_base / "resultados_correlacao_gerencias"
    
    pasta_resultados.mkdir(parents=True, exist_ok=True)
    
    # 1. Carrega as bases
    print("[2] Carregando as bases mensais de 2026 e a matriz de Canais...")
    arquivos_service_now = sorted(
        caminho_base.glob("2026.[0-9][0-9]_x_mrsls_csc_sla.xlsx")
    )
    if not arquivos_service_now:
        raise FileNotFoundError(
            f"Nenhuma base mensal de 2026 foi encontrada em: {caminho_base}"
        )

    bases_service_now = []
    for arquivo in arquivos_service_now:
        df_mes = pd.read_excel(arquivo, dtype=str)
        competencia = arquivo.stem[:7].replace('.', '-')
        df_mes['Competência_Mensal'] = competencia
        bases_service_now.append(df_mes)

    df_sn = pd.concat(bases_service_now, ignore_index=True)
    df_canais = pd.read_excel(arquivo_canais, sheet_name='2026', dtype=str)
    df_rh = pd.read_excel(arquivo_rh, sheet_name='Ativos RH', dtype=str)
    df_desligados = pd.read_excel(arquivo_desligados, dtype=str)
    
    # 2. Higieniza e padroniza e-mails na base de Canais
    def limpar_email(email):
        if pd.isna(email): return ""
        email = str(email).strip().lower()
        if '@' not in email and email != "":
            email = email + "@mrs.com.br"
        return email

    df_canais['Grupo_Atribuicao_Clean'] = df_canais['Grupo de atribuição'].fillna('').astype(str).str.strip().str.upper()
    df_canais['Coordenador_Clean'] = df_canais['Coordenador'].apply(limpar_email)
    df_canais['Gerente_Clean'] = df_canais['Gerente'].apply(limpar_email)
    df_canais['Gerencia_Clean'] = df_canais['Gerência'].fillna('').astype(str).str.strip().str.upper()

    # A gerência hierárquica do chamado é a gerência do gestor do solicitante.
    # O gestor pode aparecer na coluna Coordenador ou Gerente da matriz de canais.
    gestores_canais = pd.concat(
        [
            df_canais[['Coordenador_Clean', 'Gerencia_Clean']].rename(
                columns={'Coordenador_Clean': 'Email_Gestor_RH_Clean'}
            ),
            df_canais[['Gerente_Clean', 'Gerencia_Clean']].rename(
                columns={'Gerente_Clean': 'Email_Gestor_RH_Clean'}
            ),
        ],
        ignore_index=True,
    )
    gestores_canais = gestores_canais[
        gestores_canais['Email_Gestor_RH_Clean'].ne('')
    ].drop_duplicates()

    df_rh['Matricula_Clean'] = (
        df_rh['Matricula']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )
    df_rh['Nome_Gestor_RH_Clean'] = df_rh['Nome Gestor'].fillna('').astype(str).str.strip().str.upper()
    df_rh['Email_Gestor_RH_Clean'] = df_rh['Email Gestor'].apply(limpar_email)
    df_rh['Nome_Clean'] = df_rh['Nome'].fillna('').astype(str).str.strip().str.upper()
    df_desligados['Matricula_Clean'] = (
        df_desligados['Nº pessoal']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )
    df_desligados['Nome_Desligado_Clean'] = (
        df_desligados['Nome'].fillna('').astype(str).str.strip().str.upper()
    )
    df_desligados['Unidade_Desligado_Clean'] = (
        df_desligados['Desc.Unidade de Lotação']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # 3. Limpeza Inicial do ServiceNow e formatação de datas
    print("[3] Higienizando colunas e convertendo datas...")
    df_sn['Grupo_Atribuicao_Clean'] = df_sn['Grupo de atribuição'].fillna('').astype(str).str.strip().str.upper()
    df_sn.rename(columns={'Criação de': 'Criacao_De_Original'}, inplace=True)
    
    # Processa e extrai a data de criação
    if 'Criação em' in df_sn.columns:
        df_sn['Data_Criacao'] = pd.to_datetime(df_sn['Criação em'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    else:
        df_sn['Data_Criacao'] = pd.to_datetime(df_sn.iloc[:, 1], format='%d/%m/%Y %H:%M:%S', errors='coerce')

    # A competência vem do nome do arquivo mensal, não da data interna do chamado.
    # Isso preserva chamados do arquivo de cada mês, mesmo que outro campo contenha
    # uma referência a 2025 ou a outro período.
    print(f"    -> {len(df_sn)} chamados carregados das bases mensais de 2026.")

    # 4. Extração segura da Matrícula de quem abriu (Criação de)
    print("[4] Extraindo matrículas de quem abriu o chamado...")
    def extrair_matricula(texto):
        if pd.isna(texto): return ""
        match = re.search(r'\b\d+\b', str(texto))
        return match.group(0) if match else str(texto).strip().upper()

    df_sn['Matricula_Original'] = df_sn['Criacao_De_Original'].apply(extrair_matricula)
    
    def extrair_nome_sem_matricula(texto):
        if pd.isna(texto): return ""
        texto_limpo = re.sub(r'\b\d+\b', '', str(texto)).strip().upper()
        texto_limpo = re.sub(r'[^A-Z\s]', '', texto_limpo)
        return re.sub(r'\s+', ' ', texto_limpo).strip()

    df_sn['Nome_Ativo_RH'] = df_sn['Criacao_De_Original'].apply(extrair_nome_sem_matricula)
    
    df_sn['Solicitante_Original'] = df_sn['Nome do Solicitante'].fillna('').astype(str).str.strip().str.upper()
    df_sn['Solicitante_Original'] = df_sn['Solicitante_Original'].where(
        df_sn['Solicitante_Original'].ne(''),
        df_sn['Solicitante'].fillna('').astype(str).str.strip().str.upper()
    )
    df_sn['Formulário'] = df_sn['Formulário'].fillna('').astype(str).str.strip()

    df_sn['Matricula_Clean'] = (
        df_sn['Matrícula']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )
    df_sn = pd.merge(
        df_sn,
        df_rh[
            [
                'Matricula_Clean',
                'Nome',
                'Gerência Geral',
                'Email Usuario',
                'Nome_Gestor_RH_Clean',
                'Email_Gestor_RH_Clean'
            ]
        ],
        on='Matricula_Clean',
        how='left',
        suffixes=('', '_RH')
    )
    df_sn['Nome_Ativo_RH'] = df_sn['Nome'].fillna(df_sn['Nome_Ativo_RH']).astype(str).str.strip().str.upper()

    # A partir de agosto, as exportações mensais não trazem a matrícula.
    # Nesses arquivos, o solicitante é identificado pelo nome.
    rh_por_nome = (
        df_rh[df_rh['Nome_Clean'].ne('')]
        .drop_duplicates('Nome_Clean')
        .set_index('Nome_Clean')
    )
    nome_solicitante = df_sn['Solicitante_Original']
    for coluna in ['Nome', 'Nome_Gestor_RH_Clean', 'Email_Gestor_RH_Clean']:
        valores_por_nome = nome_solicitante.map(rh_por_nome[coluna])
        df_sn[coluna] = df_sn[coluna].fillna(valores_por_nome)
    df_sn['Nome_Ativo_RH'] = df_sn['Nome'].fillna(
        df_sn['Nome_Ativo_RH']
    ).astype(str).str.strip().str.upper()

    # Pessoas desligadas não possuem gestor na base de desligados. Quando
    # a unidade de lotação identifica a regional, a gerência é atribuída
    # diretamente, sem preencher um gestor que não foi informado.
    desligados_por_matricula = (
        df_desligados[df_desligados['Matricula_Clean'].ne('')]
        .drop_duplicates('Matricula_Clean')
        .set_index('Matricula_Clean')
    )
    df_sn['Nome_Desligado_Clean'] = df_sn['Matricula_Clean'].map(
        desligados_por_matricula['Nome_Desligado_Clean']
    )
    df_sn['Unidade_Desligado_Clean'] = df_sn['Matricula_Clean'].map(
        desligados_por_matricula['Unidade_Desligado_Clean']
    )
    desligados_por_nome = (
        df_desligados[df_desligados['Nome_Desligado_Clean'].ne('')]
        .drop_duplicates('Nome_Desligado_Clean')
        .set_index('Nome_Desligado_Clean')
    )
    nome_solicitante = df_sn['Solicitante_Original']
    df_sn['Unidade_Desligado_Clean'] = df_sn['Unidade_Desligado_Clean'].fillna(
        nome_solicitante.map(desligados_por_nome['Unidade_Desligado_Clean'])
    )
    df_sn['Gerencia_Desligado_Clean'] = ''
    regional_mg = df_sn['Unidade_Desligado_Clean'].str.contains(
        r'COORD\. SERVICOS COMPARTILHADOS \(MG\)',
        na=False,
        regex=True,
    )
    df_sn.loc[regional_mg, 'Gerencia_Desligado_Clean'] = 'SERVIÇOS REGIONAIS - MG'
    regional_sp = df_sn['Unidade_Desligado_Clean'].isin(
        [
            'COORD. SERVICOS COMPARTILHADOS (SP)',
            'GER. SERVICOS REGIONAIS CSC - SP',
        ]
    )
    df_sn.loc[regional_sp, 'Gerencia_Desligado_Clean'] = (
        'GER SERVICOS REGIONAIS CSC - SP'
    )
    regional_rj = df_sn['Unidade_Desligado_Clean'].isin(
        [
            'COORD. SERVICOS COMPARTILHADOS (RJ)',
            'GER. SERVICOS REGIONAIS CSC - RJ',
        ]
    )
    df_sn.loc[regional_rj, 'Gerencia_Desligado_Clean'] = (
        'GER SERVICOS REGIONAIS CSC - RJ'
    )

    # 5. Identifica a gerência hierárquica pelo gestor do solicitante
    print("[5] Cruzando solicitante, gestor do RH e gerência...")
    df_merged = pd.merge(
        df_sn,
        gestores_canais,
        on='Email_Gestor_RH_Clean',
        how='left'
    )
    df_merged['Gerencia_Clean'] = df_merged['Gerencia_Clean'].fillna(
        df_merged['Gerencia_Desligado_Clean']
    )
    
    # 6. Organização e Exportação Final
    print("[6] Preparando arquivo final de exportação...")
    colunas_finais = [
        'Número', 'Competência_Mensal', 'Data_Criacao', 
        'Matricula_Original', 'Nome_Ativo_RH', 'Solicitante_Original', 
        'Grupo_Atribuicao_Clean', 'Formulário',
        'Nome_Gestor_RH_Clean', 'Email_Gestor_RH_Clean',
        'Gerencia_Clean'
    ]
    colunas_disponiveis = [col for col in colunas_finais if col in df_merged.columns]
    
    df_final = df_merged[colunas_disponiveis].copy()
    
    df_final.rename(columns={
        'Grupo_Atribuicao_Clean': 'Grupo_Atribuicao',
        'Nome_Gestor_RH_Clean': 'Nome_Gestor',
        'Email_Gestor_RH_Clean': 'Email_Gestor_RH',
        'Gerencia_Clean': 'Gerência Geral'
    }, inplace=True)
    
    df_final.drop_duplicates(subset=['Número'], keep='first', inplace=True)
    
    arquivo_saida = pasta_resultados / "1_Chamados_Classificados_por_Gerencia.xlsx"
    df_final.to_excel(arquivo_saida, index=False)
    
    print(f"\n[SUCESSO] Base final exportada para:")
    print(f"Caminho: {arquivo_saida}")
    print(f"Total de Registros (bases mensais de 2026): {len(df_final)}")

if __name__ == "__main__":
    main()