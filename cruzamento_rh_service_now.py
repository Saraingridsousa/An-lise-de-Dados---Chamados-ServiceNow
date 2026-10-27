import pandas as pd
from pathlib import Path
import re

def main():
    print("[1] Iniciando cruzamento e purificação de dados: ServiceNow + RH")
    
    # Define caminhos
    caminho_base = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla")
    arquivo_canais = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - Chamados Service Now\Canal por Coordenação.xlsx")
    arquivo_rh = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - sys_user\Ativos RH 3.xlsx")
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
    
    # 2. Higieniza e padroniza e-mails na base de Canais
    def limpar_email(email):
        if pd.isna(email): return ""
        email = str(email).strip().lower()
        if '@' not in email and email != "":
            email = email + "@mrs.com.br"
        return email

    df_canais['Grupo_Atribuicao_Clean'] = df_canais['Grupo de atribuição'].fillna('').astype(str).str.strip().str.upper()
    df_canais['Nome_Gestor_Clean'] = df_canais['Coordenador'].fillna('').astype(str).str.strip().str.upper()
    df_canais['Email_Gestor_Clean'] = df_canais['Coordenador'].apply(limpar_email)
    df_canais['Gerencia_Geral_Clean'] = df_canais['Gerência'].fillna('').astype(str).str.strip().str.upper()

    df_rh['Matricula_Clean'] = (
        df_rh['Matricula']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )
    df_rh['Nome_Gestor_RH_Clean'] = df_rh['Nome Gestor'].fillna('').astype(str).str.strip().str.upper()
    df_rh['Email_Gestor_RH_Clean'] = df_rh['Email Gestor'].apply(limpar_email)

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

    # 5. Cruzamento Principal via Canal
    print("[5] Cruzando dados para atribuir a Gestão de cada Canal...")
    df_merged = pd.merge(
        df_sn,
        df_canais[['Grupo_Atribuicao_Clean', 'Nome_Gestor_Clean', 'Email_Gestor_Clean', 'Gerencia_Geral_Clean']],
        left_on='Grupo_Atribuicao_Clean',
        right_on='Grupo_Atribuicao_Clean',
        how='left'
    )
    
    # 6. Organização e Exportação Final
    print("[6] Preparando arquivo final de exportação...")
    colunas_finais = [
        'Número', 'Competência_Mensal', 'Data_Criacao', 
        'Matricula_Original', 'Nome_Ativo_RH', 'Solicitante_Original', 
        'Grupo_Atribuicao_Clean', 'Formulário',
        'Nome_Gestor_RH_Clean', 'Email_Gestor_RH_Clean',
        'Nome_Gestor_Clean', 'Email_Gestor_Clean', 'Gerencia_Geral_Clean'
    ]
    colunas_disponiveis = [col for col in colunas_finais if col in df_merged.columns]
    
    df_final = df_merged[colunas_disponiveis].copy()
    
    df_final.rename(columns={
        'Nome_Gestor_Clean': 'Nome_Gestor',
        'Email_Gestor_Clean': 'Email_Gestor',
        'Grupo_Atribuicao_Clean': 'Grupo_Atribuicao',
        'Nome_Gestor_RH_Clean': 'Nome_Gestor_RH',
        'Email_Gestor_RH_Clean': 'Email_Gestor_RH',
        'Gerencia_Geral_Clean': 'Gerência Geral'
    }, inplace=True)
    
    df_final.drop_duplicates(subset=['Número'], keep='first', inplace=True)
    
    arquivo_saida = pasta_resultados / "1_Chamados_Classificados_por_Gerencia.xlsx"
    df_final.to_excel(arquivo_saida, index=False)
    
    print(f"\n[SUCESSO] Base final exportada para:")
    print(f"Caminho: {arquivo_saida}")
    print(f"Total de Registros (bases mensais de 2026): {len(df_final)}")

if __name__ == "__main__":
    main()