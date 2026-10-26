import pandas as pd
from pathlib import Path
import re

def main():
    print("[1] Iniciando cruzamento e purificação de dados: ServiceNow + RH")
    
    # Define caminhos
    caminho_base = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla")
    arquivo_service_now = caminho_base / "x_mrsls_csc_sla.xlsx"
    arquivo_canais = caminho_base / "Canal por Coordenação.xlsx"
    pasta_resultados = caminho_base / "resultados_correlacao_gerencias"
    
    pasta_resultados.mkdir(parents=True, exist_ok=True)
    
    # 1. Carrega as bases
    print("[2] Carregando a base bruta do ServiceNow e matriz de Canais...")
    df_sn = pd.read_excel(arquivo_service_now, dtype=str)
    df_canais = pd.read_excel(arquivo_canais, dtype=str)
    
    # 2. Higieniza e padroniza e-mails na base de Canais
    def limpar_email(email):
        if pd.isna(email): return ""
        email = str(email).strip().lower()
        if '@' not in email and email != "":
            email = email + "@mrs.com.br"
        return email

    df_canais['Email_Gestor_Clean'] = df_canais['Email Gestor'].apply(limpar_email)
    df_canais['Nome_Gestor_Clean'] = df_canais['Gestor '].fillna('').astype(str).str.strip().str.upper()
    df_canais['Canal_Atendimento_Clean'] = df_canais['Canal de Atendimento'].fillna('').astype(str).str.strip().str.upper()

    # 3. Limpeza Inicial do ServiceNow e formatação de datas
    print("[3] Higienizando colunas e convertendo datas para competência 2026...")
    df_sn['Canal_Clean'] = df_sn['Subcategoria'].fillna('').astype(str).str.strip().str.upper()
    df_sn.rename(columns={'Criação de': 'Criacao_De_Original'}, inplace=True)
    
    # Processa e extrai a data de criação
    if 'Criado' in df_sn.columns:
        df_sn['Data_Criacao'] = pd.to_datetime(df_sn['Criado'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    else:
        df_sn['Data_Criacao'] = pd.to_datetime(df_sn.iloc[:, 1], format='%d/%m/%Y %H:%M:%S', errors='coerce')

    # Filtra apenas o ano de 2026 e cria coluna YYYY-MM
    tamanho_antes = len(df_sn)
    df_sn = df_sn[df_sn['Data_Criacao'].dt.year == 2026].copy()
    df_sn['Competência_Mensal'] = df_sn['Data_Criacao'].dt.strftime('%Y-%m')
    tamanho_depois = len(df_sn)
    print(f"    -> {tamanho_antes - tamanho_depois} chamados descartados (Fora de 2026 ou data inválida).")

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
    
    df_sn['Solicitante_Original'] = df_sn['Solicitante'].fillna('').astype(str).str.strip().str.upper()
    df_sn['Formulário'] = df_sn['Item'].fillna('').astype(str).str.strip()

    # 5. Cruzamento Principal via Canal
    print("[5] Cruzando dados para atribuir a Gestão de cada Canal...")
    df_merged = pd.merge(
        df_sn,
        df_canais[['Canal_Atendimento_Clean', 'Nome_Gestor_Clean', 'Email_Gestor_Clean', 'Gerência Geral']],
        left_on='Canal_Clean',
        right_on='Canal_Atendimento_Clean',
        how='left'
    )
    
    # 6. Organização e Exportação Final
    print("[6] Preparando arquivo final de exportação...")
    colunas_finais = [
        'Número', 'Competência_Mensal', 'Data_Criacao', 
        'Matricula_Original', 'Nome_Ativo_RH', 'Solicitante_Original', 
        'Canal_Clean', 'Formulário', 'Nome_Gestor_Clean', 'Email_Gestor_Clean', 'Gerência Geral'
    ]
    colunas_disponiveis = [col for col in colunas_finais if col in df_merged.columns]
    
    df_final = df_merged[colunas_disponiveis].copy()
    
    df_final.rename(columns={
        'Nome_Gestor_Clean': 'Nome_Gestor',
        'Email_Gestor_Clean': 'Email_Gestor',
        'Canal_Clean': 'Canal'
    }, inplace=True)
    
    df_final.drop_duplicates(subset=['Número'], keep='first', inplace=True)
    
    arquivo_saida = pasta_resultados / "1_Chamados_Classificados_por_Gerencia.xlsx"
    df_final.to_excel(arquivo_saida, index=False)
    
    print(f"\n[SUCESSO] Base final exportada para:")
    print(f"Caminho: {arquivo_saida}")
    print(f"Total de Registros (2026): {len(df_final)}")

if __name__ == "__main__":
    main()