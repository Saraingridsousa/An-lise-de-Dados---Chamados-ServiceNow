import pandas as pd
from pathlib import Path
import re
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# 1. FUNÇÕES DE NORMALIZAÇÃO
# ==========================================
def normalizar_texto(texto):
    """Remove acentos, espaços extras e converte para minúsculas."""
    if pd.isna(texto) or str(texto).strip() == "":
        return ""
    texto = str(texto).strip().lower()
    # Tratamento básico de acentos em português
    mapa_acentos = str.maketrans("áàãâéêíóôõúüç", "aaaaeeiooouuc")
    texto = texto.translate(mapa_acentos)
    # Remove espaços duplos entre as palavras
    texto = re.sub(r'\s+', ' ', texto)
    return texto

def normalizar_matricula(mat):
    """Garante que a matrícula seja um texto limpo, sem o '.0' que o pandas inventa."""
    if pd.isna(mat) or str(mat).strip() == "":
        return ""
    mat_str = str(mat).strip().upper()
    if mat_str.endswith('.0'):
        mat_str = mat_str[:-2]
    return mat_str

def encontrar_coluna(df, possiveis_nomes):
    """Busca dinâmica por colunas ignorando maiúsculas e espaços"""
    colunas_df = {normalizar_texto(c): c for c in df.columns}
    for nome in possiveis_nomes:
        nome_norm = normalizar_texto(nome)
        if nome_norm in colunas_df:
            return colunas_df[nome_norm]
    return None

# ==========================================
# 2. PREPARAÇÃO DAS BASES DO RH (DICIONÁRIOS)
# ==========================================
def carregar_base_rh(caminho, mapeamento_colunas):
    """Lê a base do RH, mapeia as colunas e cria dicionários de busca rápida e segura."""
    if not Path(caminho).exists():
        print(f"[Erro] Base de RH não encontrada: {caminho}")
        return {}, {}
    
    # Lê forçando as colunas principais como string para não perder zeros à esquerda
    df = pd.read_excel(caminho, dtype=str)
    
    # Renomeia as colunas do RH para um padrão interno para facilitar o manuseio
    colunas_presentes = {}
    for col_interna, nomes_possiveis in mapeamento_colunas.items():
        col_real = encontrar_coluna(df, nomes_possiveis)
        if col_real:
            colunas_presentes[col_interna] = col_real
            
    if 'nome' not in colunas_presentes or 'matricula' not in colunas_presentes:
        print(f"[Erro] Colunas de Nome ou Matrícula não encontradas no arquivo: {caminho}")
        return {}, {}

    dict_nome = {}
    dict_mat = {}

    for _, row in df.iterrows():
        nome_norm = normalizar_texto(row[colunas_presentes['nome']])
        mat_norm = normalizar_matricula(row[colunas_presentes['matricula']])
        
        # Puxa todos os dados disponíveis dessa linha
        dados_linha = {k: row[v] if pd.notna(row[v]) else "" for k, v in colunas_presentes.items()}
        
        # Estratégia Anti-Homônimos (Se o nome já existe, marca como AMBIGUO)
        if nome_norm != "":
            if nome_norm in dict_nome:
                dict_nome[nome_norm] = "AMBIGUO"
            else:
                dict_nome[nome_norm] = dados_linha
                
        if mat_norm != "":
            dict_mat[mat_norm] = dados_linha

    return dict_nome, dict_mat

# ==========================================
# 3. LÓGICA DE CRUZAMENTO (O MOTOR)
# ==========================================
def buscar_funcionario(nome_orig, mat_orig, rh_ativos, rh_desligados):
    """Aplica a regra de negócio exata da prioridade 1 a 5."""
    nome_norm = normalizar_texto(nome_orig)
    mat_norm = normalizar_matricula(mat_orig)
    
    dict_ati_nome, dict_ati_mat = rh_ativos
    dict_des_nome, dict_des_mat = rh_desligados
    
    # Retornos padrão vazios
    dados_encontrados = {}
    status = "Não localizado"
    
    if nome_norm == "" and mat_norm == "":
        return dados_encontrados, "Dados insuficientes (Nome e Matrícula vazios)"

    # ETAPA 1: Ativos por Nome
    if nome_norm != "" and nome_norm in dict_ati_nome:
        if dict_ati_nome[nome_norm] != "AMBIGUO":
            return dict_ati_nome[nome_norm], "Ativo - Nome"
            
    # ETAPA 2: Ativos por Matrícula
    if mat_norm != "" and mat_norm in dict_ati_mat:
        return dict_ati_mat[mat_norm], "Ativo - Matrícula"
        
    # ETAPA 3: Desligados por Nome
    if nome_norm != "" and nome_norm in dict_des_nome:
        if dict_des_nome[nome_norm] != "AMBIGUO":
            return dict_des_nome[nome_norm], "Desligado - Nome"
            
    # ETAPA 4: Desligados por Matrícula
    if mat_norm != "" and mat_norm in dict_des_mat:
        return dict_des_mat[mat_norm], "Desligado - Matrícula"
        
    # ETAPA 5: Inconclusivo / Não Localizado
    if nome_norm in dict_ati_nome and dict_ati_nome[nome_norm] == "AMBIGUO":
        status = "Pendente de validação (Homônimo no RH Ativo sem matrícula exata)"
    elif nome_norm in dict_des_nome and dict_des_nome[nome_norm] == "AMBIGUO":
        status = "Pendente de validação (Homônimo no RH Desligado sem matrícula exata)"

    return dados_encontrados, status

# ==========================================
# 4. PIPELINE PRINCIPAL
# ==========================================
def executar_modulo_rh():
    caminho_base = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla")
    pasta_resultados = caminho_base / "resultados_cruzamento_rh"
    pasta_resultados.mkdir(exist_ok=True)
    
    caminho_ativos = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - sys_user\Ativos RH 3.xlsx")
    caminho_desligados = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - sys_user\RH- Relatório de desligados.xlsx")

    print("Carregando bases do RH...")
    map_ativos = {
        'nome': ['Nome'], 'matricula': ['Matricula'], 'diretoria': ['Diretoria'], 
        'ger_geral': ['Gerência Geral'], 'gerencia': ['Gerência'], 'unidade': ['Unidade de Lotação'],
        'funcao': ['Função'], 'admissao': ['Admissão'], 'cc': ['CC'], 'email': ['Email Usuario'],
        'nome_gestor': ['Nome Gestor'], 'funcao_gestor': ['Função Gestor'], 'email_gestor': ['Email Gestor']
    }
    rh_ativos = carregar_base_rh(caminho_ativos, map_ativos)
    
    map_desligados = {
        'nome': ['Nome', 'Desc Funcionário'], 'matricula': ['Nº pessoal', 'Cód Funcionário'],
        'cod_unidade': ['Cód. Unidade de Lotação'], 'desc_unidade': ['Desc.Unidade de Lotação'],
        'admissao': ['Dt. Admissão'], 'localidade': ['Localidade'], 'desligamento': ['Data de desligamento'],
        'cc': ['Centro de Custo'], 'desc_cc': ['Desc. Centro de Custo']
    }
    rh_desligados = carregar_base_rh(caminho_desligados, map_desligados)

    print("Varrendo planilhas de chamados mensais...")
    arquivos_validos = []
    nomes_excluidos = {'ativo_x_mrsls_csc_sla', '2025_x_mrsls_csc_sla'}
    
    registros_analisados = []
    estatisticas = {
        'Total Analisado': 0, 'Ativo - Nome': 0, 'Ativo - Matrícula': 0,
        'Desligado - Nome': 0, 'Desligado - Matrícula': 0, 'Não localizado': 0,
        'Pendente de validação': 0, 'Dados insuficientes': 0
    }

    for arquivo in caminho_base.iterdir():
        if not arquivo.is_file() or arquivo.name.startswith('~$') or arquivo.suffix.lower() not in ['.xlsx', '.xls', '.csv'] or arquivo.stem in nomes_excluidos:
            continue

        try:
            if arquivo.suffix.lower() == '.csv':
                df = pd.read_csv(arquivo, low_memory=False, dtype=str)
            else:
                df = pd.read_excel(arquivo, dtype=str)
                
            col_nome = encontrar_coluna(df, ['Solicitante'])
            col_mat = encontrar_coluna(df, ['Criação de', 'Criacao de'])
            col_id = encontrar_coluna(df, ['Número', 'Numero', 'ID', 'Ticket'])
            
            if not col_nome and not col_mat:
                continue

            for _, row in df.iterrows():
                estatisticas['Total Analisado'] += 1
                
                nome_chamado = row[col_nome] if col_nome and pd.notna(row[col_nome]) else ""
                mat_chamado = row[col_mat] if col_mat and pd.notna(row[col_mat]) else ""
                id_chamado = row[col_id] if col_id and pd.notna(row[col_id]) else ""
                
                # O Motor em Ação
                dados_rh, status = buscar_funcionario(nome_chamado, mat_chamado, rh_ativos, rh_desligados)
                
                # Conta estatística genérica (agrupa subtipos de pendências na chave mãe)
                if "Pendente de validação" in status:
                    estatisticas['Pendente de validação'] += 1
                elif "Dados insuficientes" in status:
                    estatisticas['Dados insuficientes'] += 1
                else:
                    estatisticas[status] += 1
                
                # Monta a linha estruturada preservando rastreabilidade
                registro = {
                    'Mês/Arquivo_Origem': arquivo.name,
                    'ID_Chamado': id_chamado,
                    'Solicitante_Original': nome_chamado,
                    'Matricula_Original': mat_chamado,
                    'Situação_da_Busca': status,
                    
                    # Campos consolidados do RH
                    'Nome_Localizado_RH': dados_rh.get('nome', ''),
                    'Matricula_RH': dados_rh.get('matricula', ''),
                    
                    # Campos específicos de Ativos
                    'Ativo_Diretoria': dados_rh.get('diretoria', ''),
                    'Ativo_Gerencia_Geral': dados_rh.get('ger_geral', ''),
                    'Ativo_Gerencia': dados_rh.get('gerencia', ''),
                    'Ativo_Funcao': dados_rh.get('funcao', ''),
                    'Ativo_CC': dados_rh.get('cc', ''),
                    'Ativo_Email': dados_rh.get('email', ''),
                    'Ativo_Gestor': dados_rh.get('nome_gestor', ''),
                    
                    # Campos específicos de Desligados
                    'Desligado_Unidade': dados_rh.get('desc_unidade', dados_rh.get('unidade', '')),
                    'Desligado_Admissao': dados_rh.get('admissao', ''),
                    'Desligado_Data_Desligamento': dados_rh.get('desligamento', '')
                }
                registros_analisados.append(registro)
                
        except Exception as e:
            print(f"[Erro] Falha no arquivo {arquivo.name}: {e}")

    # ==========================================
    # 5. GERAÇÃO DOS ENTREGÁVEIS (ARQUIVOS)
    # ==========================================
    if not registros_analisados:
        print("Nenhum registro encontrado para cruzar.")
        return

    df_final = pd.DataFrame(registros_analisados)
    
    # Arquivo 1: Chamados Enriquecidos (Apenas cruzamentos bem sucedidos)
    df_sucesso = df_final[df_final['Situação_da_Busca'].str.contains("Ativo|Desligado")]
    df_sucesso.to_excel(pasta_resultados / '1_Chamados_Enriquecidos_RH.xlsx', index=False)
    
    # Arquivo 2: Pendências e Não Localizados
    df_pendentes = df_final[~df_final['Situação_da_Busca'].str.contains("Ativo|Desligado")]
    df_pendentes.to_excel(pasta_resultados / '2_Registros_Nao_Localizados_e_Pendentes.xlsx', index=False)
    
    # Arquivo 3: Auditoria (Resumo executivo)
    df_auditoria = pd.DataFrame(list(estatisticas.items()), columns=['Categoria', 'Quantidade'])
    df_auditoria.to_excel(pasta_resultados / '3_Auditoria_Cruzamento_RH.xlsx', index=False)

    print("\n--- RESUMO DO CRUZAMENTO ---")
    for k, v in estatisticas.items():
        print(f"{k}: {v}")
    print(f"\nOs 3 arquivos foram gerados com sucesso na pasta:\n{pasta_resultados}")

if __name__ == "__main__":
    executar_modulo_rh()