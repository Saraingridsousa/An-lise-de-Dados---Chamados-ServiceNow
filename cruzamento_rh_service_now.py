import pandas as pd
from pathlib import Path
import re
import warnings

# Suprime avisos de formato do Excel
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# 1. FUNÇÕES DE NORMALIZAÇÃO
# ==========================================
def normalizar_texto(texto):
    """Remove acentos, espaços extras e converte para minúsculas."""
    if pd.isna(texto) or str(texto).strip() == "":
        return ""
    texto = str(texto).strip().lower()
    mapa_acentos = str.maketrans("áàãâéêíóôõúüç", "aaaaeeiooouuc")
    texto = texto.translate(mapa_acentos)
    return re.sub(r'\s+', ' ', texto)

def normalizar_matricula(mat):
    """Remove decimais artificiais e padroniza a matrícula."""
    if pd.isna(mat) or str(mat).strip() == "":
        return ""
    mat_str = str(mat).strip().upper()
    if mat_str.endswith('.0'):
        mat_str = mat_str[:-2]
    return mat_str

def normalizar_email(email):
    """Padroniza e-mails corporativos, adicionando o domínio de forma segura."""
    if pd.isna(email) or str(email).strip() == "":
        return ""
    email_str = str(email).strip().lower()
    if '@' not in email_str and ' ' not in email_str:
        email_str = f"{email_str}@mrs.com.br"
    return email_str

def encontrar_coluna(df, possiveis_nomes):
    """Localiza colunas de forma dinâmica, tolerando diferenças de espaçamento e case."""
    colunas_df = {normalizar_texto(c): c for c in df.columns}
    for nome in possiveis_nomes:
        nome_norm = normalizar_texto(nome)
        if nome_norm in colunas_df:
            return colunas_df[nome_norm]
    return None

def extrair_competencia(nome_arquivo):
    """Deduz a competência a partir da nomenclatura do arquivo."""
    match = re.search(r'(20\d{2})\.(\d{2})', nome_arquivo)
    if match:
        return pd.Period(f"{match.group(1)}-{match.group(2)}", freq='M')
    return None

# ==========================================
# 2. CARREGAMENTO DOS DICIONÁRIOS BASE
# ==========================================
def carregar_ativos_rh(caminho_rh):
    """Mapeia os funcionários ativos para extrair o e-mail do gestor."""
    dict_nome, dict_mat = {}, {}
    if not Path(caminho_rh).exists():
        print(f"[Erro] Base de Ativos não encontrada: {caminho_rh}")
        return dict_nome, dict_mat

    df = pd.read_excel(caminho_rh, dtype=str)
    col_nome = encontrar_coluna(df, ['Nome'])
    col_mat = encontrar_coluna(df, ['Matricula', 'Matrícula'])
    col_email_gestor = encontrar_coluna(df, ['Email Gestor'])
    col_nome_gestor = encontrar_coluna(df, ['Nome Gestor'])

    if not col_nome or not col_mat or not col_email_gestor:
        print("[Erro] Colunas obrigatórias ausentes na base de Ativos RH.")
        return dict_nome, dict_mat

    for _, row in df.iterrows():
        nome_norm = normalizar_texto(row[col_nome])
        mat_norm = normalizar_matricula(row[col_mat])
        
        email_gestor = normalizar_email(row[col_email_gestor])
        nome_gestor = str(row[col_nome_gestor]).strip() if pd.notna(row[col_nome_gestor]) else ""
        
        dados = {'email_gestor': email_gestor, 'nome_gestor': nome_gestor, 'nome_rh': str(row[col_nome]).strip(), 'mat_rh': str(row[col_mat]).strip()}
        
        if nome_norm:
            dict_nome[nome_norm] = "AMBIGUO" if nome_norm in dict_nome else dados
        if mat_norm:
            dict_mat[mat_norm] = dados

    return dict_nome, dict_mat

def carregar_canal_coordenacao(caminho_canal):
    """Cria um dicionário de mapeamento Email do Gestor -> Gerência, buscando em todas as abas."""
    dict_coord = {}
    if not Path(caminho_canal).exists():
        print(f"[Erro] Base de Canal por Coordenação não encontrada: {caminho_canal}")
        return dict_coord

    xls = pd.ExcelFile(caminho_canal)
    df = None
    
    # -------------------------------------------------------------
    # NOVO: Varredura de abas para encontrar onde estão os dados reais
    # -------------------------------------------------------------
    for aba in xls.sheet_names:
        df_temp = pd.read_excel(xls, sheet_name=aba, dtype=str)
        colunas_norm = [normalizar_texto(c) for c in df_temp.columns]
        
        # Se a aba tiver a palavra coordenador ou gerente, é a aba correta
        if normalizar_texto("Coordenador") in colunas_norm or normalizar_texto("Gerente") in colunas_norm:
            df = df_temp
            print(f"[Sucesso] Aba correta encontrada no Excel: '{aba}'")
            break
            
    if df is None:
        # Fallback: Se não achou pelos nomes, pega a aba que tiver mais colunas
        for aba in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=aba, dtype=str)
            if len(df_temp.columns) >= 7:
                df = df_temp
                print(f"[Aviso] Usando a aba '{aba}' por ter as colunas F e G preenchidas.")
                break
                
    if df is None:
        df = pd.read_excel(caminho_canal, dtype=str) # Volta pro padrão se tudo falhar

    col_gerencia = encontrar_coluna(df, ['Gerência', 'Gerencia'])
    col_coord = encontrar_coluna(df, ['Coordenador']) 
    col_gerente = encontrar_coluna(df, ['Gerente'])

    # Captura posicional caso os nomes tenham formatação oculta
    if not col_coord and len(df.columns) >= 6:
        col_coord = df.columns[5] # Índice 5 = Coluna F
        print(f"[Aviso] Coluna F capturada à força para Coordenador. (Nome: '{col_coord}')")

    if not col_gerente and len(df.columns) >= 7:
        col_gerente = df.columns[6] # Índice 6 = Coluna G
        print(f"[Aviso] Coluna G capturada à força para Gerente. (Nome: '{col_gerente}')")

    if not col_gerencia or (not col_coord and not col_gerente):
        print(f"[Erro] Falha crítica ao ler colunas do Canal por Coordenação.")
        print(f"Colunas lidas pelo Pandas: {df.columns.tolist()}")
        return dict_coord

    def processar_email(email_bruto, gerencia_linha, tipo_coluna):
        email_norm = normalizar_email(email_bruto)
        if not email_norm:
            return
            
        gerencia_str = str(gerencia_linha).strip() if pd.notna(gerencia_linha) and str(gerencia_linha).strip() != "" else "MISSING"

        if email_norm in dict_coord:
            gerencia_existente = dict_coord[email_norm].get('gerencia')
            if gerencia_existente != gerencia_str and gerencia_existente != "AMBIGUO":
                dict_coord[email_norm] = {'gerencia': 'AMBIGUO', 'coluna': 'Múltiplas'}
        else:
            dict_coord[email_norm] = {'gerencia': gerencia_str, 'coluna': tipo_coluna}

    for _, row in df.iterrows():
        gerencia_val = row[col_gerencia] if col_gerencia else ""
        if col_coord:
            processar_email(row[col_coord], gerencia_val, 'Coordenador')
        if col_gerente:
            processar_email(row[col_gerente], gerencia_val, 'Gerente')

    return dict_coord

# ==========================================
# 3. MOTOR DE CRUZAMENTO
# ==========================================
def executar_correlacao_gerencial():
    dir_base = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - x_mrsls_csc_sla")
    caminho_ativos = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - sys_user\Ativos RH 3.xlsx")
    caminho_coord = Path(r"C:\Users\99006526\MRS Logística SA\Otimização e Gestão Processos - Chamados Service Now\Canal por Coordenação.xlsx")
    
    pasta_resultados = dir_base / "resultados_correlacao_gerencias"
    pasta_resultados.mkdir(exist_ok=True)

    print("Carregando bases de governança (RH e Coordenação)...")
    dict_ati_nome, dict_ati_mat = carregar_ativos_rh(caminho_ativos)
    dict_coord = carregar_canal_coordenacao(caminho_coord)
    
    if not dict_coord:
        print("[Abortado] A carga da planilha de coordenação falhou. Verifique os cabeçalhos e abas.")
        return

    forms_inspecao = ["inspecoes csc regional", "inspecoes csc regional 1.0", "inspecao csc"]
    nomes_excluidos = {'ativo_x_mrsls_csc_sla', '2025_x_mrsls_csc_sla'}
    
    auditoria = {
        '1. Total Analisados': 0,
        '2. Excluídos (Inspeções CSC)': 0,
        '3. Sem correspondência na base de ativos': 0,
        '4. Sem e-mail de gestor': 0,
        '5. Gestor não localizado na coordenação': 0,
        '6. Gerência não informada': 0,
        '7. Correspondência ambígua ou conflitante': 0,
        '8. Classificados por gerência (Sucesso)': 0
    }
    
    registros = []

    print("Processando planilhas mensais...")
    for arquivo in dir_base.iterdir():
        if not arquivo.is_file() or arquivo.name.startswith('~$') or arquivo.suffix.lower() not in ['.xlsx', '.xls', '.csv'] or arquivo.stem in nomes_excluidos:
            continue

        try:
            mes_arquivo = extrair_competencia(arquivo.name)
            df = pd.read_csv(arquivo, low_memory=False, dtype=str) if arquivo.suffix.lower() == '.csv' else pd.read_excel(arquivo, dtype=str)
            
            col_form = encontrar_coluna(df, ['Formulário', 'Formulario', 'Item'])
            col_nome = encontrar_coluna(df, ['Solicitante'])
            col_mat = encontrar_coluna(df, ['Criação de', 'Criacao de'])
            col_id = encontrar_coluna(df, ['Número', 'Numero', 'ID', 'Ticket'])
            col_data = encontrar_coluna(df, ['Aberto(a)', 'Aberto', 'Data de Abertura', 'Created', 'Data'])
            
            if not col_nome and not col_mat:
                continue

            if col_data:
                df['Data_Real'] = pd.to_datetime(df[col_data], errors='coerce', dayfirst=False)
                df['Mes_Apuracao'] = df['Data_Real'].dt.to_period('M')
                df['Mes_Apuracao'] = df['Mes_Apuracao'].fillna(mes_arquivo)
            else:
                df['Mes_Apuracao'] = mes_arquivo

            if col_id:
                df = df.drop_duplicates(subset=[col_id])
                
            for _, row in df.iterrows():
                auditoria['1. Total Analisados'] += 1
                
                form_str = str(row[col_form]).strip() if col_form and pd.notna(row[col_form]) else ""
                if normalizar_texto(form_str) in forms_inspecao:
                    auditoria['2. Excluídos (Inspeções CSC)'] += 1
                    continue

                nome_orig = str(row[col_nome]).strip() if col_nome and pd.notna(row[col_nome]) else ""
                mat_orig = str(row[col_mat]).strip() if col_mat and pd.notna(row[col_mat]) else ""
                id_orig = str(row[col_id]).strip() if col_id and pd.notna(row[col_id]) else ""
                
                nome_norm = normalizar_texto(nome_orig)
                mat_norm = normalizar_matricula(mat_orig)

                dados_rh = None
                if nome_norm and nome_norm in dict_ati_nome and dict_ati_nome[nome_norm] != "AMBIGUO":
                    dados_rh = dict_ati_nome[nome_norm]
                elif mat_norm and mat_norm in dict_ati_mat:
                    dados_rh = dict_ati_mat[mat_norm]

                status = ""
                gerencia_final = ""
                coluna_coord = ""
                
                if not dados_rh:
                    status = "Sem correspondência na base de ativos"
                    auditoria['3. Sem correspondência na base de ativos'] += 1
                else:
                    email_gestor = dados_rh.get('email_gestor', '')
                    if not email_gestor:
                        status = "Sem e-mail de gestor"
                        auditoria['4. Sem e-mail de gestor'] += 1
                    else:
                        dados_coord = dict_coord.get(email_gestor)
                        if not dados_coord:
                            status = "Gestor não localizado na coordenação"
                            auditoria['5. Gestor não localizado na coordenação'] += 1
                        else:
                            coluna_coord = dados_coord['coluna']
                            gerencia_encontrada = dados_coord['gerencia']
                            
                            if gerencia_encontrada == "AMBIGUO":
                                status = "Correspondência ambígua ou conflitante"
                                auditoria['7. Correspondência ambígua ou conflitante'] += 1
                            elif gerencia_encontrada == "MISSING":
                                status = "Gerência não informada"
                                auditoria['6. Gerência não informada'] += 1
                            else:
                                status = "Classificado por gerência"
                                gerencia_final = gerencia_encontrada
                                auditoria['8. Classificados por gerência (Sucesso)'] += 1

                registros.append({
                    'Competência_Mensal': row.get('Mes_Apuracao'),
                    'Arquivo_Origem': arquivo.name,
                    'ID_Chamado': id_orig,
                    'Formulário': form_str,
                    'Solicitante_Original': nome_orig,
                    'Matricula_Original': mat_orig,
                    'Nome_Ativo_RH': dados_rh.get('nome_rh', '') if dados_rh else "",
                    'Matricula_Ativo_RH': dados_rh.get('mat_rh', '') if dados_rh else "",
                    'Email_Gestor': dados_rh.get('email_gestor', '') if dados_rh else "",
                    'Nome_Gestor': dados_rh.get('nome_gestor', '') if dados_rh else "",
                    'Coluna_Coordenação': coluna_coord,
                    'Gerência_Identificada': gerencia_final,
                    'Situação_Correlação': status
                })
                
        except Exception as e:
            print(f"[Erro] Falha no processamento de {arquivo.name}: {e}")

    # ==========================================
    # 4. EXPORTAÇÕES E AUDITORIA
    # ==========================================
    if not registros:
        print("Nenhum dado processado após filtros.")
        return

    df_final = pd.DataFrame(registros)
    df_final['Competência_Mensal'] = df_final['Competência_Mensal'].astype(str)
    
    df_final.to_excel(pasta_resultados / '1_Chamados_Classificados_por_Gerencia.xlsx', index=False)

    df_sucesso = df_final[df_final['Gerência_Identificada'] != ""]
    if not df_sucesso.empty:
        df_resumo = df_sucesso.groupby(['Competência_Mensal', 'Gerência_Identificada']).size().reset_index(name='Total_Chamados')
        df_resumo = df_resumo.sort_values(by=['Competência_Mensal', 'Total_Chamados'], ascending=[True, False])
        df_resumo.to_excel(pasta_resultados / '2_Resumo_Mensal_por_Gerencia.xlsx', index=False)
    else:
        print("[Aviso] Nenhum chamado classificado com sucesso para gerar o resumo mensal.")

    df_auditoria = pd.DataFrame(list(auditoria.items()), columns=['Categoria', 'Quantidade'])
    df_auditoria.to_excel(pasta_resultados / '3_Auditoria_Correlacao.xlsx', index=False)

    print("\n--- RESUMO DE EXECUÇÃO ---")
    for k, v in auditoria.items():
        print(f"{k}: {v}")
    print(f"\nArquivos exportados para: {pasta_resultados}")

if __name__ == "__main__":
    executar_correlacao_gerencial()