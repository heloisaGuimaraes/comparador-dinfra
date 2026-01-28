# Função para ler e preparar dados da planilha
import re
import pandas as pd


# from comparador import identificar_valor_total_planilhas_df #identificar_valor_total_planilhas
DESCRICAO_PADRAO_TOTAL_SEM_BDI = "Total sem BDI"  # ajuste para o que aparece na sua planilha
DESCRICAO_PADRAO_TOTAL_BDI = "Total do BDI"  # ajuste para o que aparece na sua planilha
DESCRICAO_PADRAO_TOTAL_GERAL = "Total Geral"  # ajuste para o que aparece na sua planilha
COLUNAS_PLANILHA = ["Item","Código","Banco","Descrição","Und","Quant.","Valor Unit","Valor Unit com BDI","Total","Peso (%)"]


def texto_para_numero(valor): # Função para converter texto em número float, tirando os textos de BDI da coluna Valor unitário com BDI
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor)

    # procura primeiro padrão tipo "1.234.567,89" ou "1234,56" ou "1234"
    m = re.search(r"\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?", s)
    if not m:
        return 0.0

    token = m.group(0)          # pega o primeiro número reconhecido
    token = token.replace(".", "")   # remove separadores de milhar
    token = token.replace(",", ".")  # transforma vírgula decimal em ponto

    try:
        return float(token)
    except ValueError:
        return 0.0

def texto_para_float(valor):
    """
    Limpa e converte uma string de valor monetário complexa para um float padrão.
    Lida com formatos brasileiros (milhar.decimal,) e texto extra (BDI, %).
    """
    
    # 1. Trata valores nulos ou não-string
    if valor is None or valor == False:
        return 0.0
    
    if isinstance(valor, (float, int)):
        return float(valor)

    try:
        texto = str(valor).strip()
        
        # 2. Remoção de Texto Extra (incluindo parênteses e BDI)
        # Exemplo: "141.279,97 (BDI 15,21%)" -> "141.279,97"
        
        # Remove qualquer coisa entre parênteses
        texto = re.sub(r'\s*\(.*\)\s*', '', texto).strip()
        
        # Remove caracteres/símbolos comuns
        texto = texto.upper().replace('R$', '').replace('%', '').replace('BDI', '').strip()
        
        # 3. Tratamento de Separadores
        
        # Verifica se o formato é brasileiro (PONTO de milhar, VÍRGULA decimal)
        # Ex: 141.279,97
        if re.search(r'\d\.\d{3},\d{2}$', texto) or re.search(r'\d,\d{2}$', texto):
            # Substitui ponto de milhar por nada
            texto = texto.replace('.', '')
            # Substitui a vírgula decimal por ponto
            texto = texto.replace(',', '.')
        
        # Verifica se é formato brasileiro simples (sem milhar, vírgula decimal)
        # Ex: 1.500,00
        elif texto.count(',') == 1 and texto.count('.') == 0:
             texto = texto.replace(',', '.')
             
        # Se for o formato americano (milhar, ponto.decimal), remove as vírgulas
        # Ex: 1,526.20
        elif texto.count('.') == 1 and texto.count(',') >= 1:
            texto = texto.replace(',', '')
        
        # 4. Limpeza final e Conversão
        texto = texto.replace(' ', '')
        
        return float(texto)
        
    except ValueError:
        # Retorna 0.0 se a string final não puder ser convertida (ex: "NA")
        return 0.0

def encontrar_linhas_bdi_diferente(df, coluna_texto): #Função para encontrar as linhas que contém BDI diferente em valor unitário com BDI
    linhas_bdi = []
    
    for idx, row in df.iterrows():
        texto = str(row[coluna_texto]) if row[coluna_texto] is not None else ""
        if "BDI" in texto.upper():   # procura "BDI" ignorando maiúsc/minúsc
            # usa v2 para extrair o valor principal
            linhas_bdi.append(row.to_dict().copy())  # salva uma cópia da linha inteira
            valor = texto_para_numero(texto)
            df.at[idx, coluna_texto] = valor

    #print(f"Linhas com BDI diferente encontradas: {len(linhas_bdi)}")
    return pd.DataFrame(linhas_bdi)

def normaliza_planilha(df, colunas_esperadas=COLUNAS_PLANILHA):
    header_idx = None
    
    # procura a linha do cabeçalho
    for i, row in df.iterrows():
        valores_linha = [str(v).strip() for v in row.values if v is not None]
        # print(valores_linha)
        if all(col in valores_linha for col in colunas_esperadas):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Cabeçalho esperado não encontrado na planilha")

    # pega apenas as linhas a partir do cabeçalho
    df = df.iloc[header_idx:].copy()

    # define a primeira linha como cabeçalho
    df.columns = df.iloc[0]
    df = df[1:]  # remove a linha do cabeçalho do conteúdo

    # remove linhas/colunas completamente vazias
    df = df.dropna(how="all", axis=1)
    df = df.dropna(how="all")

    # normaliza nomes das colunas
    df.columns = [str(c).strip().lower() for c in df.columns]

    # opcional: resetar índice
    df = df.reset_index(drop=True)

    return df
     

def identificar_valor_total_planilhas_df(df, descricao_padrao, nome_planilha=""):
    """
    Identifica a linha que contém a descrição padrão no DataFrame,
    removendo espaços extras e normalizando o texto.
    """
    # 1. Normaliza a descrição: remove espaços e reduz múltiplos espaços internos a um só
    desc_norm = " ".join(descricao_padrao.strip().upper().split())

    # Percorre de baixo para cima
    for idx in reversed(df.index):
        row = df.loc[idx]
        valores = row.tolist()
        
        # 2. Limpa cada valor: remove espaços de cada célula e ignora NaNs
        valores_limpos = [str(v).strip() for v in valores if pd.notna(v)]
        
        # 3. Junta tudo em um único texto, também normalizando espaços internos
        texto_linha = " ".join(valores_limpos).upper()
        texto_linha_normalizado = " ".join(texto_linha.split())
        
        # 4. Comparação
        if desc_norm in texto_linha_normalizado:
            # Retorna o índice e os valores originais da linha (sem os NaNs)
            linha_filtrada = [v for v in valores if pd.notna(v)]
            return idx, linha_filtrada
    #Lançar um erro se não encontrar
    raise ValueError(f"O valor do {descricao_padrao} não foi encontrado na planilha {nome_planilha}. Verifique se a planilha está completa.")


# --------------------------------------------------------------------
# Função principal para carregar e preparar a planilha
# --------------------------------------------------------------------
def carregar_planilha(caminho):
    
    abas = pd.ExcelFile(caminho).sheet_names

    nome_planilha = caminho.name
    # print(f"Caminho: {nome_planilha}")
    # print(f"Abas encontradas na planilha: {abas}")


    # Definimos uma lista com as variações comuns dos nomes das abas (todas em minúsculo)
    termos_procurados = ["orçamento", "orcamento", "ocamento", "sintético"]

    # Procuramos a aba: se o termo estiver dentro do nome da aba (convertida para minúsculo)
    aba_analisada = next(
        (s for s in abas if any(termo in s.lower() for termo in termos_procurados)), 
        None
    )
    
    # Validação
    if aba_analisada is None:
        raise ValueError(f"Não foi possível encontrar uma aba de Orçamento Sintético. Abas disponíveis: {abas}. Verifique se a planilha está correta.")  
    
    df = pd.read_excel(caminho, sheet_name=aba_analisada, header=None, usecols="A:J")
    df = normaliza_planilha(df, COLUNAS_PLANILHA)
    
    #Acessando os valores totais do fim da planilha
    idx_total_geral, valores_total_geral = identificar_valor_total_planilhas_df(df, DESCRICAO_PADRAO_TOTAL_GERAL, nome_planilha)
    # print(valores_total_geral)

    idx_total_bdi, valores_total_bdi = identificar_valor_total_planilhas_df(df, DESCRICAO_PADRAO_TOTAL_BDI)
    # print(valores_total_bdi)

    idx_total_sem_bdi, valores_total_sem_bdi = identificar_valor_total_planilhas_df(df, DESCRICAO_PADRAO_TOTAL_SEM_BDI, nome_planilha)
    # print(valores_total_sem_bdi)

    # print(valores_total_geral, valores_total_sem_bdi, valores_total_bdi)


    linhas_com_totais = [valores_total_sem_bdi, valores_total_bdi, valores_total_geral]
    
    dict_totais = {item[0]: item[1] for item in linhas_com_totais}
    # print("Totais encontrados na planilha:", aba_analisada, dict_totais)
    
    linhas_remover = [idx_total_geral, idx_total_bdi, idx_total_sem_bdi]
    linhas_df = [i for i in linhas_remover]

    df = df.drop(linhas_df) #Retirando as linhas com os valores totais

    

    df = df.drop(['código', 'banco', 'und', 'peso (%)'], axis=1) # Removendo as colunas com os textos de descrição

    df = df.rename(columns={
        'Item': 'item',
        'descrição': 'descricao',
        'quant.' : 'quantidade',
        'valor unit': 'valor_unit',
        'valor unit com bdi': 'valor_unit_bdi',
        'total': 'valor_total'
    })
    


    df_valores_bdi_diferente = encontrar_linhas_bdi_diferente(df, 'valor_unit_bdi')

    df['item'] = df['item'].astype(str).str.strip()
    df['descricao'] = df['descricao'].astype(str).str.strip() #ok
    df['quantidade'] = df['quantidade'].astype(float) #ok
    df['valor_unit'] = df['valor_unit'].astype(float) 
    df['valor_unit_bdi'] = df['valor_unit_bdi'].apply(texto_para_float)
    df['valor_total'] = df['valor_total'].apply(texto_para_float)
    
    df = df.dropna(subset=['valor_total']) #Para remover os itens nulos
    

    return df, dict_totais, df_valores_bdi_diferente
