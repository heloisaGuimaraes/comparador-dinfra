# Função para ler e preparar dados da planilha
import re
import math
import pandas as pd
from comparador import identificar_valor_total_planilhas_df #identificar_valor_total_planilhas
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
    
    

def carregar_planilha(caminho):
    
    abas = pd.ExcelFile(caminho).sheet_names
    aba_analisada = "Orçamento Sintético"

    if aba_analisada not in abas:
        raise ValueError(f"A aba '{aba_analisada}' não foi encontrada no arquivo. Abas disponíveis: {abas}")    
    
    df = pd.read_excel(caminho, sheet_name=aba_analisada, header=None)
    df = normaliza_planilha(df, COLUNAS_PLANILHA)
    
    
    idx_total_sem_bdi, valores_total_sem_bdi = identificar_valor_total_planilhas_df(df, DESCRICAO_PADRAO_TOTAL_SEM_BDI)
    idx_total_bdi, valores_total_bdi = identificar_valor_total_planilhas_df(df, DESCRICAO_PADRAO_TOTAL_BDI)
    idx_total_geral, valores_total_geral = identificar_valor_total_planilhas_df(df, DESCRICAO_PADRAO_TOTAL_GERAL)
    # print(valores_total_geral, valores_total_sem_bdi, valores_total_bdi)


    linhas_com_totais = [valores_total_sem_bdi, valores_total_bdi, valores_total_geral]
    # dict_totais = {item[0]: item[1] for item in linhas_com_totais}
    dict_totais = {
    linha_filtrada[0]: linha_filtrada[1]
    for linha in linhas_com_totais
    if (linha_filtrada := [v for v in linha if v is not None and (not isinstance(v, float) or not math.isnan(v))]) and len(linha_filtrada) >= 2
    }
    # print("Totais encontrados na planilha:", dict_totais)
    
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
    # df['valor_unit_bdi'] = df['valor_unit_bdi'].apply(texto_para_numero)
    df['valor_total'] = df['valor_total'].astype(float) #ok
    
    df = df.dropna(subset=['valor_total']) #Para remover os itens nulos
    

    return df, dict_totais, df_valores_bdi_diferente
