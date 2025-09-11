# Lógica principal de comparação

from difflib import SequenceMatcher
from openpyxl import load_workbook
import unicodedata
import re
import pandas as pd


# =============================================================AUXILIARRES=============================================================

def verifica_item(item_prop, item_ref):
    if item_prop != item_ref:
        return False # Caso não possuam valores iguais
    return True # Caso possuam valores iguais

def verifica_descricao(descricao_ref, props_dict):
    """
    Verifica se descricao_ref bate com alguma chave do props_dict.
    Se bater, remove o item do dicionário e retorna a linha.
    Se não, retorna None.
    """
    descricao_ref_clean = descricao_ref.lower().replace(" ", "")
    for descricao_prop in list(props_dict.keys()):  # usar lista para iterar sem problemas
        descricao_prop_clean = descricao_prop.lower().replace(" ", "")
        if descricao_ref_clean == descricao_prop_clean:
            valor = props_dict.pop(descricao_prop)  # remove do dicionário original
            return valor
    return None

def verifica_quantidade(qtd_ref, qtd_prop, limiar=0.95):
    return qtd_ref == qtd_prop  # Considera iguais se forem exatamente iguais

def verifica_valor_total(valor_total_prop, quantidade_prop, valor_unit_bdi_prop): # Função para verrificar o valor total linha a linha
    return valor_total_prop == (quantidade_prop * valor_unit_bdi_prop)

def normaliza_desconto(desconto):
    return (f"{desconto:.2f}%")

def verifica_desconto (valor_total_prop, valor_total_ref, limiar=25):
    desconto = (1-(valor_total_prop/valor_total_ref))*100
    return desconto,  ((desconto >= 0 and desconto <= limiar) if True else False)

def normaliza(texto: str) -> str: # Função para normalizar os textos
    if not isinstance(texto, str):
        texto = str(texto)
    # remove acentos
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    # coloca em maiúsculas
    texto = texto.upper()
    # substitui múltiplos espaços por um só
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()

# =============================================================COMPARADORES=============================================================

# def identificar_valor_total_planilhas(arquivo, descricao_padrao, sheet_name="Orçamento Sintético"): #Função que idetifica os valores totais da planinha baseado na descrição padrão
#     wb = load_workbook(arquivo, data_only=True)#Usa a planilha em formato xlsx
#     ws = wb[sheet_name] if sheet_name else wb.active  # pega aba ativa ou a escolhida
    
#     desc_norm = normaliza(descricao_padrao)
    
#     for row in range(ws.max_row, 0, -1):  # percorre da última linha até a primeira
#         valores = [cell.value for cell in ws[row]]
#         # junta todos os valores da linha em um só texto
#         texto_linha = " ".join(str(v) for v in valores if v is not None)
#         if desc_norm in normaliza(texto_linha):
#             linha_filtrada = [x for x in valores if x is not None]
#             return row, linha_filtrada  # retorna o indice e a linha filtrada (sem os nulos)

#     return None

def identificar_valor_total_planilhas_df(df, descricao_padrao):
    """
    Identifica a linha que contém a descrição padrão no DataFrame,
    percorrendo de baixo para cima.
    
    Parâmetros:
    - df: DataFrame carregado da planilha
    - descricao_padrao: string que identifica a linha desejada
    
    Retorna:
    - indice da linha no df
    - lista com os valores não nulos da linha
    """
    desc_norm = descricao_padrao.strip().upper()  # normaliza descrição

    # percorre de baixo para cima
    for idx in reversed(df.index):
        row = df.loc[idx]
        valores = row.tolist()
        # junta os valores não nulos em um texto
        texto_linha = " ".join(str(v) for v in valores if v is not None)
        if desc_norm in texto_linha.upper():
            # retorna o índice e apenas os valores não nulos
            linha_filtrada = [v for v in valores if v is not None]
            return idx, linha_filtrada

    return None

def resumo_descontos(df):
    df_filtrado = df[df["desconto_prop"] != 0]

    maiores_10 = df_filtrado.nlargest(10, "desconto_prop")
    menores_10 = df_filtrado.nsmallest(10, "desconto_prop")

    return {
        "maiores_10": maiores_10,
        "menores_10": menores_10
    }


def calcula_desconto_total_final(valor_ref, valor_prop):

    if valor_ref == 0:
        return 0.0

    desconto_valor = valor_ref - valor_prop
    desconto_percentual = (desconto_valor / valor_ref) * 100

    return desconto_percentual

def comparar_planilhas(df_ref, df_prop):
    relatorio = []

    # Criar um dicionário para acesso rápido às propostas pela descrição
    props_dict = {row['descricao']: row for _, row in df_prop.iterrows()}
    # descricoes_ref = set(df_ref['descricao'])


    for _, row_ref in df_ref.iterrows():
        item_ref = row_ref['item']
        descricao_ref = row_ref['descricao']
        quantidade_ref = row_ref['quantidade']
        valor_unit_ref = row_ref['valor_unit']
        valor_unit_bdi_ref = row_ref['valor_unit_bdi']
        valor_total_ref = row_ref['valor_total']
        
        # Preparando o resultado inicial
        resultado = {
            'item_ref': item_ref,
            'descricao_ref': descricao_ref,
            'quantidade_ref': quantidade_ref,
            'valor_unit_ref': valor_unit_ref,
            'valor_unit_bdi_ref': valor_unit_bdi_ref,
            'valor_total_ref': valor_total_ref,
        }
       
        row_prop = verifica_descricao(descricao_ref, props_dict)
        if row_prop is not None:
            # Acessando os valores da proposta
            item_prop = row_prop['item']
            descricao_prop = row_prop['descricao']
            quantidade_prop = row_prop['quantidade']
            valor_unit_prop = row_prop['valor_unit']
            valor_unit_bdi_prop = row_prop['valor_unit_bdi']
            valor_total_prop = row_prop['valor_total']
           

            
            # Montando o relatório com as propostas
            resultado['presente'] = True
            resultado['item_prop'] = item_prop
            resultado['descricao_prop'] = descricao_prop
            resultado['quantidade_prop'] = quantidade_prop
            resultado['valor_unit_prop'] = valor_unit_prop
            resultado['valor_unit_bdi_prop'] = valor_unit_bdi_prop
            resultado['valor_total_prop'] = valor_total_prop
            
            # Realizando os comparativos
            resultado['item_ok'] = bool(verifica_item(item_prop, item_ref))
            resultado['quantidade_ok'] = verifica_quantidade(quantidade_ref, quantidade_prop)
            resultado['valor_total_ok'] = verifica_valor_total(valor_total_prop, quantidade_prop, valor_unit_bdi_prop)
            resultado['desconto_prop'], resultado['desconto_ok'] = verifica_desconto (valor_total_prop, valor_total_ref, limiar=25)


        else:
            # Se o item não está presente na proposta, preenche com valores padrão
            resultado['presente'] = False
            resultado['item_prop'] = '-'
            resultado['descricao_prop'] = '-'
            resultado['quantidade_prop'] = 0
            resultado['valor_unit_prop'] = 0.0
            resultado['valor_unit_bdi_prop'] = 0.0
            resultado['valor_total_prop'] = 0.0
            resultado['desconto_prop'] = 0.0
            
            # Definindo os campos de comparação como False
            resultado['item_ok'] = False
            resultado['quantidade_ok'] = False
            resultado['valor_total_ok'] = False
            resultado['desconto_ok'] = True  # True para não contar como desconto com problema
            
        
        relatorio.append(resultado)
       
    df_relatorio = pd.DataFrame(relatorio)
    dict_resumo_descontos = resumo_descontos(df_relatorio)
    df_relatorio['desconto_prop']=df_relatorio['desconto_prop'].apply(normaliza_desconto)
    
    # -------------------
    # Somatório Total das colunas de Valor Total #TODO Fazzer um dict
    # -------------------
    soma_valor_global_prop = df_prop.loc[~df_prop["item"].astype(str).str.contains("\."), "valor_total"].sum()
    soma_valor_global_ref = df_ref.loc[~df_ref["item"].astype(str).str.contains("\."), "valor_total"].sum()

    # -------------------
    # Itens extras (para analisar, pois estão com problema)
    # -------------------
    extras_prop = pd.DataFrame(props_dict.values())

    # -------------------
    # Itens faltando na proposta 
    # -------------------
    ausentes_prop = df_relatorio.loc[df_relatorio['presente'] == False]
    ausentes_prop = ausentes_prop[['item_ref', 'descricao_ref', 'quantidade_ref', 'valor_unit_ref', 'valor_unit_bdi_ref', 'valor_total_ref']]

    # -------------------
    # Itens com desconto fora do padrão
    # -------------------
    descontos_prop = df_relatorio.loc[df_relatorio['desconto_ok'] == False]

    # -------------------
    # Tratando as linhas nulas pela coluna valor total
    # -------------------
    
    df_relatorio = df_relatorio.dropna(subset=['valor_total_ref']) #Para remover os itens nulos
    extras_prop = extras_prop.dropna(subset=['valor_total']) #Para remover os itens nulos
    ausentes_prop = ausentes_prop.dropna(subset=['valor_total_ref']) #Para remover os itens nulos
    descontos_prop = descontos_prop.dropna(subset=['valor_total_ref']) #Para remover os itens nulos





    return df_relatorio, extras_prop, ausentes_prop, descontos_prop, soma_valor_global_prop, soma_valor_global_ref, dict_resumo_descontos


















# =============================================================TESTES=============================================================


# def checar_quantidades_planilhas(df_ref, df_prop):
    # # Verifica se as quantidades de cada itens estão corretas
    
    # total_ref = len(df_ref.dropna(subset=['valor_total']))
    # total_prop = len(df_prop.dropna(subset=['valor_total']))

    # total_etapas_gerais_ref = [valor for valor in df_ref['item'] if str(valor).find('.') == -1]
    # total_etapas_gerais_prop = [valor for valor in df_prop['item'] if str(valor).find('.') == -1]

    # subetapas_ref = []
    # subetapas_prop = []
    # #Listar todas as subetapas
    # for etapa in total_etapas_gerais_ref:
    #     subetapas_ref = [valor for valor in df_ref['item'] if str(valor).startswith(f"{etapa}.")]
        
    # for etapa in total_etapas_gerais_prop:
    #     subetapas_prop = [valor for valor in df_prop['item'] if str(valor).startswith(f"{etapa}.")]
    
    # for subetapa in subetapas_ref:
        # subetapas_servicos_ref = [valor for valor in df_ref['item'] if str(valor).startswith(f"{subetapa}.")]
        
import re
from collections import defaultdict

# regex para identificar formatos
re_item      = re.compile(r'^\d+$')
re_subitem   = re.compile(r'^(\d+)\.(\d+)$')
re_servico   = re.compile(r'^(\d+)\.(\d+)\.(\d+)$')

def _chave_sort(s: str):
    """ordena numericamente strings tipo '10.2.3' → (10,2,3)"""
    return tuple(int(p) for p in s.split("."))

def analisar_coluna_item(df, coluna="item"):
    """
    Faz o raio-x da coluna 'item' de um DataFrame.
    
    Retorna dict com:
      - resumo hierárquico (itens → subitens → qtd de serviços)
      - contagem geral por tipo
      - lista de entradas inválidas
    """
    itens = defaultdict(lambda: {"subitens": defaultdict(lambda: {"servicos": set()})})
    contagem = {"geral": 0, "subitem": 0, "servico": 0, "invalido": 0}
    invalidos = []
    vistos = set()

    for raw in df[coluna]:
        s = str(raw).strip()
        if not s or s in vistos:
            continue
        vistos.add(s)

        m3 = re_servico.match(s)
        m2 = re_subitem.match(s)
        m1 = re_item.match(s)

        if m3:
            i, j, k = m3.groups()
            chave_item = i
            chave_sub  = f"{i}.{j}"
            itens[chave_item]["subitens"][chave_sub]["servicos"].add(s)
            contagem["servico"] += 1
        elif m2:
            i, j = m2.groups()
            chave_item = i
            chave_sub  = f"{i}.{j}"
            _ = itens[chave_item]["subitens"][chave_sub]
            contagem["subitem"] += 1
        elif m1:
            i = m1.group()
            _ = itens[i]  # garante existência
            contagem["geral"] += 1
        else:
            contagem["invalido"] += 1
            invalidos.append(s)

    # monta resumo hierárquico
    resumo = {}
    for item, dados in itens.items():
        subitens = dados["subitens"]
        resumo[item] = {
            "qtd_subitens": len(subitens),
            "subitens": {
                sub: {
                    "qtd_servicos": len(info["servicos"])
                }
                for sub, info in sorted(subitens.items(), key=lambda kv: _chave_sort(kv[0]))
            }
        }

    resumo = dict(sorted(resumo.items(), key=lambda kv: int(kv[0])))

    return {"resumo": resumo, "contagem": contagem, "invalidos": invalidos}

def relatorio_hierarquico(resultado):
    """
    Gera texto legível do raio-x
    """
    linhas = []
    for item, bloco in resultado["resumo"].items():
        linhas.append(f"Item {item} tem {bloco['qtd_subitens']} subitens.")
        for sub, info in bloco["subitens"].items():
            linhas.append(f"  Subitem {sub} tem {info['qtd_servicos']} serviços.")
    if resultado["invalidos"]:
        linhas.append(f"\nInválidos ({len(resultado['invalidos'])}): {', '.join(resultado['invalidos'])}")
    return "\n".join(linhas)

