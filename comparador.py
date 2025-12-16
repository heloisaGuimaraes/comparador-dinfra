# Lógica principal de comparação

from difflib import SequenceMatcher
from openpyxl import load_workbook
import unicodedata
import re
import pandas as pd
from rapidfuzz import fuzz, process


# =============================================================AUXILIARRES=============================================================

def verifica_item(item_prop, item_ref):
    if item_prop != item_ref:
        return False # Caso não possuam valores iguais
    return True # Caso possuam valores iguais

def limpar_descricao(descricao):
    """Normaliza a descrição para comparação (minúsculas, sem espaços)."""
    return str(descricao).lower().replace(" ", "").strip()


def verifica_descricao_fuzzy(descricao_ref, props_list, limite_similaridade=90):
    """
    Busca correspondência para descricao_ref em props_list:
    1. Tenta correspondência exata (100%).
    2. Se não, tenta a melhor correspondência fuzzy.

    Retorna:
    - Se Match Exato: (match_row, True, None)
    - Se Match Fuzzy: (None, False, {dados da divergência})
    - Se Sem Match: (None, False, None)
    """
    descricao_ref_clean = limpar_descricao(descricao_ref)
    
    # ----------------------------------------------------
    # 1. Tenta Comparação Exata
    # ----------------------------------------------------
    
    # Lista de descrições limpas da proposta para busca exata
    descricoes_prop_limpas = [limpar_descricao(desc) for desc, row in props_list]
    
    try:
        # Tenta achar o índice do match exato
        idx_exato = descricoes_prop_limpas.index(descricao_ref_clean)
        
        # Match Exato (100%) encontrado: remove e retorna a linha
        match_row = props_list.pop(idx_exato)[1]
        return match_row, True, None # (linha, is_exact, divergencia_data)
        
    except ValueError:
        # Não houve match exato, continua para a busca fuzzy
        pass 

    # ----------------------------------------------------
    # 2. Busca o Melhor Match Fuzzy
    # ----------------------------------------------------
    
    # Lista das descrições originais da proposta para a busca fuzzy
    descricoes_prop_originais = [desc for desc, row in props_list]
    
    if not descricoes_prop_originais:
        return None, False, None # Lista da proposta vazia
        
    # Usa process.extractOne para encontrar o melhor match
    best_match = process.extractOne(
        descricao_ref, # Usa a descrição original, que é melhor para o token_set_ratio
        descricoes_prop_originais, 
        scorer=fuzz.token_set_ratio 
    )

    melhor_desc_prop, score, idx_na_props_list = best_match

    if score >= limite_similaridade:
        # Match Fuzzy encontrado (acima do limite)
        melhor_match_row = props_list[idx_na_props_list][1]
        
        divergencia_data = {
            'Melhor_Match_Prop': melhor_desc_prop,
            'Melhor_Match_Row': melhor_match_row, # <--- Linha completa da Proposta
            'Similaridade': f"{score:.2f}%",
            'Validado_OK': False,
            'idx_na_props_list': idx_na_props_list # Guarda o índice para futura remoção
        }
        return None, False, divergencia_data # Retorna a divergência para validação
    
    # 3. Nenhum match (nem exato nem fuzzy aceitável)
    return None, False, None

#Função para verificar a descrição antiga (removida para acrescentar o RapidFuzz)
def verifica_descricao_simples(descricao_ref, props_list):
    """
    Verifica se descricao_ref bate com alguma descricao em props_list.
    Se bater, remove a primeira ocorrência da lista e retorna a linha.
    Se não, retorna None.
    """
    descricao_ref_clean = descricao_ref.lower().replace(" ", "")
    
    for i, (descricao_prop, row) in enumerate(props_list):
        descricao_prop_clean = descricao_prop.lower().replace(" ", "")
        if descricao_ref_clean == descricao_prop_clean:
            return props_list.pop(i)[1]  # remove e retorna a linha (row)
    
    return None






def verifica_quantidade(qtd_ref, qtd_prop, limiar=0.95):
    return qtd_ref == qtd_prop  # Considera iguais se forem exatamente iguais

def verifica_valor_total(valor_total_prop, quantidade_prop, valor_unit_bdi_prop): # Função para verrificar o valor total linha a linha
    return valor_total_prop == (quantidade_prop * valor_unit_bdi_prop)

def normaliza_desconto(desconto):
    return (f"{desconto:.2f}%")

def verifica_desconto (valor_total_prop, valor_total_ref, limiar=25):
    try:
        desconto = (1-(valor_total_prop/valor_total_ref))*100
        return desconto, ((desconto >= 0 and desconto <= limiar) if True else False)
    except ZeroDivisionError: #TODO lançar mensagem de aviso/erro
        # raise ValueError("Valor total de referência é zero, não é possível calcular desconto. Valores zerados para fins de comparação.")
        return 0.0, False #Desconto zero se algum dos valores for zero, valor não ok
   
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

    if valor_ref == 0.0 or valor_prop == 0.0:
        return 0.0

    desconto_valor = valor_ref - valor_prop
    desconto_percentual = (desconto_valor / valor_ref) * 100

    return desconto_percentual

def comparar_planilhas(df_ref, df_prop):
    relatorio = []

    # Criar uma listas das linhas para acesso rápido às propostas pela descrição
    props_list = [[row['descricao'], row] for _, row in df_prop.iterrows()]
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
       
        row_prop = verifica_descricao_fuzzy(descricao_ref, props_list)
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
    extras_prop = pd.DataFrame([row for _, row in props_list])

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
    # extras_prop = extras_prop.dropna(subset=['valor_total'])
    extras_prop = extras_prop.dropna(subset=['valor_total']) if not extras_prop.empty else extras_prop  #Para remover os itens nulos caso não esteja vazio
    ausentes_prop = ausentes_prop.dropna(subset=['valor_total_ref']) if not ausentes_prop.empty else ausentes_prop  #Para remover os itens nulos caso não esteja vazio
    descontos_prop = descontos_prop.dropna(subset=['valor_total_ref']) if not descontos_prop.empty else descontos_prop  #Para remover os itens nulos caso não esteja vazio


    return df_relatorio, extras_prop, ausentes_prop, descontos_prop, soma_valor_global_prop, soma_valor_global_ref, dict_resumo_descontos


#Adaptação para lidar com os states

def iniciar_comparacao_fuzzy(df_ref, df_prop):
    """
    Realiza a comparação inicial para identificar matches exatos,
    e lista as divergências fuzzy para validação manual.
    
    Retorna:
    - divergencias_para_validar: Lista de dicionários para o st.session_state
    - props_list_ajustada: Lista de propostas com matches exatos REMOVIDOS.
    """
    
    COLUNA_DESCRICAO = 'descricao'
    
    # Lista de tuplas (descricao, linha_completa) da proposta para busca e remoção
    # Usamos .to_dict('records') e list() para fácil acesso aos dados
    props_list_ajustada = list(zip(df_prop[COLUNA_DESCRICAO], df_prop.to_dict('records')))
    
    divergencias_para_validar = []
    
    # Itera sobre a referência
    for _, row_ref_series in df_ref.iterrows():
        # Converte a Series para um dicionário ou objeto acessível
        row_ref = row_ref_series.to_dict() 
        descricao_ref = row_ref[COLUNA_DESCRICAO]
        
        # NOTE: A função busca_descricao_fuzzy remove o item da props_list_ajustada 
        # se houver match exato.
        match_row, is_exact, divergencia_data_fuzzy = verifica_descricao_fuzzy(
            descricao_ref, 
            props_list_ajustada
            # O limite de similaridade pode ser passado aqui, se for dinâmico.
        )

        if is_exact:
            # Match Exato: Item foi removido de props_list_ajustada. OK.
            pass
        elif divergencia_data_fuzzy:
            # Match Fuzzy: Adiciona para validação. O item AINDA está em props_list_ajustada.
            # Match Fuzzy: Adiciona os detalhes da Referência e Proposta ao dicionário
            melhor_match_row = divergencia_data_fuzzy.pop('Melhor_Match_Row') # Remove o registro completo
            
            divergencia = {
                # Dados da Referência (usando row_ref)
                'Descricao_Ref': descricao_ref,
                'Item_Ref': row_ref['item'],
                'Quantidade_Ref': row_ref['quantidade'],
                'Valor_Total_Ref': row_ref['valor_total'],
                
                # Dados da Proposta (usando melhor_match_row)
                'Melhor_Match_Prop': divergencia_data_fuzzy['Melhor_Match_Prop'],
                'Item_Prop': melhor_match_row['item'], 
                'Quantidade_Prop': melhor_match_row['quantidade'],
                'Valor_Total_Prop': melhor_match_row['valor_total'],
                
                # Detalhes do Match
                'Similaridade': divergencia_data_fuzzy['Similaridade'],
                'Validado_OK': divergencia_data_fuzzy['Validado_OK'],
                'idx_na_props_list': divergencia_data_fuzzy['idx_na_props_list'],
            }
            divergencias_para_validar.append(divergencia) 
            
        
        # Se não houver match (fuzzy ou exato), o item é considerado AUSENTE neste momento.

    return divergencias_para_validar, props_list_ajustada


def gerar_relatorio_final(df_ref, df_prop, divergencias_validadas):
    """
    Gera o relatório final e os DFs de resumo, após a validação fuzzy do usuário.
    """
    
    COLUNA_DESCRICAO = 'descricao'
    
    # --------------------------------------------------------
    # 1. PRÉ-PROCESSAMENTO E DEFINIÇÃO DE CONJUNTOS
    # --------------------------------------------------------
    
    # a) Recria a props_list com os itens da proposta
    props_list_original = list(zip(df_prop[COLUNA_DESCRICAO], df_prop.to_dict('records')))
    
    # b) Identifica quais descrições foram validadas OK (Chaves da Proposta)
    descricoes_validadas_ok_prop = {
        d['Melhor_Match_Prop'] 
        for d in divergencias_validadas 
        if d['Validado_OK']
    }
    # c) Mapeamento (Referência -> Proposta) para itens fuzzy validados
    map_fuzzy_validado = {
        d['Descricao_Ref']: d['Melhor_Match_Prop']
        for d in divergencias_validadas
        if d['Validado_OK']
    }

    # d) Define a função de busca simples (para buscar na lista ORIGINAL)
    # NOTE: Esta função agora precisa garantir que os itens validados OK não sejam removidos
    # antes do final, ou teremos problemas de contagem/preenchimento. 
    def verifica_descricao_simples(descricao_ref, props_list_simples):
        descricao_ref_clean = descricao_ref.lower().replace(" ", "")
        
        for i, (descricao_prop, row) in enumerate(props_list_simples):
            descricao_prop_clean = descricao_prop.lower().replace(" ", "")
            
            # Checa match exato OU match fuzzy validado
            if descricao_ref_clean == descricao_prop_clean:
                # Match Exato: Remove o item e retorna
                return props_list_simples.pop(i)[1]
            
            # Se for um item da Referência que foi validado OK, NÃO TENTA BUSCAR AQUI, 
            # lidamos com ele na lógica principal abaixo.
            
        return None
        
    # --------------------------------------------------------
    # 2. GERAÇÃO DO RELATÓRIO E CRUZAMENTO
    # --------------------------------------------------------
    relatorio = []
    
    # props_list_busca: Cópias da lista para remover itens EXATOS encontrados
    props_list_busca_exata = props_list_original[:]
    
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
        
        # --- Lógica de Busca ---
        
        # 1. Tenta Match Exato (usa a função de busca simples que remove o item)
        row_prop = verifica_descricao_simples(descricao_ref, props_list_busca_exata)
        
        # 2. Verifica Match Fuzzy Validado (Se o item não foi encontrado por match exato)
        is_fuzzy_validado = False
        if row_prop is None and descricao_ref in map_fuzzy_validado:
            # Encontramos o item da Proposta que corresponde ao Match Fuzzy Validado
            prop_desc_validada = map_fuzzy_validado[descricao_ref]
            
            # Encontrar o row_prop correspondente na props_list_original (sem remover)
            for desc, row in props_list_original:
                if desc == prop_desc_validada:
                    row_prop = row # Encontramos os dados da Proposta!
                    is_fuzzy_validado = True
                    break
            
            # Remove o item da lista de busca exata para garantir que não seja contado 
            # como um "extra" ou "ausente" no futuro
            # Esta remoção é complexa e perigosa. Simplificando: Apenas trate o item como presente.

        
        # --- Preenchimento do Relatório ---
        
        if row_prop is not None:
            # Item Encontrado (Exato OU Fuzzy Validado)
            
            # Acessando os valores da proposta
            item_prop = row_prop['item'] # <-- CORRIGIDO: row_prop agora é um dict/registro
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
            
            
            # Realizando os comparativos (usando as funções auxiliares que você já tem)
            resultado['item_ok'] = bool(verifica_item(item_prop, item_ref))
            resultado['quantidade_ok'] = verifica_quantidade(quantidade_ref, quantidade_prop)
            resultado['valor_total_ok'] = verifica_valor_total(valor_total_prop, quantidade_prop, valor_unit_bdi_prop)
            resultado['desconto_prop'], resultado['desconto_ok'] = verifica_desconto (valor_total_prop, valor_total_ref, limiar=25)
            
            # Marcador de Match
            if is_fuzzy_validado:
                resultado['descricao_match_tipo'] = 'Fuzzy Validado'
            else:
                resultado['descricao_match_tipo'] = 'Exato'


        else:
            # Item Ausente (Não encontrado e não foi match fuzzy validado)
            resultado['presente'] = False
            resultado['item_prop'] = '-'
            resultado['descricao_prop'] = '-'
            resultado['quantidade_prop'] = 0
            resultado['valor_unit_prop'] = 0.0
            resultado['valor_unit_bdi_prop'] = 0.0
            resultado['valor_total_prop'] = 0.0
            resultado['desconto_prop'] = 0.0 # Garante que a coluna existe
            resultado['descricao_match_tipo'] = 'Ausente'
                    
            # Definindo os campos de comparação como False
            resultado['item_ok'] = False
            resultado['quantidade_ok'] = False
            resultado['valor_total_ok'] = False
            resultado['desconto_ok'] = True  # True para não contar como desconto com problema
            
        relatorio.append(resultado)
        
    df_relatorio = pd.DataFrame(relatorio)
    
    # --------------------------------------------------------
    # 3. TRATAMENTO DE EXTRAS E FINALIZAÇÃO
    # --------------------------------------------------------
    
    # O que sobrou da PROPOSTA (props_list_busca_exata) são itens que não deram match exato.
    extras_prop_nao_match_exato = pd.DataFrame([row for _, row in props_list_busca_exata])
    
    # 3.1. REMOVER itens que foram Match Fuzzy Validado
    if not descricoes_validadas_ok_prop:
        df_itens_extras_prop = extras_prop_nao_match_exato.copy()
    else:
        # Remove os itens validados OK, pois eles NÃO SÃO EXTRAS, apenas matches flexíveis.
        df_itens_extras_prop = extras_prop_nao_match_exato[
            ~extras_prop_nao_match_exato[COLUNA_DESCRICAO].isin(descricoes_validadas_ok_prop)
        ].copy()


    # 3.2. Finalização (Seu código original)
    
    dict_resumo_descontos = resumo_descontos(df_relatorio)
    df_relatorio['desconto_prop']=df_relatorio['desconto_prop'].apply(normaliza_desconto)
    
    # ... (Seu código de somatório, ausentes, descontos_prop, e dropna - idêntico) ...

     # -------------------
    # Somatório Total das colunas de Valor Total #TODO Fazzer um dict
    # -------------------
    soma_valor_global_prop = df_prop.loc[~df_prop["item"].astype(str).str.contains("\."), "valor_total"].sum()
    soma_valor_global_ref = df_ref.loc[~df_ref["item"].astype(str).str.contains("\."), "valor_total"].sum()

    # -------------------
    # Itens extras (para analisar, pois estão com problema)
    # -------------------
    extras_prop = pd.DataFrame([row for _, row in props_list_original])

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
    # extras_prop = extras_prop.dropna(subset=['valor_total'])
    extras_prop = extras_prop.dropna(subset=['valor_total']) if not extras_prop.empty else extras_prop  #Para remover os itens nulos caso não esteja vazio
    ausentes_prop = ausentes_prop.dropna(subset=['valor_total_ref']) if not ausentes_prop.empty else ausentes_prop  #Para remover os itens nulos caso não esteja vazio
    descontos_prop = descontos_prop.dropna(subset=['valor_total_ref']) if not descontos_prop.empty else descontos_prop  #Para remover os itens nulos caso não esteja vazio

    
    # # Ajustando variáveis para retorno
    # df_itens_extras_prop = extras_prop_final
    # ausentes_prop = df_relatorio.loc[df_relatorio['presente'] == False]
    
    
    # Você precisará dos totais e BDI dos DFs originais, que foram perdidos aqui.
    # Assumindo que você pode re-obter esses valores, ou eles foram passados como argumentos:
    
    # Placeholder (Adapte para sua função original de carregamento/extração de totais)
    dict_totais_ref = {} 
    dict_totais_prop = {}
    df_valores_bdi_diferente_ref = pd.DataFrame()
    df_valores_bdi_diferente_prop = pd.DataFrame()
    
    # Retorno completo (com a adição dos retornos necessários ao app.py)
    return (
        df_relatorio, df_itens_extras_prop, ausentes_prop, descontos_prop, 
        soma_valor_global_prop, soma_valor_global_ref, dict_resumo_descontos,
        dict_totais_ref, dict_totais_prop, df_valores_bdi_diferente_ref, df_valores_bdi_diferente_prop
    )


