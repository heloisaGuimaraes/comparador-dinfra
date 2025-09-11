# Função para gerar relatório
  
import pandas as pd

def organizar_relatorio(relatorio):    
    df_relatorio = pd.DataFrame(relatorio)
    ordem_colunas = [
    'item_ref',
    'descricao_ref',
    'quantidade_ref',
    'valor_unit_ref',
    'valor_unit_bdi_ref',
    'valor_total_ref',
    
    'presente',
    'item_ok',
    # 'descricao_ok',
    'quantidade_ok',
    'valor_total_ok',
    'desconto_ok',
    'desconto_prop',
   
    'item_prop',
    'descricao_prop',
    'quantidade_prop',
    'valor_unit_prop',
    'valor_unit_bdi_prop',
    'valor_total_prop',
]
    # Só manter as colunas que realmente existem no DataFrame (caso falte alguma opcional)
    colunas_existentes = [col for col in ordem_colunas if col in df_relatorio.columns]
    return df_relatorio[colunas_existentes]
    
    
    

def salvar_relatorio(relatorio, caminho_saida):
    df = organizar_relatorio (relatorio)
    df.to_excel(caminho_saida, index=False)
    print(f'Relatório salvo em: {caminho_saida}')


def destacar_itens(df, coluna):

    def highlight(v, color):
        return f"color: {color};" if v == False else None

    styled = df.style.map(highlight, color='red', subset=[coluna])
    return styled



def construir_df_resumo_totais_globais(dict_totais_ref, dict_totais_prop, soma_valor_global_ref, soma_valor_global_prop, valor_comprasnet, desconto_total_final):
    resumo = {
        "Descrição": [
            "Valor Global da planilha de referência apresentado",
            "Valor Global da planilha proposta apresentado",
            "Valor Global da planilha proposta calculado",
            "Valor apresentado no Comprasnet",
            "Desconto Total Final"
        ],
        "Valor": [
            soma_valor_global_ref,
            dict_totais_prop.get("Total Geral", 0),
            soma_valor_global_prop,
            valor_comprasnet,
            desconto_total_final
        ]
    }

    df_resumo = pd.DataFrame(resumo)
    return df_resumo
