# ================ IMPORTAÇÕES ================
import streamlit as st
import math
import re
import pandas as pd
import pyexcel_ods
from comparador import comparar_planilhas, calcula_desconto_total_final, iniciar_comparacao_fuzzy, gerar_relatorio_final
from leitor_planilha import carregar_planilha
from relatorio import organizar_relatorio, destacar_itens, construir_df_resumo_totais_globais
from io import BytesIO

# ================ FUNÇÕES AUXILIARES ================
def ler_planilha(uploaded_file):
    """
    Lê um arquivo enviado pelo Streamlit, suportando XLSX e ODS.
    Retorna um DataFrame ou None em caso de erro.
    """
    if not uploaded_file:
        return None
    
    filename = uploaded_file.name.lower()
    
    try:
        if filename.endswith(".xlsx"):
            # lê XLSX normalmente
            df = pd.read_excel(uploaded_file)
            return df
        elif filename.endswith(".ods"):
            # lê ODS via pyexcel_ods, pega a primeira aba
            data = pyexcel_ods.get_data(uploaded_file)
            first_sheet_name = list(data.keys())[0]
            df = pd.DataFrame(data[first_sheet_name])
            return df
        else:
            st.error("Formato de arquivo não suportado. Use XLSX ou ODS.")
            return None
    except Exception as e:
        st.error(f"Erro ao abrir o arquivo: {e}")
        return None

def num_para_real(valor):
    # Trunca para 2 casas decimais
    valor_truncado = math.trunc(valor * 100) / 100
    # Formata em Real
    return f"R$ {valor_truncado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def num_para_percentual(valor):
    return f"{valor:.2f}%".replace(".", ",")  
    
    
def texto_para_float(valor):
    """
    Converte uma string com valores monetários em float.
    Suporta formatos como:
    - "R$ 1.234,56"
    - "1234.56"
    - "1.23456"
    - "1234,56"
    - "  1 234,56  "
    """
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return math.floor(float(valor) * 100) / 100  # trunca para 2 casas decimais
    
    s = str(valor)
    # Remove o "R$" e espaços
    s = s.replace("R$", "").replace(" ", "")
    
    # Troca vírgula por ponto se houver
    if "," in s and "." in s:
        # assume que o formato é "1.234,56" -> 1234.56
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    
    # Remove qualquer outro caractere que não seja número ou ponto
    s = re.sub(r"[^0-9.]", "", s)
    
    try:
        numero = float(s)
        # trunca para 2 casas decimais
        return math.floor(numero * 100) / 100
    except ValueError:
        return 0.0


        # Função helper para card colorido


def metric_card(title, value, color, height):
    st.markdown(
        f"""
        <div style="
            background-color:{color};
            padding:20px;
            border-radius:10px;
            text-align:center;
            color:white;
            font-weight:bold;
            min-height:{height}px;
            display:flex;
            flex-direction:column;
            justify-content:center;">
            <div style="font-size:18px;">{title}</div>
            <div style="font-size:28px;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def renderiza_card(cards_valores):
    max_len_valores = max(len(str(title)) + len(str(value)) for title, value, _ in cards_valores)
    base_height_valores = 100
    extra_per_char_valores = 2
    common_height_valores = base_height_valores + (max_len_valores * extra_per_char_valores)

    # 🔹 Renderiza os cards lado a lado
    cols_valores = st.columns(len(cards_valores))
    for col, (title, value, color) in zip(cols_valores, cards_valores):
        with col:
            metric_card(title, value, color, common_height_valores)
    
    
# ================ APLICATIVO STREAMLIT ================ 
    
# ----------------------------------------------------------------------
# GESTÃO DE ESTADO E INICIALIZAÇÃO
# ----------------------------------------------------------------------

st.set_page_config(page_title="DINFRA - Comparador de Orçamentos", layout="wide")

# Inicializa o estado da sessão para controlar a navegação e dados
if 'etapa' not in st.session_state:
    st.session_state.etapa = 'upload' # 'upload', 'validacao_divergencias', 'resumo_final'
if 'divergencias_desc' not in st.session_state:
    st.session_state.divergencias_desc = []
if 'df_ref' not in st.session_state:
    st.session_state.df_ref = None
if 'df_prop' not in st.session_state:
    st.session_state.df_prop = None
if 'props_list_fuzzy_pendente' not in st.session_state:
    st.session_state.props_list_fuzzy_pendente = None # Lista da proposta após matches exatos
if 'valor_comprasnet' not in st.session_state:
    st.session_state.valor_comprasnet = 0.0  
    
    
# ----------------------------------------------------------------------
# FUNÇÕES DE FLUXO (CONTROLE DE TELA)
# ----------------------------------------------------------------------

def etapa_comparacao_e_fuzzy(ref_file, prop_file):
    """
    Etapa 1: Carrega os dados e inicia a busca fuzzy, definindo a próxima tela.
    """
    try:
        with st.spinner("Processando planilhas..."):
            # 1. Carrega os DataFrames (usa a função do módulo lógico)
            st.session_state.df_ref, _, _ = carregar_planilha(ref_file)
            st.session_state.df_prop, _, _ = carregar_planilha(prop_file)

            # 2. Executa a Lógica Fuzzy (usa a função do módulo lógico)
            divergencias, props_pendentes = iniciar_comparacao_fuzzy(
                st.session_state.df_ref, 
                st.session_state.df_prop
            )
            
            st.session_state.divergencias_desc = divergencias
            st.session_state.props_list_fuzzy_pendente = props_pendentes
            
            st.success("Processamento inicial concluído.")

    except Exception as e:
        st.error(f"❌ Erro na etapa de pré-processamento/fuzzy match: {e}")
        st.session_state.etapa = 'upload'
        return

    # 3. Define a Próxima Etapa
    if st.session_state.divergencias_desc:
        st.session_state.etapa = 'validacao_divergencias'
    else:
        st.session_state.etapa = 'resumo_final'
        
    st.rerun()


def etapa_validacao_divergencias():
    """
    Etapa 2: Interface para o usuário validar as descrições suspeitas.
    """
    st.title("⚠️ Validação Manual de Descrições com divergências")

    if not st.session_state.divergencias_desc:
        st.success("Não foram encontradas divergências para validação manual. Avançando.")
        st.session_state.etapa = 'resumo_final'
        st.rerun()
        return

    st.warning(f"Total de {len(st.session_state.divergencias_desc)} itens para validação. Analise cuidadosamente se a descrição da Proposta é aceitável para o item da Referência.")

    novas_divergencias = st.session_state.divergencias_desc[:]
    
    # --- Loop dos Itens Divergentes ---
    for i, item in enumerate(novas_divergencias):
        
        # Usa um container para delimitar visualmente cada item divergente
        with st.container(border=True):
            
            # --- Linha de Ação e Título (Topo) ---
            col_titulo, col_acao = st.columns([0.8, 0.2])
            col_titulo.markdown(f"**Item {i+1}** | Grau de Similaridade: **{item['Similaridade']}**")
            
            # Checkbox de validação (na coluna de ação)
            validado = col_acao.checkbox(
                "Aceitável (OK)", 
                value=item['Validado_OK'],
                key=f"valida_{i}",
                label_visibility="visible"
            )
            novas_divergencias[i]['Validado_OK'] = validado
            
            st.markdown("---") # Separador para o bloco de dados

            # --- Estrutura de Comparação Linha-a-Linha ---
            
            # 1. Descrição
            c_label_desc, c_ref_desc, c_prop_desc = st.columns([0.15, 0.4, 0.45])
            c_label_desc.markdown("**Descrição**")
            c_ref_desc.text(item['Descricao_Ref'])
            c_prop_desc.text(item['Melhor_Match_Prop'])
            
            # 2. Item
            c_label_item, c_ref_item, c_prop_item = st.columns([0.15, 0.4, 0.45])
            c_label_item.markdown("**Item**")
            c_ref_item.text(item.get('Item_Ref', '-'))
            c_prop_item.text(item.get('Item_Prop', '-'))
            
            # 3. Qtd
            c_label_qtd, c_ref_qtd, c_prop_qtd = st.columns([0.15, 0.4, 0.45])
            c_label_qtd.markdown("**Qtd**")
            c_ref_qtd.text(item.get('Quantidade_Ref', '-'))
            c_prop_qtd.text(item.get('Quantidade_Prop', '-'))
            
            # 4. Valor
            c_label_valor, c_ref_valor, c_prop_valor = st.columns([0.15, 0.4, 0.45])
            c_label_valor.markdown("**Valor Total**")
            c_ref_valor.text(num_para_real(item.get('Valor_Total_Ref', 0)))
            c_prop_valor.text(num_para_real(item.get('Valor_Total_Prop', 0)))

            # st.markdown("---") # Fim do bloco de dados (dentro do container)


    # Atualiza o estado da sessão com as validações
    st.session_state.divergencias_desc = novas_divergencias
    
    # Botão para finalizar (fora do loop)
    if st.button("Finalizar Validação e Gerar Resumo"):
        st.session_state.etapa = 'resumo_final'
        st.rerun()


def etapa_resumo_final():
    """
    Etapa 3: Executa a comparação final e exibe o resumo completo.
    """
    st.title("Geração de Relatório e Resumo Final")
    
    # 1. Executa a Lógica de Comparação Final (usa a função do módulo lógico)
    try:
        with st.spinner("Gerando relatório completo e cruzamento de todos os dados..."):
                        
            # Chama a função de lógica que realiza todo o cruzamento final, 
            # usando os DFs originais e os resultados da validação
            (
                df_relatorio, df_itens_extras_prop, df_itens_ausentes_prop, 
                df_descontos_problema, soma_valor_global_prop, soma_valor_global_ref, 
                dict_resumo_descontos, dict_totais_ref, dict_totais_prop, 
                df_valores_bdi_diferente_ref, df_valores_bdi_diferente_prop
            ) = gerar_relatorio_final(
                st.session_state.df_ref, 
                st.session_state.df_prop, 
                st.session_state.divergencias_desc # Passa o resultado da validação
            )
            
            # Funções de resumo (usadas para exibição no app.py)
            df_relatorio = organizar_relatorio(df_relatorio)
            desconto_total_final = calcula_desconto_total_final(soma_valor_global_ref, soma_valor_global_prop)
            
    except Exception as e:
        st.error(f"❌ Erro ao gerar o relatório final: {e}.")
        return

    # 2. CONSTRUÇÃO E EXIBIÇÃO DO RESUMO (código original de cards e dataframes)
    
    # Bloco de Resumo de Totais
        
    df_resumo_totais_globais = construir_df_resumo_totais_globais(
            dict_totais_ref,
            dict_totais_prop,
            soma_valor_global_ref,
            soma_valor_global_prop,
            st.session_state.valor_comprasnet, 
            desconto_total_final
            
        )
        
    # -------------------
    # Painel de resumo
    # -------------------
    total_itens = len(df_relatorio)
    ausentes = len(df_itens_ausentes_prop)
    para_analise = len(df_itens_extras_prop)
    descontos_altos = len(df_descontos_problema)
    
    st.subheader("📌 Resumo da Verificação")

    # 🔹 Prepara os cards (título, valor, cor)
    info_cards = [
        ("Itens Totais da planilha de referência", total_itens, "#4CAF50"),
        ("Itens Ausentes na planilha proposta", ausentes, "#F44336" if ausentes > 0 else "#4CAF50"),
        ("Itens a mais ou com alguma divergência na descrição", para_analise, "#FF9800" if para_analise > 0 else "#4CAF50"),
        ("Itens com desconto fora do padrão", descontos_altos, "#F44336" if descontos_altos > 0 else "#4CAF50"),
    ]

    renderiza_card(info_cards)


    # -------------------
    # Mostrar resultado
    # -------------------

    st.write("### 📊 Resultado da Comparação")
    st.dataframe(df_relatorio, use_container_width=True)
    st.caption("Obs: Os descontos dos itens ausentes na planilha de proposta foram marcados como Verdadeiro para fins de análise.")

    # Exibir a validação fuzzy (para registro)
    st.subheader("📜 Resultados da Validação Manual de Descrições")
    if st.session_state.divergencias_desc:
        df_validacao = pd.DataFrame(st.session_state.divergencias_desc)
        st.dataframe(df_validacao, hide_index=True)
    else:
        st.write("Nenhuma divergência foi validada manualmente.")

    st.write("## ⚠️ Itens para Análise")

    
    st.subheader("📌 Valores Totais")
    info_cards = [
        ("Valor Global da planilha referência apresentado", num_para_real(soma_valor_global_ref), "#4CAF50"),
        ("Valor Global da planilha proposta apresentado", num_para_real(dict_totais_prop.get("Total Geral", 0)), "#F44336" if dict_totais_prop.get("Total Geral", 0) > st.session_state.valor_comprasnet else "#4CAF50"),
        ("Valor Global da planilha proposta calculado", num_para_real(soma_valor_global_prop), "#F44336" if soma_valor_global_prop > st.session_state.valor_comprasnet else "#4CAF50"),
        ("Valor apresentado no Comprasnet", num_para_real(st.session_state.valor_comprasnet), "#F44336" if st.session_state.valor_comprasnet > soma_valor_global_prop else "#4CAF50"),
        ("Valor Global do Desconto", num_para_percentual(desconto_total_final), "#F44336" if (desconto_total_final > 25 or desconto_total_final < 0) else "#4CAF50"),
    ]

    renderiza_card(info_cards)


    st.write("### 🟡 Itens de referência ausentes na Planilha de Proposta")
    if (not df_itens_ausentes_prop.empty):
        st.dataframe(df_itens_ausentes_prop, use_container_width=True)   

    else:
        st.write("Nenhum item da planilha de referência ausente na planilha de proposta.")
    
    
    st.write("### 🟡 Planilha de Proposta: Itens a mais ou com alguma divergência na descrição")
    if (not df_itens_extras_prop.empty):
        st.dataframe(df_itens_extras_prop, use_container_width=True)   
    else:
        st.write("Nenhum item a mais ou divergente encontrado na planilha de proposta.")

    st.write("### 🟡 Planilha de Proposta: Itens com desconto fora do padrão")
    if (not df_descontos_problema.empty):
        st.dataframe(df_descontos_problema, use_container_width=True)   
    else:
        st.write("Nenhum item com desconto fora do padrão foi encontrado na planilha de proposta.")
    
    st.divider()
    
    st.write("## 📋 Detalhes Adicionais")        
            
    # st.write("#### 🟡 Planilha de Referência: BDI sinalizado com valores diferentes")
    # if (not df_valores_bdi_diferente_ref.empty):
    #     st.dataframe(df_valores_bdi_diferente_ref) 
    # else:
    #     st.write("Nenhum valor diferente encontrado na planilha de referência.")

    # st.write("#### 🟡 Planilha de Proposta: BDI sinalizado com valores diferentes")
    # if (not df_valores_bdi_diferente_prop.empty):
    #     st.dataframe(df_valores_bdi_diferente_prop)
    # else:
    #     st.write("Nenhum valor diferente encontrado na planilha de proposta.")

    # st.divider()
    
    st.write("#### 📋 Planilha de Proposta: 10 Maiores descontos")
    st.dataframe(dict_resumo_descontos.get("maiores_10"))

    st.write("#### 📋  Planilha de Proposta: 10 Menores descontos")
    st.dataframe(dict_resumo_descontos.get("menores_10"))

    st.divider()
    
    st.write("### 📋 Planilha de Referência utilizada")
    st.dataframe(st.session_state.df_ref)

    st.write("### 📋 Planilha de Proposta utilizada")
    st.dataframe(st.session_state.df_prop)

    
    # Exemplo do Painel de Resumo
    total_itens = len(df_relatorio)
    ausentes = len(df_itens_ausentes_prop)
    
    # A contagem de 'para_analise' deve ser ajustada na função 'gerar_relatorio_final' 
    # para descontar os itens validados OK.
    para_analise = len(df_itens_extras_prop) 
    descontos_altos = len(df_descontos_problema)
        

    st.divider()
    # Exportar resultado
    df_relatorio = destacar_itens(df_relatorio, "desconto_prop")
                    
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_relatorio.to_excel(writer, index=False, sheet_name="Analise Completa")
        df_itens_ausentes_prop.to_excel(writer, index=False, sheet_name="Ausentes na Proposta")
        df_itens_extras_prop.to_excel(writer, index=False, sheet_name="Itens Extras ou Divergentes")
        df_descontos_problema.to_excel(writer, index=False, sheet_name="Descontos Problemáticos")
        (dict_resumo_descontos.get("maiores_10")).to_excel(writer, index=False, sheet_name="10 maiores descontos proposta")
        (dict_resumo_descontos.get("menores_10")).to_excel(writer, index=False, sheet_name="10 menores descontos proposta")
        df_resumo_totais_globais.to_excel(writer, index=False, sheet_name="Resumo Totais Globais")
        df_valores_bdi_diferente_ref.to_excel(writer, index=False, sheet_name="BDI Diferente - Referencia")
        df_valores_bdi_diferente_prop.to_excel(writer, index=False, sheet_name="BDI Diferente - Proposta")
        # Planilhas originais
        st.session_state.df_ref.to_excel(writer, index=False, sheet_name="Orçamento de Referência")
        st.session_state.df_prop.to_excel(writer, index=False, sheet_name="Orçamento de Proposta")


    # Volta o ponteiro para o início do arquivo
    output.seek(0)

    st.download_button(
        label="⬇️ Baixar Relatório Completo em Excel",
        data=output.getvalue(),
        file_name="relatorio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # csv = df_relatorio.to_csv(index=False).encode("utf-8")
    # st.download_button(
    #     label="⬇️ Baixar Relatório em CSV",
    #     data=csv,
    #     file_name="relatorio.csv",
    #     mime="text/csv"
    # )
    
    # resultado = analisar_coluna_item(df_itens_extras_prop, coluna="item")
    # st.write(relatorio_hierarquico(resultado))
    

    if st.button("Reiniciar Análise"):
        st.session_state.etapa = 'upload'
        st.session_state.divergencias_desc = []
        st.rerun() 
    
    
    
    
# ----------------------------------------------------------------------
# FUNÇÃO PRINCIPAL DE CONTROLE DE FLUXO
# ----------------------------------------------------------------------

def main():
    st.title("📊 DINFRA - Comparador de Orçamentos")
    st.subheader("Módulo Orçamento Sintético")
    
    # --- Input do valor global ---
    valor_texto = st.text_input(
        "Digite o valor global da proposta:",
        placeholder="R$00,00", width=350, value=None
    )
    valor_comprasnet = texto_para_float(valor_texto)
    st.session_state.valor_comprasnet = valor_comprasnet

    # --- Upload dos arquivos ---
    ref_file = st.file_uploader("Carregar planilha de referência", type=["xlsx"], key='ref_up')
    prop_file = st.file_uploader("Carregar planilha de proposta", type=["xlsx"], key='prop_up')
    
    pronto_para_comparar = (
        ref_file and prop_file and 
        valor_comprasnet is not None and valor_comprasnet > 0.00
    )
    
    # --- Controle de Etapas ---
    
    if pronto_para_comparar:
        if st.session_state.etapa == 'upload':
            st.write("✅ Clique em **Comparar** para iniciar o processamento e a validação.")
            if st.button("🔎 Comparar planilhas"):
                etapa_comparacao_e_fuzzy(ref_file, prop_file)
                    
        elif st.session_state.etapa == 'validacao_divergencias':
            etapa_validacao_divergencias()

        elif st.session_state.etapa == 'resumo_final':
            etapa_resumo_final()
            
    elif st.session_state.etapa != 'upload':
        st.session_state.etapa = 'upload'
        st.rerun()


if __name__ == "__main__":
    main()    
    