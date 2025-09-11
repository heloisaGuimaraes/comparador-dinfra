import streamlit as st
import math
import re
import pandas as pd
import pyexcel_ods
from comparador import comparar_planilhas, calcula_desconto_total_final
from leitor_planilha import carregar_planilha
from relatorio import organizar_relatorio, destacar_itens, construir_df_resumo_totais_globais
from io import BytesIO

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
    
    
    
    
st.set_page_config(page_title="DINFRA - Comparador de Orçamentos", layout="wide")

st.title("📊 DINFRA - Comparador de Orçamentos")
st.subheader("Módulo Orçamento Sintético")

# Input do valor global da proposta vindo do Cmprasnet
st.write("##### Valor global da proposta conforme consta no Comprasnet:")
valor_texto = st.text_input(
    "Digite o valor:",
    placeholder="R$00,00", width=350, value=None
)

# converte para float
valor_comprasnet = texto_para_float(valor_texto)
# print(f"Valor Comprasnet: {valor_comprasnet}")

# Upload dos arquivos
st.write("##### Planilha de Referência:")
ref_file = st.file_uploader("Carregar planilha de referência", type=["xlsx"])
st.write("##### Planilha de Proposta:")
prop_file = st.file_uploader("Carregar planilha de proposta", type=["xlsx"])



if (ref_file and prop_file and valor_comprasnet != None and valor_comprasnet > 0.00):
    st.write("✅ Clique em **Comparar** para processar.")
    
    if st.button("🔎 Comparar"):
# Processa planilha de referência
        try:
            with st.spinner("Processando a planilha de referência, por favor aguarde..."):
                df_ref, dict_totais_ref, df_valores_bdi_diferente_ref = carregar_planilha(ref_file)
        except ValueError as e:
            st.error(f"❌ Erro ao processar a planilha de referência: {e}. Revise se o arquivo está correto.")
            st.stop()

        # Processa planilha de proposta
        try:
            with st.spinner("Processando a planilha de proposta, por favor aguarde..."):
                df_prop, dict_totais_prop, df_valores_bdi_diferente_prop = carregar_planilha(prop_file)
        except ValueError as e:
            st.error(f"❌ Erro ao processar a planilha de proposta: {e}. Revise se o arquivo está correto.")
            st.stop()

        # Comparação
        try:
            with st.spinner("Comparando planilhas, por favor aguarde..."):
                df_relatorio, df_itens_extras_prop, df_itens_ausentes_prop, df_descontos_problema, soma_valor_global_prop, soma_valor_global_ref, dict_resumo_descontos = comparar_planilhas(df_ref, df_prop)
                df_relatorio = organizar_relatorio(df_relatorio)
                desconto_total_final = calcula_desconto_total_final(soma_valor_global_ref, soma_valor_global_prop)
                
        except Exception as e:
            st.error(f"❌ Erro ao comparar as planilhas: {e}.")
            st.stop()

        st.success("✅ Processamento concluído!")
        df_resumo_totais_globais = construir_df_resumo_totais_globais(
            dict_totais_ref,
            dict_totais_prop,
            soma_valor_global_ref,
            soma_valor_global_prop,
            valor_comprasnet, 
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

        st.write("## ⚠️ Itens para Análise")
   
        
        st.subheader("📌 Valores Totais")
        info_cards = [
            ("Valor Global da planilha referência apresentado", num_para_real(soma_valor_global_ref), "#4CAF50"),
            ("Valor Global da planilha proposta apresentado", num_para_real(dict_totais_prop.get("Total Geral", 0)), "#F44336" if dict_totais_prop.get("Total Geral", 0) > valor_comprasnet else "#4CAF50"),
            ("Valor Global da planilha proposta calculado", num_para_real(soma_valor_global_prop), "#F44336" if soma_valor_global_prop > valor_comprasnet else "#4CAF50"),
            ("Valor apresentado no Comprasnet", num_para_real(valor_comprasnet), "#F44336" if valor_comprasnet > soma_valor_global_prop else "#4CAF50"),
            ("Valor Global do Desconto", num_para_percentual(desconto_total_final), "#F44336" if (desconto_total_final > 25 or desconto_total_final < 0) else "#4CAF50"),
        ]

        renderiza_card(info_cards)


        st.write("### 🟡 Itens de referência ausentes na Planilha de Proposta")
        st.dataframe(df_itens_ausentes_prop, use_container_width=True)   
        
        st.write("### 🟡 Planilha de Proposta: Itens a mais ou com alguma divergência na descrição")
        st.dataframe(df_itens_extras_prop, use_container_width=True)   

        st.write("### 🟡 Planilha de Proposta: Itens com desconto fora do padrão")
        st.dataframe(df_descontos_problema, use_container_width=True)   
        
        st.divider()
        
        st.write("## 📋 Detalhes Adicionais")        
               
        st.write("#### 🟡 Planilha de Referência: BDI sinalizado com valores diferentes")
        if (not df_valores_bdi_diferente_ref.empty):
            st.dataframe(df_valores_bdi_diferente_ref) 
        else:
            st.write("Nenhum valor diferente encontrado na planilha de referência.")

        st.write("#### 🟡 Planilha de Proposta: BDI sinalizado com valores diferentes")
        if (not df_valores_bdi_diferente_prop.empty):
            st.dataframe(df_valores_bdi_diferente_prop)
        else:
            st.write("Nenhum valor diferente encontrado na planilha de proposta.")

        st.divider()
        
        st.write("#### 📋 Planilha de Proposta: 10 Maiores descontos")
        st.dataframe(dict_resumo_descontos.get("maiores_10"))

        st.write("#### 📋  Planilha de Proposta: 10 Menores descontos")
        st.dataframe(dict_resumo_descontos.get("menores_10"))

        st.divider()
        
        st.write("### 📋 Planilha de Referência utilizada")
        st.dataframe(df_ref)

        st.write("### 📋 Planilha de Proposta utilizada")
        st.dataframe(df_prop)


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
            df_ref.to_excel(writer, index=False, sheet_name="Orçamento de Referência")
            df_prop.to_excel(writer, index=False, sheet_name="Orçamento de Proposta")


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
        

        
        