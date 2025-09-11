# Ponto de entrada da CLI


import argparse
from leitor_planilha import carregar_planilha
from comparador import comparar_planilhas
from relatorio import salvar_relatorio

def main():
    parser = argparse.ArgumentParser(description="Comparador de propostas orçamentárias.")
    parser.add_argument('--ref', required=True, help='Arquivo de referência')
    parser.add_argument('--prop', required=True, help='Arquivo da proposta')
    parser.add_argument('--margem', type=float, default=0.05, help='Margem de tolerância percentual (padrão: 0.05)')
    parser.add_argument('--saida', default='relatorio.xlsx', help='Arquivo de saída do relatório')

    args = parser.parse_args()

    df_ref = carregar_planilha(args.ref)
    df_prop = carregar_planilha(args.prop)
    relatorio = comparar_planilhas(df_ref, df_prop, margem=args.margem)    
    salvar_relatorio(relatorio, args.saida)

if __name__ == '__main__':
    main()
