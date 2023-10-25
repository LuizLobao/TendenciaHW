def calcular_e_filtrar(df, produtos, segmento, gest):
    print(df, 'VL', produtos, segmento, gest)

df='a'
produtos = ['FIBRA', 'NOVA FIBRA']
segmentos = ['VAREJO', 'EMPRESARIAL']
gestao = ['RSE', 'RCS', 'RNN', 'TLV', 'WEB', 'OUTROS NACIONAIS']

for produto in produtos:
    for segmento in segmentos:
        for gest in gestao:
            calcular_e_filtrar(df, produto, segmento, gest)