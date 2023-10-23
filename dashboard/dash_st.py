import dash
from dash import html, dcc
import plotly.express as px
import pyodbc
import pandas as pd
import warnings

# Ignorar o aviso específico sobre o tipo de conexão com banco de dados
warnings.filterwarnings('ignore', message="pandas only supports SQLAlchemy connectable .*")


def criar_conexao():
    dados_conexao = (
        "Driver={SQL Server};"
        "Server=SQLPW90DB03\DBINST3, 1443;"
        "Database=BDintelicanais;"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(dados_conexao)

def puxa_dados_real():
    conexao = criar_conexao()
    comando_sql = f'''
                   SELECT DS_TIPO, DS_PRODUTO, DS_INDICADOR, DS_DET_INDICADOR, DT_REFERENCIA, DS_VELOCIDADE, DT_ANOMES, DS_UNIDADE_NEGOCIO,  DS_SEGMENTACAO, DS_CLASS_CELULA, SUM(QTD) AS QTD
                    FROM [BDINTELICANAIS].[DBO].[TBL_CDO_FISICOS_REAL] AS A
                    WHERE [DT_ANOMES] = '202310'
                    GROUP BY DS_TIPO, DS_PRODUTO, DS_INDICADOR, DS_DET_INDICADOR, DT_REFERENCIA, DS_VELOCIDADE, DT_ANOMES, DS_UNIDADE_NEGOCIO, DS_SEGMENTACAO, DS_CLASS_CELULA
                    order by DT_REFERENCIA
                '''
    conexao.execute(comando_sql)
    
    df=pd.read_sql(comando_sql, conexao)
    conexao.close()
    return df

def consolidar_dados(df, indicador):
    df_filtered = df[df['DS_INDICADOR'] == indicador]
    df_consolidado = df_filtered.groupby(['DT_REFERENCIA', 'DS_PRODUTO'])['QTD'].sum().reset_index()
    return df_consolidado


df=puxa_dados_real()
df['DT_REFERENCIA'] = df['DT_REFERENCIA'].str.replace('/', '-')
df['DT_REFERENCIA'] = pd.to_datetime(df['DT_REFERENCIA'], format='%d-%m-%Y')


# Consolidando dados para diferentes indicadores
df_consolidado_vl = consolidar_dados(df, 'VL')
total_por_dia = df_consolidado_vl.groupby('DT_REFERENCIA')['QTD'].sum().to_dict()

df_consolidado_vll = consolidar_dados(df, 'VLL')
total_por_diavll = df_consolidado_vll.groupby('DT_REFERENCIA')['QTD'].sum().to_dict()

df_consolidado_gross = consolidar_dados(df, 'GROSS')
total_por_diagross = df_consolidado_gross.groupby('DT_REFERENCIA')['QTD'].sum().to_dict()


# Criando os gráficos
grafico1 = px.bar(df_consolidado_vl, 
                  x='DT_REFERENCIA', 
                  y='QTD', 
                  color='DS_PRODUTO', 
                  text='QTD',
                  text_auto='.2s',
                  title="VL por dia e Produto",
                  labels={'DT_REFERENCIA': 'Data', 'QTD': 'Quantidade'},
                  color_discrete_map={ # replaces default color mapping by value
                                "FIBRA": "#00d318", "NOVA FIBRA": "#009911"
                                    },
                  template="simple_white"
                )
grafico2 = px.bar(df_consolidado_vll, 
                  x='DT_REFERENCIA', 
                  y='QTD', 
                  color='DS_PRODUTO', 
                  text_auto='.2s',
                  title='VLL por dia e Produto', 
                  text='QTD',
                  labels={'DT_REFERENCIA': 'Data', 'QTD': 'Quantidade'},
                  color_discrete_map={ # replaces default color mapping by value
                                "FIBRA": "#00d318", "NOVA FIBRA": "#009911"
                                    },
                  template="simple_white")
grafico3 = px.bar(df_consolidado_gross, 
                  x='DT_REFERENCIA', 
                  y='QTD', 
                  color='DS_PRODUTO', 
                  title='GROSS', 
                  text='QTD',
                  text_auto='.2s',
                  labels={'DT_REFERENCIA': 'Data', 'QTD': 'Quantidade'},
                  color_discrete_map={ # replaces default color mapping by value
                                "FIBRA": "#00d318", "NOVA FIBRA": "#009911"
                                    },
                  template="simple_white")



# Adicionar anotações com os totais por dia
for dia, total in total_por_dia.items():
    grafico1.add_annotation(
        x=dia,  # Posição x (dia)
        y=total,  # Posição y (total do dia)
        text=f'{total:,.0f}'.replace(',', '.'),  # Texto da anotação
        showarrow=False,  # Não mostrar seta
        xshift=0,  # Ajustar a posição horizontal do texto
        font=dict(size=10),  # Tamanho da fonte
        textangle=0,
        yshift=20
        )
for dia, total in total_por_diavll.items():
    grafico2.add_annotation(
        x=dia,  # Posição x (dia)
        y=total,  # Posição y (total do dia)
        text=f'{total:,.0f}'.replace(',', '.'),  # Texto da anotação
        showarrow=False,  # Não mostrar seta
        xshift=0,  # Ajustar a posição horizontal do texto
        font=dict(size=10),  # Tamanho da fonte
        textangle=0,
        yshift=20
        )
for dia, total in total_por_diagross.items():
    grafico3.add_annotation(
        x=dia,  # Posição x (dia)
        y=total,  # Posição y (total do dia)
        text=f'{total:,.0f}'.replace(',', '.'),  # Texto da anotação
        showarrow=False,  # Não mostrar seta
        xshift=0,  # Ajustar a posição horizontal do texto
        font=dict(size=10),  # Tamanho da fonte
        textangle=0,
        yshift=20
        )

grafico1.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-90, showlegend=False)
grafico1.update_traces(textfont_size=14, textangle=270, cliponaxis=False)

grafico2.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-90, showlegend=False)
grafico2.update_traces(textfont_size=14, textangle=270, cliponaxis=False)

grafico3.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-90, showlegend=False)
grafico3.update_traces(textfont_size=14, textangle=270, cliponaxis=False)

# Inicialização da aplicação Dash
app = dash.Dash(__name__)

# Layout da aplicação
app.layout = html.Div(children=[
    html.H1("Meus Relatórios"),
    html.H2("Meus gráficos de resultados"),
    
    html.Div(children=[
        html.Div(children=[
            dcc.Graph(figure=grafico1)
        ], style={'width': '33%', 'display': 'inline-block'}),
        
        html.Div(children=[
            dcc.Graph(figure=grafico2)
        ], style={'width': '33%', 'display': 'inline-block'}),

        html.Div(children=[
            dcc.Graph(figure=grafico3)
        ], style={'width': '33%', 'display': 'inline-block'})
    ]),
    
    html.Div(children=[
        html.Div(children=[
            dcc.Graph(figure=grafico3)
        ], style={'width': '50%', 'display': 'inline-block'}),
        
        html.Div(children=[
            dcc.Graph(figure=grafico3)
        ], style={'width': '50%', 'display': 'inline-block'})
    ])
])


if __name__ == '__main__':
    app.run(debug=True)