import pandas as pd
import pyodbc
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import statsmodels.api as sm
from statsmodels.tsa.seasonal import seasonal_decompose
import time
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
import shutil
import segredos
from winotify import Notification, audio
import win32com.client as win32
import warnings
# Ignorar o aviso específico sobre o tipo de conexão com banco de dados
warnings.filterwarnings('ignore', message="pandas only supports SQLAlchemy connectable .*")



# Obter a data de hoje
data_referencia = (datetime.today()- timedelta(days=1))
AAAAMMDD_referencia = (datetime.today()- timedelta(days=0)).strftime('%Y%m%d') 
AMD_ref = (data_referencia).strftime('%Y-%m-%d') 
ultimo_dia_do_mes = (data_referencia.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
dias_faltando = (ultimo_dia_do_mes - data_referencia).days
AAAAMM_ant = (datetime.today()- relativedelta(months=1)).strftime('%Y%m')
AAAAMM = (datetime.today()).strftime('%Y%m')
AAAAMMDD = (datetime.today()).strftime('%Y%m%d')
hoje = (datetime.today()- timedelta(days=0)).strftime('%d/%m/%Y') 

print(f'Data Referencia: {data_referencia}')
print(f'AAAAMMDD_Referencia: {AAAAMMDD_referencia}')
print(f'AMD_ref: {AMD_ref}')
print(f'Ultimo dia do mês: {ultimo_dia_do_mes}')
print(f'Dias Faltando: {dias_faltando}')
print(f'AnoMes Ant: {AAAAMM_ant}')
print(f'Anomes: {AAAAMM}')






def criar_conexao():
    dados_conexao = (
        "Driver={SQL Server};"
        "Server=SQLPW90DB03\DBINST3, 1443;"
        "Database=BDintelicanais;"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(dados_conexao)

def montaExcelTendVlVll():
    print('montaExcelTendVlVll')
    notificacao = Notification(app_id="Tendência HoltWinters", title="Montando Excel VL e VLL",
                           msg = "Preparando excel com Tendencia de VL e VLL para enviar por e-mail",
                           duration = 'short',
                           icon=r"C://Users/oi066724/Documents/Python/selfie.png")
    #notificacao.set_audio(audio.LoopingAlarm, loop=False)
    notificacao.show()
    
    
    comando_sql = f'''select DS_PRODUTO,
                            DS_INDICADOR,
                            DS_UNIDADE_NEGOCIO,
                            NO_CURTO_TERRITORIO as FILIAL,
                            sum(qtd) as QTD,
                            TS_ATUALIZACAO
                    FROM TBL_CDO_fisicos_tendencia
                    where DT_ANOMES = '{AAAAMM}' and DS_DET_INDICADOR <> 'MIG BASE' and DS_INDICADOR <> 'gross'
                    group by DS_PRODUTO,
                        DS_INDICADOR,
                        DS_UNIDADE_NEGOCIO,
                        NO_CURTO_TERRITORIO,
                        TS_ATUALIZACAO'''
    
    comando_sql_2 = '''select * from TBL_CDO_APOIO_TENDENCIA_VL_VLL'''

    conexao = criar_conexao()

    #print("Conectado")
    cursor = conexao.cursor()
    df=pd.read_sql(comando_sql, conexao)
    df2=pd.read_sql(comando_sql_2, conexao)
    
    pt_tabdin = df.pivot_table(
                                    values="QTD", 
                                    index=["FILIAL"], 
                                    columns=["DS_INDICADOR","DS_PRODUTO"], 
                                    aggfunc=sum,
                                    fill_value=0,
                                    margins=True, margins_name="TOTAL",
                                    )

    pt_tabdin2 = df2.pivot_table(
                                values ='QTD_FINAL',
                                index=['DS_INDICADOR','DS_PRODUTO','GESTAO'],
                                columns=['DATA'],
                                aggfunc=sum,
                                fill_value=0,
                                margins=True, margins_name="TOTAL",
    )
    dest_filename = (f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Tend_VL_VLL_Fibra_NovaFibra_{AAAAMMDD}.xlsx')
    with pd.ExcelWriter(dest_filename) as writer:
            df.to_excel(writer, sheet_name="DADOS",startcol=0, startrow=0, index=0)
            pt_tabdin.to_excel(writer, sheet_name="TEND",startcol=0, startrow=0, index=1)
            df2.to_excel(writer, sheet_name="DADOS_DIARIO",startcol=0, startrow=0, index=0)
            pt_tabdin2.to_excel(writer, sheet_name="TEND_DIARIA",startcol=0, startrow=0, index=1)

def enviaEmaileAnexo():	
    print('enviaEmaileAnexo')	
    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = segredos.lista_email_vll_nf_to
    email.Cc = segredos.lista_email_vll_nf_cc
    email.Subject = f"Projeção VL e VLL - FIBRA e NOVA FIBRA - {hoje}"
    email.HTMLBody = f"""
    <p>Caros,</p>

    <p>Segue o arquivo atualizado com a projeção de VL e de VLL para Fibra Legado e Nova Fibra calculada hoje: {hoje}</p>
    <p></p>
    <p></p>

    <p>Att,</p>
    <p>Lobão, Luiz</p>
    """
    anexo = (f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Tend_VL_VLL_Fibra_NovaFibra_{AAAAMMDD}.xlsx')
    email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviado")

def demonstrativo_gross():
    print('demonstrativo_gross')	
    shutil.copy(rf"Y:\\Demonstrativo Gross_Analitico_{AAAAMM}.csv", fr'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_Analitico_{AAAAMMDD}.csv')

    in_file = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_Analitico_{AAAAMMDD}.csv'
    dest_filename = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsx'

    df = pd.read_csv(in_file, decimal=',',sep=';', quotechar='"')
    df1=df[df['TIPO'].str.contains("INSTALACAO")]
    pt_instalacao = df1.query('TIPO == "INSTALACAO"'and 'MERCADO in ("EMPRESARIAL", "VAREJO")').pivot_table(
                                                                                                        values="PROJ", 
                                                                                                        index=["UF"], 
                                                                                                        columns="MERCADO", 
                                                                                                        aggfunc=sum,
                                                                                                        fill_value=0,
                                                                                                        margins=True, 
                                                                                                        margins_name="INSTALACAO",
                                                                                                        )
    df2=df[df['TIPO'].str.contains("MIGRACAO")]
    pt_migracao = df2.query('TIPO == "MIGRACAO"'and 'MERCADO in ("EMPRESARIAL","VAREJO")').pivot_table(
                                                                                                        values=["PROD","PROJ"], 
                                                                                                        #index=["UF"], 
                                                                                                        index="MERCADO", 
                                                                                                        aggfunc=sum,
                                                                                                        fill_value=0,
                                                                                                        margins=True, 
                                                                                                        margins_name="MIGRACAO",
                                                                                                        )
    with pd.ExcelWriter(dest_filename) as writer:
        pt_instalacao.to_excel(writer, sheet_name="TabDin",startcol=0, startrow=0)
        pt_migracao.to_excel(writer, sheet_name="TabDin",startcol=6, startrow=0)

def puxa_dados_real():
    print('puxa_dados_real')	
    notificacao = Notification(app_id="Tendência HoltWinters", title="Puxando Dados Real",
                           msg = "Puxando Dados REAL",
                           duration = 'short',
                           icon=r"C://Users/oi066724/Documents/Python/selfie.png"
                           )
    #notificacao.set_audio(audio.LoopingAlarm, loop=False)
    notificacao.show()
    
    
    conexao = criar_conexao()
    conexao.execute('delete from dbo.TBL_CDO_APOIO_TENDENCIA')
    conexao.commit()

    comando_sql = f'''
                    INSERT into dbo.TBL_CDO_APOIO_TENDENCIA
                    SELECT CONVERT(DATE,a.[DT_REFERENCIA],103) AS DATA
                        ,sum(CAST([QTD] AS FLOAT)) as qtd
                        ,a.[DS_INDICADOR]
                        ,a.[DS_PRODUTO]
                        ,a.[DS_UNIDADE_NEGOCIO]
                        ,CASE when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('TLV RECEPTIVO','TLV ATIVO','TLV PP') then 'TLV'
                              when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('WEB') then 'WEB'
                              when b.DS_GESTAO = 'GESTAO REGIONAL' then c.COD_REGIONAL
                              else b.DS_GESTAO
                              END  AS GESTAO
                    FROM [BDintelicanais].[dbo].[TBL_CDO_FISICOS_REAL] as A
                        left join [dbo].[TBL_CDO_DE_PARA_CANAL] as B on A.DS_CANAL_BOV = b.[DS_DESCRICAO_CANAL_BOV]
                        LEFT JOIN [dbo].[TBL_CDO_DE_PARA_REGIONAL] AS c ON A.NO_CURTO_TERRITORIO = C.NO_CURTO_TERRITORIO
                    where [DS_DET_INDICADOR] in ('NOVOS CLIENTES','MIG AQUISICAO')
                        and [dt_anomes] = '{AAAAMM}'
                        and DS_INDICADOR = 'VL'
                    group by [DS_PRODUTO]
                        ,[DS_INDICADOR]
                        ,CONVERT(DATE,[DT_REFERENCIA],103)
                        ,[DS_UNIDADE_NEGOCIO]
                        ,CASE when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('TLV RECEPTIVO','TLV ATIVO','TLV PP') then 'TLV'
                              when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('WEB') then 'WEB'
                              when b.DS_GESTAO = 'GESTAO REGIONAL' then c.COD_REGIONAL
                              else b.DS_GESTAO
                              END                              
                '''
    
    comando_sql2 = '''select * from dbo.TBL_CDO_APOIO_TENDENCIA order by DATA'''
    conexao.execute(comando_sql)
    conexao.execute(comando_sql2)

    conexao.commit()

    
    df=pd.read_sql(comando_sql2, conexao, parse_dates =['DATA'])
    df.to_csv(f'C:\\Users\\oi066724\\Documents\\Python\\Tendencia\\TEND_DEFLAC\\tendencia_{AAAAMMDD_referencia}.csv', sep=';',header = True, index = False, mode = 'w')

    #df.to_sql('dbo.tbl_cdo_teste66724_vl', conexao, index=False, if_exists='replace')
    
    conexao.close()

def puxa_dados_para_simular():
    print('puxa_dados_para_simular')	
    
    notificacao = Notification(app_id="Tendência HoltWinters", title="Puxando Dados",
                           msg = "Puxando a base de Dados para Simulações",
                           duration = 'short',
                           icon=r"C://Users/oi066724/Documents/Python/selfie.png")
    #notificacao.set_audio(audio.LoopingAlarm, loop=False)
    notificacao.show()
    
    conexao = criar_conexao()
    comando_sql = '''
                    select 
                        CONVERT(DATE,[DATA],103) AS DATA
                        ,sum(QTD) as qtd
                        ,DS_INDICADOR
                        ,DS_PRODUTO
                        ,DS_UNIDADE_NEGOCIO
                        ,GESTAO
                    from [BDintelicanais].[dbo].[TBL_CDO_FISICOS_REAL_PROFORMA_PARA_TEND_VL]
                    where CONVERT(DATE,[DATA],103) >= '2023-06-01'
                    group by 
                        CONVERT(DATE,[DATA],103)
                        ,DS_INDICADOR
                        ,DS_PRODUTO
                        ,DS_UNIDADE_NEGOCIO
                        ,GESTAO
                    order by 1
    '''
    
    df=pd.read_sql(comando_sql, conexao, parse_dates =['DATA'])
    
    #df.to_csv(f'tendencia_{AAAAMMDD_referencia}.csv', sep=';',header = True, index = False, mode = 'w')
    
    df['qtd'] = df['qtd'].astype(float)
    #print('Cabeçalho BASE FULL')
    #print(df)
    conexao.close()
    return df
    
def filtra_df (base, indicador, PRODUTO=None, UNIDADE_NEGOCIO = None, GESTAO = None):
    print('filtra_df')	
    
    df_filtrada=base.query(f'DS_INDICADOR == "{indicador}"')
    #print(f'FILTRANDO INDICADOR = {indicador}')
    #print(df_filtrada)
    
    if PRODUTO != None:
        df_filtrada=df_filtrada.query(f'DS_PRODUTO == "{PRODUTO}"')
        #print(f'FILTRANDO PRODUTO = {PRODUTO}')
        #print(df_filtrada)
    if UNIDADE_NEGOCIO != None:
        df_filtrada=df_filtrada.query(f'DS_UNIDADE_NEGOCIO == "{UNIDADE_NEGOCIO}"')
        #print(f'FILTRANDO UNID NEGOCIO = {UNIDADE_NEGOCIO}')
        #print(df_filtrada)
    if GESTAO != None:
        df_filtrada=df_filtrada.query(f'GESTAO == "{GESTAO}"')
        #print(f'FILTRANDO GESTAO = {GESTAO}')
        #print(df_filtrada)

    #print(f'Cabeçalho BASE FILTRADA: {indicador}, {PRODUTO}, {UNIDADE_NEGOCIO}, {GESTAO}')
    #print(df_filtrada)
    
    #base apenas com DATA e valor
    df_filtrada=df_filtrada[['DATA','qtd']]
    
    #Soma por DATA
    df_a=df_filtrada.groupby('DATA').sum()

    #diario 'D' > mensal 'MS'
    df_b = df_a.resample(rule='D').sum()
    return df_b

def CalculaTendencia(df, dias_ate_fim_mes, indicador, PRODUTO, UNIDADE_NEGOCIO, GESTAO):
    print(f'CalculaTendencia - {indicador}-{PRODUTO}-{UNIDADE_NEGOCIO}-{GESTAO}')
    #seasonal_decompose(df_b, model='additve', period=7).plot();
    #localiza o INDICE de uma data especifica para criar a base de TREINO e de TESTE
    #indice = df_b.index.get_loc('2023-08-01')
    #train = df_b[:indice]
    #test = df_b[indice:]

    final_model=ExponentialSmoothing(df.qtd, trend='additive', seasonal='add', seasonal_periods=7).fit()
    pred=final_model.forecast(dias_ate_fim_mes)

    dff = pred.to_frame()
    dff = dff.rename(columns = {0:'VALOR'})
    dff['INDICADOR'] = indicador
    dff['PRODUTO'] = PRODUTO
    dff['UNIDADE'] = UNIDADE_NEGOCIO
    dff['GESTAO'] = GESTAO
    
    dff.to_csv(f'C:\\Users\\oi066724\\Documents\\Python\\Tendencia\\TEND_DEFLAC\\tendencia_{AAAAMMDD_referencia}.csv', sep=';',header = False, index = True, mode = 'a')

    # Create a new column with index values
    dff['DATA'] = dff.index
    # Using reset_index() to set index into column
    dff=dff.reset_index()
    
    dff['DATA_Ajustada'] = dff['DATA'].dt.strftime('%Y-%m-%d')
    #print(dff)

    conexao = criar_conexao()
    #insere no banco os dados de deflação
    for index, linha in dff.iterrows():
        conexao.execute('Insert into dbo.TBL_CDO_APOIO_TENDENCIA (DATA, qtd, DS_INDICADOR, DS_PRODUTO, DS_UNIDADE_NEGOCIO, GESTAO) values(?,?,?,?,?,?)',
                        linha.DATA_Ajustada, linha.VALOR, linha.INDICADOR, linha.PRODUTO, linha.UNIDADE, linha.GESTAO)
    conexao.commit()
    
def filtraDF_e_CalculaTendencia(base, indicador, PRODUTO, UNIDADE_NEGOCIO, GESTAO):
    print('filtraDF_e_CalculaTendencia')
    #FILTRANDO
    df_b = filtra_df (base, indicador,PRODUTO, UNIDADE_NEGOCIO, GESTAO)
    #print('base para simulação')
    #print(df_b.head())
    dias_ate_fim_mes=dias_faltando
    CalculaTendencia(df_b, dias_ate_fim_mes, indicador, PRODUTO, UNIDADE_NEGOCIO, GESTAO)

def executa_procedure_sql_combinada(nome_procedure, param=None):
    print(f'executa_procedure_sql_combinada - {nome_procedure}')
    notificacao = Notification(app_id="Tendência HoltWinters", title="Executando Procedure SQL",
                           msg = f"Executando Procedure {nome_procedure}",
                           duration = 'short',
                           icon=r"C://Users/oi066724/Documents/Python/selfie.png")
    #notificacao.set_audio(audio.LoopingAlarm, loop=False)
    notificacao.show()
    
    conexao = criar_conexao()
   
    try:
        cursor = conexao.cursor()

        if param:
            # Executa a procedure com parâmetro
            cursor.execute(f'SET NOCOUNT ON; EXEC {nome_procedure} {param}')
        else:
            # Executa a procedure sem parâmetro
            cursor.execute(f'SET NOCOUNT ON; EXEC {nome_procedure}')
        conexao.commit()
    
    except Exception as e:
        print(f"Erro ao executar a procedure: {e}")
    
    finally:
        conexao.close()
        print('\x1b[1;33;41m' + 'Conexão Fechada'+ '\x1b[0m')

def puxa_deflac_ref():
    print('puxa_deflac_ref')
    conexao = criar_conexao()
    comando_sql = f'''
                    SELECT 
                        -- a.[dt_anomes]
                        a.[DS_PRODUTO]
                        ,a.[DS_UNIDADE_NEGOCIO]
                        ,CASE when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('TLV RECEPTIVO','TLV ATIVO','TLV PP') then 'TLV'
                            when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('WEB') then 'WEB'
                            when b.DS_GESTAO = 'GESTAO REGIONAL' then c.COD_REGIONAL
                            else b.DS_GESTAO
                            END  AS GESTAO
                        ,CASE WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('01','02','03','04','05','06','07') THEN 'S1'
                            WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('08','09','10','11','12','13','14') THEN 'S2'
                            WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('15','16','17','18','19','20','21') THEN 'S3'
                            WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('22','23','24','25','26','27','28') THEN 'S4'
                            ELSE 'S5' 
                            END AS SEMANA
                        ,sum(CASE WHEN DS_INDICADOR = 'VL' THEN CAST([QTD] AS FLOAT) ELSE 0 END) as QTD_VL
                        ,sum(CASE WHEN DS_INDICADOR = 'VLL' THEN CAST([QTD] AS FLOAT) ELSE 0 END) as QTD_VLL
                        ,(sum(CASE WHEN DS_INDICADOR = 'VLL' THEN CAST([QTD] AS FLOAT) ELSE 0 END)) / (sum(CASE WHEN DS_INDICADOR = 'VL' THEN CAST([QTD] AS FLOAT) ELSE 0 END)) -1 as pct
                    FROM [BDintelicanais].[dbo].[TBL_CDO_FISICOS_REAL] as A
                        left join [dbo].[TBL_CDO_DE_PARA_CANAL] as B on A.DS_CANAL_BOV = b.[DS_DESCRICAO_CANAL_BOV]
                        LEFT JOIN [dbo].[TBL_CDO_DE_PARA_REGIONAL] AS c ON A.NO_CURTO_TERRITORIO = C.NO_CURTO_TERRITORIO
                    where [DS_DET_INDICADOR] in ('NOVOS CLIENTES','MIG AQUISICAO')
                        and [dt_anomes] in ( '202308' )
                        and DS_INDICADOR IN ('VL','VLL')
                    group by  
                        --a.[dt_anomes]
                        [DS_PRODUTO]
                        ,[DS_UNIDADE_NEGOCIO]
                        ,CASE when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('TLV RECEPTIVO','TLV ATIVO','TLV PP') then 'TLV'
                            when b.DS_GESTAO = 'GESTAO NACIONAL' and DS_CANAL_FINAL in ('WEB') then 'WEB'
                            when b.DS_GESTAO = 'GESTAO REGIONAL' then c.COD_REGIONAL
                            else b.DS_GESTAO
                            END 
                        ,CASE WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('01','02','03','04','05','06','07') THEN 'S1'
                            WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('08','09','10','11','12','13','14') THEN 'S2'
                            WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('15','16','17','18','19','20','21') THEN 'S3'
                            WHEN RIGHT(CONVERT(DATE,a.[DT_REFERENCIA],103),2) IN ('22','23','24','25','26','27','28') THEN 'S4'
                            ELSE 'S5' 
                            END
                    order by 1,2,3,4 
                '''
    df=pd.read_sql(comando_sql, conexao)
    df.to_csv(f'C:\\Users\\oi066724\\Documents\\Python\\Tendencia\\TEND_DEFLAC\\deflac_ref_{AAAAMMDD_referencia}.csv', sep=';',header = True, index = False, mode = 'w')
    
    #deleta a tabela de deflação antes de gravar os novos dados
    #print(df)
    conexao.execute('delete from dbo.TBL_CDO_APOIO_DEFLAC_REF_TEND')
    conexao.commit()
    
    #insere no banco os dados de deflação
    for index, linha in df.iterrows():
        conexao.execute('Insert into dbo.TBL_CDO_APOIO_DEFLAC_REF_TEND (DS_PRODUTO, DS_UNIDADE_NEGOCIO, GESTAO, SEMANA, QTD_VL, QTD_VLL, pct) values(?,?,?,?,?,?,?)',
                        linha.DS_PRODUTO, linha.DS_UNIDADE_NEGOCIO, linha.GESTAO, linha.SEMANA, linha.QTD_VL, linha.QTD_VLL, linha.pct)
    conexao.commit()
    conexao.close()

def atualiza_TB_VALIDA_CARGA_TENDENCIA():
    print('atualiza_TB_VALIDA_CARGA_TENDENCIA')
    comando_sql='update TB_VALIDA_CARGA_TENDENCIA set DATA_CARGA = convert(varchar, getdate(), 120 )'

    conexao = criar_conexao()
    print("Conectado ao banco para dar update")
    cursor = conexao.cursor()
    cursor.execute(comando_sql)
    conexao.commit()
    conexao.close()
    print('Conexão Fechada')
        
puxa_deflac_ref()
df_real = puxa_dados_real()
executa_procedure_sql_combinada('SP_CDO_PREPARA_BASE_TEND_VL')

df = puxa_dados_para_simular()

produtos = ['FIBRA', 'NOVA FIBRA']
segmentos = ['VAREJO', 'EMPRESARIAL']
gestao = ['RSE', 'RCS', 'RNN', 'TLV', 'WEB', 'OUTROS NACIONAIS']

for produto in produtos:
    for segmento in segmentos:
        for gest in gestao:
            filtraDF_e_CalculaTendencia(df, 'VL', produto, segmento, gest)


executa_procedure_sql_combinada('SP_CDO_PREPARA_BASE_TEND_VL_VLL')

montaExcelTendVlVll()
enviaEmaileAnexo()
#demonstrativo_gross()
executa_procedure_sql_combinada('SP_CDO_TEND_VL_VLL_LEGADA_IGUAL_CDO')
#atualiza_TB_VALIDA_CARGA_TENDENCIA()


notificacao = Notification(app_id="Tendência HoltWinters", title="CONCLUIDO",
                           msg = "Etapas Concluidas",
                           duration = 'long',
                           icon=r"C://Users/oi066724/Documents/Python/selfie.png")
notificacao.set_audio(audio.LoopingAlarm, loop=False)
notificacao.show()