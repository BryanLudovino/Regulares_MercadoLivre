import pandas as pd
import openpyxl
from pandasql import sqldf
import numpy as np


def formatacaoGeral(df_read):
     
    operacao = df_read.iat[0,3]
    quinzena = df_read.iat[0,5]
    idFatura = df_read.iat[0,4]

    df_read.columns = df_read.iloc[1]
    df = df_read.drop([0,1]).reset_index(drop=True)

    tabela = df.dropna(subset=['Descrição'])

    indices_del = tabela[tabela['Descrição'].str.contains('total', case=False)].index
    tabela = tabela.drop(indices_del)

    return tabela, operacao, quinzena, idFatura


def FormatacaoFirstMile(Dados, Operação, Periodo, IDFatura):

    Ajudantes = Dados[Dados['Descrição'].str.contains('Helpers')].copy()
    Ajudantes = Ajudantes[['ID da rota','Total']]
    Ajudantes = Ajudantes.rename(columns={'Total':'Ajudante'})
    Ajudantes['Ajudante'] = Ajudantes[['Ajudante']].astype(float)

    Penalidades = Dados[Dados['Descrição'].str.contains('Penalty')].copy()
    if Penalidades.empty:
        Penalidades = pd.DataFrame()
    else:
        Penalidades = Penalidades[['Descrição','ID da rota','Placa','Motorista','Total']]
        Penalidades[['Descrição', 'Descricao_2']] = Penalidades['Descrição'].str.split(': ', expand=True, n=1)
        Penalidades[['Placa','date','PacoteID']] = Penalidades['Descricao_2'].str.split(' ',expand=True,n=2)
        Penalidades = Penalidades[['Descrição','ID da rota','PacoteID','Placa','Motorista','Total']]
        for coluna in ['ID da rota','PacoteID']:
            Penalidades[coluna] = pd.to_numeric(Penalidades[coluna], errors='coerce')

    Pedagio = Dados[Dados['Descrição'].str.contains('Peajes')].copy()
    Pedagio = Pedagio[['ID da rota','Data de início','Data de término','Total']]
    Pedagio['ID da rota'] = pd.to_numeric(Pedagio['ID da rota'], errors='coerce')

    RotasFiltrar = Dados[Dados['Descrição'].str.contains('SVC', na=False)].copy()
    Rotas = RotasFiltrar[~RotasFiltrar['Descrição'].str.contains('Helpers', na=False)].copy()

    Rotas['Ambulance'] = Rotas['Descrição'].str.contains('AMBULANCE', na=False).map({True:'Yes', False:'No'})
    Rotas['holiday'] = Rotas['Descrição'].str.contains('HOLIDAY', na=False).map({True:'Yes', False:'No'})
    Rotas['part of time'] = Rotas['Descrição'].str.contains('TIME', na=False).map({True:'Yes', False:'No'})

    Rotas['Operação'] = Operação
    Rotas['Quinzena'] = Periodo
    Rotas['IDFatura'] = IDFatura

    Rotas[['Descricao_Base','SVC_Info']] = Rotas['Descrição'].str.split(' - SVC: ', expand=True, n=1)
    Rotas[['SVC_Parte','KMs_Range']] = Rotas['SVC_Info'].str.split(' - KMs RANGE:', expand=True, n=1)
    Rotas[['SVC_Parte_1','SVC_Parte_2']] = Rotas['SVC_Parte'].str.split(' - ', expand=True, n=1)

    Rotas = Rotas.drop(columns=['SVC_Info','SVC_Parte','Descrição','Total'])
    Rotas = Rotas.rename(columns={
        'Descricao_Base':'Descrição',
        'KMs_Range':'KM',
        'SVC_Parte_1':'Service',
        'SVC_Parte_2':'Obs',
        'Preço unitário':'Total'
    })

    Rotas = Rotas[['IDFatura','ID da rota','Data de início','Data de término',
                   'Operação','Quinzena','Descrição','Service','Placa',
                   'Motorista','KM','Total','Ambulance','part of time','holiday']]

    Rotas['KM'] = Rotas['KM'].replace("None","").fillna("").replace("","0/0")

    Rotas = pd.merge(Rotas, Ajudantes, how='left', on='ID da rota').fillna(0)

    for col in ['Data de início','Data de término']:
        Rotas[col] = pd.to_datetime(Rotas[col], format='%d/%m/%Y')

    for col in ['IDFatura','ID da rota','Total']:
        Rotas[col] = pd.to_numeric(Rotas[col], errors='coerce')

    return Rotas, Penalidades, Pedagio


def FormatacaoMeliOne(Dados, Operação, Periodo, IDFatura):

    Penalidades = Dados[Dados['Descrição'].str.contains('Penalty')].copy()
    if Penalidades.empty:
        Penalidades = pd.DataFrame()
    else:
        Penalidades = Penalidades[['Descrição','ID da rota','Placa','Motorista','Total']]
        Penalidades['Quinzena de Faturamento'] = Periodo
        Penalidades[['Descrição','Descricao_2']] = Penalidades['Descrição'].str.split(': ', expand=True, n=1)
        Penalidades[['Placa','date','ID Pacote']] = Penalidades['Descricao_2'].str.split(' ', expand=True, n=2)
        Penalidades = Penalidades[['Descrição','ID da rota','Quinzena de Faturamento',
                                   'ID Pacote','Placa','Motorista','Total']]
        for col in ['ID da rota','ID Pacote']:
            Penalidades[col] = pd.to_numeric(Penalidades[col], errors='coerce')

    Rotas = Dados[Dados['Descrição'].str.contains('SVC', na=False)].copy()
    Rotas = Rotas[~Rotas['Descrição'].str.contains('Helpers', na=False)]

    Rotas['Ambulance'] = Rotas['Descrição'].str.contains('AMBULANCE', na=False).map({True:'Yes', False:'No'})
    Rotas['holiday'] = Rotas['Descrição'].str.contains('HOLIDAY', na=False).map({True:'Yes', False:'No'})
    Rotas['part of time'] = Rotas['Descrição'].str.contains('TIME', na=False).map({True:'Yes', False:'No'})

    Rotas['Operação'] = Operação
    Rotas['Quinzena'] = Periodo
    Rotas['IDFatura'] = IDFatura

    Rotas[['Descricao_Base','Service']] = Rotas['Descrição'].str.split(' - SVC: ', expand=True, n=1)
    Rotas = Rotas.drop(columns=['Descrição','Total'])
    Rotas = Rotas.rename(columns={'Descricao_Base':'Descrição','Preço unitário':'Total'})

    Rotas = Rotas[['IDFatura','ID da rota','Data de início','Data de término',
                   'Operação','Quinzena','Descrição','Service','Placa',
                   'Motorista','Total','Ambulance','part of time','holiday']]

    for col in ['Data de início','Data de término']:
        Rotas[col] = pd.to_datetime(Rotas[col], format='%d/%m/%Y')

    for col in ['IDFatura','ID da rota','Total']:
        Rotas[col] = pd.to_numeric(Rotas[col], errors='coerce')

    return Rotas, Penalidades, pd.DataFrame()


def formatarLastMile(Dados, Operacao, Periodo, IDFatura):

    # ===== PARADAS =====
    Paradas = Dados[Dados['Descrição'].str.contains('addresses')].copy()
    Paradas = Paradas[['ID da rota', 'Quantidade', 'Total']]
    Paradas = Paradas.astype({'Total': float, 'Quantidade': int})
    Paradas = Paradas.rename(columns={
        'Quantidade': 'QTD Paradas',
        'Total': 'R$ Valor'
    })

    # ===== PENALIDADES =====
    Penalidades = Dados[Dados['Descrição'].str.contains('Penalty')].copy()
    if Penalidades.empty:
        Penalidades = pd.DataFrame()
    else:
        Penalidades = Penalidades[['Descrição','ID da rota','Placa','Motorista','Total']]
        Penalidades['ID Pacote'] = Penalidades['Descrição'].str[-11:]

        Penalidades[['Descrição','Descricao_2']] = Penalidades['Descrição'].str.split(': ', expand=True, n=1)
        Penalidades = Penalidades.drop(columns=['Descricao_2'])
        Penalidades['Descrição'] = Penalidades['Descrição'].str.replace(' Penalty', "", regex=False)

        naoVisitados = Dados[Dados['Descrição'].str.contains('Not Visited')].copy()
        naoVisitados = naoVisitados[['Descrição','ID da rota','Placa','Motorista','Total']]
        naoVisitados = naoVisitados.drop(columns=['Descrição'])
        naoVisitados['Descrição'] = "Not Visited"

        Penalidades = pd.concat([Penalidades, naoVisitados], ignore_index=True)

        Penalidades['ID Pacote'] = Penalidades[['ID Pacote']].fillna("")
        for col in ['ID da rota','ID Pacote']:
            Penalidades[col] = pd.to_numeric(Penalidades[col], errors='coerce')

    # ===== ROTAS =====
    Rotas = Dados[Dados['Descrição'].str.contains('SVC', na=False)].copy()
    Rotas = Rotas[~Rotas['Descrição'].str.contains('Helpers', na=False)]

    Rotas['Ambulance'] = Rotas['Descrição'].str.contains('AMBULANCE', na=False).map({True:'Yes', False:'No'})
    Rotas['holiday'] = Rotas['Descrição'].str.contains('HOLIDAY', na=False).map({True:'Yes', False:'No'})
    Rotas['part of time'] = Rotas['Descrição'].str.contains('TIME', na=False).map({True:'Yes', False:'No'})

    Rotas['Operação'] = Operacao
    Rotas['Quinzena'] = Periodo
    Rotas['IDFatura'] = IDFatura

    Rotas[['Descricao_Base','SVC_Info']] = Rotas['Descrição'].str.split(' - SVC: ', expand=True, n=1)
    Rotas[['SVC_Parte','KMs_Range']] = Rotas['SVC_Info'].str.split(' - KMs RANGE:', expand=True, n=1)
    Rotas[['SVC_Parte_1','SVC_Parte_2']] = Rotas['SVC_Parte'].str.split(' - ', expand=True, n=1)

    Rotas = Rotas.drop(columns=['SVC_Info','SVC_Parte'])
    Rotas = Rotas.rename(columns={
        'Descricao_Base':'Descriçao',
        'KMs_Range':'KM',
        'SVC_Parte_1':'Service',
        'SVC_Parte_2':'Obs'
    })

    Rotas = Rotas[['IDFatura','ID da rota','Data de início','Data de término',
                   'Operação','Quinzena','Descriçao','Service','Placa',
                   'Motorista','KM','Total com Tax','Total com ISS','Total com ICMS','Total','Ambulance','part of time','holiday']]

    # ===== AJUSTES =====
    Rotas['KM'] = Rotas['KM'].replace("None","").fillna("")
    Rotas['KM'] = Rotas['KM'].astype(str).str.split(' - ', n=1).str[0]
    Rotas['KM'] = Rotas['KM'].replace("", "0/0")

    Rotas['Placa'] = Rotas['Placa'].str.replace('SDD-', '', regex=False)

    Rotas = pd.merge(Rotas, Paradas, how='left', on='ID da rota').fillna(0)

    for col in ['Data de início','Data de término']:
        Rotas[col] = pd.to_datetime(Rotas[col], format='%d/%m/%Y')

    for col in ['IDFatura','ID da rota']:
        Rotas[col] = pd.to_numeric(Rotas[col], errors='coerce')

    for col in ['Total com Tax','Total com ISS','Total com ICMS','Total','QTD Paradas','R$ Valor']:
        Rotas[col] = pd.to_numeric(Rotas[col], errors='coerce')

    return Rotas, Penalidades, pd.DataFrame()

def ConciliarLineHaul(Rotas, Tabela1):

    consultasql = """
    SELECT *
    FROM Rotas
    LEFT JOIN Tabela1
        ON TRIM(Rotas."Descrição") = TRIM(Tabela1."Perfil")
        AND TRIM(Rotas."Service") = TRIM(Tabela1."Service")
    """

    resultado = sqldf(consultasql, locals())

    resultado = resultado.drop(columns=['Perfil'], errors='ignore')

    for col in ['Data de início','Data de término']:
        resultado[col] = pd.to_datetime(resultado[col])

    resultado['Total'] = pd.to_numeric(resultado['Total'], errors='coerce')

    return resultado


def ConciliarMeliOne(Rotas, Tabela1):

    consultasql = """
    SELECT *
    FROM Rotas
    LEFT JOIN Tabela1
        ON Rotas."Descrição" = Tabela1."Perfil"
    """

    resultado = sqldf(consultasql, locals())

    resultado = resultado.drop(columns=['Faixa de Km','Perfil'], errors='ignore')

    for col in ['Data de início','Data de término']:
        resultado[col] = pd.to_datetime(resultado[col])

    resultado['Total'] = pd.to_numeric(resultado['Total'], errors='coerce')

    return resultado


def ConciliarFirstMile(Rotas, Tabela1, Tabela2, Regiao):

    consultasql = """
    SELECT
        tabelacompleta.*,
        COALESCE(R1."Valor Frete", R2."Valor Frete", 'Não Registrado') AS "Valor Tabela"
    FROM
    (
        SELECT *
        FROM Rotas AS RT
        LEFT JOIN Regiao
            ON TRIM(RT."Service") = TRIM(Regiao."base")
    ) AS tabelacompleta

    LEFT JOIN Tabela1 AS R1
        ON TRIM(tabelacompleta."Descrição") = TRIM(R1."Perfil")
        AND TRIM(tabelacompleta."KM") = TRIM(R1."Faixa de Km")

    LEFT JOIN Tabela2 AS R2
        ON TRIM(tabelacompleta."Descrição") = TRIM(R2."Perfil")
        AND TRIM(tabelacompleta."KM") = TRIM(R2."KM")
        AND TRIM(tabelacompleta."Estado") = TRIM(R2."UF")
    """

    resultado = sqldf(consultasql, locals())

    for col in ['Data de início','Data de término']:
        resultado[col] = pd.to_datetime(resultado[col])

    resultado = resultado.drop(columns=['Estado','base'], errors='ignore')

    resultado['Total'] = pd.to_numeric(resultado['Total'], errors='coerce')

    return resultado


def ConciliarLastMile(Rotas, Tabela1, Tabela2, TabelaParada):

    consultasql = """
    SELECT *
    FROM (
        SELECT
            R.*,
            COALESCE(T1."Valor Frete", T2."Valor Frete", 'Não Registrado') AS "Valor Frete Final"
        FROM Rotas AS R
        
        LEFT JOIN Tabela1 AS T1
            ON R."Descrição" = T1.Perfil
            AND R.KM = T1."Faixa de Km"

        LEFT JOIN Tabela2 AS T2
            ON R."Descrição" = T2.Perfil
            AND R.KM = T2."Faixa de Km"
            AND R.Service = T2.Base
    ) AS TableConc

    LEFT JOIN TabelaParada
        ON TRIM(TableConc.Service) = TRIM(TabelaParada.SVC)
    """

    resultado = sqldf(consultasql, locals())

    colunas_calculo = ['QTD Paradas', '0 - 60', '61 - 90', '>91']
    resultado[colunas_calculo] = resultado[colunas_calculo].fillna(0)

    condicoes = [
        resultado['QTD Paradas'] <= 60,
        (resultado['QTD Paradas'] > 60) & (resultado['QTD Paradas'] <= 90)
    ]

    valores = [
        resultado['0 - 60'] * resultado['QTD Paradas'],
        (resultado['0 - 60'] * 60) + resultado['61 - 90'] * (resultado['QTD Paradas'] - 60)
    ]

    valor_default = (
        (resultado['0 - 60'] * 60) +
        (resultado['61 - 90'] * 30) +
        (resultado['QTD Paradas'] - 90) * resultado['>91']
    )

    resultado['Tabela Paradas'] = np.select(
        condicoes,
        valores,
        default=valor_default
    )

    resultado = resultado.drop(
        columns=['SVC','Região Meli','0 - 60','61 - 90','>91'],
        errors='ignore'
    )

    for col in ['Data de início','Data de término']:
        resultado[col] = pd.to_datetime(resultado[col])

    return resultado


def LoadDataLastMile():
    Tabela1 = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Table 1'
    )

    Tabela2 = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Table 2'
    )

    TabelaParada = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Payment Stops Table'
    )

    return Tabela1, Tabela2, TabelaParada


def LoadDataMeliOne():
    Tabela1 = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Table 4'
    )

    return Tabela1

def LoadDataFirstMile():
    Tabela1 = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Table 5'
    )

    Tabela2 = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Table 6'
    )

    Regiao = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Regiao First'
    )

    return Tabela1, Tabela2, Regiao


def LoadDataLineHaul():
    Tabela1 = pd.read_excel(
        'bases/Tabela Frete python.xlsx',
        sheet_name='Table 3'
    )

    return Tabela1
