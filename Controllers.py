import Métodos
import pandas as pd


def OutputTable(Tabela, Operacao, Quinzena, IDFatura):

    # Cria tudo vazio primeiro
    Rotas = pd.DataFrame()
    Penalidades = pd.DataFrame()
    Pedagio = pd.DataFrame()

    # FIRST MILE → Rotas + Penalidades + Pedágio
    if Operacao == 'First Mile':
        Rotas, Penalidades, Pedagio = Métodos.FormatacaoFirstMile(
            Tabela, Operacao, Quinzena, IDFatura
        )
        return Rotas, Penalidades, Pedagio

    # MELI ONE → Rotas + Penalidades
    elif Operacao == 'MeliOne':
        Rotas, Penalidades, _ = Métodos.FormatacaoMeliOne(
            Tabela, Operacao, Quinzena, IDFatura
        )
        Pedagio = pd.DataFrame()
        return Rotas, Penalidades, Pedagio

    # LINE HAUL → Apenas Rotas
    elif Operacao == 'Line Haul':
        Rotas, _, _ = Métodos.FormatacaoLineHaul(
            Tabela, Operacao, Quinzena, IDFatura
        )
        Penalidades = pd.DataFrame()
        Pedagio = pd.DataFrame()
        return Rotas, Penalidades, Pedagio

    # LAST MILE → Rotas + Penalidades
    elif Operacao == 'Last Mile':
        Rotas, Penalidades, _ = Métodos.formatarLastMile(
            Tabela, Operacao, Quinzena, IDFatura
        )
        Pedagio = pd.DataFrame()
        return Rotas, Penalidades, Pedagio

    else:
        print("❌ Erro: Operação inválida")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
