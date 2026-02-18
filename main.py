import Métodos
import pandas as pd
import Controllers
import re


try:
    df_read = pd.read_csv("C:/Users/bryan.souza/Downloads/Detalhes da pré-fatura #4960323 (1).csv", sep=";")
    print("Iniciando leitura do arquivo...")
except:
    print("❌ Falha na leitura do arquivo!")
    exit()


Tabela, Operacao, Quinzena, IDFatura = Métodos.formatacaoGeral(df_read)

Viagem, Penalidades, Pedagio = Controllers.OutputTable(
    Tabela, Operacao, Quinzena, IDFatura
)






nome = f"Regular_{Operacao}_Quinzena_{Quinzena}_ID_{IDFatura}"
nome = re.sub(r'[\\\\/*?:\"<>|]', "_", nome)
nome_planilha = nome + ".xlsx"

print("Arquivo será salvo como:", nome_planilha)

with pd.ExcelWriter(nome_planilha) as writer:

    if not Viagem.empty:
        Viagem.to_excel(writer, sheet_name="Rotas", index=False)

    if not Penalidades.empty:
        Penalidades.to_excel(writer, sheet_name="Descontos", index=False)

    if not Pedagio.empty:
        Pedagio.to_excel(writer, sheet_name="Pedágio", index=False)
