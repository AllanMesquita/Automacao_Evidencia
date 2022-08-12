"""
    Main Script
"""

import openpyxl as xl
from datetime import datetime
from Modulos import validacao, atualizar_v2
import os


now = datetime.now()

path = "C:\\Users\\allan.mesquita\\OneDrive - NTT\\Documents\\Projetos\\Automacao_Evidencias\\Arquivos Testes\\"
file_name = "Copy of EvidenciaExpedicaoJR RIO 01102021 - 2039 (002).xlsx"

wb = xl.open(path + file_name)
wb.active
sheet = wb.sheetnames
aba = wb[sheet[0]]

verif_evid = aba["A2"].value  # Verificador do tipo de evidência


# Variáveis

type_evidencia = ""
temp_val = ""
qtd_linhas = len(aba["A"])
resultado = ''


if bool(verif_evid) is False:
    print("Erro!")
    # Alguma ação mais apropriada
elif len(verif_evid) == 44:
    type_evidencia = "Recebimento"
    # verificador de repetição de caracteres
elif len(verif_evid) == 24 and "E" in verif_evid:
    type_evidencia = "Expedição"
elif len(verif_evid) != 44:
    print("Erro!")
    # Alguma ação mais apropriada


inicio = now
print(now)


if type_evidencia == "Recebimento":
    resultado = validacao.rec_validation(aba, qtd_linhas)
if type_evidencia == "Expedição":
    resultado = validacao.exp_validacao(aba, qtd_linhas)

wb.save(path + file_name)

if resultado == 'Erro nos dados':
    os.replace(path + file_name, "C:\\Users\\allan.mesquita\\OneDrive - NTT\\Documents\\Projetos\\Automacao_Evidencias\\Script\\" + file_name)
    # Verificar o caso se haver o mesmo arquivo no destino
elif resultado == 'Sucesso':
    atualizar_v2.popular_V17(aba, qtd_linhas, type_evidencia)

print(inicio)
print(f'Tempo total: {datetime.now() - now}')
