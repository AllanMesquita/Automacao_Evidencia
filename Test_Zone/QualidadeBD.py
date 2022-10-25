import json
import Modulos.validacaoBD

repeticao_rfid = []
repeticao_serial = ['E00000000000000000314981']
erros = []
nfs = []
validacao = []

data = open("C:\\Users\\allan.mesquita\\Downloads\\teste.json")

obj = json.load(data)

for item in obj:
    if item['ChaveNF_Entrada'] not in nfs:
        nfs.append(item['ChaveNF_Entrada'])
    # if item['RFID_Produto'] not in repeticao:
    #     repeticao.append(item['RFID_Produto'])
    else:
        pass
    repeticao_rfid.append(item['RFID_Produto'])
    repeticao_serial.append(item['SerialNumber'])
    # print(item['ChaveNF_Entrada'])

for nf in nfs:
    for item in obj:
        if nf == item['ChaveNF_Entrada']:
            validacao.append(item)
        else:
            pass
    print(validacao)
    for linha in validacao:
        print(linha)
    'Chama código validação'
    Modulos.validacaoBD.rec_validation(validacao, repeticao_rfid, repeticao_serial)
    'processo se erro'
    validacao.clear()

# print(repeticao)
