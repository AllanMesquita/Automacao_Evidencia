import json
import Modulos.validacaoBD

tipo_evidencia = ''
repeticao_rfid = []
repeticao_serial = []
erros = []
nfs = []
validacao = []
resultado = ''

data = open("C:\\Users\\allan.mesquita\\Downloads\\teste.json")

obj = json.load(data)

for item in obj:
    if len(item) > 9:
        tipo_evidencia = 'Recebimento'
    else:
        tipo_evidencia = 'Expedição'

if tipo_evidencia == 'Recebimento':
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
        resultado = Modulos.validacaoBD.rec_validation(validacao, repeticao_rfid, repeticao_serial)
        'processo se erro'
        print(bool(resultado))
        if bool(resultado) is True:
            erros.append(resultado)
        validacao.clear()
    # resultado = Modulos.validacaoBD.rec_validation(obj, repeticao_rfid, repeticao_serial)

if tipo_evidencia == 'Expedição':
    for item in obj:
        if item['ChaveNF_Saida'] not in nfs:
            nfs.append(item['ChaveNF_Saida'])
        else:
            pass
        repeticao_rfid.append(item['RFID_Produto'])

    for nf in nfs:
        for item in obj:
            if nf == item['ChaveNF_Saida']:
                validacao.append(item)
            else:
                pass
        resultado = Modulos.validacaoBD.exp_validation(validacao, repeticao_rfid)
        validacao.clear()

# erros.append(resultado)
print(erros)
# print(repeticao)
