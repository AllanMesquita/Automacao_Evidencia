import json

repeticao = []
erros = []
validacao = []

data = open("C:\\Users\\allan.mesquita\\Downloads\\teste.json")

obj = json.load(data)

for item in obj:
    if item['ChaveNF_Entrada'] not in repeticao:
        repeticao.append(item['ChaveNF_Entrada'])
    else:
        pass
    print(item['ChaveNF_Entrada'])

for nf in repeticao:
    for item in obj:
        if nf == item['ChaveNF_Entrada']:
            validacao.append(item)
        else:
            pass
    print(validacao)
    for linha in validacao:
        print(linha)
    'Chama código validação'
    'processo se erro'
    validacao.clear()

print(repeticao)
