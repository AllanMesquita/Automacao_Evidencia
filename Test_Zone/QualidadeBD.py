import json

repeticao = []

data = open("C:\\Users\\allan.mesquita\\Downloads\\teste.json")

obj = json.load(data)

for item in obj:
    repeticao.append(item['RFID_Produto'])
    print(item)

if repeticao.count('E00000000000000000314981') > 1:
    print('Teste')
