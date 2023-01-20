import json

with open("C:\\Users\\allan.mesquita\\Downloads\\eventos.json", encoding="utf8") as data:
    reader = json.load(data)

for item in reader:
    print(item['CNPJ do Destinatário'])
