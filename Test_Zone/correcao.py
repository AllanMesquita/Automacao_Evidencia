import psycopg2
import json


# Update connection string information
host = "psql-itlatam-logisticcontrol.postgres.database.azure.com"
dbname = "logistic-control"
user = "logisticpsqladmin@psql-itlatam-logisticcontrol"
password = "EsjHSrS69295NzHu342ap6P!N"
sslmode = "require"
# Construct connection string
conn_string = "host={0} user={1} dbname={2} password={3} sslmode={4}".format(host, user, dbname, password,
                                                                             sslmode)
conn = psycopg2.connect(conn_string)
print("Connection established")
cursor = conn.cursor()

with open("C:\\Users\\allan.mesquita\\Downloads\\correcao.json", encoding="utf8") as data:
    reader = json.load(data)

cnpj = ''

for item in reader:
    # print(item)

    if item['CNPJ do Destinatário'] == '00000000000':
        if item['Razão Social Destinatário'] == 'CISCO SYSTEMS, INC.':
            # print(item)
            cnpj = '402'
            # print(cnpj)
            cursor.execute(f"UPDATE material_management.master_saf_saida SET destinatario = '{cnpj}' WHERE chave_acesso = '{item['Chave de Acesso']}'")
            conn.commit()
        else:
            cursor.execute(
                f"SELECT cnpj FROM material_management.dados_juridicos WHERE razao_social = '{item['Razão Social Destinatário']}'")
            resultsado = cursor.fetchall()
            for dado in resultsado:
                cnpj = dado[0]
            cursor.execute(f"UPDATE material_management.master_saf_saida SET destinatario = '{cnpj}' WHERE chave_acesso = '{item['Chave de Acesso']}'")
            conn.commit()
    else:
        print('Não', item['ID_BD'])





cursor.close()
conn.close()
