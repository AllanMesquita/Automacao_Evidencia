from dateutil.parser import parse
import pandas as pd
import openpyxl as xl
import psycopg2

arquivo = xl.open('C:\\Users\\allan.mesquita\\Downloads\\nf_valor.xlsx')
arquivo_sheets = arquivo.sheetnames
aba = arquivo[arquivo_sheets[0]]

linhas = len(aba['A'])
linha = 2
print(type(aba['B3'].value))
# print(parse(str(aba['G2'].value)))

# print(type(aba['A2'].value))
# while linha != linhas:
#     print(linha)
#     df = pd.read_excel("C:\\Users\\allan.mesquita\\Downloads\\2022 á 2027 - Nfs Entrada Mastersaf.xlsx", sheet_name='Listagem de NF-e Recebidas')
#     item = df.loc[df['Chave de Acesso'] == aba[f'A{linha}'].value]
#     data = item['Data e Hora da Emissão']
#     print(parse(str(data.at[item.index[0]])))
#     aba[f'G{linha}'] = parse(str(data.at[item.index[0]]))
#     linha += 1
#
# arquivo.save('C:\\Users\\allan.mesquita\\Downloads\\nfs_datas.xlsx')

# con = psycopg2.connect(
#     host="psql-itlatam-logisticcontrol.postgres.database.azure.com",
#     dbname="logistic-control",
#     user="logisticpsqladmin@psql-itlatam-logisticcontrol",
#     password="EsjHSrS69295NzHu342ap6P!N",
#     sslmode="require"
# )
#
# cur = con.cursor()
#
# while linha != linhas + 1:
#     print(linha)
#     cur.execute(f"UPDATE material_management.master_saf_entrada SET total_nfe = '{aba[f'B{linha}'].value}' WHERE chave_acesso = '{aba[f'A{linha}'].value}'")
#     con.commit()
#     linha += 1
#
# cur.close()
# con.close()
