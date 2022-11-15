import psycopg2
import json
from Modulos.class_pesquisa_v2_1 import Pesquisa as pesquisa
from dateutil.parser import parse
import openpyxl

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

# data = open("C:\\Users\\allan.mesquita\\Downloads\\teste.json")
#
# req_body = json.load(data)

wb = openpyxl.load_workbook("C:\\Users\\allan.mesquita\\Downloads\\2022 á 2027 - NFs Saída Mastersaf.xlsx")
sheet = wb.sheetnames
aba = wb[sheet[1]]
linhas = len(aba['A'])

linha = 19655

dicionario = {}
dicionario_bd = {}

error_bd = False

while linha != linhas + 1:
    chave_acesso = aba[f'R{linha}'].value
    quant_com = aba[f'AC{linha}'].value if bool(aba[f'AC{linha}'].value) else '00'

    if chave_acesso not in dicionario:
        dicionario[chave_acesso] = int(str(quant_com.split(',')[0]).replace(',', '').replace('.', ''))
    else:
        dicionario[chave_acesso] += int(str(quant_com.split(',')[0]).replace(',', '').replace('.', ''))

    linha += 1

print(dicionario)

for chave in dicionario.items():
    quantidade_banco = 0

    chave_acesso = chave[0]

    # PESQUISA NOTA NO BANCO
    cursor.execute(f"SELECT quantidade_com FROM material_management.master_saf_saida_itens WHERE chave_acesso = '{chave_acesso}'")
    resultado = cursor.fetchall()
    if bool(resultado):
        for dado in resultado:
            for quant in dado:
                quantidade_banco += quant
        print('quant banco: ', quantidade_banco)
    if quantidade_banco == 0:
        dicionario_bd[chave_acesso] = 'naoEncontrado'
    elif chave[1] == quantidade_banco:
        dicionario_bd[chave_acesso] = 'correto'
    elif chave[1] != quantidade_banco:
        dicionario_bd[chave_acesso] = 'divergente'

print(dicionario_bd)

linha = 19655

while linha != linhas + 1:

    chave_acesso = aba[f'R{linha}'].value
    print(chave_acesso)

    if dicionario_bd[chave_acesso] == 'correto':
        print('pulou.')
        linha += 1
        continue
    elif dicionario_bd[chave_acesso] == 'divergente':
        print('Texto para e-mail')
        linha += 1
        continue
    elif dicionario_bd[chave_acesso] == 'naoEncontrado':
        print('Colocar banco')

        data_emissao = aba[f'M{linha}'].value
        descricao_produto = aba[f'V{linha}'].value
        cod_produto = aba[f'W{linha}'].value
        numero_pedido = aba[f'X{linha}'].value
        cean = aba[f'Y{linha}'].value
        cean_trib = aba[f'Z{linha}'].value
        unid_com = aba[f'AA{linha}'].value
        valor_unitario = aba[f'AB{linha}'].value
        quantidade = aba[f'AC{linha}'].value if bool(aba[f'AC{linha}'].value) else '00'
        valor_total = aba[f'AD{linha}'].value
        origem = aba[f'AE{linha}'].value
        base_icms = aba[f'AF{linha}'].value
        cst_icms = aba[f'AG{linha}'].value if bool(aba[f'AG{linha}'].value) else '00'
        aliq_icms = aba[f'AH{linha}'].value
        valor_icms = aba[f'AI{linha}'].value
        perc_icms = aba[f'AJ{linha}'].value
        base_icms_st = aba[f'AK{linha}'].value
        valor_icms_st = aba[f'AL{linha}'].value
        aliq_icms_st = aba[f'AM{linha}'].value
        valor_pis = aba[f'AN{linha}'].value
        cst_pis = aba[f'AO{linha}'].value
        valor_cofins = aba[f'AP{linha}'].value
        cst_cofins = aba[f'AQ{linha}'].value
        valor_ipi = aba[f'AR{linha}'].value
        cst_ipi = aba[f'AS{linha}'].value
        aliq_ipi = aba[f'AT{linha}'].value
        ncm = aba[f'AU{linha}'].value
        cfop = aba[f'AV{linha}'].value

        # PESQUISAS
        # CFOP
        teste = pesquisa(cursor, conn, chave_acesso, aba, linha)
        id_cfop = teste.cfop_saida()
        print(id_cfop)

        cursor.execute(
            'INSERT INTO material_management.master_saf_saida_itens ('
            'chave_acesso, data_emissao, descricao_produto, cod_produto, numero_pedido, cean, cean_trib, unid_com, '
            'valor_unitario, quantidade_com, valor_total, origem, base_calculo_icms, cst_icms_csosn, aliq_icms, valor_icms,'
            'perc_margem_icms_st, base_calc_icms_st, valor_icms_st, aliq_icms_st, valor_pis, cst_pis, valor_cofins, '
            'cst_cofins, valor_ipi, cst_ipi, aliq_ipi, ncm, cfop'
            ')'
            'VALUES ('
            '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, '
            '%s, %s'
            ')',
            (
                chave_acesso, parse(str(data_emissao)), descricao_produto, cod_produto, numero_pedido, cean, cean_trib, unid_com,
                valor_unitario.replace('.', '').replace(',', '.'), quantidade.split(',')[0].replace('.', ''),
                valor_total.replace('.', '').replace(',', '.'), origem, base_icms.replace('.', '').replace(',', '.'),
                cst_icms.replace('.', '').replace(',', '.'), aliq_icms.replace('.', '').replace(',', '.'),
                valor_icms.replace('.', '').replace(',', '.'), perc_icms.replace('.', '').replace(',', '.'),
                base_icms_st.replace('.', '').replace(',', '.'), valor_icms_st.replace('.', '').replace(',', '.'),
                aliq_icms_st.replace('.', '').replace(',', '.'), valor_pis.replace('.', '').replace(',', '.'),
                cst_pis.replace('.', '').replace(',', '.'), valor_cofins.replace('.', '').replace(',', '.'),
                cst_cofins.replace('.', '').replace(',', '.'), valor_ipi.replace('.', '').replace(',', '.'),
                cst_ipi.replace('.', '').replace(',', '.'), aliq_ipi.replace('.', '').replace(',', '.'), ncm, id_cfop
            )
        )
        conn.commit()

    linha += 1

cursor.close()
conn.close()
