import psycopg2
# import datetime
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

wb = openpyxl.load_workbook("C:\\Users\\allan.mesquita\\Downloads\\2022 á 2027 - Nfs Entrada Mastersaf.xlsx")
sheet = wb.sheetnames
aba = wb[sheet[0]]
linhas = len(aba['A'])

linha = 2

while linha != 4:
    # print(dic)

    numero_nfe = aba[f'A{linha}'].value
    serie = aba[f'B{linha}'].value
    tipo_nfe = aba[f'C{linha}'].value if bool(aba[f'C{linha}'].value) else 99
    destinatario = aba[f'D{linha}'].value
    fornecedor = aba[f'L{linha}'].value
    natureza_cfop = aba[f'T{linha}'].value
    chave_acesso = aba[f'V{linha}'].value
    situacao = aba[f'W{linha}'].value
    descricao_retorno = aba[f'X{linha}'].value
    data_hora_emissao = aba[f'Y{linha}'].value if bool(aba[f'Y{linha}'].value) else "01/01/2001 00:00:01"
    data_hora_saida = aba[f'Z{linha}'].value if bool(aba[f'Z{linha}'].value) else "01/01/2001 00:00:01"
    data_hora_autorizacao = aba[f'AA{linha}'].value if bool(aba[f'AA{linha}'].value) else "01/01/2001 00:00:01"
    protocolo_autorizacao = aba[f'AB{linha}'].value
    data_hora_cancelamento = aba[f'AC{linha}'].value if bool(aba[f'AC{linha}'].value) else "01/01/2001 00:00:01"
    protocolo_cancelamento = aba[f'AD{linha}'].value
    motivo_cancelamento = aba[f'AE{linha}'].value
    tipo_manifestacao = aba[f'AF{linha}'].value
    data_hora_manifestacao = aba[f'AG{linha}'].value if bool(aba[f'AG{linha}'].value) else "01/01/2001 00:00:01"
    transportadora = aba[f'AH{linha}'].value
    tipo_frete = aba[f'AM{linha}'].value
    codigo_antt = aba[f'AN{linha}'].value
    quantidade = aba[f'AO{linha}'].value if bool(aba[f'AO{linha}'].value) else '00'
    especie = aba[f'AP{linha}'].value
    marca = aba[f'AQ{linha}'].value
    numeracao_volume = aba[f'AR{linha}'].value if bool(aba[f'AR{linha}'].value) else '00'
    peso_liquido = aba[f'AS{linha}'].value if bool(aba[f'AS{linha}'].value) else '00'
    peso_bruto = aba[f'AT{linha}'].value if bool(aba[f'AT{linha}'].value) else '00'
    base_icms = aba[f'AU{linha}'].value if bool(aba[f'AU{linha}'].value) else '00'
    total_icms = aba[f'AV{linha}'].value if bool(aba[f'AV{linha}'].value) else '00'
    total_icms_deson = aba[f'AW{linha}'].value if bool(aba[f'AW{linha}'].value) else '00'
    total_fcp = aba[f'AX{linha}'].value if bool(aba[f'AX{linha}'].value) else '00'
    total_fcp_uf_dest = aba[f'AY{linha}'].value if bool(aba[f'AY{linha}'].value) else '00'
    total_icms_dest = aba[f'AZ{linha}'].value if bool(aba[f'AZ{linha}'].value) else '00'
    total_icms_uf_remet = aba[f'BA{linha}'].value if bool(aba[f'BA{linha}'].value) else '00'
    base_icms_st = aba[f'BB{linha}'].value if bool(aba[f'BB{linha}'].value) else '00'
    total_icms_st = aba[f'BC{linha}'].value if bool(aba[f'BC{linha}'].value) else '00'
    total_fcp_st = aba[f'BD{linha}'].value if bool(aba[f'BD{linha}'].value) else '00'
    total_fcp_st_ret = aba[f'BE{linha}'].value if bool(aba[f'BE{linha}'].value) else '00'
    total_produtos_servicos = aba[f'BF{linha}'].value if bool(aba[f'BF{linha}'].value) else '00'
    total_frete = aba[f'BG{linha}'].value if bool(aba[f'BG{linha}'].value) else '00'
    total_seguro = aba[f'BH{linha}'].value if bool(aba[f'BH{linha}'].value) else '00'
    total_desconto = aba[f'BI{linha}'].value if bool(aba[f'BI{linha}'].value) else '00'
    total_ii = aba[f'BJ{linha}'].value if bool(aba[f'BJ{linha}'].value) else '00'
    total_ipi = aba[f'BK{linha}'].value if bool(aba[f'BK{linha}'].value) else '00'
    total_ipi_devolvido = aba[f'BL{linha}'].value if bool(aba[f'BL{linha}'].value) else '00'
    total_pis = aba[f'BM{linha}'].value if bool(aba[f'BM{linha}'].value) else '00'
    total_cofins = aba[f'BN{linha}'].value if bool(aba[f'BN{linha}'].value) else '00'
    total_outras_despesas = aba[f'BO{linha}'].value if bool(aba[f'BO{linha}'].value) else '00'
    total_nfe = aba[f'BP{linha}'].value if bool(aba[f'BP{linha}'].value) else '00'
    vl_aprox_tot_trib = aba[f'BQ{linha}'].value if bool(aba[f'BQ{linha}'].value) else '00'
    placa = aba[f'BR{linha}'].value
    uf = aba[f'BS{linha}'].value
    informacoes_fisco = aba[f'BT{linha}'].value
    informacoes_complementares = aba[f'BU{linha}'].value
    data_consulta = aba[f'BW{linha}'].value if bool(aba[f'BW{linha}'].value) else '01/01/2001 00:00:01'

    print(chave_acesso)
    # print(numero_nfe, '\n',
    #       serie, '\n',
    #       tipo_nfe, '\n',
    #       destinatario, '\n',
    #       fornecedor, '\n',
    #       natureza_cfop, '\n',
    #       chave_acesso, '\n',
    #       situacao, '\n',
    #       descricao_retorno, '\n',
    #       data_hora_emissao, '\n',
    #       data_hora_saida, '\n',
    #       data_hora_autorizacao, '\n',
    #       protocolo_autorizacao, '\n',
    #       data_hora_cancelamento, '\n',
    #       protocolo_cancelamento, '\n',
    #       motivo_cancelamento, '\n',
    #       tipo_manifestacao, '\n',
    #       data_hora_manifestacao, '\n',
    #       transportadora, '\n',
    #       tipo_frete, '\n',
    #       codigo_antt, '\n',
    #       quantidade, '\n',
    #       especie, '\n',
    #       marca, '\n',
    #       numeracao_volume, '\n',
    #       peso_liquido, '\n',
    #       peso_bruto, '\n',
    #       base_icms, '\n',
    #       total_icms, '\n',
    #       total_icms_deson, '\n',
    #       total_fcp, '\n',
    #       total_fcp_uf_dest, '\n',
    #       total_icms_dest, '\n',
    #       total_icms_uf_remet, '\n',
    #       base_icms_st, '\n',
    #       total_icms_st, '\n',
    #       total_fcp_st, '\n',
    #       total_fcp_st_ret, '\n',
    #       total_produtos_servicos, '\n',
    #       total_frete, '\n',
    #       total_seguro, '\n',
    #       total_desconto, '\n',
    #       total_ii, '\n',
    #       total_ipi, '\n',
    #       total_ipi_devolvido, '\n',
    #       total_pis, '\n',
    #       total_cofins, '\n',
    #       total_outras_despesas, '\n',
    #       total_nfe, '\n',
    #       vl_aprox_tot_trib, '\n',
    #       placa, '\n',
    #       uf, '\n',
    #       informacoes_fisco, '\n',
    #       informacoes_complementares, '\n',
    #       data_consulta
    #       )

    # PESQUISA NOTA NO BANCO
    cursor.execute(f"SELECT chave_acesso FROM material_management.master_saf_entrada WHERE chave_acesso = '{chave_acesso}'")
    resultado = cursor.fetchall()
    if bool(resultado):
        print('Já no banco')
        linha += 1
        continue
    else:
    # PESQUISAS
        # Destinatário
        teste = pesquisa(cursor, conn, destinatario, aba, linha)
        p_destinatario = teste.destinatario()
        # Fornecedor
        teste = pesquisa(cursor, conn, fornecedor, aba, linha)
        p_fornecedor = teste.fornecedor()
        # Natura e CFOP
        teste = pesquisa(cursor, conn, natureza_cfop, aba, linha)
        p_natureza = teste.natureza()
        # Transportadora
        teste = pesquisa(cursor, conn, transportadora, aba, linha)
        p_transportadora = teste.transportadora()

    print(p_destinatario, p_fornecedor, p_natureza, p_transportadora)

    cursor.execute(
        'INSERT INTO material_management.master_saf_entrada ('
        'numero_nfe, serie, tipo_nfe, destinatario, fornecedor, natureza_cfop, chave_acesso, situacao, '
        'descricao_retorno, data_hora_emissao, data_hora_saida, data_hora_autorizacao, protocolo_autorizacao, '
        'data_hora_cancelamento, protocolo_cancelamento, motivo_cancelamento, tipo_manifestacao, data_hora_manifestacao,'
        'transportadora, tipo_frete, codigo_antt, quantidade, especie, marca, numeracao_volume, peso_liquido, '
        'peso_bruto, base_icms, total_icms, total_icms_deson, total_fcp, total_fcp_uf_dest, total_icms_dest, '
        'total_icms_uf_remet, base_icms_st, total_icms_st, total_fcp_st, total_fcp_st_ret, total_produtos_servicos,'
        'total_frete, total_seguro, total_desconto, total_ii, total_ipi, total_ipi_devolvido, total_pis, total_cofins,'
        'total_outras_despesas, total_nfe, vl_aprox_tot_trib, placa, uf, informacoes_fisco, informacoes_complementares,'
        'data_consulta'
        ')'
        'VALUES ('
        '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, '
        '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s'
        ')',
        (
            numero_nfe, serie, tipo_nfe, p_destinatario, p_fornecedor, p_natureza, chave_acesso, situacao,
            descricao_retorno, parse(str(data_hora_emissao)), parse(str(data_hora_saida)), parse(str(data_hora_autorizacao)), protocolo_autorizacao,
            parse(str(data_hora_cancelamento)), protocolo_cancelamento, motivo_cancelamento, tipo_manifestacao,
            parse(str(data_hora_manifestacao)), p_transportadora, tipo_frete, codigo_antt, quantidade.split(',')[0], especie, marca,
            numeracao_volume, str(peso_liquido).replace('.', '').replace(',', '.'), str(peso_bruto).replace('.', '').replace(',', '.'),
            str(base_icms).replace('.', '').replace(',', '.'), str(total_icms).replace('.', '').replace(',', '.'),
            str(total_icms_deson).replace('.', '').replace(',', '.'), str(total_fcp).replace('.', '').replace(',', '.'),
            str(total_fcp_uf_dest).replace('.', '').replace(',', '.'), str(total_icms_dest).replace('.', '').replace(',', '.'),
            str(total_icms_uf_remet).replace('.', '').replace(',', '.'), str(base_icms_st).replace('.', '').replace(',', '.'),
            str(total_icms_st).replace('.', '').replace(',', '.'), str(total_fcp_st).replace('.', '').replace(',', '.'),
            str(total_fcp_st_ret).replace('.', '').replace(',', '.'), str(total_produtos_servicos).replace('.', '').replace(',', '.'),
            str(total_frete).replace('.', '').replace(',', '.'), str(total_seguro).replace('.', '').replace(',', '.'),
            str(total_desconto).replace('.', '').replace(',', '.'), str(total_ii).replace('.', '').replace(',', '.'),
            str(total_ipi).replace('.', '').replace(',', '.'), str(total_ipi_devolvido).replace('.', '').replace(',', '.'),
            str(total_pis).replace('.', '').replace(',', '.'), str(total_cofins).replace('.', '').replace(',', '.'),
            str(total_outras_despesas).replace('.', '').replace(',', '.'), str(total_nfe).replace('.', '').replace(',', '.'),
            str(vl_aprox_tot_trib).replace('.', '').replace(',', '.'), placa,
            uf, informacoes_fisco, informacoes_complementares, parse(str(data_consulta))
        )
    )
    conn.commit()

    linha += 1

# cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'nf_entrada2'")
# resuldado = cursor.fetchall()
#
# print(resuldado)

cursor.close()
conn.close()
