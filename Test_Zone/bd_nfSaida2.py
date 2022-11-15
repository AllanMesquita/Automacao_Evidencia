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

wb = openpyxl.load_workbook("C:\\Users\\allan.mesquita\\Downloads\\2022 á 2027 - NFs Saída Mastersaf.xlsx")
sheet = wb.sheetnames
aba = wb[sheet[0]]
linhas = len(aba['A'])

linha = 4119

while linha != linhas + 1:
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
    usuario_cancelamento = aba[f'AF{linha}'].value
    transportadora = aba[f'AG{linha}'].value
    tipo_frete = aba[f'AL{linha}'].value
    codigo_antt = aba[f'AM{linha}'].value
    quantidade = aba[f'AN{linha}'].value if bool(aba[f'AN{linha}'].value) else '00'
    especie = aba[f'AO{linha}'].value
    marca = aba[f'AP{linha}'].value
    numeracao_volume = aba[f'AQ{linha}'].value if bool(aba[f'AQ{linha}'].value) else '00'
    peso_liquido = aba[f'AR{linha}'].value if bool(aba[f'AR{linha}'].value) else '00'
    peso_bruto = aba[f'AS{linha}'].value if bool(aba[f'AS{linha}'].value) else '00'
    base_icms = aba[f'AT{linha}'].value if bool(aba[f'AT{linha}'].value) else '00'
    total_icms = aba[f'AU{linha}'].value if bool(aba[f'AU{linha}'].value) else '00'
    total_icms_deson = aba[f'AV{linha}'].value if bool(aba[f'AV{linha}'].value) else '00'
    total_fcp = aba[f'AW{linha}'].value if bool(aba[f'AW{linha}'].value) else '00'
    total_fcp_uf_dest = aba[f'AX{linha}'].value if bool(aba[f'AX{linha}'].value) else '00'
    total_icms_uf_dest = aba[f'AY{linha}'].value if bool(aba[f'AY{linha}'].value) else '00'
    total_icms_uf_remet = aba[f'AZ{linha}'].value if bool(aba[f'AZ{linha}'].value) else '00'
    base_icms_st = aba[f'BA{linha}'].value if bool(aba[f'BA{linha}'].value) else '00'
    total_icms_st = aba[f'BB{linha}'].value if bool(aba[f'BB{linha}'].value) else '00'
    total_fcp_st = aba[f'BC{linha}'].value if bool(aba[f'BC{linha}'].value) else '00'
    total_fcp_st_ret = aba[f'BD{linha}'].value if bool(aba[f'BD{linha}'].value) else '00'
    total_produtos_servicos = aba[f'BE{linha}'].value if bool(aba[f'BE{linha}'].value) else '00'
    total_frete = aba[f'BF{linha}'].value if bool(aba[f'BF{linha}'].value) else '00'
    total_seguro = aba[f'BG{linha}'].value if bool(aba[f'BG{linha}'].value) else '00'
    total_desconto = aba[f'BH{linha}'].value if bool(aba[f'BH{linha}'].value) else '00'
    total_ii = aba[f'BI{linha}'].value if bool(aba[f'BI{linha}'].value) else '00'
    total_ipi = aba[f'BJ{linha}'].value if bool(aba[f'BJ{linha}'].value) else '00'
    total_ipi_devolvido = aba[f'BK{linha}'].value if bool(aba[f'BK{linha}'].value) else '00'
    total_pis = aba[f'BL{linha}'].value if bool(aba[f'BL{linha}'].value) else '00'
    total_cofins = aba[f'BM{linha}'].value if bool(aba[f'BM{linha}'].value) else '00'
    total_outras_despesas = aba[f'BN{linha}'].value if bool(aba[f'BN{linha}'].value) else '00'
    total_nfe = aba[f'BO{linha}'].value if bool(aba[f'BO{linha}'].value) else '00'
    vl_aprox_tot_trib = aba[f'BP{linha}'].value if bool(aba[f'BP{linha}'].value) else '00'
    placa = aba[f'BQ{linha}'].value
    uf = aba[f'BR{linha}'].value
    informacoes_fisco = aba[f'BS{linha}'].value
    informacoes_complementares = aba[f'BT{linha}'].value
    data_consulta = aba[f'BV{linha}'].value if bool(aba[f'BV{linha}'].value) else '01/01/2001 00:00:01'

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
    #       usuario_cancelamento, '\n',
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
    #       total_icms_uf_dest, '\n',
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

    print(chave_acesso)

    # PESQUISA NOTA NO BANCO
    cursor.execute(f"SELECT chave_acesso FROM material_management.master_saf_saida WHERE chave_acesso = '{chave_acesso}'")
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
        p_transportadora = teste.transportadora_saida()

    print(p_destinatario, p_fornecedor, p_natureza, p_transportadora)
    #
    cursor.execute(
        'INSERT INTO material_management.master_saf_saida ('
        'numero_nfe, serie, tipo_nfe, destinatario, fornecedor, natureza_cfop, chave_acesso, situacao, '
        'descricao_retorno, data_hora_emissao, data_hora_saida, data_hora_autorizacao, protocolo_autorizacao, '
        'data_hora_cancelamento, protocolo_cancelamento, motivo_cancelamento, usuario_cancelamento, '
        'transportadora, tipo_frete, codigo_antt, quantidade, especie, marca, numeracao_volume, peso_liquido, '
        'peso_bruto, base_icms, total_icms, total_icms_deson, total_fcp, total_fcp_uf_dest, total_icms_uf_dest, '
        'total_icms_uf_remet, base_icms_st, total_icms_st, total_fcp_st, total_fcp_st_ret, total_produtos_servicos,'
        'total_frete, total_seguro, total_desconto, total_ii, total_ipi, total_ipi_devolvido, total_pis, total_cofins,'
        'total_outras_despesas, total_nfe, vl_aprox_tot_trib, placa, uf, informacoes_fisco, informacoes_complementares,'
        'data_consulta'
        ')'
        'VALUES ('
        '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, '
        '%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s'
        ')',
        (
            numero_nfe, serie, tipo_nfe, p_destinatario, p_fornecedor, p_natureza, chave_acesso, situacao,
            descricao_retorno, parse(str(data_hora_emissao)), parse(str(data_hora_saida)), parse(str(data_hora_autorizacao)), protocolo_autorizacao,
            parse(str(data_hora_cancelamento)), protocolo_cancelamento, motivo_cancelamento, usuario_cancelamento,
            p_transportadora, tipo_frete, codigo_antt, quantidade.split(',')[0].replace('.', ''), especie, marca,
            numeracao_volume, peso_liquido.replace('.', '').replace(',', '.'), peso_bruto.replace('.', '').replace(',', '.'),
            base_icms.replace('.', '').replace(',', '.'), total_icms.replace('.', '').replace(',', '.'),
            total_icms_deson.replace('.', '').replace(',', '.'), total_fcp.replace('.', '').replace(',', '.'),
            total_fcp_uf_dest.replace('.', '').replace(',', '.'), total_icms_uf_dest.replace('.', '').replace(',', '.'),
            total_icms_uf_remet.replace('.', '').replace(',', '.'), base_icms_st.replace('.', '').replace(',', '.'),
            total_icms_st.replace('.', '').replace(',', '.'), total_fcp_st.replace('.', '').replace(',', '.'),
            total_fcp_st_ret.replace('.', '').replace(',', '.'), total_produtos_servicos.replace('.', '').replace(',', '.'),
            total_frete.replace('.', '').replace(',', '.'), total_seguro.replace('.', '').replace(',', '.'),
            total_desconto.replace('.', '').replace(',', '.'), total_ii.replace('.', '').replace(',', '.'),
            total_ipi.replace('.', '').replace(',', '.'), total_ipi_devolvido.replace('.', '').replace(',', '.'),
            total_pis.replace('.', '').replace(',', '.'), total_cofins.replace('.', '').replace(',', '.'),
            total_outras_despesas.replace('.', '').replace(',', '.'), total_nfe.replace('.', '').replace(',', '.'),
            vl_aprox_tot_trib.replace('.', '').replace(',', '.'), placa,
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
