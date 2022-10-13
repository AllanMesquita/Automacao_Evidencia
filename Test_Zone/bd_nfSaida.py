import psycopg2
# import datetime
import json
from Modulos.class_pesquisa_v2 import Pesquisa as pesquisa
from dateutil.parser import parse

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

data = open("C:\\Users\\allan.mesquita\\Downloads\\teste.json")

req_body = json.load(data)

for dic in req_body:
    # print(dic)

    numero_nfe = dic['NÃºmero da NF-e']
    serie = dic['SÃ©rie']
    tipo_nfe = dic['Tipo NF-e'] if bool(dic['Tipo NF-e']) else 99
    destinatario = dic['CNPJ do DestinatÃ¡rio']
    fornecedor = dic['CNPJ/CPF do Fornecedor']
    natureza_cfop = dic['Natureza da OperaÃ§Ã£o']
    chave_acesso = dic['Chave de Acesso']
    situacao = dic['SituaÃ§Ã£o']
    descricao_retorno = dic['DescriÃ§Ã£o no Retorno']
    data_hora_emissao = dic['Data e Hora da EmissÃ£o'] if bool(dic['Data e Hora da EmissÃ£o']) else "01/01/2001 00:00:01"
    data_hora_saida = dic['Data e Hora da SaÃ\xadda'] if bool(dic['Data e Hora da SaÃ\xadda']) else "01/01/2001 00:00:01"
    data_hora_autorizacao = dic['Data e Hora da AutorizaÃ§Ã£o'] if bool(dic['Data e Hora da AutorizaÃ§Ã£o']) else "01/01/2001 00:00:01"
    protocolo_autorizacao = dic['Protocolo AutorizaÃ§Ã£o']
    data_hora_cancelamento = dic['Data e hora Cancelamento'] if bool(dic['Data e hora Cancelamento']) else "01/01/2001 00:00:01"
    protocolo_cancelamento = dic['Protocolo Cancelamento']
    motivo_cancelamento = dic['Motivo Cancelamento']
    usuario_cancelamento = dic['UsuÃ¡rio Cancelamento']
    transportadora = dic['Cnpj Transportadora']
    tipo_frete = dic['Tipo Frete']
    codigo_antt = dic['CÃ³digo ANTT']
    quantidade = dic['Quantidade'] if bool(dic['Quantidade']) else '00'
    especie = dic['EspÃ©cie']
    marca = dic['Marca']
    numeracao_volume = dic['NumeraÃ§Ã£o do Volume'] if bool(dic['NumeraÃ§Ã£o do Volume']) else '00'
    peso_liquido = dic['Peso LÃ\xadquido'] if bool(dic['Peso LÃ\xadquido']) else '00'
    peso_bruto = dic['Peso Bruto'] if bool(dic['Peso Bruto']) else '00'
    base_icms = dic['Base Calculo ICMS'] if bool(dic['Base Calculo ICMS']) else '00'
    total_icms = dic['Total ICMS'] if bool(dic['Total ICMS']) else '00'
    total_icms_deson = dic['Total ICMS Deson_x002e_'] if bool(dic['Total ICMS Deson_x002e_']) else '00'
    total_fcp = dic['Total FCP'] if bool(dic['Total FCP']) else '00'
    total_fcp_uf_dest = dic['Total FCP UF Dest_x002e_'] if bool(dic['Total FCP UF Dest_x002e_']) else '00'
    total_icms_uf_dest = dic['Total ICMS UF Dest_x002e_'] if bool(dic['Total ICMS UF Dest_x002e_']) else '00'
    total_icms_uf_remet = dic['Total ICMS UF Remet_x002e_'] if bool(dic['Total ICMS UF Remet_x002e_']) else '00'
    base_icms_st = dic['Base de CÃ¡lculo ICMS ST'] if bool(dic['Base de CÃ¡lculo ICMS ST']) else '00'
    total_icms_st = dic['Total ICMS ST'] if bool(dic['Total ICMS ST']) else '00'
    total_fcp_st = dic['Total FCP ST'] if bool(dic['Total FCP ST']) else '00'
    total_fcp_st_ret = dic['Total FCP ST Ret'] if bool(dic['Total FCP ST Ret']) else '00'
    total_produtos_servicos = dic['Total Produtos e ServiÃ§os'] if bool(dic['Total Produtos e ServiÃ§os']) else '00'
    total_frete = dic['Total Frete'] if bool(dic['Total Frete']) else '00'
    total_seguro = dic['Total Seguro'] if bool(dic['Total Seguro']) else '00'
    total_desconto = dic['Total Desconto'] if bool(dic['Total Desconto']) else '00'
    total_ii = dic['Total II'] if bool(dic['Total II']) else '00'
    total_ipi = dic['Total IPI'] if bool(dic['Total IPI']) else '00'
    total_ipi_devolvido = dic['Total IPI Devolvido'] if bool(dic['Total IPI Devolvido']) else '00'
    total_pis = dic['Total PIS'] if bool(dic['Total PIS']) else '00'
    total_cofins = dic['Total COFINS'] if bool(dic['Total COFINS']) else '00'
    total_outras_despesas = dic['Total Outras Despesas'] if bool(dic['Total Outras Despesas']) else '00'
    total_nfe = dic['Total da NF-e'] if bool(dic['Total da NF-e']) else '00'
    vl_aprox_tot_trib = dic['Vl_x002e_ Aprox_x002e_ Tot_x002e_ Trib_x002e_'] if bool(dic['Vl_x002e_ Aprox_x002e_ Tot_x002e_ Trib_x002e_']) else '00'
    placa = dic['Placa']
    uf = dic['UF']
    informacoes_fisco = dic['InformaÃ§Ãµes Adicionais do Fisco']
    informacoes_complementares = dic['InformaÃ§Ãµes Complementares']
    data_consulta = dic['Data Consulta'] if bool(dic['Data Consulta']) else '01/01/2001 00:00:01'

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

    # PESQUISA NOTA NO BANCO
    cursor.execute(f"SELECT chave_acesso FROM public.nf_saida WHERE chave_acesso = '{chave_acesso}'")
    resultado = cursor.fetchall()
    if bool(resultado):
        continue
    else:
    # PESQUISAS
        # Destinatário
        teste = pesquisa(cursor, conn, destinatario, dic)
        p_destinatario = teste.destinatario()
        # Fornecedor
        teste = pesquisa(cursor, conn, fornecedor, dic)
        p_fornecedor = teste.fornecedor()
        # Natura e CFOP
        teste = pesquisa(cursor, conn, natureza_cfop, dic)
        p_natureza = teste.natureza()
        # Transportadora
        teste = pesquisa(cursor, conn, transportadora, dic)
        p_transportadora = teste.transportadora()

    print(p_destinatario, p_fornecedor, p_natureza, p_transportadora)

    cursor.execute(
        'INSERT INTO public.nf_saida ('
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
            descricao_retorno, parse(data_hora_emissao), parse(data_hora_saida), parse(data_hora_autorizacao), protocolo_autorizacao,
            parse(data_hora_cancelamento), protocolo_cancelamento, motivo_cancelamento, usuario_cancelamento,
            p_transportadora, tipo_frete, codigo_antt, quantidade.split(',')[0], especie, marca,
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
            uf, informacoes_fisco, informacoes_complementares, parse(data_consulta)
        )
    )
    conn.commit()


# cursor.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'nf_entrada2'")
# resuldado = cursor.fetchall()
#
# print(resuldado)

cursor.close()
conn.close()
