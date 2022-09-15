import psycopg2
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

    chave_acesso = dic['Chave de Acesso']
    data_emissao = dic['Data EmissÃ£o']
    descricao_produto = dic['DescriÃ§Ã£o do Produto']
    cod_produto = dic['CÃ³d_x002e_ Produto']
    numero_pedido = dic['NÃºm_x002e_ Pedido']
    cean = dic['cEAN']
    cean_trib = dic['cEANTrib']
    unid_com = dic['Unid_x002e_ Com_x002e_']
    valor_unitario = dic['Valor UnitÃ¡rio Comercial']
    quantidade = dic['Qtde_x002e_ Com_x002e_']
    valor_total = dic['Valor Total']
    origem = dic['Origem']
    base_icms = dic['Base de CÃ¡lculo ICMS']
    cst_icms = dic['CST ICMS / CSOSN']
    aliq_icms = dic['AlÃ\xadq_x002e_ ICMS']
    valor_icms = dic['Valor ICMS']
    perc_icms = dic['Perc_x002e_ Margem ICMS ST']
    base_icms_st = dic['Base de CÃ¡lc_x002e_ ICMS ST']
    valor_icms_st = dic['Valor ICMS ST']
    aliq_icms_st = dic['AlÃ\xadq_x002e_ ICMS ST']
    valor_pis = dic['Valor PIS']
    cst_pis = dic['CST PIS']
    valor_cofins = dic['Valor COFINS']
    cst_cofins = dic['CST COFINS']
    valor_ipi = dic['Valor IPI']
    cst_ipi = dic['CST IPI']
    aliq_ipi = dic['AlÃ\xadq_x002e_ IPI']
    ncm = dic['NCM']
    cfop = dic['CFOP']

    # PESQUISA NOTA NO BANCO
    cursor.execute(f"SELECT chave_acesso FROM public.nf_entrada_itens2 WHERE chave_acesso = '{chave_acesso}'")
    resultado = cursor.fetchall()
    if bool(resultado):
        continue
    else:
    # PESQUISAS
        # CFOP
        teste = pesquisa(cursor, conn, chave_acesso, dic)
        id_cfop = teste.cfop()
        print(id_cfop)

    cursor.execute(
        'INSERT INTO public.nf_entrada_itens2 ('
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
            chave_acesso, parse(data_emissao), descricao_produto, cod_produto, numero_pedido, cean, cean_trib, unid_com,
            valor_unitario.replace('.', '').replace(',', '.'), quantidade.split(',')[0],
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

cursor.close()
conn.close()
