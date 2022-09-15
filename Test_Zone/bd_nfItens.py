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

cursor.close()
conn.close()

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
