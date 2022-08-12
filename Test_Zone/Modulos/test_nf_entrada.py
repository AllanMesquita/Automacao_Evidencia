import psycopg2
from datetime import datetime
import win32com.client
import logging
import class_pesquisa

tempo = datetime.now()

win32 = win32com.client.Dispatch('Excel.Application')
win32.Visible = False
arquivo = win32.Workbooks.Open("C:\\Users\\allan.mesquita\\NTT\\@AM BR Services and Operations - Privado\\GESTÃO DE ESTOQUE\\100 BcoDados\\003 Evidencias\\06 Lixeira\\Testes\\01 Processamento\\20220527_RECEBIMENTO_NFE.xlsx")
aba_arquivo = arquivo.Worksheets('Listagem de NF-e Recebidas')

linha = 2
qtd_linhas = aba_arquivo.UsedRange.Rows.Count

        ### Connection SQL
con = psycopg2.connect(
    host="psql-itlatam-logisticcontrol.postgres.database.azure.com",
    dbname="logistic-control",
    user="logisticpsqladmin@psql-itlatam-logisticcontrol",
    password="EsjHSrS69295NzHu342ap6P!N",
    sslmode="require"
)

cur = con.cursor()

    ###VARIAVEIS
while linha != qtd_linhas + 1:
    print(linha)
    numero_nfe = str(aba_arquivo.Range(f'A{linha}'))
    serie = str(aba_arquivo.Range(f'B{linha}'))

    tipo_nfe = str(aba_arquivo.Range(f'C{linha}'))
    tipo_nfe = tipo_nfe if bool(tipo_nfe) else '00'

    destinatario = str(aba_arquivo.Range(f'D{linha}'))  # CNPJ destinatário para pesquisa
    fornecedor = str(aba_arquivo.Range(f'L{linha}'))  # CNPJ fornecedor para pesquisa
    natureza_operacao = str(aba_arquivo.Range(f'T{linha}'))  # para pesquisa
    chave_acesso = str(aba_arquivo.Range(f'V{linha}'))
    situacao = str(aba_arquivo.Range(f'W{linha}'))
    descricao_retorno = str(aba_arquivo.Range(f'X{linha}'))

    data_hora_emissao = str(aba_arquivo.Range(f'Y{linha}'))
    data_hora_emissao = data_hora_emissao if bool(data_hora_emissao) else '01/01/0101 01:01:01'

    data_hora_saida = str(aba_arquivo.Range(f'Z{linha}'))
    data_hora_saida = data_hora_saida if bool(data_hora_saida) else '01/01/0101 01:01:01'

    data_hora_autorizacao = str(aba_arquivo.Range(f'AA{linha}'))
    data_hora_autorizacao = data_hora_autorizacao if bool(data_hora_autorizacao) else '01/01/0101 01:01:01'

    protocolo_autorizacao = str(aba_arquivo.Range(f'AB{linha}'))

    data_hora_cancelamento = str(aba_arquivo.Range(f'AC{linha}'))
    data_hora_cancelamento = data_hora_cancelamento if bool(data_hora_cancelamento) else '01/01/0101 01:01:01'

    protocolo_cancelamento = str(aba_arquivo.Range(f'AD{linha}'))
    motivo_cancelamento = str(aba_arquivo.Range(f'AE{linha}'))
    ciencia_manifestacao = str(aba_arquivo.Range(f'AF{linha}'))

    data_hora_manifestacao = str(aba_arquivo.Range(f'AG{linha}'))
    data_hora_manifestacao = data_hora_manifestacao if bool(data_hora_manifestacao) else '01/01/0101 01:01:01'

    transportadora = str(aba_arquivo.Range(f'AH{linha}'))  # CNPJ transportador para pesquisa

    tipo_frete = str(aba_arquivo.Range(f'AM{linha}'))
    tipo_frete = tipo_frete if bool(tipo_frete) else '00'

    codigo_antt = str(aba_arquivo.Range(f'AN{linha}'))

    quantidade = str(aba_arquivo.Range(f'AO{linha}'))
    quantidade = quantidade if bool(quantidade) else '00'

    especie = str(aba_arquivo.Range(f'AP{linha}'))
    marca = str(aba_arquivo.Range(f'AQ{linha}'))
    numeracao_volume = str(aba_arquivo.Range(f'AR{linha}'))

    peso_liquido = str(aba_arquivo.Range(f'AS{linha}'))
    peso_liquido = peso_liquido if bool(peso_liquido) else '00'

    peso_bruto = str(aba_arquivo.Range(f'AT{linha}'))
    peso_bruto = peso_bruto if bool(peso_bruto) else '00'

    base_calculo_icms = str(aba_arquivo.Range(f'AU{linha}'))
    total_icms = str(aba_arquivo.Range(f'AV{linha}'))
    total_icms_deson = str(aba_arquivo.Range(f'AW{linha}'))
    total_fcp = str(aba_arquivo.Range(f'AX{linha}'))
    total_fcp_uf_dest = str(aba_arquivo.Range(f'AY{linha}'))
    total_icms_uf_dest = str(aba_arquivo.Range(f'AZ{linha}'))
    total_icms_uf_remet = str(aba_arquivo.Range(f'BA{linha}'))
    base_calculo_icms_st = str(aba_arquivo.Range(f'BB{linha}'))
    total_icms_st = str(aba_arquivo.Range(f'BC{linha}'))
    total_fcp_st = str(aba_arquivo.Range(f'BD{linha}'))
    total_fcp_st_ret = str(aba_arquivo.Range(f'BE{linha}'))
    total_produtos_servicos = str(aba_arquivo.Range(f'BF{linha}'))
    total_frete = str(aba_arquivo.Range(f'BG{linha}'))
    total_seguro = str(aba_arquivo.Range(f'BH{linha}'))
    total_desconto = str(aba_arquivo.Range(f'BI{linha}'))
    total_ii = str(aba_arquivo.Range(f'BJ{linha}'))
    total_ipi = str(aba_arquivo.Range(f'BK{linha}'))
    total_ipi_devolvido = str(aba_arquivo.Range(f'BL{linha}'))
    total_pis = str(aba_arquivo.Range(f'BM{linha}'))
    total_cofins = str(aba_arquivo.Range(f'BN{linha}'))
    total_outras_despesas = str(aba_arquivo.Range(f'BO{linha}'))
    total_nfe = str(aba_arquivo.Range(f'BP{linha}'))
    vl_aprox_tot_trib = str(aba_arquivo.Range(f'BQ{linha}'))
    placa = str(aba_arquivo.Range(f'BR{linha}'))
    uf = str(aba_arquivo.Range(f'BS{linha}'))
    informacoes_adicionais_fisco = str(aba_arquivo.Range(f'BT{linha}'))
    informacoes_complementares = str(aba_arquivo.Range(f'BU{linha}'))
    usuario_consulta = str(aba_arquivo.Range(f'BV{linha}'))

    data_consulta = str(aba_arquivo.Range(f'BW{linha}'))
    data_consulta = data_consulta if bool(data_consulta) else '01/01/0101'

    cfop = str(aba_arquivo.Range(f'U{linha}'))  # para pesquisa

    # print(numero_nfe, serie, tipo_nfe, destinatario, fornecedor, natureza_operacao, chave_acesso, situacao,
    #       descricao_retorno, data_hora_emissao, data_hora_saida, data_hora_autorizacao, protocolo_autorizacao,
    #       data_hora_cancelamento, protocolo_cancelamento, motivo_cancelamento, ciencia_manifestacao,
    #       data_hora_manifestacao, transportadora, tipo_frete, codigo_antt, quantidade, especie, marca, numeracao_volume,
    #       peso_liquido, peso_bruto, base_calculo_icms, total_icms, total_icms_deson, total_fcp, total_fcp_uf_dest,
    #       total_icms_uf_dest, total_icms_uf_remet, base_calculo_icms_st, total_icms_st, total_fcp_st, total_fcp_st_ret,
    #       total_produtos_servicos, total_frete, total_seguro, total_desconto, total_ii, total_ipi, total_ipi_devolvido,
    #       total_pis, total_cofins, total_outras_despesas, total_nfe, vl_aprox_tot_trib, placa, uf,
    #       informacoes_adicionais_fisco, informacoes_complementares, usuario_consulta, data_consulta, cfop)

    ### Validação de Linha em Branco
    if bool(numero_nfe) is False:
        linha += 1
        print("Linha em branco")
        continue
    else:

        ### Pesquisa da nota no banco
        cur.execute(f"SELECT chave_acesso FROM nf_entrada WHERE chave_acesso = '{chave_acesso}'")
        pesquisa_banco = cur.fetchall()


        # Caso for encontrado
        if bool(pesquisa_banco):
            linha += 1
            print('Dado no banco.')
            continue
        # Caso não for encontrado
        else:
            ### Pesquisa Destinatário
            id_destinatario = class_pesquisa.Pesquisa(cur, con, destinatario, aba_arquivo, linha).destinatario()

            ## Pesquisa Fornecedor
            id_fornecedor = class_pesquisa.Pesquisa(cur, con, fornecedor, aba_arquivo, linha).fornecedor()

            ### Pesquisa Natureza Operação
            id_natureza = class_pesquisa.Pesquisa(cur, con, natureza_operacao, aba_arquivo, linha).natureza_operacao()

            ### Pesquisa CFOP
            id_cfop = class_pesquisa.Pesquisa(cur, con, cfop, aba_arquivo, linha).cfop()

            ### Pesquisa Transportadora
            id_transportadora = class_pesquisa.Pesquisa(cur, con, transportadora, aba_arquivo, linha).transportadora()

            print(datetime.strptime(data_hora_saida, '%d/%m/%Y %H:%M:%S').strftime('%m/%d/%Y'))

            cur.execute("INSERT INTO nf_entrada (numero_nfe, serie, tipo_nfe, destinatario, fornecedor, natureza_operacao, chave_acesso, situacao, descricao_retorno, data_hora_emissao, data_hora_saida, data_hora_autorizacao, protocolo_autorizacao, data_hora_cancelamento, protocolo_cancelamento, motivo_cancelamento, ciencia_manifestacao, data_hora_manifestacao, transportadora, tipo_frete, codigo_antt, quantidade, especie, marca, numeracao_volume, peso_liquido, peso_bruto, base_calculo_icms, total_icms, total_icms_deson, total_fcp, total_fcp_uf_dest, total_icms_uf_dest, total_icms_uf_remet, base_calculo_icms_st, total_icms_st, total_fcp_st, total_fcp_st_ret, total_produtos_servicos, total_frete, total_seguro, total_desconto, total_ii, total_ipi, total_ipi_devolvido, total_pis, total_cofins, total_outras_despesas, total_nfe, vl_aprox_tot_trib, placa, uf, informacoes_adicionais_fisco, informacoes_complementares, usuario_consulta, data_consulta, cfop) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", (numero_nfe, serie, tipo_nfe, id_destinatario, id_fornecedor, id_natureza, chave_acesso, situacao, descricao_retorno, datetime.strptime(data_hora_emissao, '%d/%m/%Y %H:%M:%S').strftime('%m/%d/%Y'), datetime.strptime(data_hora_saida, '%d/%m/%Y %H:%M:%S').strftime('%m/%d/%Y'), datetime.strptime(data_hora_autorizacao, '%d/%m/%Y %H:%M:%S'), protocolo_autorizacao, datetime.strptime(data_hora_cancelamento, '%d/%m/%Y %H:%M:%S').strftime('%m/%d/%Y'), protocolo_cancelamento, motivo_cancelamento, ciencia_manifestacao, datetime.strptime(data_hora_manifestacao, '%d/%m/%Y %H:%M:%S').strftime('%m/%d/%Y'), id_transportadora, tipo_frete, codigo_antt, quantidade.split(',')[0], especie, marca, numeracao_volume, peso_liquido.replace('.', '').replace(',', '.'), peso_bruto.replace('.', '').replace(',', '.'), base_calculo_icms.replace('.', '').replace(',', '.'), total_icms.replace('.', '').replace(',', '.'), total_icms_deson.replace('.', '').replace(',', '.'), total_fcp.replace('.', '').replace(',', '.'), total_fcp_uf_dest.replace('.', '').replace(',', '.'), total_icms_uf_dest.replace('.', '').replace(',', '.'), total_icms_uf_remet.replace('.', '').replace(',', '.'), base_calculo_icms_st.replace('.', '').replace(',', '.'), total_icms_st.replace('.', '').replace(',', '.'), total_fcp_st.replace('.', '').replace(',', '.'), total_fcp_st_ret.replace('.', '').replace(',', '.'), total_produtos_servicos.replace('.', '').replace(',', '.'), total_frete.replace('.', '').replace(',', '.'), total_seguro.replace('.', '').replace(',', '.'), total_desconto.replace('.', '').replace(',', '.'), total_ii.replace('.', '').replace(',', '.'), total_ipi.replace('.', '').replace(',', '.'), total_ipi_devolvido.replace('.', '').replace(',', '.'), total_pis.replace('.', '').replace(',', '.'), total_cofins.replace('.', '').replace(',', '.'), total_outras_despesas.replace('.', '').replace(',', '.'), total_nfe.replace('.', '').replace(',', '.'), vl_aprox_tot_trib.replace('.', '').replace(',', '.'), placa, uf, informacoes_adicionais_fisco, informacoes_complementares, usuario_consulta, datetime.strptime(data_consulta[0:10], '%d/%m/%Y').strftime('%m/%d/%Y'), id_cfop))
            con.commit()
            print('Dados da nota de entrada inserido no banco')

            print(id_destinatario)

    linha += 1

# cur.execute("SELECT column_name FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'nf_entrada'")
# row = cur.fetchall()
#
# for lista in row:
#     for dado in lista:
#         print(dado, end=',')

# print(row)

cur.close()

con.close()

arquivo.Save()
win32.Application.Quit()

print(datetime.now() - tempo)
