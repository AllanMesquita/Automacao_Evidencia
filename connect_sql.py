import psycopg2
from datetime import datetime
import win32com.client
import logging

name_log = str('C:\\Users\\allan.mesquita\\OneDrive - NTT\\Documents\\Projetos\\Automacao_Evidencias\\Script\\Error_Log\\Error_Log_' + datetime.strftime(datetime.today(), '%d-%m-%Y %H.%M') + '.txt')

tempo = datetime.now()

# win32 = win32com.client.Dispatch('Excel.Application')
# win32.Visible = False
# tblrec = win32.Workbooks.Open("C:\\Users\\allan.mesquita\\OneDrive - NTT\\Documents\\Projetos\\Automacao_Evidencias\\Backup V17\\Backup V17.1\\tblEvidenciaRecebimento - 07.04.2022.xlsm")
# aba_tblrec = tblrec.Worksheets('Evidencias')
#
# linha = 2
# qtd_linhas = aba_tblrec.UsedRange.Rows.Count



        ### Connection SQL
con = psycopg2.connect(
    host = "psql-itlatam-logisticcontrol.postgres.database.azure.com",
    dbname = "logistic-control",
    user = "logisticpsqladmin@psql-itlatam-logisticcontrol",
    password = "EsjHSrS69295NzHu342ap6P!N",
    sslmode = "require"
)

cur = con.cursor()



##################################################################################################
        ### Insert line

try:
    # while linha != qtd_linhas:
    #
    #     chavenf_entrada = str(aba_tblrec.Range(f'A{linha}'))
    #     pedidocompra = int(aba_tblrec.Range(f'B{linha}'))
    #     rfid_cxmaster = str(aba_tblrec.Range(f'C{linha}'))
    #     partnumber = str(aba_tblrec.Range(f'D{linha}'))
    #     rfid_produto = str(aba_tblrec.Range(f'E{linha}'))
    #     serialnumber = str(aba_tblrec.Range(f'F{linha}'))
    #     local = str(aba_tblrec.Range(f'G{linha}'))
    #     dataevidencia = str(aba_tblrec.Range(f'H{linha}'))
    #     usuario = str(aba_tblrec.Range(f'I{linha}'))
    #     obsrecebimento = str(aba_tblrec.Range(f'J{linha}'))
    #     chaverelacionamento = str(aba_tblrec.Range(f'K{linha}'))
    #     lctobd_data = str(aba_tblrec.Range(f'L{linha}'))
    #     lctobd_usuario = str(aba_tblrec.Range(f'M{linha}'))
    #
    #     data_banco = ''
    #     cur.execute(f"SELECT dataevidencia FROM tblrecebimento WHERE chaverelacionamento = '{chaverelacionamento}'")
    #     databco = cur.fetchall()
    #     if bool(databco):
    #         for lista in databco:
    #             for dado in lista:
    #                 data_banco = dado
    #         dataevidencia_convertida = datetime.strptime(dataevidencia[0:10], '%Y-%m-%d')
    #         data_banco_convertida = datetime.strptime(str(data_banco), '%Y-%m-%d')
    #         if dataevidencia_convertida < data_banco_convertida or dataevidencia_convertida == data_banco_convertida:
    #             linha += 1
    #             continue
    #         else:
    #             cur.execute(f'INSERT INTO tblrecebimento ("chavenf_entrada", '
    #                                                       '"pedidocompra", '
    #                                                       '"rfid_cxmaster", '
    #                                                       '"partnumber", '
    #                                                       '"rfid_produto", '
    #                                                       '"serialnumber", '
    #                                                       '"local", '
    #                                                       '"dataevidencia", '
    #                                                       '"usuario", '
    #                                                       '"obsrecebimento", '
    #                                                       '"chaverelacionamento", '
    #                                                       '"lctobd_data", '
    #                                                       '"lctobd_usuario") '
    #                                                       'VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)',
    #                                                       (chavenf_entrada,
    #                                                        pedidocompra,
    #                                                        rfid_cxmaster,
    #                                                        partnumber,
    #                                                        rfid_produto,
    #                                                        serialnumber,
    #                                                        local,
    #                                                        dataevidencia,
    #                                                        usuario,
    #                                                        obsrecebimento,
    #                                                        chaverelacionamento,
    #                                                        lctobd_data,
    #                                                        lctobd_usuario))
    #
    #             con.commit()
    #             linha += 1
    #     else:
    #         cur.execute(f'INSERT INTO tblrecebimento ("chavenf_entrada", '
    #                     '"pedidocompra", '
    #                     '"rfid_cxmaster", '
    #                     '"partnumber", '
    #                     '"rfid_produto", '
    #                     '"serialnumber", '
    #                     '"local", '
    #                     '"dataevidencia", '
    #                     '"usuario", '
    #                     '"obsrecebimento", '
    #                     '"chaverelacionamento", '
    #                     '"lctobd_data", '
    #                     '"lctobd_usuario") '
    #                     'VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)',
    #                     (chavenf_entrada,
    #                      pedidocompra,
    #                      rfid_cxmaster,
    #                      partnumber,
    #                      rfid_produto,
    #                      serialnumber,
    #                      local,
    #                      dataevidencia,
    #                      usuario,
    #                      obsrecebimento,
    #                      chaverelacionamento,
    #                      lctobd_data,
    #                      lctobd_usuario))
    #
    #         con.commit()
    #         linha += 1
    #
    # tblrec.Save()
    # win32.Application.Quit()

    # cnpj = '05437734000318'
    # inscricao_estadual = '082623910'
    # razao_social = 'NTT BRASIL COMERCIO E SERVICOS DE TECNOLOGIA LTDA'
    # endereco = 'RODOVIA GOVERNADOR MARIO COVAS, 3101'
    # bairro = 'PADRE MATHIAS'
    # cep = '29157100'
    # municipio = 'Cariacica'
    # uf = 'ES'
    #
    # cur.execute(f'INSERT INTO destinatario (cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)', (int(cnpj), inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf))
    # con.commit()

    # numero_nfe = '4052420'
    # serie = '1'
    # tipo_nfe = '1'
    # cnpj_destinatario = '00'
    # inscricao_estadual_destinatario = '00'
    # razao_social_destinatario = '1'
    # endereco_destinatario = '00'
    # bairro_destinatario = '00'
    # cep_destinatario = '00'
    # municipio_destinatario = '00'
    # uf_destinatario = '00'
    #
    # cur.execute(f'INSERT INTO nf_entrada (numero_nfe, serie, tipo_nfe, cnpj_destinatario, inscricao_estadual_destinatario, razao_social_destinatario, endereco_destinatario, bairro_destinatario, cep_destinatario, municipio_destinatario, uf_destinatario) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)', (numero_nfe, serie, tipo_nfe, cnpj_destinatario, inscricao_estadual_destinatario, razao_social_destinatario, endereco_destinatario, bairro_destinatario, cep_destinatario, municipio_destinatario, uf_destinatario))
    # con.commit()

        ########################################################################################################################
                ### Select data

    # cur.execute("SELECT dataevidencia FROM tblrecebimento WHERE chaverelacionamento = 'E00000000000000000067137TERCA VIX'")
    #
    # resultado = cur.fetchall()
    # dataa = ''
    # quant = 0
    # for res in resultado:
    #     quant += 1
    #     for dado in res:
    #         dataa = dado
    #
    # print(quant)
    # data = str(aba_tblrec.Range(f'H{linha}'))
    # data2 = datetime.strptime(str(aba_tblrec.Range(f'H{linha}'))[0:10], '%Y-%m-%d')
    # print(data2 == datetime.strptime(str(dado), '%Y-%m-%d'))
    # print(type(dataa))

        #######################################################################################################################
                ## Select specific data
    # qtde = 0
    # chave = '35220572381189001001550040001989671118636423'
    #
    # pesquisa = cur.execute(f"SELECT id FROM nf_entrada_itens WHERE chave_acesso = '{chave}'")
    # # select = select[0][0]
    # row = cur.fetchall()
    #
    # # for d in row:
    # #     # print(d)
    # #     for dado in d:
    # #         qtde += dado
    #
    # # print(sum(row))
    # print(row)



        #########################################################################################################################
                ### Select with filter

        # cur.execute("SELECT rfid_produto FROM tblrecebimento WHERE serialnumber = 'T3AA3N0930'")
        #
        # row = cur.fetchall()
        #
        # # for d in row:
        # #     print(f'ID: {d[0]}\n'
        # #           f'RFID: {d[1]}\n'
        # #           f'Serial: {d[2]}')
        #
        # print(row)



        #########################################################################################################################
                ### Delete data
        
    id = 1040
    #
    while id != 1296:

        cur.execute(f"DELETE FROM public.nf_entrada_itens2 WHERE id = '{id}'")

        con.commit()

        id += 1



        #########################################################################################################################
                ### Update data
    # numero = 11111
    # serie = 'VENDA MERCAD RECEB ADQ TERCEIR'
    # var = 39728562888
    # cnpf = 15462589000126
    # valor = '440,65'
    # razao = 'Allan de Oliveira Mesquita'
    # qtd = '1,0'
    # data = "03/05/2022 13:51:03"
    # cur.execute("UPDATE stockcontrol SET id = '1' WHERE id = '12'")
    #
    # if bool(row):
    #     if row[50][0] == chave:
    #         print("Dado no banco")
    #     else:
    #         print("não tem")
    # else:
    #     pass
        # cur.execute(f'INSERT INTO nf_entrada_itens (natureza_operacao, serie, cnpj_cpf_destinatario, razao_social_destinatario, cnpj_emitente, valor_nf_e, data_emissao, qtde_com)'
        #             'VALUES (%s, %s, %s, %s, %s, %s, %s, %s)',
        #             (serie, numero, var, razao, cnpf, valor.replace(',', '.'), data, qtd)
        #             )
    # con.commit()

    # print(type(row[0][0]))

        #########################################################################################################################

except Exception as error:
    cur.close()
    con.close()
#     print(f'linha {linha} - rfid {rfid_produto}')
    print(f'id {id}')
    print(datetime.now() - tempo)
    logging.basicConfig(filename=name_log, filemode='w', format='%(asctime)s %(message)s')
    logging.critical(f'- {error}', exc_info=True)

finally:    
    cur.close()

    con.close()

    print(datetime.now() - tempo)