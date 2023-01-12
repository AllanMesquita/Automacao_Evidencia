
dict_return = {}


def rec_validation(lista, repeticao_rfid, repeticao_serial):
    # Imports
    from datetime import datetime
    from Modulos.class_errosBD import Error
    from dateutil.parser import parse
    import psycopg2

    global error_chave, data

    dict_return_rec = {}
    dict_error = {}

    print('Início da validação - Recebimento')

    for item in lista:

        error = Error()

        error_chave = 0
        error_PO = 0
        error_PN = 0
        error_RFID = 0
        error_SN = 0
        error_Date = 0
        error_ChaveRel = 0
        error_local = 0

        ### VALIDAÇÃO DA CHAVE DE NOTA FISCAL

        cell_range = item['ChaveNF_Entrada']

        con = psycopg2.connect(
            host="psql-itlatam-logisticcontrol.postgres.database.azure.com",
            dbname="logistic-control",
            user="logisticpsqladmin@psql-itlatam-logisticcontrol",
            password="EsjHSrS69295NzHu342ap6P!N",
            sslmode="require"
        )

        cur = con.cursor()

        if bool(cell_range) is False:
            error.empty()
            error_chave += 1
        else:
            cur.execute(f"SELECT chave_acesso FROM material_management.master_saf_entrada WHERE chave_acesso = '{cell_range}'")
            pesquisa = cur.fetchall()
            if bool(pesquisa) is False:
                error.chave_bd()
                error_chave += 1
            else:
                try:
                    cell_range = int(cell_range)
                except:
                    error.chave()
                    error_chave += 1
                finally:
                    pass
                if len(str(cell_range)) != 44:
                    error.chave()
                    error_chave += 1
                for c in str(cell_range):
                    if str(cell_range).count(c) == 44:
                        error.chave()
                        error_chave += 1
                        break
                    else:
                        continue
                if len(str(cell_range)) == 44:
                    var = str(cell_range)[:43]
                    multiplicador = 2
                    somatoria = 0
                    for c in var[::-1]:
                        somatoria = (int(c) * multiplicador) + somatoria
                        if multiplicador == 9:
                            multiplicador = 2
                        else:
                            multiplicador += 1
                    if (somatoria % 11) == 1 or (somatoria % 11) == 0:
                        if str(cell_range)[-1] != '0':
                            error.chave_validador()
                            error_chave += 1
                    else:
                        if str(cell_range)[-1] != str(11 - (somatoria % 11)):
                            error.chave_validador()
                            error_chave += 1

        ### VALIDAÇÃO DO PEDIDO DE COMPRA (PO)

        cell_range = item['PedidoCompra']
        if bool(cell_range) is False:
            error_PO += 1
            error.empty()
        try:
            int_test = int(cell_range[1:])
        except:
            error.po()
            error_PO += 1
        finally:
            pass
        if 'K' in cell_range[0] or 'k' in cell_range[0]:
            if len(cell_range[1:]) > 5 or len(cell_range[1:]) < 5:
                error_PO += 1
                error.po()
            item['PedidoCompra'] = cell_range[1:]
        elif len(str(cell_range)) > 5 or len(str(cell_range)) < 5:
            error_PO += 1
            error.po()
        else:
            pass

        ### VALIDAÇÃO DO PART-NUMBER

        cell_range = item['PartNumber']
        if bool(cell_range) is False or cell_range is None:
            error.empty()
            error_PN += 1
        elif "!" in cell_range or \
                "@" in cell_range or \
                "$" in cell_range or \
                "%" in cell_range or \
                "&" in cell_range or \
                "*" in cell_range or \
                ")" in cell_range or \
                "'" in cell_range or \
                ":" in cell_range or \
                ";" in cell_range:
            error.part_number()
            error_PN += 1
        else:
            pass

        ### VALIDAÇÃO RFID DO PRODUTO

        cell_range = item['RFID_Produto']
        if bool(cell_range) is False:
            error.empty()
            error_RFID += 1
        elif len(cell_range) != 24:
            error.rfid()
            error_RFID += 1
        elif "E" != cell_range[0]:
            error.rfid()
            error_RFID += 1
        if repeticao_rfid.count(cell_range) > 1:
            error.rfid_repetido()
            error_RFID += 1
        else:
            pass

        ### VALIDAÇÃO DO SERIAL NUMBER

        cell_range = item['SerialNumber']
        if bool(cell_range) is False:
            error.empty()
            error_SN += 1
        if "!" in str(cell_range) or \
                "@" in str(cell_range) or \
                "$" in str(cell_range) or \
                "%" in str(cell_range) or \
                "&" in str(cell_range) or \
                "*" in str(cell_range) or \
                "(" in str(cell_range) or \
                ")" in str(cell_range) or \
                "'" in str(cell_range) or \
                ":" in str(cell_range) or \
                "/" in str(cell_range):
            error.serial_number()
            error_SN += 1
        if repeticao_serial.count(cell_range) > 1:
            error.serial_number_repetido()
            error_RFID += 1
        else:
            pass

        ### VALIDAÇÃO LOCAL

        cell_range = str(item['Local']).strip()
        if cell_range == 'TERCA VIX' or cell_range == 'AGS RIO' or cell_range == 'NEXUS SAO':
            pass
        else:
            error.local()
            error_local += 1

        ### VALIDAÇÃO DA DATA

        cell_range = item['DataEvidencia']

        try:
            parse(cell_range)
            data = parse(cell_range).strptime(cell_range, '%d/%m/%Y')
            # if data.day <= 12:
            #     data = datetime.strptime(datetime.strftime(parse(cell_range), "%m/%d/%Y"), "%d/%m/%Y")
            if data > datetime.today():
                error.data_maior()
                error_Date += 1
            else:
                pass
        except Exception as erros:
            print(erros)
            error.data()
            error_Date += 1
        finally:
            pass

        ### CHAVE DE RELACIONAMENTO

        # Select na tabela Recebimento com base na data
        # item['ChaveRelacionamento'] = str(item['RFID_Produto']).strip() + str(item['Local']).strip()
        #
        # # Update connection string information
        # host = "psql-itlatam-logisticcontrol.postgres.database.azure.com"
        # dbname = "logistic-control"
        # user = "logisticpsqladmin@psql-itlatam-logisticcontrol"
        # password = "EsjHSrS69295NzHu342ap6P!N"
        # sslmode = "require"
        # # Construct connection string
        # conn_string = "host={0} user={1} dbname={2} password={3} sslmode={4}".format(host, user, dbname, password,
        #                                                                              sslmode)
        # conn = psycopg2.connect(conn_string)
        # print("Connection established")
        # cursor = conn.cursor()
        #
        # cursor.execute(f"SELECT data FROM public.recb_test WHERE chave_relacionamento = '{item['ChaveRelacionamento']}'")
        # pesquisa = cursor.fetchall()
        #
        # # data = datetime.date(data)
        #
        # if bool(pesquisa) is False:
        #     pass
        # else:
        #     for dado in pesquisa:
        #         for date_dado in dado:
        #             if parse(str(date_dado)) >= data:
        #                 error.chave_relacionamento()
        #                 error_ChaveRel += 1

        if error_chave > 0 or error_PO > 0 or error_PN > 0 \
                or error_RFID > 0 or error_SN > 0 or error_Date > 0 or error_ChaveRel:
            if dict[item['RFID_Produto']] in dict_error:
                dict_error[item['RFID_Produto']] += error.retornar()
            else:
                dict_error[item['RFID_Produto']] = error.retornar()

            dict_return_rec[item['ChaveNF_Entrada']] = dict_error
        # dict_error.clear()

    # if bool(dict_return) is False:
    #     Insert(lista).rec_insert()
    #     return dict_return
    # else:
    return dict_return_rec


def exp_validation(lista, repeticao_rfid):
    # Imports
    from datetime import datetime
    from Modulos.class_errosBD import Error
    from dateutil.parser import parse
    import psycopg2

    global error_chave, data

    dict_error = {}

    print('Início da validação - Expedição')

    for item in lista:

        error = Error()

        error_chave = 0
        error_PO = 0
        error_PN = 0
        error_RFID = 0
        error_SN = 0
        error_Date = 0
        error_ChaveRel = 0
        error_local = 0

        ### VALIDAÇÃO RFID DO PRODUTO

        cell_range = item['RFID_Produto']
        if bool(cell_range) is False:
            error.empty()
            error_RFID += 1
        elif len(cell_range) != 24:
            error.rfid()
            error_RFID += 1
        elif "E" != cell_range[0]:
            error.rfid()
            error_RFID += 1
        if repeticao_rfid.count(cell_range) > 1:
            error.rfid_repetido()
            error_RFID += 1
        else:
            pass

        ### VALIDAÇÃO DA CHAVE DE NOTA FISCAL
        cell_range = item['ChaveNF_Saida']

        con = psycopg2.connect(
            host="psql-itlatam-logisticcontrol.postgres.database.azure.com",
            dbname="logistic-control",
            user="logisticpsqladmin@psql-itlatam-logisticcontrol",
            password="EsjHSrS69295NzHu342ap6P!N",
            sslmode="require"
        )

        cur = con.cursor()

        if bool(cell_range) is False:
            error.empty()
            error_chave += 1
        else:
            cur.execute(
                f"SELECT chave_acesso FROM material_management.master_saf_saida WHERE chave_acesso = '{cell_range}'")
            pesquisa = cur.fetchall()
            if bool(pesquisa) is False:
                error.chave_bd()
                error_chave += 1
            else:
                try:
                    cell_range = int(cell_range)
                except:
                    error.chave()
                    error_chave += 1
                finally:
                    pass
                if len(str(cell_range)) != 44:
                    error.chave()
                    error_chave += 1
                for c in str(cell_range):
                    if str(cell_range).count(c) == 44:
                        error.chave()
                        error_chave += 1
                        break
                    else:
                        continue
                if len(str(cell_range)) == 44:
                    var = str(cell_range)[:43]
                    multiplicador = 2
                    somatoria = 0
                    for c in var[::-1]:
                        somatoria = (int(c) * multiplicador) + somatoria
                        if multiplicador == 9:
                            multiplicador = 2
                        else:
                            multiplicador += 1
                    if (somatoria % 11) == 1 or (somatoria % 11) == 0:
                        if str(cell_range)[-1] != '0':
                            error.chave_validador()
                            error_chave += 1
                    else:
                        if str(cell_range)[-1] != str(11 - (somatoria % 11)):
                            error.chave_validador()
                            error_chave += 1

        ### VALIDAÇÃO LOCAL

        cell_range = str(item['Local']).strip()
        if cell_range == 'TERCA VIX' or cell_range == 'AGS RIO' or cell_range == 'NEXUS SAO':
            pass
        else:
            error.local()
            error_local += 1

        ### VALIDAÇÃO DA DATA

        cell_range = item['DataEvidencia']

        try:
            parse(cell_range)
            data = parse(cell_range).strptime(cell_range, '%d/%m/%Y')
            # if data.day <= 12:
            #     data = datetime.strptime(datetime.strftime(parse(cell_range), "%m/%d/%Y"), "%d/%m/%Y")
            if data > datetime.today():
                error.data_maior()
                error_Date += 1
            else:
                pass
        except Exception as erros:
            print(erros)
            error.data()
            error_Date += 1
        finally:
            pass

        ### CHAVE DE RELACIONAMENTO

        # Select na tabela Recebimento com base na data
        item['ChaveRelacionamento'] = str(item['RFID_Produto']).strip() + str(item['Local']).strip()

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

        cursor.execute(
            f"SELECT data FROM public.exp_test WHERE chave_relacionamento = '{item['ChaveRelacionamento']}'")
        pesquisa = cursor.fetchall()

        # data = datetime.date(data)

        if bool(pesquisa) is False:
            pass
        else:
            for dado in pesquisa:
                for date_dado in dado:
                    if parse(str(date_dado)) >= data:
                        error.chave_relacionamento()
                        error_ChaveRel += 1

        if error_chave > 0 or error_PO > 0 or error_PN > 0 \
                or error_RFID > 0 or error_SN > 0 or error_Date > 0 or error_ChaveRel:
            if dict[item['RFID_Produto']] in dict_error:
                dict_error[item['RFID_Produto']] += error.retornar()
            else:
                dict_error[item['RFID_Produto']] = error.retornar()

            dict_return[item['ChaveNF_Saida']] = dict_error
        # dict_error.clear()

    if bool(dict_return) is False:
        Insert(lista).exp_insert()
        return dict_return
    else:
        return dict_return


class Insert:

    def __init__(self, lista):
        self.lista = lista

    def exp_insert(self):
        import psycopg2
        from datetime import datetime
        from dateutil.parser import parse

        print(self.lista)

        con = psycopg2.connect(
            host="psql-itlatam-logisticcontrol.postgres.database.azure.com",
            dbname="logistic-control",
            user="logisticpsqladmin@psql-itlatam-logisticcontrol",
            password="EsjHSrS69295NzHu342ap6P!N",
            sslmode="require"
        )

        cur = con.cursor()

        for item in self.lista:
            data = parse(item['DataEvidencia'])
            cur.execute(f"INSERT INTO public.exp_test ("
                        f"rfid,"
                        f"chave_nf,"
                        f"ov,"
                        f"local,"
                        f"data,"
                        f"usuario,"
                        f"obs,"
                        f"chave_relacionamento,"
                        f"data_lancamento"
                        f") "
                        f"VALUES ("
                        f"%s, %s, %s, %s, %s, %s, %s, %s, %s"
                        f")",
                        (
                            item['RFID_Produto'],
                            item['ChaveNF_Saida'],
                            item['OrdemVenda'],
                            item['Local'],
                            parse(item['DataEvidencia']),
                            item['Usuario(email)'],
                            item['ObsExpedicao'],
                            item['ChaveRelacionamento'],
                            datetime.today()
                        )
                        )
            con.commit()

        cur.close()
        con.close()

        print(self.lista)

    def rec_insert(self):

        import psycopg2
        from datetime import datetime

        print(self.lista)

        con = psycopg2.connect(
            host="psql-itlatam-logisticcontrol.postgres.database.azure.com",
            dbname="logistic-control",
            user="logisticpsqladmin@psql-itlatam-logisticcontrol",
            password="EsjHSrS69295NzHu342ap6P!N",
            sslmode="require"
        )

        cur = con.cursor()

        for item in self.lista:
            cur.execute(f"INSERT INTO public.recb_test ("
                                                        f"chave_nf,"
                                                        f"po,"
                                                        f"cx_master,"
                                                        f"part_number,"
                                                        f"rfid,"
                                                        f"serial,"
                                                        f"local,"
                                                        f"data,"
                                                        f"usuario,"
                                                        f"obs,"
                                                        f"chave_relacionamento,"
                                                        f"data_lancamento"
                                                        f") "
                        f"VALUES ("
                                 f"%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s"
                        f")",
                        (
                            item['ChaveNF_Entrada'],
                            item['PedidoCompra'],
                            item['RFID_CxMaster/TagAtivo'],
                            item['PartNumber'],
                            item['RFID_Produto'],
                            item['SerialNumber'],
                            item['Local'],
                            item['DataEvidencia'],
                            item['Usuario(email)'],
                            item['ObsRecebimento'],
                            item['ChaveRelacionamento'],
                            datetime.today()
                        )
                        )
            con.commit()

        cur.close()
        con.close()

        print(self.lista)
