import dateutil.parser


class Error:

    def __init__(self):
        self.lista = []
        self.retorno = ''
        self.dic_erros = {'Célula sem dado': 0,
                          'Chave de Nota Fiscal fora do padrão': 0,
                          'Número de PO fora do padrão': 0,
                          'Part Number com caractere especial': 0,
                          'RFID fora do padrão': 0,
                          'RFID repetido no arquivo': 0,
                          'RFID consta na V17': 0,
                          'Serial Number com caractere especial': 0,
                          'Serial Number com "13s" no início do serial': 0,
                          'Serial Number repetido no arquivo': 0,
                          'Serial Number consta na V17': 0,
                          'Data fora do padrão': 0,
                          'Data maior que a data atual': 0,
                          'Chave de Relacionamento consta na Tbl.Recebimento/Expedição': 0,
                          'Quantidade do RFID diferente da Nota Fiscal': 0
                          }

    def empty(self):
        erro = 'Célula sem dado'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def chave(self):
        erro = 'Chave de Nota Fiscal fora do padrão'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def po(self):
        erro = 'Número de PO fora do padrão'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def part_number(self):
        erro = 'Part Number com caractere especial'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def rfid(self):
        erro = 'RFID fora do padrão'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def rfid_repetido(self):
        erro = 'RFID repetido no arquivo'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def rfid_v17(self):
        erro = 'RFID consta na V17'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number(self):
        erro = 'Serial Number com caractere especial'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number_13s(self):
        erro = 'Serial Number com "13s" no início do serial'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number_repetido(self):
        erro = 'Serial Number repetido no arquivo'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number_v17(self):
        erro = 'Serial Number consta na V17'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def data(self):
        erro = 'Data fora do padrão'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def data_maior(self):
        erro = 'Data maior que a data atual'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def chave_relacionamento(self):
        erro = 'Chave de Relacionamento consta na Tbl.Recebimento/Expedição'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def quantidade(self):
        erro = 'Quantidade do RFID diferente da Nota Fiscal'
        self.dic_erros[erro] += 1
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def retornar(self):
        for erro in self.lista:
            if bool(self.retorno) is False:
                self.retorno = erro
            else:
                self.retorno = self.retorno + ' / ' + erro

        return self.retorno


class SaveError:
    def __init__(self, aba, linha, tipo, erros, file_name):
        self.aba = aba
        self.linha = linha
        self.tipo = tipo
        self.erros = erros
        self.file_name = file_name

    def connect(self):

        global chave_nf, data_evidencia, local, lctobd_data
        import psycopg2
        from datetime import datetime
        from dateutil.parser import parse

        con = psycopg2.connect(
            host="psql-itlatam-logisticcontrol.postgres.database.azure.com",
            dbname="logistic-control",
            user="logisticpsqladmin@psql-itlatam-logisticcontrol",
            password="EsjHSrS69295NzHu342ap6P!N",
            sslmode="require"
        )

        cur = con.cursor()

        # # chave_nf = self.aba[f'A{self.linha}'].value  # dado vazio?
        # chave_nf = 'NULL' if self.aba[f'A{self.linha}'].value is None else self.aba[f'A{self.linha}'].value
        # # local = self.aba[f'G{self.linha}'].value  # dado vazio
        # local = 'NULL' if self.aba[f'G{self.linha}'].value is None else self.aba[f'G{self.linha}'].value
        # # data_evidencia = datetime.strptime(self.aba[f'H{self.linha}'].value, '%d/%m/%Y')  # dado vazio
        # cell_range = self.aba[f'H{self.linha}'].value
        # if type(cell_range) is datetime:
        #     data_evidencia = datetime.strptime('01/01/2001', '%d/%m/%Y') if cell_range is None else datetime.strftime(cell_range, '%m/%d/%Y')
        # else:
        #     data_evidencia = datetime.strptime('01/01/2001', '%d/%m/%Y') if cell_range is None else datetime.strftime(datetime.strptime(str(cell_range), '%d/%m/%Y'), '%m/%d/%Y')
        # lctobd_data = datetime.strptime(str(self.aba[f'L{self.linha}'].value), '%d/%m/%Y %H:%M')

        if self.tipo == 'Recebimento':
            chave_nf = self.file_name if self.aba[f'A{self.linha}'].value is None else self.aba[f'A{self.linha}'].value
            local = 'NULL' if self.aba[f'G{self.linha}'].value is None else self.aba[f'G{self.linha}'].value
            # if type(self.aba[f'H{self.linha}'].value) is datetime:
            #     data_evidencia = datetime.strptime('01/01/2001', '%d/%m/%Y') if self.aba[f'H{self.linha}'].value is None else self.aba[f'H{self.linha}'].value
            # else:
            #     data_evidencia = datetime.strptime('01/01/2001', '%d/%m/%Y') if self.aba[f'H{self.linha}'].value is None else datetime.strptime(str(self.aba[f'H{self.linha}'].value), '%d/%m/%Y')
            try:
                parse(str(self.aba[f'H{self.linha}'].value))
                data = parse(str(self.aba[f'H{self.linha}'].value))
                if data.day <= 12:
                    data_evidencia = datetime.strptime(datetime.strftime(data, "%m/%d/%Y"), "%d/%m/%Y")
                else:
                    data_evidencia = data
            except dateutil.parser.ParserError:
                data_evidencia = datetime.strptime("01/01/2001", "%d/%m/%Y")
            lctobd_data = datetime.strptime(str(self.aba[f'L{self.linha}'].value), '%d/%m/%Y %H:%M')
        if self.tipo == 'Expedição':
            chave_nf = self.file_name if self.aba[f'B{self.linha}'].value is None else self.aba[f'B{self.linha}'].value
            local = 'NULL' if self.aba[f'D{self.linha}'].value is None else self.aba[f'D{self.linha}'].value
            # if type(self.aba[f'E{self.linha}'].value) is datetime:
            #     data_evidencia = datetime.strptime('01/01/2001', '%d/%m/%Y') if self.aba[f'E{self.linha}'].value is None else self.aba[f'E{self.linha}'].value
            # else:
            #     data_evidencia = datetime.strptime('01/01/2001', '%d/%m/%Y') if self.aba[f'E{self.linha}'].value is None else datetime.strptime(str(self.aba[f'E{self.linha}'].value), '%d/%m/%Y')
            try:
                parse(str(self.aba[f'E{self.linha}'].value))
                data = parse(str(self.aba[f'E{self.linha}'].value))
                if data.day <= 12:
                    data_evidencia = datetime.strptime(datetime.strftime(data, "%m/%d/%Y"), "%d/%m/%Y")
                else:
                    data_evidencia = data
            except dateutil.parser.ParserError:
                data_evidencia = datetime.strptime("01/01/2001", "%d/%m/%Y")
            lctobd_data = datetime.strptime(str(self.aba[f'I{self.linha}'].value), '%d/%m/%Y %H:%M')

        # if self.erros['Célula sem dado'] == 6:
        #     cur.execute(
        #         'INSERT INTO material_management.error_recebimento (tipo_evidencia, chave_nf, data_evidencia, local, '
        #         'erro, '
        #         'responsabilidade, quantidade_erros, data_processamento) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)',
        #         (self.tipo, chave_nf, data_evidencia, local, 'Linha em braco', 'NTT', '1', lctobd_data)
        #     )
        #     con.commit()
        if self.tipo == 'Recebimento' and self.erros['Célula sem dado'] >= 4:
            cur.execute(
                'INSERT INTO material_management.error_recebimento (tipo_evidencia, chave_nf, data_evidencia, local, '
                'erro,'
                'responsabilidade, quantidade_erros, data_processamento) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)',
                (self.tipo, chave_nf, str(data_evidencia), local, 'linha em branco', 'NTT', '1', lctobd_data)
            )
            con.commit()
        elif self.tipo == 'Expedição' and self.erros['Célula sem dado'] >= 2:
            cur.execute(
                'INSERT INTO material_management.error_recebimento (tipo_evidencia, chave_nf, data_evidencia, local,'
                'erro, responsabilidade, quantidade_erros, data_processamento) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)',
                (self.tipo, chave_nf, data_evidencia, local, 'linha em branco', 'NTT', '1', lctobd_data)
            )
            con.commit()
        else:
            for erro in self.erros.items():
                if erro[1] > 0:
                    if erro[0] == 'RFID consta na V17' or erro[0] == 'Serial Number consta na V17' or erro[0] == \
                            'Chave de Relacionamento consta na Tbl.Recebimento/Expedição':
                        cur.execute(
                            'INSERT INTO material_management.error_recebimento (tipo_evidencia, chave_nf, '
                            'data_evidencia, local, erro, '
                            'responsabilidade, quantidade_erros, data_processamento) VALUES (%s, %s, %s, %s, %s, %s, '
                            '%s, %s)',
                            (self.tipo, chave_nf, str(data_evidencia), local, erro[0], 'NTT', erro[1], lctobd_data)
                        )
                        con.commit()
                    else:
                        cur.execute(
                            'INSERT INTO material_management.error_recebimento (tipo_evidencia, chave_nf, '
                            'data_evidencia, local, erro, '
                            'responsabilidade, quantidade_erros, data_processamento) VALUES (%s, %s, %s, %s, %s, %s, '
                            '%s, %s)',
                            (self.tipo, chave_nf, str(data_evidencia), local, erro[0], 'Armazem', erro[1], lctobd_data)
                        )
                        con.commit()

        cur.close()
        con.close()
