class Pesquisa:

    def __init__(self, cur, conn, pesquisa, json, linha):
        self.cur = cur
        self.conn = conn
        self.pesquisa = pesquisa
        self.json = json
        self.linha = linha

    def destinatario(self):
        destinatario = ''
        classificacao = 'Cliente'
        cnpjs = [
            '00447484000111',
            '05437734000156',
            '31546914000186',
            '00447484000200',
            '05437734000318',
            '05437734000407',
            '05437734000580',
            '00447484000626',
            '05437734000660'
            ]

        # DADO VAZIO
        if bool(self.pesquisa) is False:
            destinatario = 'NULL'
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT cnpj FROM material_management.dados_juridicos WHERE cnpj = '{str(self.pesquisa)}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        destinatario = dado
            else:
                # DADO FORA DO BANCO
                cnpj = self.json[f'D{self.linha}'].value
                inscricao_estadual = self.json[f'E{self.linha}'].value
                razao_social = self.json[f'F{self.linha}'].value
                endereco = self.json[f'G{self.linha}'].value
                bairro = self.json[f'H{self.linha}'].value
                cep = self.json[f'I{self.linha}'].value
                municipio = self.json[f'J{self.linha}'].value
                uf = self.json[f'K{self.linha}'].value

                if cnpj in cnpjs:
                    classificacao = 'NTT'

                self.cur.execute(
                    "INSERT INTO material_management.dados_juridicos "
                    "("
                    "classificacao, cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf"
                    ")"
                    "VALUES "
                    "("
                    "%s, %s, %s, %s, %s, %s, %s, %s, %s"
                    ")",
                    (
                        classificacao, cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf
                    )
                )
                self.conn.commit()
                destinatario = cnpj

        return destinatario

    def fornecedor(self):
        fornecedor = ''
        classificacao = 'Fornecedor'
        cnpjs = [
            '00447484000111',
            '05437734000156',
            '31546914000186',
            '00447484000200',
            '05437734000318',
            '05437734000407',
            '05437734000580',
            '00447484000626',
            '05437734000660'
        ]

        # DADO VAZIO
        if bool(self.pesquisa) is False:
            fornecedor = 'NULL'
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT cnpj FROM material_management.dados_juridicos WHERE cnpj = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        fornecedor = dado
            else:
                # DADO FORA DO BANCO
                cnpj = self.json[f'L{self.linha}'].value
                inscricao_estadual = self.json[f'M{self.linha}'].value
                razao_social = self.json[f'N{self.linha}'].value
                endereco = self.json[f'O{self.linha}'].value
                bairro = self.json[f'P{self.linha}'].value
                cep = self.json[f'Q{self.linha}'].value
                municipio = self.json[f'R{self.linha}'].value
                uf = self.json[f'S{self.linha}'].value

                if cnpj in cnpjs:
                    classificacao = 'NTT'

                self.cur.execute(
                    "INSERT INTO material_management.dados_juridicos "
                    "("
                    "classificacao, cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf"
                    ")"
                    "VALUES "
                    "("
                    "%s, %s, %s, %s, %s, %s, %s, %s, %s"
                    ")",
                    (
                        classificacao, cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf
                    )
                )
                self.conn.commit()
                fornecedor = cnpj

        return fornecedor

    def natureza(self):
        id_natureza = ''
        # DADO VAZIO
        if bool(self.pesquisa) is False:
            id_natureza = 1
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT id FROM material_management.natureza_cfop WHERE natureza = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        id_natureza = dado
            else:
                # DADO FORA DO BANCO
                natureza = self.json[f'T{self.linha}'].value
                cfop = self.json[f'U{self.linha}'].value

                self.cur.execute(
                    "INSERT INTO material_management.natureza_cfop "
                    "("
                    "natureza, cfop"
                    ")"
                    "VALUES "
                    "("
                    "%s, %s"
                    ")",
                    (
                        natureza, cfop
                    )
                )
                self.conn.commit()

                self.cur.execute(f"SELECT id FROM material_management.natureza_cfop WHERE natureza = '{self.pesquisa}'")
                id_natureza = self.cur.fetchall()
                for lista in id_natureza:
                    for dado in lista:
                        id_natureza = dado

        return id_natureza

    def transportadora(self):
        transportadora = ''
        # DADO VAZIO
        if bool(self.pesquisa) is False:
            transportadora = 'NULL'
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT cnpj FROM material_management.dados_juridicos WHERE cnpj = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        transportadora = dado
            else:
                # DADO FORA DO BANCO
                cnpj = self.json[f'AH{self.linha}'].value
                inscricao_estadual = ""
                razao_social = self.json[f'AI{self.linha}'].value
                endereco = self.json[f'AJ{self.linha}'].value
                bairro = ""
                cep = ""
                municipio = self.json[f'AK{self.linha}'].value
                uf = self.json[f'AL{self.linha}'].value

                self.cur.execute(
                    "INSERT INTO material_management.dados_juridicos "
                    "("
                    "classificacao, cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf"
                    ")"
                    "VALUES "
                    "("
                    "%s, %s, %s, %s, %s, %s, %s, %s, %s"
                    ")",
                    (
                        'Transportadora', cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf
                    )
                )
                self.conn.commit()
                transportadora = cnpj

        return transportadora

    def transportadora_saida(self):
        transportadora = ''
        # DADO VAZIO
        if bool(self.pesquisa) is False:
            transportadora = 'NULL'
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT cnpj FROM material_management.dados_juridicos WHERE cnpj = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        transportadora = dado
            else:
                # DADO FORA DO BANCO
                cnpj = self.json[f'AG{self.linha}'].value
                inscricao_estadual = ""
                razao_social = self.json[f'AH{self.linha}'].value
                endereco = self.json[f'AI{self.linha}'].value
                bairro = ""
                cep = ""
                municipio = self.json[f'AJ{self.linha}'].value
                uf = self.json[f'AK{self.linha}'].value

                self.cur.execute(
                    "INSERT INTO material_management.dados_juridicos "
                    "("
                    "classificacao, cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf"
                    ")"
                    "VALUES "
                    "("
                    "%s, %s, %s, %s, %s, %s, %s, %s, %s"
                    ")",
                    (
                        'Transportadora', cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf
                    )
                )
                self.conn.commit()
                transportadora = cnpj

        return transportadora

    def cfop(self):
        id_cfop = ''
        # DADO VAZIO
        if bool(self.pesquisa) is False:
            id_cfop = 1
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT natureza_cfop FROM material_management.master_saf_entrada WHERE chave_acesso = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        id_cfop = dado
            else:
                # DADO FORA DO BANCO
                natureza = self.json[f'D{self.linha}'].value
                cfop = self.json[f'AV{self.linha}'].value

                self.cur.execute(
                    "INSERT INTO material_management.natureza_cfop "
                    "("
                    "natureza, cfop"
                    ")"
                    "VALUES "
                    "("
                    "%s, %s"
                    ")",
                    (
                        natureza, cfop
                    )
                )
                self.conn.commit()

                self.cur.execute(f"SELECT id FROM material_management.natureza_cfop WHERE natureza = '{natureza}'")
                id_natureza = self.cur.fetchall()
                for lista in id_natureza:
                    for dado in lista:
                        id_cfop = dado

        return id_cfop


    def cfop_saida(self):
        id_cfop = ''
        # DADO VAZIO
        if bool(self.pesquisa) is False:
            id_cfop = 1
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT natureza_cfop FROM material_management.master_saf_saida WHERE chave_acesso = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        id_cfop = dado
            else:
                # DADO FORA DO BANCO
                natureza = self.json[f'D{self.linha}'].value
                cfop = self.json[f'AV{self.linha}'].value

                self.cur.execute(
                    "INSERT INTO material_management.natureza_cfop "
                    "("
                    "natureza, cfop"
                    ")"
                    "VALUES "
                    "("
                    "%s, %s"
                    ")",
                    (
                        natureza, cfop
                    )
                )
                self.conn.commit()

                self.cur.execute(f"SELECT id FROM material_management.natureza_cfop WHERE natureza = '{natureza}'")
                id_natureza = self.cur.fetchall()
                for lista in id_natureza:
                    for dado in lista:
                        id_cfop = dado

        return id_cfop
