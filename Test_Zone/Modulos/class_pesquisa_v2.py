class Pesquisa:

    def __init__(self, cur, conn, pesquisa, json):
        self.cur = cur
        self.conn = conn
        self.pesquisa = pesquisa
        self.json = json

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
            self.cur.execute(f"SELECT cnpj FROM public.dados_juridicos WHERE cnpj = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        destinatario = dado
            else:
                # DADO FORA DO BANCO
                cnpj = self.json['CNPJ do DestinatÃ¡rio']
                inscricao_estadual = self.json['InscriÃ§Ã£o Estadual DestinatÃ¡rio']
                razao_social = self.json['RazÃ£o Social DestinatÃ¡rio']
                endereco = self.json['EndereÃ§o DestinatÃ¡rio']
                bairro = self.json['Bairro DestinatÃ¡rio']
                cep = self.json['CEP DestinatÃ¡rio']
                municipio = self.json['MunicÃ\xadpio DestinatÃ¡rio']
                uf = self.json['UF DestinatÃ¡rio']

                if cnpj in cnpjs:
                    classificacao = 'NTT'

                self.cur.execute(
                    "INSERT INTO public.dados_juridicos "
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
            self.cur.execute(f"SELECT cnpj FROM public.dados_juridicos WHERE cnpj = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        fornecedor = dado
            else:
                # DADO FORA DO BANCO
                cnpj = self.json['CNPJ/CPF do Fornecedor']
                inscricao_estadual = self.json['InscriÃ§Ã£o Estadual Fornecedor']
                razao_social = self.json['RazÃ£o Social Fornecedor']
                endereco = self.json['EndereÃ§o Fornecedor']
                bairro = self.json['Bairro Fornecedor']
                cep = self.json['CEP Fornecedor']
                municipio = self.json['MunicÃ\xadpio Fornecedor']
                uf = self.json['UF Fornecedor']

                if cnpj in cnpjs:
                    classificacao = 'NTT'

                self.cur.execute(
                    "INSERT INTO public.dados_juridicos "
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
            id_natureza = 'NULL'
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT id FROM public.natureza_cfop WHERE natureza = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        id_natureza = dado
            else:
                # DADO FORA DO BANCO
                natureza = self.json['Natureza da OperaÃ§Ã£o']
                cfop = self.json['CFOP']

                self.cur.execute(
                    "INSERT INTO public.natureza_cfop "
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

                self.cur.execute(f"SELECT id FROM public.natureza_cfop WHERE natureza = '{self.pesquisa}'")
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
            self.cur.execute(f"SELECT cnpj FROM public.dados_juridicos WHERE cnpj = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        transportadora = dado
            else:
                # DADO FORA DO BANCO
                cnpj = self.json['Cnpj Transportadora']
                inscricao_estadual = ""
                razao_social = self.json['RazÃ£o Social Transportadora']
                endereco = self.json['EndereÃ§o Transportadora']
                bairro = ""
                cep = ""
                municipio = self.json['MunicÃ\xadpio Transportadora']
                uf = self.json['UF Transportadora']

                self.cur.execute(
                    "INSERT INTO public.dados_juridicos "
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
            id_cfop = 'NULL'
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT natureza_cfop FROM public.nf_entrada2 WHERE chave_acesso = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        id_cfop = dado
            else:
                # DADO FORA DO BANCO
                natureza = self.json['Natureza da OperaÃ§Ã£o']
                cfop = self.json['CFOP']

                self.cur.execute(
                    "INSERT INTO public.natureza_cfop "
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

                self.cur.execute(f"SELECT id FROM public.natureza_cfop WHERE natureza = '{natureza}'")
                id_natureza = self.cur.fetchall()
                for lista in id_natureza:
                    for dado in lista:
                        id_cfop = dado

        return id_cfop


    def cfop_saida(self):
        id_cfop = ''
        # DADO VAZIO
        if bool(self.pesquisa) is False:
            id_cfop = 'NULL'
        else:
            # DADO NO BANCO
            self.cur.execute(f"SELECT natureza_cfop FROM public.nf_saida WHERE chave_acesso = '{self.pesquisa}'")
            resultado = self.cur.fetchall()
            if bool(resultado):
                for lista in resultado:
                    for dado in lista:
                        id_cfop = dado
            else:
                # DADO FORA DO BANCO
                natureza = self.json['Natureza da OperaÃ§Ã£o']
                cfop = self.json['CFOP']

                self.cur.execute(
                    "INSERT INTO public.natureza_cfop "
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

                self.cur.execute(f"SELECT id FROM public.natureza_cfop WHERE natureza = '{natureza}'")
                id_natureza = self.cur.fetchall()
                for lista in id_natureza:
                    for dado in lista:
                        id_cfop = dado

        return id_cfop