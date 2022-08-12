class Pesquisa:

    def __init__(self, cur, con, pesquisa, arquivo, linha):
        self.cur = cur
        self.con = con
        self.pesquisa = pesquisa
        self.arquivo = arquivo
        self.linha = linha

    def destinatario(self):
        if bool(self.pesquisa) is False:
            id_destinatario = 3
            print(f"Dado vazio - código - {id_destinatario}")
        else:
            # Pesquisa na tabela 'destinatario'
            self.cur.execute(f"SELECT id FROM destinatario WHERE cnpj = '{self.pesquisa}'")
            id_destinatario = self.cur.fetchall()
            # Caso for encontrado
            if bool(id_destinatario):
                for lista in id_destinatario:
                    for dado in lista:
                        id_destinatario = dado
                        print(f"Dado destinatario encontrado - {id_destinatario}")  # print temporário
            # Caso não for encotrado
            else:
                # Coleta dos novos dados
                cnpj = str(self.arquivo.Range(f'D{self.linha}'))
                inscricao_estadual = str(self.arquivo.Range(f'E{self.linha}'))
                razao_social = str(self.arquivo.Range(f'F{self.linha}'))
                endereco = str(self.arquivo.Range(f'G{self.linha}'))
                bairro = str(self.arquivo.Range(f'H{self.linha}'))
                cep = str(self.arquivo.Range(f'I{self.linha}'))
                municipio = str(self.arquivo.Range(f'J{self.linha}'))
                uf = str(self.arquivo.Range(f'K{self.linha}'))

                # Comando para inserir os dados no banco
                self.cur.execute(
                    "INSERT INTO destinatario (cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                    (cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf))
                self.con.commit()
                print('Novos dados inseridos')  # print temporario

                # Pesquisa para coletar o 'id' do novo dado destinatario
                self.cur.execute(f"SELECT id FROM destinatario WHERE cnpj = '{self.pesquisa}'")
                id_destinatario = self.cur.fetchall()
                for lista in id_destinatario:
                    for dado in lista:
                        id_destinatario = dado
                        print(f"Dado destinatario coletado - {id_destinatario}")  # print temporario

        return id_destinatario

    def fornecedor(self):
        if bool(self.pesquisa) is False:
            id_fornecedor = 4
            print(f"Dado vazio - código - {id_fornecedor}")
        else:
            self.cur.execute(f"SELECT id FROM fornecedor WHERE cnpj = '{self.pesquisa}'")
            id_fornecedor = self.cur.fetchall()
            # Caso for encontrado
            if bool(id_fornecedor):
                for lista in id_fornecedor:
                    for dado in lista:
                        id_fornecedor = dado
                        print(f"Dado fornecedor encontrado - {id_fornecedor}")  # print temporário
            # Caso não for encotrado
            else:
                # Coleta dos novos dados
                cnpj = str(self.arquivo.Range(f'L{self.linha}'))

                inscricao_estadual = str(self.arquivo.Range(f'M{self.linha}'))
                # inscricao_estadual = inscricao_estadual if inscricao_estadual != 'ISENTO' else '00'
                if inscricao_estadual == 'ISENTO' or bool(inscricao_estadual) is False:
                    inscricao_estadual = '00'
                else:
                    inscricao_estadual = inscricao_estadual

                razao_social = str(self.arquivo.Range(f'N{self.linha}'))

                endereco = str(self.arquivo.Range(f'O{self.linha}'))
                endereco = endereco if bool(endereco) else '00'

                bairro = str(self.arquivo.Range(f'P{self.linha}'))
                bairro = bairro if bool(bairro) else '00'

                cep = str(self.arquivo.Range(f'Q{self.linha}'))
                cep = cep if bool(cep) else '00'

                municipio = str(self.arquivo.Range(f'R{self.linha}'))
                municipio = municipio if bool(municipio) else '00'

                uf = str(self.arquivo.Range(f'S{self.linha}'))
                uf = uf if bool(uf) else '00'

                # Comando para inserir os dados no banco
                self.cur.execute(
                    "INSERT INTO fornecedor (cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                    (cnpj, inscricao_estadual, razao_social, endereco, bairro, cep, municipio, uf))
                self.con.commit()
                print('Novos dados inseridos')  # print temporario

                # Pesquisa para coletar o 'id' do novo dado fornecedor
                self.cur.execute(f"SELECT id FROM fornecedor WHERE cnpj = '{self.pesquisa}'")
                id_fornecedor = self.cur.fetchall()
                for lista in id_fornecedor:
                    for dado in lista:
                        id_fornecedor = dado
                        print(f"Dado fornecedor coletado - {id_fornecedor}")  # print temporario

        return id_fornecedor

    def natureza_operacao(self):
        if bool(self.pesquisa) is False:
            id_natureza = 3
            print(f"Dado vazio - código - {id_natureza}")
        else:
            self.cur.execute(f"SELECT id FROM natureza_operacao_nfe WHERE natureza = '{self.pesquisa}'")
            id_natureza = self.cur.fetchall()
            # Caso for encontrado
            if bool(id_natureza):
                for lista in id_natureza:
                    for dado in lista:
                        id_natureza = dado
                        print(f"Dado natureza operacao encontrado - {id_natureza}")  # print temporário
            # Caso não for encotrado
            else:
                # Coleta dos novos dados

                # Comando para inserir os dados no banco
                self.cur.execute(f"INSERT INTO natureza_operacao_nfe (natureza) VALUES ('{self.pesquisa}')")
                self.con.commit()
                print('Novos dados inseridos')  # print temporario

                # Pesquisa para coletar o 'id' do novo dado fornecedor
                self.cur.execute(f"SELECT id FROM natureza_operacao_nfe WHERE natureza = '{self.pesquisa}'")
                id_natureza = self.cur.fetchall()
                for lista in id_natureza:
                    for dado in lista:
                        id_natureza = dado
                        print(f"Dado natureza operacao coletado - {id_natureza}")  # print temporario

        return id_natureza

    def cfop(self):
        if bool(self.pesquisa) is False:
            id_cfop = 3
            print(f"Dado vazio - código - {id_cfop}")
        else:
            self.cur.execute(f"SELECT id FROM cfop WHERE codigo = '{self.pesquisa}'")
            id_cfop = self.cur.fetchall()
            # Caso for encontrado
            if bool(id_cfop):
                for lista in id_cfop:
                    for dado in lista:
                        id_cfop = dado
                        print(f"Dado cfop encontrado - {id_cfop}")  # print temporário
            # Caso não for encotrado
            else:
                # Coleta dos novos dados

                # Comando para inserir os dados no banco
                self.cur.execute(f"INSERT INTO cfop (codigo) VALUES ('{self.pesquisa}')")
                self.con.commit()
                print('Novos dados inseridos')  # print temporario

                # Pesquisa para coletar o 'id' do novo dado fornecedor
                self.cur.execute(f"SELECT id FROM cfop WHERE codigo = '{self.pesquisa}'")
                id_cfop = self.cur.fetchall()
                for lista in id_cfop:
                    for dado in lista:
                        id_cfop = dado
                        print(f"Dado cfop coletado - {id_cfop}")  # print temporario

        return id_cfop

    def transportadora(self):
        if bool(self.pesquisa) is False:
            id_transportadora = 2
            print(f"Dado vazio - código - {id_transportadora}")
        else:
            self.cur.execute(f"SELECT id FROM transportadora WHERE cnpj = '{self.pesquisa}'")
            id_transportadora = self.cur.fetchall()
            # Caso for encontrado
            if bool(id_transportadora):
                for lista in id_transportadora:
                    for dado in lista:
                        id_transportadora = dado
                        print(f"Dado transportadora encontrado - {id_transportadora}")  # print temporário
            # Caso não for encotrado
            else:
                # Coleta dos novos dados
                cnpj_transportadora = str(self.arquivo.Range(f'AH{self.linha}'))
                razao_social_transportadora = str(self.arquivo.Range(f'AI{self.linha}'))
                endereco_transportadora = str(self.arquivo.Range(f'AJ{self.linha}'))
                municipio_transportadora = str(self.arquivo.Range(f'AK{self.linha}'))
                uf_transportadora = str(self.arquivo.Range(f'AL{self.linha}'))

                # Comando para inserir os dados no banco
                self.cur.execute(f"INSERT INTO transportadora (cnpj, razao_social, endereco, municipio, uf) VALUES (%s, %s, %s, %s, %s)", (cnpj_transportadora, razao_social_transportadora, endereco_transportadora, municipio_transportadora, uf_transportadora))
                self.con.commit()
                print('Novos dados inseridos')  # print temporario

                # Pesquisa para coletar o 'id' do novo dado transportadora
                self.cur.execute(f"SELECT id FROM transportadora WHERE cnpj = '{self.pesquisa}'")
                id_transportadora = self.cur.fetchall()
                for lista in id_transportadora:
                    for dado in lista:
                        id_transportadora = dado
                        print(f"Dado transportadora coletado - {id_transportadora}")  # print temporario

        return id_transportadora
