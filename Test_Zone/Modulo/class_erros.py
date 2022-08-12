class Error:

    def __init__(self):
        self.lista = []
        self.retorno = ''

    def empty(self):
        erro = 'Célula sem dado'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def chave(self):
        erro = 'Chave de Nota Fiscal fora do padrão'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def po(self):
        erro = 'Número de PO fora do padrão'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def part_number(self):
        erro = 'Part Number com caractere especial'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def rfid(self):
        erro = 'RFID fora do padrão'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def rfid_repetido(self):
        erro = 'RFID repetido no arquivo'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def rfid_v17(self):
        erro = 'RFID consta na V17'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number(self):
        erro = 'Serial Number com caractere especial'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number_13s(self):
        erro = 'Serial Number com "13s" no início do serial'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number_repetido(self):
        erro = 'Serial Number repetido no arquivo'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def serial_number_v17(self):
        erro = 'Serial Number consta na V17'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def data(self):
        erro = 'Data fora do padrão'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def data_maior(self):
        erro = 'Data maior que a data atual'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def chave_relacionamento(self):
        erro = 'Chave de Relacionamento consta na Tbl.Recebimento/Expedição'
        if erro in self.lista:
            pass
        else:
            self.lista.append(erro)

    def quantidade(self):
        erro = 'Quantidade do RFID diferente da Nota Fiscal'
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
