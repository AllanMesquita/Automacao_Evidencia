def atualizar(tblPA, aba_tblPA, path, file_name, df_mastersaf, v17):

    ### IMPORTS
    from Modulos import validacao, atualizar_v2
    import openpyxl as xl
    from openpyxl.styles import PatternFill
    import os
    from datetime import datetime

    ###
    qtd_linhas_tblPA = len(aba_tblPA['A'])

    aba_tblPA[f'A{qtd_linhas_tblPA + 1}'] = file_name.split('_')[0]
    aba_tblPA[f'B{qtd_linhas_tblPA + 1}'] = 'UpdatePlanEstoque'
    aba_tblPA[f'C{qtd_linhas_tblPA + 1}'] = datetime.strftime(datetime.now(), '%d/%m/%Y %H:%M')
    aba_tblPA[f'E{qtd_linhas_tblPA + 1}'] = 'EmProcessamento'

    tblPA.save(
        "C:\\Users\\allan.mesquita\\OneDrive - NTT\\Privado\\GESTÃO DE ESTOQUE\\100 BcoDados\\003 Evidencias\\"
        "06 Lixeira\\Testes\\02 Tabela\\tblProcessamentoAutomacoes.xlsx")

    # print(file_name)
    wb = xl.open(path + file_name)
    wb.active
    sheet = wb.sheetnames
    aba = wb[sheet[0]]

    verif_evid = str(aba["A2"].value)  # Verificador do tipo de evidência

    # Variáveis

    type_evidencia = ""
    qtd_linhas = 0
    resultado = ''
    validacao_local = False

    while bool(aba[f'A{qtd_linhas + 1}'].value) is True:
        qtd_linhas += 1

    if bool(verif_evid) is False:
        print("Erro! Célula vazia")
        os.replace(path + file_name,
                   "C:\\Users\\allan.mesquita\\NTT\\@AM BR Services and Operations - Privado\\GESTÃO DE ESTOQUE\\"
                   "100 BcoDados\\003 Evidencias\\03 Erro\\" + file_name)
    elif len(verif_evid) == 44:
        type_evidencia = "Recebimento"
    # verificador de repetição de caracteres
    elif len(verif_evid) == 24 and "E" in verif_evid:
        type_evidencia = "Expedição"
    elif len(verif_evid) != 44:
        print("Erro! Diferente de 44.")
        os.replace(path + file_name,
                   "C:\\Users\\allan.mesquita\\NTT\\@AM BR Services and Operations - Privado\\GESTÃO DE ESTOQUE\\"
                   "100 BcoDados\\003 Evidencias\\03 Erro\\" + file_name)

    if type_evidencia == 'Recebimento':
        if aba['G2'].value == 'TERCA VIX' or aba['G2'].value == 'JR SAO' or aba['G2'].value == 'JR RIO' or aba['G2'].value == 'AGS RIO' or aba['G2'].value == 'NEXUS SAO':
            pass
        else:
            aba['G2'].fill = PatternFill(fill_type="solid", fgColor="FF0000")
            validacao_local = True
            resultado = 'Erro local'
    elif type_evidencia == 'Expedição':
        if aba['D2'].value == 'TERCA VIX' or aba['D2'].value == 'JR SAO' or aba['D2'].value == 'JR RIO' or aba['D2'].value == 'AGS RIO' or aba['D2'].value == 'NEXUS SAO':
            pass
        else:
            aba['D2'].fill = PatternFill(fill_type="solid", fgColor="FF0000")
            validacao_local = True
            resultado = 'Erro local'

    # inicio = now
    # print(now)

    if validacao_local is False:
        if type_evidencia == "Recebimento":
            resultado = validacao.rec_validation(aba, qtd_linhas)
        if type_evidencia == "Expedição":
            resultado = validacao.exp_validacao(aba, qtd_linhas)

    wb.save(path + file_name)

    if resultado == 'Erro nos dados' or resultado == 'Erro local':
        os.replace(path + file_name,
                   "C:\\Users\\allan.mesquita\\OneDrive - NTT\\Privado\\GESTÃO DE ESTOQUE\\100 BcoDados\\"
                   "003 Evidencias\\06 Lixeira\\Testes\\03 Erro\\" + file_name)
                    # Verificar o caso se haver o mesmo arquivo no destino

    elif resultado == 'Sucesso':
        atualizar_v2.popular_V17(aba, qtd_linhas, type_evidencia, df_mastersaf, v17)
        os.replace(path + file_name,
                   "C:\\Users\\allan.mesquita\\OneDrive - NTT\\Privado\\GESTÃO DE ESTOQUE\\100 BcoDados\\"
                   "003 Evidencias\\06 Lixeira\\Testes\\04 Fluig\\" + file_name)

    aba_tblPA[f'D{qtd_linhas_tblPA + 1}'] = datetime.strftime(datetime.now(), '%d/%m/%Y %H:%M')
    aba_tblPA[f'E{qtd_linhas_tblPA + 1}'] = resultado

    tblPA.save("C:\\Users\\allan.mesquita\\OneDrive - NTT\\Privado\\GESTÃO DE ESTOQUE\\100 BcoDados\\"
               "003 Evidencias\\06 Lixeira\\Testes\\02 Tabela\\tblProcessamentoAutomacoes.xlsx")

    return resultado
