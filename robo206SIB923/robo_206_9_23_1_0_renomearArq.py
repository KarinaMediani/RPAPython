import datetime
from openpyxl import load_workbook
import locale
import os 
import logging
import xml.etree.ElementTree as ET


def renomearArq():
    try:
        tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
        root2 = tree.getroot()
        for child2 in root2:
                for x2 in root2.findall(child2.tag):
                        moniop = x2.find('moniop').text
        data_atual = datetime.datetime.now()
        data = data_atual.date()
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        ano = data.strftime("%Y")
        mes = data.strftime("%m")
        anoMes = data.strftime("%Y%m") 
        caminho  =moniop


        lista_arquivo = os.listdir(caminho)

        lista_data = []
        for arquivo in lista_arquivo:#pegando arquivo mais recente
            if arquivo.endswith('.xlsm'):
                data = os.path.getmtime(f"{caminho}/{arquivo}")
                lista_data.append((arquivo))
        # print(lista_data)

        lista_data.sort(reverse=True)
        ultimo_arquivo = lista_data[0]#variavel com o arquivo recente
        # print(ultimo_arquivo)
        wb = load_workbook(caminho +  ultimo_arquivo)#abrindo arquivo mais recente
        # ws = wb['Sheet']


        # RENOMEANDO COM MES E ANO ATUAL
        data_atual = datetime.datetime.now()
        data = data_atual.date()
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        data_br = data.strftime("%m%Y")

        atual = str(ultimo_arquivo).split()

        atual.pop()#retirando data antiga
        atual.pop()#retirando data antiga

        atual.append(data_br)#adicionando data atual ao nome
        arquivo_atual = (atual[0] + ' ' + atual[1] + ' ' +  atual[2] + ' ' +  atual[3] + ' ' + atual[4] + ' ' + atual[5]+ ' ' + atual[6])
        # print(arquivo_atual)
        wb.save(caminho + '\\' + arquivo_atual + ".xlsm")
    except Exception as e:
        logging.error('| Ocorreu um erro: | 3')
        logging.exception(str(e))     
