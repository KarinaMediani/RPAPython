# from openpyxl import Workbook
import os
from zipfile import ZipFile
import logging
import xml.etree.ElementTree as ET
import datetime
import locale

#VÁLIDO PARA AS TRATATIVAS BRADESCO SEGUROS 571 E BRADESCO OPERADORA 571#
        ## NÃO É UTILIZADO NA TRATATIVA DA MEDISERVICE ##

def rpa206_971():

        try:
                tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
                root2 = tree.getroot()
                for child2 in root2:
                        for x2 in root2.findall(child2.tag):
                                suop= x2.find('suop').text
                data_atual = datetime.datetime.now()
                data = data_atual.date()
                locale.setlocale(locale.LC_ALL, '')
                ano = data.strftime("%Y")
                mesAno = data.strftime("%m%Y")
                anoMes = data.strftime("%Y%m")
                #Abertura de arquivo SU05:
                path_su05 = suop
                lista_arquivos_su05 = os.listdir(path_su05)
                        # print(lista_arquivos_su05)
                lista_data = []
                for arquivo in lista_arquivos_su05:
                        if '.zip' in arquivo:
                                data = os.path.getmtime(f"{path_su05}/{arquivo}")
                                lista_data.append((arquivo))
                # print(arquivo)
                lista_data.sort(reverse=True)
                ultimo_arquivo = lista_data[0]
                # print(path_su05 + ultimo_arquivo) 
                zip = ZipFile(path_su05 + '\\' + ultimo_arquivo)
                zip.extractall(path_su05)
        except Exception as e:
                logging.error('| Ocorreu um erro: | 3')
                logging.exception(str(e))      
                        
        