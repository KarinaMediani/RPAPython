import logging
import xml.etree.ElementTree as ET
import os
from zipfile import ZipFile
def descompac():
    try:
        tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
        
        root2 = tree.getroot()
        for child2 in root2:
                for x2 in root2.findall(child2.tag):
                        confeop = x2.find('confeop').text

                        
#9.2.1 abertura de pasta "Arquivo Conferencia CIA 571 "e descompactação de arquivos 
        dir_cia571 = confeop
        lista_cia571 = os.listdir(dir_cia571)
            # print(lista_cia571)
        for arquivo in lista_cia571:
                # print(arquivo)
            if arquivo.startswith('ArqConf') and arquivo.endswith('.zip'):
                # print(arquivo)
                zip = ZipFile(dir_cia571+'\\'+arquivo)
                zip.extractall(dir_cia571)
    except Exception as e:
                logging.error('| Ocorreu um erro: | 3')
                logging.exception(str(e))    
# zip.close()
       


        