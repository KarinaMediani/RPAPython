import logging
import os
from openpyxl import Workbook
import xml.etree.ElementTree as ET
def final():
    try:
        tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
        root2 = tree.getroot()
        for child2 in root2:
                for x2 in root2.findall(child2.tag):
                        inifim = x2.find('inifim').text
                        
        with open(inifim+"\\finalizado.txt" ,'w') as arquivo:
            print('')
    except Exception as e:
        logging.error('| Ocorreu um erro: | 3')
        logging.exception(str(e))