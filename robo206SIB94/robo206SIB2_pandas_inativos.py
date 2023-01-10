import pandas as pd 
import os 
import logging
import xml.etree.ElementTree as ET
# #  2
def excluirInativos():
    try:
        tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
        root2 = tree.getroot()
        for child2 in root2:
                for x2 in root2.findall(child2.tag):
                        confeop = x2.find('confeop').text
        diretorio = confeop
        lista_arquivo = os.listdir(diretorio)

        for arquivo in lista_arquivo:
            if arquivo.startswith('ArqConf') and arquivo.endswith('.xlsx'):
                inativos = pd.read_excel(diretorio+'\\'+arquivo)
                inativos.dropna(subset=['EXCLUIR'],inplace=True)
                inativos.to_excel(diretorio+'\\'+arquivo,index=False)
    except Exception as e:
                logging.error('| Ocorreu um erro: | 3')
                logging.exception(str(e))    