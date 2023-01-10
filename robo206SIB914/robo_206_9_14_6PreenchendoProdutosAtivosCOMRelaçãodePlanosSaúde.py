# from openpyxl import load_workbook
# import logging
# import xml.etree.ElementTree as ET
# def ansop_X_ProdutosAtivos2():
#     try:
#         tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
#         root2 = tree.getroot()
#         for child2 in root2:
#                 for x2 in root2.findall(child2.tag):
#                         confeop = x2.find('confeop').text
#                         proativosopmedi = x2.find('proativosopmedi').text
#         diretorio_2 =proativosop
#         wb1 = load_workbook(diretorio_2)
#         # contador = 1

#         ws1 = wb1['Situação do Plano']
#         contador = 0
#         for row in ws1:
#             contador +=1 
#             if row[1].value == None:
#                 ws1[f'B{contador}'] = 'NÃO LOCALIZADO'
#         ws1['B1'] = 'SEGMENTAÇÃO ASSISTENCIAL'
#         wb1.save(diretorio_2)
#     except Exception as e:
#             logging.error(' | Ocorreu um erro: | 3 | '+ str(e))