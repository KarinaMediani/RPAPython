# from openpyxl import load_workbook
# import logging
# import xml.etree.ElementTree as ET
# # def ansop_x_Produtos2():
# #     try:
# tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
# root2 = tree.getroot()
# for child2 in root2:
#         for x2 in root2.findall(child2.tag):
#                 proativosopmedi = x2.find('proativosopmedi').text
#                 apcia = x2.find('apcia').text
# diretorio_2 = proativosop
# wb1 = load_workbook(diretorio_2)


# ws1 = wb1['Produtos Ativos']
# contador = 0
# for row in ws1.iter_rows(min_row=3):
#         print(row[1].value) 
#         contador +=1 
#         if row[1].value == None:
#                 ws1[f'B{contador}'] = 'NÃO LOCALIZADO'
# # ws1['B1'] = 'SEGMENTAÇÃO ASSISTENCIAL'
# # ws1['B2'] = ''
# #             wb1.save(diretorio_2)
# #     except Exception as e:
# #             logging.error(' | Ocorreu um erro: | 3 | '+ str(e))