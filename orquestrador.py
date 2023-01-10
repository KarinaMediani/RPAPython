import time
import logging
import locale
import datetime
import robo206SIB92.robo206_inicio
import robo206SIB92.robo206SIBDescompactando
import robo206SIB93.robo206SIB1_txtPARAexcel #1°
import robo206SIB93.robo206SIB2_exclusão_linha_inativo #2°
import robo206SIB94.robo206SIB1_cabecalho
import robo206SIB94.robo206SIB2_pandas_inativos
import robo206SIB94.robo206SIB3_colunas
import robo206SIB94.robo206SIB4_editarCabecalho
import robo206SIB96.robo206SIB_9_6_1
import robo206SIB96.robo206SIB_9_6_2
import robo206SIB96.robo206SIB_9_6_3
import robo206SIB97.robo2069_7_1Descompac
import robo206SIB97.robo206SIB97SU05
import robo206SIB98.robo206SIB981
import robo206SIB98.robo206SIB982
import robo206SIB99.robo206SIB99
import robo206SIB910.robo206SIB_9_10
import robo206SIB911.roboSIB206_9_11_somas
import robo206SIB911.roboSIB206_9_11_ramos
import robo206SIB911.robo206_9_11_47.rpa_206_9_11_47_grupo
import robo206SIB911.robo206_9_11_48.rpa_206_9_11_48_extração_pdf_cpf_repetido
import robo206SIB911.robo206_9_11_49.rpa_206_9_11_49_IDSS_
import robo206SIB911.robo206_9_11_50.rpa_206_9_11_50_bon_dep_menor_valid
import robo206SIB912.rpa_206_9_12_trat_bases_aba_total
import robo206SIB913.robo_206_9_13_1PreenchendoProdutosAtivos
import robo206SIB913.robo_206_9_13_10PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúde
import robo206SIB913.robo_206_9_13_10PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúdeP2
import robo206SIB913.robo_206_9_13_13TratandoValores
import robo206SIB914.robo_206_9_14_1PassandoValoresDeUmaAbaParaOutra
import robo206SIB914.robo_206_9_14_4PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúde
import robo206SIB914.robo_206_9_14_6PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúde
import robo206SIB915.robo_206_9_15PassandoLinhasComConteudoVazio
import robo206SIB916.robo_206_9_16_1ComparacaoDeMesesAnteriores
import robo206SIB916.robo_206_9_16_10ComparacaoDeMesesAnteriores
import robo206SIB917.robo_206_9_17TratandoAbaOdontologico
import robo206SIB918.robo_206_9_18ProdutosCancelados
import robo206SIB923.robo_206_9_23_1_0_renomearArq
import robo206SIB923.robo_206_9_23_1_4_sibxdw
import robo206SIB923.robo_206_9_23_2_sibxdw2
import robo206SIB923.robo_206_9_23_3_sibxdw3
import robo206SIB923.robo_206_9_23_6_sibxdw6
import robo206SIB923.robo_206_9_23_10_0_sibxdw100
import robo206SIB923.robo_206_9_23_10_1_sibxdw101
import robo206SIB923.robo_206_9_23_11_0_sibxdw110
import robo206SIB923.robo_206_9_23_11_1_sibxdw111
import robo206SIB923.robo_206_9_23_14_sibxdw14
import robo206SIB923.robo_206_9_23_17_sibxdw17
import robo206SIB923.robo206_finalizacao
# import robo206SIB923.robo206_importando
import xml.etree.ElementTree as ET


tree = ET.parse("C:\\Users\\User\\Desktop\\206-BS-571\\diretório\\diretorio.xml")
root2 = tree.getroot()
for child2 in root2:
        for x2 in root2.findall(child2.tag):
                log = x2.find('log').text




locale.setlocale(locale.LC_ALL, '')
# RENOMEANDO COM MES E ANO ATUAL
data_atual = datetime.datetime.now()
data_br = data_atual.strftime("%H_%M_%S_")
dt = data_atual.strftime("%d_%m_%y_")
# logging.info('| ANDAMENTO: |'(data_br) 
log_format = '%(asctime)s:%(levelname)s:%(filename)s:%(message)s'

logging.basicConfig(filename= log + "\\" + data_br + dt +'RDA206_SEGUROS_INDICADORES_SIB_ans_' + '.log',
                    # w -> sobrescreve o arquivo a cada log
                    # a -> não sobrescreve o arquivo
                    encoding='utf-8',
                    filemode='a',
                    level=logging.INFO,
                    format=log_format)


try:
    if __name__ == "__main__":
        # start_time = time.time()
        logging.info('| INICIO: | RDA206_APURAÇÃO_DO_ARQUIVO_CONFERENCIA_SIB')
        # robo206SIB92.robo206_inicio.inicio()
        logging.info('| STATUS: |  1')
        logging.info('| INICIO: | INICIO DAS TRATATIVAS BRADESCO SEGUROS')
        logging.info('| STATUS DEMANDA: |  1')
        logging.info('| ANDAMENTO: | DESCOMPACTANDO ARQUIVOS CONFERENCIA')
        robo206SIB92.robo206SIBDescompactando.descompac()
        logging.info('| ANDAMENTO: | TRANSFORMANDO ARQUIVO CONFERENCIA TXT EM XLSX')
        robo206SIB93.robo206SIB1_txtPARAexcel.para_excel()
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS ARQUIVO CONFERENCIA XLSX')
        robo206SIB93.robo206SIB2_exclusão_linha_inativo.exlusao()
        robo206SIB94.robo206SIB1_cabecalho.parametro()
        robo206SIB94.robo206SIB2_pandas_inativos.excluirInativos()
        logging.info('| ANDAMENTO: | SEPARANDO COLUNAS')
        robo206SIB94.robo206SIB3_colunas.separandoColunas()
        robo206SIB94.robo206SIB4_editarCabecalho.editarCabecalho()
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS DOS ARQ CONFERENCIA')
        robo206SIB96.robo206SIB_9_6_1.conf()
        robo206SIB96.robo206SIB_9_6_2.conf()
        robo206SIB96.robo206SIB_9_6_3.conf()
        logging.info('| ANDAMENTO: | FINALIZOU AS PRIMEIRAS TRATATIVAS DOS ARQ CONFERENCIA')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | DESCOMPACTANDO ARQUIVO SU05')
        robo206SIB97.robo2069_7_1Descompac.rpa206_971()
        logging.info('| ANDAMENTO: | TRATATIVAS COM ARQUIVO SU05')
        robo206SIB97.robo206SIB97SU05.rpa206_975()
        logging.info('| ANDAMENTO: |  FINALIZOU TRATATIVAS UTILIZANDO ARQ SU05')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | TRATATIVAS ARQ CONFERENCIA COM XML')
        robo206SIB98.robo206SIB981.conf()
        robo206SIB98.robo206SIB982.confXxml()
        logging.info('| ANDAMENTO: |  FINALIZOU TRATATIVAS UTILIZANDO ARQ XML')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: |  CRIANDO NOVAS ABAS ARQ CONFERENCIA E REALIZANDO NOVAS TRATATIVAS')
        robo206SIB99.robo206SIB99.conf()
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS COM ARQ CONFERENCIA E APURAÇÃO')
        robo206SIB910.robo206SIB_9_10.conf()
        logging.info('| ANDAMENTO: | FINALIZOU AS TRATATIVAS COM ARQ CONFERENCIA E APURAÇÃO')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS PLANILHA APURAÇÃO')
        robo206SIB911.roboSIB206_9_11_somas.confx()
        robo206SIB911.roboSIB206_9_11_ramos.ramos()
        logging.info('| ANDAMENTO: | PERCENTUAL POR GRUPO')
        robo206SIB911.robo206_9_11_47.rpa_206_9_11_47_grupo.grupo()
        logging.info('| ANDAMENTO: | CPF REPETIDO EXTRAIDO DE PDF')
        robo206SIB911.robo206_9_11_48.rpa_206_9_11_48_extração_pdf_cpf_repetido.extração_pdf_cpf_repetido()
        logging.info('| ANDAMENTO: | CPF REPETIDO EXTRAIDO DE PDF IDSS')
        robo206SIB911.robo206_9_11_49.rpa_206_9_11_49_IDSS_.extração_pdf_cpf_repetido()
        robo206SIB911.robo206_9_11_50.rpa_206_9_11_50_bon_dep_menor_valid.bon_dep_menor_valid()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS ARQUIVO APURAÇÃO BRADESCO')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS ENTRE OS ARQUIVOS APURAÇÃO BRADESCO X PRODUTOS ATIVOS')
        robo206SIB912.rpa_206_9_12_trat_bases_aba_total.prodAtivos_X_Apuracao()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS ENTRE OS ARQUIVOS APURAÇÃO BRADESCO X PRODUTOS ATIVOS')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS ENTRE OS ARQUIVOS CONFERENCIA X PRODUTOS ATIVOS')
        logging.info('| ANDAMENTO: | PEGANDO DADOS DA PLANILHA CONFERENCIA E ADICIONANDO NA PLANILHA PRODUTOS ATIVOS')
        robo206SIB913.robo_206_9_13_1PreenchendoProdutosAtivos.Confe_x_Produtos()
        logging.info('| ANDAMENTO: | PEGANDO DADOS DA PLANILHA ansop E ADICIONANDO NA PLANILHA PRODUTOS ATIVOS')
        robo206SIB913.robo_206_9_13_10PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúde.ansop_X_Produtos()
        # robo206SIB913.robo_206_9_13_10PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúdeP2.ansop_x_Produtos2()
        logging.info('| ANDAMENTO: | TRATATIVAS COM A PLANILHA PRODUTOS ATIVOS')
        robo206SIB913.robo_206_9_13_13TratandoValores.ProdutosAtivos()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS COM A PLANILHA PRODUTOS ATIVOS')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVOS ABA PRODUTOS ATIVOS')
        robo206SIB914.robo_206_9_14_1PassandoValoresDeUmaAbaParaOutra.AbaProdutosAtivos()
        logging.info('| ANDAMENTO: | COMPARANDO ABA PRODUTOS ATIVOS COM ARQUIVO RELAÇÃO DE PLANOS SAÚDE ansop')
        robo206SIB914.robo_206_9_14_4PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúde.ansop_X_ProdutosAtivos()
        # robo206SIB914.robo_206_9_14_6PreenchendoProdutosAtivosCOMRelaçãodePlanosSaúde.ansop_X_ProdutosAtivos2()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVOS ABA PRODUTOS ATIVOS')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVO ABA SEGURADOS SEM PRODUTOS ATIVO')
        robo206SIB915.robo_206_9_15PassandoLinhasComConteudoVazio.linhas_Vazias()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVOS ABA SEGURADOS SEM PRODUTOS ATIVO')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVO ABA COMPARAÇÃO DE MESES ANTERIORES')
        robo206SIB916.robo_206_9_16_1ComparacaoDeMesesAnteriores.meses_Anteriores()
        robo206SIB916.robo_206_9_16_10ComparacaoDeMesesAnteriores.meses_Anteriores2()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVOS ABA COMPARAÇÃO DE MESES ANTERIORES')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVO ABA PRODUTOS ODONTOLOGICOS')
        robo206SIB917.robo_206_9_17TratandoAbaOdontologico.odonto()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVOS ABA PRODUTOS ODONTOLOGICOS')
        logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS DEMANDA: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVO ABA PRODUTOS CANCELADOS')
        robo206SIB918.robo_206_9_18ProdutosCancelados.cancelados()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS DO ARQUIVO PRODUTOS ATIVOS ABA PRODUTOS CANCELADOS')
        logging.info('| STATUS DEMANDA: | 2')
        # logging.info('| INICIO: | FINALIZAÇÃO DAS TRATATIVAS BRADESCO SEGUROS')
        # logging.info('| STATUS DEMANDA: | 2')
        logging.info('| STATUS: | 1')
        logging.info('| ANDAMENTO: | INICIO DAS TRATATIVAS DA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS')
        logging.info('| ANDAMENTO: | RENOMEANDO PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS COM MES E ANO ATUAL')
        robo206SIB923.robo_206_9_23_1_0_renomearArq.renomearArq()
        logging.info('| ANDAMENTO: | ADICIONANDO NOVAS COLUNAS A SEREM TRATADAS')
        robo206SIB923.robo_206_9_23_1_4_sibxdw.sibxdw()
        logging.info('| ANDAMENTO: | RETIRANDO ZERO A ESQUERDA DA PLANILHA DW BRADESCO')
        robo206SIB923.robo_206_9_23_2_sibxdw2.sibxdw2()
        logging.info('| ANDAMENTO: | ADICIONANDO VALORES NA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS DA PLANILHA DW BRADESCO')
        robo206SIB923.robo_206_9_23_3_sibxdw3.sibxdw3()
        logging.info('| ANDAMENTO: | ADICIONANDO VALORES NA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS')
        robo206SIB923.robo_206_9_23_6_sibxdw6.sibxdw6()
        logging.info('| ANDAMENTO: | ADICIONANDO VALORES NA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS')
        robo206SIB923.robo_206_9_23_10_0_sibxdw100.sibxdw100()
        logging.info('| ANDAMENTO: | ADICIONANDO VALORES -100% NA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS')
        robo206SIB923.robo_206_9_23_10_1_sibxdw101.sibxdw101()
        logging.info('| ANDAMENTO: | ADICIONANDO VALORES NA COLUNA QUANTIDADES DE VIDAS')
        robo206SIB923.robo_206_9_23_11_0_sibxdw110.sibxdw110()
        robo206SIB923.robo_206_9_23_11_1_sibxdw111.sibxdw111()
        logging.info('| ANDAMENTO: | ADICIONANDO NOVOS VALORES NA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS')
        robo206SIB923.robo_206_9_23_14_sibxdw14.sibxdw14()
        logging.info('| ANDAMENTO: | ADICIONANDO NOVOS VALORES NA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS')
        robo206SIB923.robo_206_9_23_17_sibxdw17.sibxdw()
        logging.info('| ANDAMENTO: | FINALIZAÇÃO DAS TRATATIVAS DA PLANILHA PORCENTAGEM MONITORAMENTO SIB X DW BS')
        logging.info('| STATUS: | 2')
        logging.info('| ANDAMENTO: | FINALIZANDO ')
        # logging.info('| STATUS: | 1')
        # # logging.info('| ANDAMENTO: | TRABALHANDO COM VBA')
        # robo206SIB923.robo206_finalizacao.final()
        logging.info('| ANDAMENTO: | FINALIZOU RDA206_APURAÇÃO_DO_ARQUIVO_CONFERENCIA_SIB')
        logging.info('| STATUS: | 2')
        # end_time = time.time()
        # print(end_time-start_time)
except Exception as e: # work on python 3.x
    logging.error('| Ocorreu um erro: | 3')
    logging.exception(str(e))
 