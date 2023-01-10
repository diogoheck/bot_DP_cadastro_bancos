# ROBO FAP: PASSOS - 
import pyautogui
from time import sleep
from openpyxl import load_workbook



def carregar_planilha():
    pasta_banco = load_workbook('R:\Compartilhado\DP\cadastro bancos\cadastro_bancos.xlsx')
    planilha_banco = pasta_banco['BANCOS']
    cabecalho = True
    lista_contas = []
    for linha in planilha_banco.values:
        if not cabecalho:
        # print(type(linha[0]), type(linha[1]))
            dic_fap = {}
            # print(linha)
            # dic_fap[str(linha[0])] = str(linha[1])
            dic_fap['empresa'] = str(linha[0])
            dic_fap['colaborador'] = str(linha[1])
            dic_fap['banco'] = str(linha[2])
            dic_fap['agencia'] = str(linha[3])
            dic_fap['conta'] = str(linha[4])
            dic_fap['digito'] = str(linha[5])
            lista_contas.append(dic_fap)
        cabecalho = False
    return lista_contas


def fazer_lctos_unico_fap(lista_contas):
    # UNICO FOLHA
    # CADASTROS
    lista_empresas_cadastradas = []
    primeiro_acesso = True
    sleep(5)
    pyautogui.click(227,33, duration=2)
    # COLABORADORES
    pyautogui.move(0, 90, duration=2)
    sleep(2)
    pyautogui.click()
    sleep(10)
    for dic in lista_contas:
    #     CÓDIGO EMPRESA

        

        if dic['empresa'] not in lista_empresas_cadastradas:
            pyautogui.doubleClick(255,161, duration=2)
            sleep(2)
            pyautogui.typewrite(dic['empresa'])
            sleep(2)
            pyautogui.press('enter')
            sleep(10)

    
        # colaborador
        pyautogui.typewrite(dic['colaborador'])
        sleep(3)
        pyautogui.press('enter')
        sleep(4)
    

        # Aba folha
        pyautogui.click(641,678, duration=2)
        sleep(2)

        # Aba Banco
        pyautogui.click(515,653, duration=2)
        sleep(2)

        # clicar campo banco
        pyautogui.click(230,294, duration=2)
        sleep(1)

        pyautogui.typewrite(dic['banco'])
        sleep(1)
        pyautogui.press('enter')
        sleep(1)
        
        pyautogui.typewrite(dic['agencia'])
        sleep(1)
        pyautogui.press('enter')
        sleep(1)

        pyautogui.typewrite(dic['conta'])
        sleep(1)
        pyautogui.press('enter')
        sleep(1)

        pyautogui.typewrite(dic['digito'])
        sleep(1)
        pyautogui.press('enter')
        sleep(2)

        # conta salario
        pyautogui.click(273,380, duration=2)
        sleep(2)
        pyautogui.press('pagedown')
        sleep(1)
        pyautogui.press('enter')
        sleep(2)

        # salvar
        pyautogui.click(302,205, duration=2)
        sleep(20)
        pyautogui.click(982,585, duration=2)
        sleep(2)
        
        lista_empresas_cadastradas.append(dic['empresa'])

    
    #     pyautogui.click(509,659, duration=2)
    #     sleep(2)
    #     # FAP
    #     pyautogui.doubleClick(506,377, duration=2)
    #     sleep(2)
    #     # VALOR FAP*
    #     pyautogui.typewrite(fap)
    #     sleep(2)
    #     pyautogui.press('enter')
    #     sleep(2)
    #     # SALVAR COMPET 01/2023
    #     pyautogui.click(263,125, duration=2)
    #     sleep(20)
    #     pyautogui.click(204,179, duration=2)
    #     sleep(2)
    pyautogui.alert('Finalizado com sucesso')



if __name__ == '__main__':
    lista_contas =  carregar_planilha()
    print(lista_contas)
    fazer_lctos_unico_fap(lista_contas)
    