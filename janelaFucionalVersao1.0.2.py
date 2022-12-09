import PySimpleGUI as sg
import pyautogui
import time  # importar blibliotema time(tempo)
import pyperclip as pc
from datetime import datetime
import keyboard
import win32com.client as win32
layaut = [
     [sg.Text("Sistema de Consulta ao Usuario")],
     [sg.Text('Nome de usuario'),sg.InputText(key='usuariocmd'), sg.Button("Consultar")],
     [sg.Text(key='informaçao')],
     [sg.Text(key='informaçao2')],

]

janela = sg.Window('Usuarios do MPGO Vesão 1.0.1', layaut)
#anela.bind('Enter', '-ENTER-')
while True:
    evento, valores = janela.read()
    if evento == sg.WIN_CLOSED:
        break
    if evento == "Consultar":
        print(evento)
        usuario= valores['usuariocmd']
        if usuario == '':
            janela['informaçao'].update('Digite o nome de usuario para continuar!')
        else:
            pyautogui.hotkey('winleft', "r")
            pyautogui.write('cmd')  # Digite "cmd"
            pyautogui.press('enter')  # Precione "Enter
            time.sleep(0.5)
            pyautogui.write('net user {} /domain'.format(usuario))
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.write('exit')
            pyautogui.press('enter')
            # pyautogui.write('exit')
            # pyautogui.press('enter')
            # ====================Tratando as informaaçoes obtidas e Transformando em uma lista====================
            texto = pc.paste()
            dados = texto.upper()[(texto.upper().find('CONTA ATIVA')):]  # comece a lista quando achar conta ativa
            lista1 = dados.split()  # lista criada
            # ======================================Define Cores====================================================
            vermelho = "\033[1;31m"  # COR VERMELHA
            verde = "\033[0;32m"  # COR VERDE
            RESET = "\033[0;0m"  # COR NORMAL
            # -------------------------------------------------------------------------------------------------------------
            # =======================================Status da Conta=======================================================
            texto2 = pc.paste()
            dados2 = texto2.upper()[(texto2.upper().find('PARA O DOMÍNIO')):]
            lista2 = dados2.split()
            lista01 = dados.split()
            print(lista2[2])
            if lista2[4] == 'NÃO':
                lista1 = lista2
            if lista1[2] == 'SIM':
                janela['informaçao'].update(' Usuário está ATIVO no sistema')
                print(' Usuário está {}ATIVO{} no sistema'.format(verde, RESET))
            elif lista1[2] == 'NÂO':
                janela['informaçao'].update('Usuário Não está ativo no sistema')
                print('Usuário Não está ativo no sistema')
            elif lista1[2] == 'BLOQUEADO':
                janela['informaçao'].update('Usuário está BLOQUEADO no sistema')
                print('Usuário está {}BLOQUEADO{} no sistema'.format(vermelho, RESET))
            elif lista2[2] != 'NÃO' or 'SIM' or 'BLOQUEADO':
                janela['informaçao'].update('Erro! Usuario não localizado, tente novamente')
                print('{}Erro! Usuario não localizado, tente novamente{}'.format(vermelho, RESET))
        datahojehorahoje = str(datetime.now())
        anohoje = int(datahojehorahoje[0:4])
        meshoje = int(datahojehorahoje[5:7])
        diahoje = int(datahojehorahoje[8:10])
        horahoje = int(datahojehorahoje[11:13])
        minutohoje = int(datahojehorahoje[14:16])
        segundoshoje = int(datahojehorahoje[17:19])
        # =======================================================================================
        if 6 == (dados.find('ATIVA')):
            print('Deu certo garoto')
            datasenha = dados.split()
            diasenha = int(datasenha[11][0:2])
            messenha = int(datasenha[11][3:5])
            anosenha = int(datasenha[11][6:])
            horasenha = int(datasenha[12][0:2])
            minutosenha = int(datasenha[12][3:5])
            segundosenha = int(datasenha[12][6:])
            ano = anohoje - anosenha
            if ano <= -1:
                ano = ano * -1
            ######################
            mes = messenha - meshoje
            if mes <= -1:
                mes = mes * -1
            ##################
            dia = diasenha - diahoje
            if dia <= -1:
                dia = dia * -1
            #####################
            hora = horasenha - horahoje
            if hora <= -1:
                hora = hora * -1
            ##########################
            minuto = minutosenha - minutohoje
            if minuto <= -1:
                minuto = minuto * -1
            #########################
            segundo = segundosenha - segundoshoje
            if segundo <= -1:
                segundo = segundo * -1
            #####################################
            if ano == 0:
                ano = str('')
            elif ano >= 2:
                ano = str(ano) + ' Anos,'
            elif ano == 1:
                ano = str(ano) + ' Ano,'
            #############################
            if mes == 0:
                mes = str('')
            elif mes == 1:
                mes = str(mes) + ' Mês,'
            elif mes >= 2:
                mes = str(mes) + ' Meses,'
            ##############################
            if dia == 0:
                dia = str('')
            elif dia == 1:
                dia = str(dia) + ' Dia,'
            elif dia >= 2:
                dia = str(dia) + ' Dias,'
            #############################
            if hora == 0:
                hora = str('')
            elif hora == 1:
                hora = str(hora) + ' Hora,'
            elif hora >= 2:
                hora = str(hora) + ' Horas,'
            ########################
            if minuto == 0:
                minuto = str('')
            elif minuto == 1:
                minuto = str(minuto) + ' Minuto,'
            elif minuto >= 2:
                minuto = str(minuto) + ' Minutos,'
            ######################################
            if segundo == 0:
                segundo = str('')
            elif segundo == 1:
                segundo = ' e ' + str(segundo) + ' Segundo'
            elif segundo >= 2:
                segundo = ' e ' + str(segundo) + ' Segundos'

            frase = ('Ultima definição de senha foi a {}{}{}{}{}{}'.format(ano, mes, dia, hora, minuto, segundo))
            print(frase)
            janela['informaçao2'].update(frase)
        else:
            frase2 = ('Sem dados adicionais')
            janela['informaçao2'].update(frase2)
            print(frase2)


janela.close()