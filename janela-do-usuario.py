import PySimpleGUI as sg
import pyautogui
import time  # importar blibliotema time(tempo)
import pyperclip as pc
import keyboard
import win32com.client as win32
layaut = [
     [sg.Text("Sistema de Consulta ao Usuario")],
     [sg.Text('Nome de usuario'),sg.InputText(key='usuariocmd'), sg.Button("Consultar")],
     [sg.Input(input(keyboard.is_pressed('enter')))],
     [sg.Text(key='informaçao')],
]

janela = sg.Window('Usuarios do MPGO Vesão 1.0.1', layaut)
while True:
    evento, valores = janela.read()
    if evento == sg.WIN_CLOSED:
        break
    if evento == "Consultar" :
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


janela.close()