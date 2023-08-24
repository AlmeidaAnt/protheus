import pyautogui
import time
import openpyxl


# Atalho Protheus
pyautogui.press('win')
time.sleep(5)
pyautogui.typewrite('Smartclient - atalh')
time.sleep(5)
pyautogui.press('enter')
time.sleep(7)

#Logar no sistema
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(5)

pyautogui.typewrite('antonio.vinicius')
pyautogui.press('tab')
pyautogui.typewrite('Endicon07')
time.sleep(5)

#Confirmar
pyautogui.click(-878, y=397)
time.sleep(5)
pyautogui.click(x=-865, y=435)
time.sleep(10)

#Favoritos
pyautogui.click(x=-1867, y=87)
time.sleep(5)

#mov.carros
pyautogui.click(x=-1840, y=127)
time.sleep(10)

#confirmar Ambiente
pyautogui.click(x=-857, y=388)
time.sleep(15)

#fechar Alerta
pyautogui.click(x=-180, y=-132)
time.sleep(15)

#botão incluir
pyautogui.click(x=-1853, y=-158)                                            

time.sleep(7)

planilha = openpyxl.load_workbook('C:\\temp\\importação_protheus.xlsx')
aba = planilha.active


for coluna in aba.iter_rows(min_row=2, values_only=True):
    placa = coluna[0]
    centro_custo = coluna[1]

    # incluir dados
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.typewrite(placa)
    pyautogui.press('tab')
    pyautogui.typewrite(centro_custo)
    pyautogui.click(x=-68, y=-189)
    time.sleep(5)
    pyautogui.click(x=-759, y=305)
    time.sleep(2)
    pyautogui.click(x=-1853, y=-158)
    print(placa, 'movimentada para: ', centro_custo, '.')

    time.sleep(1)

