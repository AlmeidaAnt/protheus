import pyautogui
import time
import openpyxl


# Atalho Protheus
pyautogui.press('win')
time.sleep(5)
pyautogui.typewrite('Smartclient - atalh')
time.sleep(5)
pyautogui.press('enter')
time.sleep(10)

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

#Trocar Ambiente
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.typewrite('95')

#Entrar
pyautogui.click(x=-865, y=435)
time.sleep(10)

#Favoritos
pyautogui.click(x=-1867, y=87)
time.sleep(5)

#Multa
pyautogui.click(x=-1873, y=130)
time.sleep(5)

#confirmar Ambiente
pyautogui.click(x=-857, y=388)
time.sleep(15)

#fechar Alerta
pyautogui.click(x=-180, y=-132)
time.sleep(15)

#Botão Incluir
pyautogui.click(x=-1879, y=-157)
time.sleep(5)

#confirmar
pyautogui.click(x=-911, y=274)
time.sleep(5)

planilha = openpyxl.load_workbook('C:\\temp\\Importa_Multa.xlsx')
aba = planilha.active


for coluna in aba.iter_rows(min_row=2, values_only=True):
    codigo = coluna[0]
    data = coluna[1]
    hora = coluna[2]
    AI = coluna[3]
    codinf = coluna[4]
    local = coluna[5]
    UF = coluna[6]
    orgao = coluna[7]
    placa = coluna[8]
    desconto = coluna[9]
    data_emissao = coluna[10]
    cc = coluna[11]
    item_conta = coluna[12]

#cod. Multa
    pyautogui.typewrite(codigo)

#Data Infração
    pyautogui.typewrite(data)

#Hora
    pyautogui.typewrite(hora)
    pyautogui.press('tab')

#AI
    pyautogui.typewrite(AI)
    pyautogui.press('tab')

#cod. Infração
    pyautogui.typewrite(codinf)
    pyautogui.press('tab')

#Local
    pyautogui.typewrite(local)
    pyautogui.press('tab')
    pyautogui.press('tab')
#UF
    pyautogui.typewrite(UF)

#Orgão
    pyautogui.typewrite(orgao) #detran PA

#Placa
    pyautogui.typewrite(placa)
    pyautogui.press('tab')
    time.sleep(10)

#confirmar
    pyautogui.click(x=-754, y=303)
    time.sleep(30)

#Aba Pagamento
    pyautogui.click(x=-1451, y=-141)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

#Desconto
    pyautogui.typewrite(desconto)
    pyautogui.press('tab')

#Data Emissão
    pyautogui.typewrite(data_emissao)
    pyautogui.press('tab')

#Centro de Custo
    pyautogui.typewrite(cc)

#Item conta
    pyautogui.typewrite(item_conta)

    time.sleep(10)

#Salvar e Continuar
    pyautogui.click(x=-204, y=-187)
    time.sleep(5)

#confirmar
    pyautogui.click(x=-736, y=402)
    time.sleep(5)

#ok
    pyautogui.click(x=-754, y=314)
    time.sleep(10)

#cadastro de multa
    pyautogui.click(x=-1723, y=-142)
    time.sleep(10)

#codigo multa
    pyautogui.click(x=-1878, y=-94)
    time.sleep(10)

    print(placa, ' - ' ,AI," Cadastrado")