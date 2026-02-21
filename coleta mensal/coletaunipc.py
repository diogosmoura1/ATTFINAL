import pyautogui as pd
import time

pd.PAUSE = 0.30
pd.FAILSAFE = True

#67%
#saindo do  vs code e indo para o excel
pd.moveTo(x=1159, y=1058, duration=0.15)
pd.click(x=1159, y=1058)

for i in range(1):

    for n in range(109):

        #o número do contrato já vai estar selecionado
        pd.hotkey("ctrl", "c")
        pd.press('down')

        #indo para o 360 após 
        pd.moveTo(x=1202, y=1057, duration=0.15)
        pd.click(x=1202, y=1057)


        time.sleep(0.3)
        #colando o número do contrato
        #pesquisando o numero do contrato
        pd.moveTo(x=1459, y=279, duration=0.5)
        pd.click(x=1459, y=279)
        pd.doubleClick(x=1459, y=279)
        pd.press('backspace')
        pd.hotkey("ctrl", "v")

        #entrando o contrato
        time.sleep(2.5)
        pd.moveTo(x=90, y=416)
        pd.click(x=90, y=416)

        #dentro da janela selevionando
        time.sleep(1)
        pd.moveTo(x=49, y=258, duration=0.5)
        pd.dragTo(x=1102, y=744, duration=1.0)
        pd.hotkey("ctrl", "c")

        #voltando para pág
        pd.moveTo(x=73, y=219, duration=0.5)
        pd.click(x=73, y=219)

        #saindo do  360 e indo para o excel
        pd.moveTo(x=1159, y=1058, duration=0.15)
        pd.click(x=1159, y=1058)

        #trocando de planilha
        pd.moveTo(x=254, y=992, duration=0.15)
        pd.click(x=254, y=992)

        #colando as informações na planilha
        pd.hotkey("ctrl", "v")
        pd.press('right')

        #trocando de planilha
        pd.moveTo(x=163, y=991, duration=0.15)
        pd.click(x=163, y=991)