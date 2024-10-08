import pyautogui
import time
import openpyxl
import pyperclip

# ABRIR A PLANILHA
workbook = openpyxl.load_workbook('auto_glpi01.xlsx')
sheet_equipamentos = workbook['equip']

# COPIAR AS INFORMAÇÕES DOS CAMPOS
for linha in sheet_equipamentos.iter_rows(min_row=305):
    # PATRIMONIO
    patrimonio = linha[0].value
    pyperclip.copy(patrimonio)
    pyautogui.click(3750,418, duration=0.4)
    pyautogui.write('NOTE-')
    pyautogui.hotkey('ctrl', 'v')

    # LOCALIZAÇÃO
    localizacao = linha[13].value
    pyperclip.copy(localizacao)
    pyautogui.click(3022, 361, duration=0.3)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('enter')

    # NOME DA PESSOA
    nome_pessoa = linha[10].value
    pyperclip.copy(nome_pessoa)
    pyautogui.click(3112, 689, duration=0.3)
    pyautogui.hotkey('ctrl', 'v')

    # FABRICANTE
    pyautogui.click(3463, 465, duration=0.3)
    pyautogui.PAUSE = 1.0
    pyautogui.write('DELL')
    pyautogui.hotkey('enter')

    # MODELO
    modelo = linha[3].value
    pyperclip.copy(modelo)
    pyautogui.click(3487, 546, duration=0.3)
    pyautogui.hotkey('ctrl', 'v',)
    pyautogui.hotkey('enter')

    # Nº DE SERIE
    n_serie = linha[4].value
    pyperclip.copy(n_serie)
    pyautogui.click(3451, 616, duration=0.3)
    pyautogui.hotkey('ctrl', 'v')

    # GRUPO ENCARREGADO
    pyautogui.click(3075, 542, duration=0.3)
    pyautogui.click(3070, 718, duration=0.3)

    # STATUS
    status = linha[21].value
    pyperclip.copy(status)
    pyautogui.click(3316, 223, duration=0.3)
    pyautogui.hotkey('pgdn')
    pyautogui.click(3085, 856, duration=0.3)
    pyautogui.hotkey('ctrl', 'v')

    # ADICIONAR NOTEBOOK E IR PARA O PROXIMO
    pyautogui.click(3701, 925, duration=0.3)

    # COMPONENTE -
    pyautogui.click(3073, 368, duration=0.3)
    pyautogui.click(3325, 276, duration=0.3)
    pyautogui.click(3363, 507, duration=0.3)

    # Disco Rígido
    pyautogui.click(3349, 321, duration=0.3)
    ssd = linha[7].value
    if ssd == '128GB':
        pyautogui.click(3327, 436, duration=0.3)
    elif ssd == '256GB':
        pyautogui.click(3383, 468, duration=0.3)
    elif ssd == '512GB':
        pyautogui.click(3416, 497, duration=0.3)
    else:
        pyautogui.click(3337, 417, duration=0.3)

    pyautogui.click(3547, 397, duration=0.3)
    pyautogui.hotkey('1')
    pyautogui.hotkey('enter')
    pyautogui.click(3714, 334, duration=0.3)

    # MEMÓRIA
    pyautogui.click(3315, 274, duration=0.3)
    pyautogui.write('memoria')
    pyautogui.hotkey('enter')
    pyautogui.click(3376, 311, duration=0.3)
    memoria = linha[8].value
    if memoria == '8GB':
        pyautogui.click(3342, 502, duration=0.3)
    elif memoria == '12GB':
        pyautogui.click(3335, 404, duration=0.3)
    elif memoria == '16GB':
        pyautogui.click(3335, 438, duration=0.3)
    elif memoria == '32GB':
        pyautogui.click(3337, 476, duration=0.3)

    pyautogui.click(3539, 402, duration=0.3)
    pyautogui.hotkey('1')
    pyautogui.hotkey('enter')
    pyautogui.click(3714, 334, duration=0.3)

    # PROCESSADOR
    pyautogui.click(3327, 278, duration=0.3)
    pyautogui.write('processador')
    pyautogui.hotkey('enter')
    pyautogui.click(3353, 313, duration=0.3)
    processador = linha[5].value
    pyautogui.PAUSE = 0.5
    if processador == 'i5 - 7th Gen':
        pyautogui.write('i5 - 7th Gen')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'Core i5-1135G7':
        pyautogui.write('Core i5-1135G7')
        pyautogui.PAUSE = 0.5
    elif processador == 'Core i5-1165G7':
        pyautogui.write('Core i5-1165G7')
        pyautogui.hotkey('enter')
    elif processador == 'Core i5-8250U':
        pyautogui.write('Core i5-8250U')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'Core i7-1165G7':
        pyautogui.write('Core i7-1165G7')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'Core i7-1355U':
        pyautogui.write('Core i7-1355U')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'Core i7-7500U':
        pyautogui.write('Core i7-7500U')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'Core i7-8550U':
        pyautogui.write('Core i7-8550U')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'i5 - 7th Gen':
        pyautogui.write('i5 - 7th Gen')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'i5-1135G7 - 4 core':
        pyautogui.write('i5-1135G7 - 4 core')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'i7-10510U- 4 core':
        pyautogui.write('i7-10510U- 4 core')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    elif processador == 'i7-1165G7 - 4 core':
        pyautogui.write('i7-1165G7 - 4 core')
        pyautogui.PAUSE = 0.5
        pyautogui.hotkey('enter')
    pyautogui.click(3537, 397, duration=0.5)
    pyautogui.hotkey('1')
    pyautogui.hotkey('enter')
    pyautogui.click(3713, 341, duration=0.5)
    pyautogui.click(3318, 110, duration=0.5)

    # pyautogui.click(3322,117,duration=1.5)
    # MANEIRA 1 DE PREENCHER STATUS

    # pyautogui.click(3419,294,duration=1)
    # status = linha[21].value
    # if status == 'Baixado':
    #    pyautogui.click(3482,394,duration=1)
    # elif status == 'Defeito':
    #    pyautogui.click(3398,431,duration=1)
    # elif status == 'Disponível':
    #    pyautogui.click(3403,450,duration=1)
    # elif status == 'Doar':
    #    pyautogui.click(3403,485,duration=1)
    # elif status == 'Extraviado':
    #    pyautogui.click(3480,423,duration=1)
    # elif status == 'Não Encontrado':
    #    pyautogui.click(3480,423,duration=1)
    # elif status == 'Em uso':
    #    pyautogui.click(3480,423,duration=1)
    # else:
    # status == 'Vendido'
    # pyautogui.click(3480,423,duration=1)
    # tipo_pc = linha[0].value
