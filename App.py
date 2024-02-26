import openpyxl
import pyperclip
import pyautogui
#https://cadastro-produtos-devaprender.netlify.app/etapa2.html
controle = 1
xxx = 1
woerkbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = woerkbook["Produtos"]

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1326,353,duration=0.01)
    pyautogui.hotkey("ctrl","v",)
    print(controle)
    controle = controle + 1


    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1328,427,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1332,562,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1332,656,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1332,739,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1334,836,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    pyautogui.click(1339,892,duration=0.01)
    foi1 = input("O site de merda ja carregou? ")
    

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1330,372,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1331,460,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1330,548,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1331,640,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    tamanho = linha[10].value
    pyautogui.click(1363,716,duration=0.01)
    if "Pequeno" in tamanho:
        pyautogui.click(1378,748,duration=0.01)
    elif "Médio" in tamanho:
        pyautogui.click(1385,772,duration=0.01)
    elif "Grande" in tamanho:
        pyautogui.click(1379,792,duration=0.01)


    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1329,805,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    pyautogui.click(1372,860,duration=0.01)
    foi2 = input("O site de merda ja carregou? ")


    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1330,398,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    pais_de_origem = linha[13].value
    pyperclip.copy(pais_de_origem)
    pyautogui.click(1331,480,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1327,594,duration=0.01)
    pyautogui.hotkey("ctrl","v")
 

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1328,694,duration=0.01)
    pyautogui.hotkey("ctrl","v")


    lacalizacao_armazem = linha[16].value
    pyperclip.copy(lacalizacao_armazem)
    pyautogui.click(1325,786,duration=0.01)
    pyautogui.hotkey("ctrl","v")

    pyautogui.click(1363,846,duration=0.01)
    foi3 = input("O site de merda ja carregou? ")

    pyautogui.click(1753,167,duration=0.01)
    foi4 = input("O site de merda ja carregou? ")

    pyautogui.click(1751,172,duration=0.01)
    foi5 = input("O site de merda ja carregou? ")

    pyautogui.click(1516,629,duration=0.01)
    foi6 = input("O site de merda ja carregou? ")

    if controle == 3:
        break