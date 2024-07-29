'''

Ler dados da planilha
Inserir cada célula de cada linha em um campo do sistema



************* COMO SABER O "X" E "Y" DA TELA *************
1-Abre o terminal(cmd)
2-Use o comando "pip install mouseinfo" para instalar
2,5- Se não for usa "pip3 ins..." Dependendo do python que usa
3-Usa o comando "python"
3,5- Se não for usa "python3"
4-Abrindo o python usa "from mouseinfo import mouseInfo"
5-Depois roda ele, com "mouseInfo()" 
6-Vai abrir uma tela que vai mostrar o X e Y de onde o cursor do mouse está
7-Desativa o "3 Sec. Buttom Delay"
8-Para salvar o local do cursor usa F6

'''
#Biblioteca para abrir o excel
import openpyxl

#Biblioteca para inserir coisa nova
import pyautogui

#Abrindo a planilha
workspace = openpyxl.load_workbook('vendas_de_produtos.xlsx')

#Selecionando a página da planilha
vendasSheet = workspace['vendas']

#Local onde o bot vai colocar os dados da planilha num local que vai enviar
for linha in vendasSheet.iter_rows(min_row=2, max_row=2):
    #Nome
    #Posição da tela:X   Y ; Tempo até lá
    pyautogui.click(1808,452, duration=1.5)
    pyautogui.write(linha[0].value)
    #Produto
    pyautogui.click(1815,476,duration=1.5)
    pyautogui.write(linha[1].value)
    #Quantidade
    pyautogui.click(1813,497, duration=1.5)
    pyautogui.write(str(linha[2].value))
    #Categoria
    pyautogui.click(1883,532)
    pyautogui.write(linha[3].value)
    #Enviar
    pyautogui.click(1752, 549, duration=1.5)
    #Concluir
    pyautogui.click(1256,581,duration=1.5)
    '''
    Só pra mostrar o que tem
   print(linha[0].value)
   print(linha[1].value)
   print(linha[2].value)
   print(linha[3].value)
   '''
