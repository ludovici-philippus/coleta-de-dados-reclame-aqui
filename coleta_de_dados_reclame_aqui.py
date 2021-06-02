from tkinter import *
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import xlsxwriter

janela = Tk()
nome_empresa = StringVar()
#try:
#    open(file="reclame-aqui-empresas.xlsx", mode='r')
#    workbook = xlsxwriter.Workbook("reclame-aqui-empresas.xlsx")
#except:
paragrafo = []
span = []
dados = ["Empresa"]
iteracao = -1

def enviar():
    global iteracao
    iteracao += 1
    dados = ["Empresa"] # Reseta os dados a cada chamada.
    driver = webdriver.Firefox()
    driver.get(f"https://www.reclameaqui.com.br/empresa/{nome_empresa.get().lower()}/")
    wait = WebDriverWait(driver, 20)
    try:
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/section/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/span[1]')))
        elem = driver.find_element_by_xpath('/html/body/div[1]/section/div[2]/div/div/div[2]/div[1]/h1')
        dados.append(elem.text) # Nome
        elem = driver.find_element_by_xpath('/html/body/div[1]/section/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/span[1]')
        dados.append("Descrição")
        dados.append(elem.text) # Descrição
        elem = driver.find_element_by_xpath('/html/body/div[1]/section/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/span[2]/b')
        dados.append("Avaliação")
        dados.append(elem.text) # Avaliação/Nota
        for p in driver.find_elements_by_xpath('/html/body/div[1]/section/section[1]/div[1]/div[1]/div[1]/div[2]/div[1]/p'):
            paragrafo.append(p.text)
        for s in driver.find_elements_by_xpath('/html/body/div[1]/section/section[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div/span'):
            span.append(s.text)
        dados.append([paragrafo[:], span[:]])
        driver.close()
        print(dados)
        criar_planilha_com_dados(dados)
    except:
        driver.close()

def criar_planilha_com_dados(lista):
    workbook = xlsxwriter.Workbook("reclame-aqui-empresas.xlsx")
    worksheet = workbook.add_worksheet()
    planilha = (
        [lista[0], lista[1]],
        [lista[2], lista[3]],
        [lista[4], lista[5]],
        [lista[6][0][0], lista[6][1][0]],
        [lista[6][0][1], lista[6][1][1]],
        [lista[6][0][2], lista[6][1][2]],
        [lista[6][0][3], lista[6][1][3]]
    )
    print(planilha)
    row = (0 + iteracao) * 8
    column = 0

    for item, value in planilha:
        print(f"Item: {item} | Value: {value}")
        worksheet.write(row, column, item)
        worksheet.write(row, column + 1, value)
        row += 1
    workbook.close()
    print("FIM")

lb_empresa = Label(janela, text="Digite o nome da empresa: ")
lb_empresa.place(x=0, y=0)

e_empresa = Entry(janela, textvariable=nome_empresa)
e_empresa.place(x=150, y=0)

bt_enviar = Button(janela, width=30, text="Enviar!", command=enviar)
bt_enviar.place(x=40, y=60)


janela.geometry("300x100")
janela.resizable(False, False)
janela.mainloop()