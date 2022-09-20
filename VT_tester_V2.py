from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import re, datetime, time, warnings, pyautogui
from tkinter import *
from tkinter import filedialog, messagebox

warnings.simplefilter(action='ignore', category=DeprecationWarning)

# atualizar web driver https://chromedriver.chromium.org/downloads #
root = Tk()
root.wm_withdraw()
janela = Tk()
janela.wm_withdraw()

def consultar_viagem():
    root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Selecione a base",filetypes = (("Excel","*.xlsx"),("all files","*.*")))
    df = pd.read_excel(root.filename)
    inicio = time.perf_counter()
    ti = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    messagebox.showinfo(title='Aviso',message=f'\nIniciado em: {ti}\n',parent=janela)
    pd.options.mode.chained_assignment = None

    # Link do site
    link = 'https://www.google.com/maps/dir///@-12.9766386,-38.4743728,15z/data=!4m2!4m1!3e3'
    # Abrir o Brave
    brave_path = "C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe"
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.binary_MespJc = brave_path
    driver = webdriver.Chrome(chrome_options=options, executable_path = r'D:\00 - Scripts\Python\Scripts\Teste VT\chromedriver.exe')
    driver.get(link)
    
    numerador = 0
    
    while len(driver.find_elements_by_class_name('tactile-searchbox-input')) == 0:
        time.sleep(0.5)
    # Buscar informações de rotas de ida
    for i,matricula in enumerate(df['MATRÍCULA']):
        
        driver.find_element_by_xpath('//*[@id="sb_ifc50"]/input').send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="sb_ifc51"]/input').send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        time.sleep(0.5)
        
        # Preencher endereço de origem
        driver.find_element_by_xpath('//*[@id="sb_ifc50"]/input').send_keys(df['Origem'][i])
        time.sleep(0.5)
        pyautogui.press('down')
        time.sleep(0.5)
        pyautogui.press('enter')
        # Preencher endereço de destino
        driver.find_element_by_xpath('//*[@id="sb_ifc51"]/input').send_keys(df['Destino'][i])
        time.sleep(0.5)
        pyautogui.press('down')
        time.sleep(0.5)
        pyautogui.press('enter')
        
        # Selecionar a Partida
        driver.find_element_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/span/div/div/div/div[2]').click()
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id=":1"]/div').click()
        time.sleep(0.5)
        
        # Preencher o horário de consulta
        while len(driver.find_elements_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/span[1]/input')) == 0:
            time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/span[1]/input').send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/span[1]/input').send_keys(df['HORARIO'][i][:6])
        pyautogui.press('enter')
        time.sleep(1)
        
        # Verificar se houve erro na busca de rotas
        if len(driver.find_elements_by_xpath('//*[contains(text(),"Não foi possível calcular as rotas de transporte público de ")]')) > 0:
            print(driver.find_elements_by_xpath('//*[contains(text(),"Não foi possível calcular as rotas de transporte público de ")]'))
            driver.refresh()
            df['Número de viagens Ida'].loc[i] = 'Validar'
        else:
            while len(driver.find_elements_by_class_name('MespJc')) == 0:
                time.sleep(0.5)
            # Pegar dados do transporte
            rotas = driver.find_elements_by_class_name('MespJc')
            linhas = []
            onibus = []
            buzu = []
            teste = []
            
            # Verificar se foram identificadas rotas
            if len(rotas) == 0:
                # Caso não ache rotas, preencha com validar
                df['Número de viagens Ida'].loc[i] = 'Validar'
            # Caso ache rotas, salve e selecione a com menor quantidade de linhas
            else:
                for rota in rotas:
                    rota_v2 = rota.text.split('\n')
                    try:
                        numero = len(re.findall('[0-9]+[A-Z]? ', rota_v2[2], re.IGNORECASE)) + 1
                    except:
                        numero = 1
                    linhas.append(numero)
                    teste.append(rota_v2)
                    for a, testes in enumerate(teste):
                        onibus.append(teste[a][2].count(' '))
                        buzu.append(teste[a][2])
                        indice_menor = onibus.index(min(onibus))
                df['Rota Ida'].loc[i] = buzu[indice_menor]
                # Colocar no dataframe a menor quantidade de linhas
                df['Número de viagens Ida'].loc[i] = min(linhas)
        
############################################################################################################################################################################################

############################################################################################################################################################################################

        driver.find_element_by_xpath('//*[@id="sb_ifc50"]/input').send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="sb_ifc51"]/input').send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        time.sleep(0.5)
        
        # Preencher endereço de origem
        driver.find_element_by_xpath('//*[@id="sb_ifc50"]/input').send_keys(df['Destino'][i])
        time.sleep(0.5)
        pyautogui.press('down')
        time.sleep(0.5)
        pyautogui.press('enter')
        # Preencher endereço de destino
        driver.find_element_by_xpath('//*[@id="sb_ifc51"]/input').send_keys(df['Origem'][i])
        time.sleep(0.5)
        pyautogui.press('down')
        time.sleep(0.5)
        pyautogui.press('enter')
        
        # Selecionar a Partida
        driver.find_element_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/span/div/div/div/div[2]').click()
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id=":1"]/div').click()
        time.sleep(0.5)
        
        # Preencher o horário de consulta
        while len(driver.find_elements_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/span[1]/input')) == 0:
            time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/span[1]/input').send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        time.sleep(0.5)
        driver.find_element_by_xpath('//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/span[1]/input').send_keys(df['HORARIO'][i][:6])
        pyautogui.press('enter')
        time.sleep(1)
        
        # Verificar se houve erro na busca de rotas
        if len(driver.find_elements_by_xpath('//*[contains(text(),"Não foi possível calcular as rotas de transporte público de ")]')) > 0:            
            driver.refresh()
            df['Número de viagens Volta'].loc[i] = 'Validar'
        else:
            while len(driver.find_elements_by_class_name('MespJc')) == 0:
                time.sleep(0.5)
            rotas = driver.find_elements_by_class_name('MespJc')
            linhas = []
            onibus = []
            buzu = []
            teste = []
            
            # Verificar se foram identificadas rotas
            if len(rotas) == 0:
                # Caso não ache rotas, preencha com validar
                df['Número de viagens Volta'].loc[i] = 'Validar'
            # Caso ache rotas, salve e selecione a com menor quantidade de linhas
            else:
                for rota in rotas:
                    rota_v2 = rota.text.split('\n')
                    try:
                        numero = len(re.findall('[0-9]+[A-Z]? ', rota_v2[2], re.IGNORECASE)) + 1
                    except:
                        numero = 1
                    linhas.append(numero)
                    teste.append(rota_v2)
                    for a, testes in enumerate(teste):
                        onibus.append(teste[a][2].count(' '))
                        buzu.append(teste[a][2])
                        indice_menor = onibus.index(min(onibus))
                df['Rota Volta'].loc[i] = buzu[indice_menor]
                # Colocar no dataframe a menor quantidade de linhas
                df['Número de viagens Volta'].loc[i] = min(linhas)
        df.to_excel(r'D:\00 - Scripts\Python\Scripts\Teste VT\Arquivos\Consulta_VT.xlsx', index=False)
        print(df['Número de viagens Ida'][i], ' - ', df['Rota Ida'][i])
        print(df['Número de viagens Volta'][i],' - ', df['Rota Volta'][i])
        numerador += 1
        tamanho = len(df)
        print(f'Linhas concluídas: {numerador} de {tamanho} linhas')
    final = time.perf_counter()
    tf = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    if round(final - inicio, 1) > 60:
        total = round((final - inicio) / 60)
        if total < 2:
            messagebox.showinfo(title='Aviso',message=f'\nTempo de Execução: {total} Minuto\n',parent=janela)
        else:
            messagebox.showinfo(title='Aviso',message=f'\nTempo de Execução: {total} Minutos\n',parent=janela)
    else:
        total = round(final - inicio)
        if total >= 2:
            messagebox.showinfo(title='Aviso',message=f'\nTempo de Execução: {total} Segundos\n',parent=janela)
        else:
            messagebox.showinfo(title='Aviso',message=f'\nTempo de Execução: {total} Segundo\n',parent=janela)
consultar_viagem()