import tkinter as tk
from charset_normalizer.cli import query_yes_no
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.ui import Select
from tkinter import messagebox
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


class MyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Baixar balancetes da contabilidade")
        self.root.geometry("520x220")

        self.label = tk.Label(root, text="Bem-vindo ao AtenaCommander!", font=("Arial", 12))
        self.label.grid(row=0, column=1, columnspan=4, sticky="nsew", padx=20, pady=10)

        self.email_label = tk.Label(root, text="", font=("Arial", 12))

        self.competencia_inicial_entry = tk.Label(root, text="Início da Competência:", font=("Arial", 12))
        self.competencia_inicial_entry.grid(row=4, column=1, sticky="w", padx=20, pady=10)
        self.competencia_inicial_entry = tk.Entry(root, font=("Arial", 12), width=30)
        self.competencia_inicial_entry.grid(row=4, column=2, sticky="ew", padx=20, pady=10)

        self.competencia_final_entry = tk.Label(root, text="Final da Competência:", font=("Arial", 12))
        self.competencia_final_entry.grid(row=5, column=1, sticky="w", padx=20, pady=10)
        self.competencia_final_entry = tk.Entry(root, font=("Arial", 12), width=30)
        self.competencia_final_entry.grid(row=5, column=2, sticky="ew", padx=20, pady=10)

        # Adicionando o texto "Formato ex: MMYYYY" abaixo de "Data de competência"
        self.date_format_label = tk.Label(root, text="Formato ex: MMYYYY", font=("Arial", 10))
        self.date_format_label.grid(row=6, column=2, sticky="w", padx=20, pady=5)

        self.download_button = tk.Button(root, text="Baixar Balancetes", command=self.executarDownloadDosArquivos, font=("Arial", 12))
        self.download_button.grid(row=7, column=1, columnspan=2, sticky="nsew", padx=20, pady=10)

    def get_date_inicial(self):
        return self.competencia_inicial_entry.get()

    def get_date_final(self):
        return self.competencia_final_entry.get()

    def executarDownloadDosArquivos(self):
        date_start = self.get_date_inicial()
        date_final = self.get_date_final()
        executar_download(date_start, date_final)

def inicializar_driver():
    driver = webdriver.Chrome()
    driver.get("https://financeiro.postalis.org.br/ControleAcesso/login/Login.aspx")

    driver.find_element(By.ID, "MainContent_lgnLogin_UserName").send_keys("emailAqui")
    time.sleep(1)
    driver.find_element(By.ID, "MainContent_lgnLogin_Password").send_keys("senhaAqui")
    time.sleep(1)
    driver.find_element(By.ID, "MainContent_lgnLogin_LoginButton").click()
    time.sleep(2)

    return driver

def executar_download(date_start, date_final):
    driver = inicializar_driver()
    try:
        wait = WebDriverWait(driver, 10)

        driver.find_element(By.ID, "MainContent_imbContabilidade").click()
        time.sleep(1)
        driver.find_element(By.ID, "MainContent_imbContabilidade").click()
        time.sleep(1)
        driver.find_element(By.ID, "MenuContent_tvwMenut96").click()
        time.sleep(1)
        driver.find_element(By.ID, "MenuContent_tvwMenut97").click()
        time.sleep(1)
        driver.find_element(By.ID, "MenuContent_tvwMenut99").click()
        time.sleep(1)
        driver.find_element(By.ID, "MenuContent_tvwMenut100").click()
        time.sleep(3)

        # Seleção de datas
        data_comp1 = driver.find_element(By.ID, "MainContent_MainContent_perPeriodo_dbxInicio")
        data_comp1.click()
        for i in range(6):
            data_comp1.send_keys(Keys.BACKSPACE)
        data_comp1.send_keys(date_start)

        data_comp2 = driver.find_element(By.ID, "MainContent_MainContent_perPeriodo_dbxFim")
        data_comp2.click()
        for i in range(6):
            data_comp2.send_keys(Keys.BACKSPACE)
        data_comp2.send_keys(date_final)

        driver.find_element(By.ID, "MainContent_MainContent_fpuBalancete_lnkSelecionado").click()

        # Obter todas as guias abertas
        abas = driver.window_handles
        # Trocar para a nova guia (geralmente a última aberta)
        driver.switch_to.window(abas[-1])
        # Agora você pode interagir com a nova guia
        print(driver.title)  # Exemplo: imprimir o título da nova guia
        # Se precisar voltar para a guia original
        driver.switch_to.window(abas[0])

        time.sleep(1)

        seleciona_BD = driver.find_element(By.XPATH, "/html/body/form/div[3]/div[3]/div[2]/div[1]/div[5]/div[11]/div[1]/div[1]/table/tbody/tr[2]/td[1]/select/option[1]")
        seleciona_PP = driver.find_element(By.XPATH, "/html/body/form/div[3]/div[3]/div[2]/div[1]/div[5]/div[11]/div[1]/div[1]/table/tbody/tr[2]/td[1]/select/option[3]")
        seleciona_PGA = driver.find_element(By.XPATH, "/html/body/form/div[3]/div[3]/div[2]/div[1]/div[5]/div[11]/div[1]/div[1]/table/tbody/tr[2]/td[1]/select/option[5]")
        seleciona_BPAuxiliar = driver.find_element(By.XPATH, "/html/body/form/div[3]/div[3]/div[2]/div[1]/div[5]/div[11]/div[1]/div[1]/table/tbody/tr[2]/td[1]/select/option[4]")
        time.sleep(1)

        actions = ActionChains(driver)
        actions.key_down(Keys.CONTROL).click(seleciona_BD).click(seleciona_PP).click(seleciona_PGA).click(seleciona_BPAuxiliar).key_up(Keys.CONTROL).perform()
        time.sleep(1)

        driver.find_element(By.ID, "MainContent_MainContent_fpuBalancete_lbtList_btnAdd").click()
        time.sleep(1)

        driver.find_element(By.ID, "MainContent_MainContent_fpuBalancete_btnCancelar").click()
        time.sleep(1)

        driver.switch_to.window(abas[0])
        time.sleep(1)

        driver.find_element(By.ID, "MainContent_MainContent_cblOpcoes_12").click()
        time.sleep(2)

        driver.find_element(By.ID, "MainContent_MainContent_btnExportar").click()
        time.sleep(1)

        time.sleep(6)

    finally:
        driver.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = MyApp(root)
    root.mainloop()