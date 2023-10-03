import time
from termcolor import colored
from datetime import datetime
import gspread
from selenium.webdriver.common.by import By
from seleniumbase import BaseCase

#------Conexão com o Google Sheets
ID_SHEET = '1oaEXhvwbLGqsq7B_IuIsExPnk_23ba_CiWnzHsqJJ6M'
gc = gspread.service_account(filename='key.json')
sh = gc.open_by_key(ID_SHEET)
ws = sh.worksheet('DadosCAF')   

plan_01 = sh.get_worksheet(0)
tempo_limite = 3600 
linha = 2

coluna = plan_01.col_values(2)  # 1 representa a coluna B
# Conte o número de linhas preenchidas (não vazias)
total_linhas_preenchidas = len([valor for valor in coluna if valor.strip()])
total = total_linhas_preenchidas

#bDAta e hora atual
data_hora = datetime.now()
# Formato "YYYY-MM-DD HH:MM:SS"
data_hora = data_hora.strftime("%Y-%m-%d %H:%M:%S")
#____________________________ CONSULTA CAF__________________________________
BaseCase.main(__name__, __file__);
class CoffeeCartTest(BaseCase):     
    def test_coffee_cart(self):          
        SITE = "https://sistemas.agricultura.gov.br/caf/dados-publicos/membros-ufpa"
        self.open(SITE) 
        linha = 2
        for linha in range(linha, total + 1):
            #________VARIÁVEIS DA GOOGLE SHEETS____
            SICAD = plan_01.cell(linha,1).value
            CLIENTE = plan_01.cell(linha,2).value
            CPF = plan_01.cell(linha,3).value.replace("-","").replace(".","")
            CPF_FOMARTADO = "{}.{}.{}-{}".format(CPF[:3], CPF[3:6], CPF[6:9], CPF[9:])
            #sair assim que cliente for ''
            if not CLIENTE:
                break             
            print(" ")
            print(colored(f"ABRIU O SITE: {data_hora}",'blue'))
            # Clica no elemento usando XPath e a classe By
            self.driver.find_element(By.XPATH, '//*[@id="input_NU_CPF"]').click()
            self.driver.find_element(By.XPATH, '//*[@id="input_NU_CPF"]').send_keys(CPF_FOMARTADO)
            #time.sleep(1)
            self.driver.find_element(By.XPATH, '//*[@id="input_consultar"]').click()
            self.find_element(By.XPATH,'//*[@id="lista"]/tbody/tr/td[5]/span').click()
            print(colored("Consulta CAF",'green'))
            N_CAF = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[1]/div/table/tbody/tr[1]/td[1]/span').text
            INSCRICAO =  self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[1]/div/table/tbody/tr[2]/td[1]/span').text
            VALIDADE = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[1]/div/table/tbody/tr[2]/td[2]/span').text
            SITUACAO = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[1]/div/table/tbody/tr[1]/td[2]/span').text
            CONDICAO_POSSE = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[3]/div[2]/table/tbody/tr[2]/td[1]/span').text
            TAMANHO_IMOVEL = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[3]/div[2]/table/tbody/tr[2]/td[2]/span').text
            MUNICIPIO = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[3]/div[2]/table/tbody/tr[5]/td/span').text
            ENTIDADE = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[4]/div[2]/table/tbody/tr[1]/td[1]/span').text
            CNPJ_ENTIDADE = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[4]/div[2]/table/tbody/tr[1]/td[2]/span').text
            CADASTRADOR = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[4]/div[2]/table/tbody/tr[2]/td[1]/span').text
            VERSAO_PORTAL = self.find_element(By.XPATH,'/html/body/footer/div/h5/p').text
            NUMERO_IMOVEIS = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[3]/div[2]/table/tbody/tr[1]/td/span').text
            print("")            
            nomes = self.find_element(By.XPATH,'//*[@id="wrap"]/div[3]/div/div/div[2]/div[2]/table/tbody')
            n_linhas = len(nomes.find_elements(By.TAG_NAME,"tr"))
            print(f'Nº DE MEMBROS UFPA: {n_linhas}')
            NOME_1  = self.find_element(By.XPATH,f'//*[@id="wrap"]/div[3]/div/div/div[2]/div[2]/table/tbody/tr[1]/td[1]').text
            plan_01.update(f"D{linha}",SITUACAO)
            plan_01.update(f"E{linha}",INSCRICAO)
            plan_01.update(f"F{linha}",VALIDADE)
            plan_01.update(f"G{linha}",N_CAF)
            plan_01.update(f"H{linha}",CONDICAO_POSSE)
            plan_01.update(f"I{linha}",NUMERO_IMOVEIS)
            plan_01.update(f"J{linha}",TAMANHO_IMOVEL)
            plan_01.update(f"K{linha}",MUNICIPIO)
            plan_01.update(f"L{linha}",n_linhas)
            plan_01.update(f"M{linha}",NOME_1)
            plan_01.update(f"P{linha}",ENTIDADE)
            plan_01.update(f"Q{linha}",CNPJ_ENTIDADE)
            plan_01.update(f"R{linha}",CADASTRADOR)
            plan_01.update(f"S{linha}",VERSAO_PORTAL)
            if n_linhas > 1:
                        NOME_2  = self.find_element(By.XPATH,f'//*[@id="wrap"]/div[3]/div/div/div[2]/div[2]/table/tbody/tr[2]/td[1]').text
                        plan_01.update(f"N{linha}",NOME_2)    
                        if n_linhas > 2:
                            linha2 = 3
                            for linha2 in range(3,n_linhas + 1):
                                    linha2 = 3 if linha2 >3 else linha2
                                    NOMES_UFPA_PLAN = plan_01.cell(linha,15).value                                                                           
                                    NOMES_UFPA  = self.find_element(By.XPATH,f'//*[@id="wrap"]/div[3]/div/div/div[2]/div[2]/table/tbody/tr[{linha2}]/td[1]').text           
                                    NOMES_UFPA = f'{NOMES_UFPA}:{linha2}ºMembro|'
                                    if NOMES_UFPA_PLAN is not None:        
                                        NOMES_UFPA = NOMES_UFPA_PLAN + NOMES_UFPA
                                    plan_01.update(f"O{linha}",NOMES_UFPA)

            plan_01.update(f"T{linha}",data_hora)
            print(colored(f'CONSULTA: {CLIENTE} SITUACAO:{SITUACAO} CONSULTA N°: {linha -1} FALTAM: {total - linha -1}','green'))
            self.go_back()

print(f"Processo Concluído!!! {total-1} consultas realizadas.")
                                    