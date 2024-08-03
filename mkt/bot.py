"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install -e .`
- Use the same interpreter as the one used to install the bot (`pip install -e .`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""

from botcity.web import WebBot, Browser, By
import pandas as pd
from dotenv import load_dotenv
import os #provides ways to access the Operating System and allows us to read the environment variables
import openpyxl
from datetime import datetime
import time
# Import smtplib for the actual sending function
import smtplib
# Import the email modules we'll need
import email.message
from email.message import EmailMessage

# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *

class Bot(WebBot):
    def action(self, execution=None):
        # Uncomment to silence Maestro errors when disconnected
        # if self.maestro:
        #     self.maestro.RAISE_NOT_CONNECTED = False

        # Configure whether or not to run on headless mode
        self.headless = False
   
        # Uncomment to change the default Browser to Firefox
        # self.browser = Browser.FIREFOX
        
        # Carregando variáveis de ambiente
        load_dotenv()
        caminho_chrome_driver = os.getenv("CAMINHO_CHROME_DRIVER")
        email = os.getenv("EMAIL")
        password = os.getenv("PASSWORD")
        caminho_arquivo_empresas = os.getenv("CAMINHO_ARQUIVO_EMPRESAS")
        caminho_arquivo_empresas_queue = os.getenv("CAMINHO_ARQUIVO_EMPRESAS_QUEUE")
        caminho_arquivo_input_palavras_chave = os.getenv("CAMINHO_ARQUIVO_INPUT_PALAVRAS_CHAVE")
        # Uncomment to set the WebDriver path
        self.driver_path = caminho_chrome_driver

        # Fetch the Activity ID from the task:
        # task = self.maestro.get_task(execution.task_id)
        # activity_id = task.activity_id
        print("Inicio processo: ")
        inicio = time.time()
        
        prepararArquivo(caminho_arquivo_empresas, caminho_arquivo_empresas_queue)
        login(self, email, password)
        extrairLinkEmpresas(self, caminho_arquivo_input_palavras_chave, caminho_arquivo_empresas_queue)
        extrairInfoEmpresas(self,caminho_arquivo_empresas_queue, caminho_arquivo_empresas)
        logout(self)
        
    
        duracao_segundos, duracao_formatada = finalizar_contagem_tempo(inicio)

        print(f"Tempo total de execução: {duracao_segundos} segundos ({duracao_formatada})")
        
        # Wait for 10 seconds before closing
        #self.wait(100000)

        # Stop the browser and clean up
        #self.stop_browser()
    
def login(self, credential_email, credential_password):
 
    # Opens the BotCity website.
    self.browse("https://www.linkedin.com")
    self.maximize_window()
    # Uncomment to mark this task as finished on BotMaestro
    # self.maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )

    email = self.find_element(selector="/html/body/main/section[1]/div/div/form/div[1]/div[1]/div/div/input", by = By.XPATH)
    password = self.find_element(selector="/html/body/main/section[1]/div/div/form/div[1]/div[2]/div/div/input", by = By.XPATH)
    loginButton = self.find_element(selector="/html/body/main/section[1]/div/div/form/div[2]/button", by = By.XPATH)

    if email and password and loginButton:
        email.send_keys(credential_email)
        password.send_keys(credential_password)
        loginButton.click()
    else:
        print("Não encontrado elemento de login, iniciando fluxo alternativo de login")
        botaoEntrar = self.find_element(selector="/html/body/main/section[1]/div/div/a", by = By.XPATH)
        botaoEntrar.click()
        alternativeEmail = self.find_element(selector="/html/body/div/main/div[2]/div[1]/form/div[1]/input", by = By.XPATH) 
        alternativePassword = self.find_element(selector="/html/body/div/main/div[2]/div[1]/form/div[2]/input", by = By.XPATH)
        alternativeLoginButton = self.find_element(selector="/html/body/div/main/div[2]/div[1]/form/div[3]/button", by = By.XPATH) 

        alternativeEmail.send_keys(credential_email)
        alternativePassword.send_keys(credential_password) 
        alternativeLoginButton.click()


def logout(self):
    botaoEu = self.find_element(selector="//span[contains(.,'Eu')]", by = By.XPATH)
    botaoEu.click()
    self.wait(2000)
    botaoSair = self.find_element(selector="//a[contains(.,'Sair')]", by = By.XPATH)
    botaoSair.click()

    botaoSairDefinitivo = self.find_element(selector="/html/body/div[3]/div/div/div[2]/section/footer/button[2]", by = By.XPATH)

    if botaoSairDefinitivo:
        botaoSairDefinitivo.click()
    

def extrairLinkEmpresas(self, caminho_arquivo_input_palavras_chave, caminho_arquivo_queue):
   
   if not os.path.isfile(caminho_arquivo_input_palavras_chave):
     #ENVIAR EMAIL 
     raise Exception("O arquivo de entrada contendo as palavras-chave não foi encontrado.") 
   
   df = pd.read_excel(caminho_arquivo_queue)
   df_palavras_chave = pd.read_excel(caminho_arquivo_input_palavras_chave)
   df_palavras_chave = df_palavras_chave[['Empresa']].dropna()

   print(df_palavras_chave)

   for index, row in df_palavras_chave.iterrows():
        try:
            palavraChave = row['Empresa']
            # Reinicializa o campo de pesquisa, com exceção da primeira ( que já está vazio )
            if index > 0:
                self.browse("https://www.linkedin.com")

            self.wait(5000)
            barraPesquisa = self.find_element(selector="/html/body/div[5]/header/div/div/div/div[1]/input", by = By.XPATH)
            barraPesquisa.send_keys(palavraChave)
            self.enter()
            buttonEmpresas = self.find_element(selector="//button[contains(.,'Empresas')]", by = By.XPATH) 
            buttonEmpresas.click()    
            buttonTamanhoDaEmpresa = self.find_element(selector="//button[contains(.,'Tamanho da empresa')]", by = By.XPATH)  
            buttonTamanhoDaEmpresa.click() 
            self.wait(5000)
            
            caixa201_500 = self.find_element(selector="//span[contains(.,'201-500')]", by = By.XPATH) 
            caixa501_1000 = self.find_element(selector="//span[contains(.,'501-1.000')]", by = By.XPATH) 
            caixa1001_5000 = self.find_element(selector="//span[contains(.,'1.001-5.000')]", by = By.XPATH) 
            caixa5001_10000 = self.find_element(selector="//span[contains(.,'5.001-10.000')]", by = By.XPATH) 
            caixa10001 = self.find_element(selector="//span[contains(.,'+ de 10.001')]", by = By.XPATH) 
            caixa201_500.click()
            caixa501_1000.click()
            caixa1001_5000.click()
            caixa5001_10000.click()
            caixa10001.click()
            self.tab()
            self.tab()
            self.enter()

            nenhumResultado = self.find_element(selector="/html/body/div[5]/div[3]/div[2]/div/div[1]/main/div/div/div/section/h2", by = By.XPATH) 

            if nenhumResultado:
                # ENVIAR EMAIL BUSINESS RULE EXCEPETION - NENHUM RESULTADO PARA PALAVRA-CHAVE
                print("Nenhum resultado encontrado para " + palavraChave)
            else:
                
                self.scroll_down(clicks=10)
                isAvancarActivated = isElementEnabled(self,'//button[@aria-label="Avançar"]')
                listaEmpresas = self.find_elements(selector='/html/body/div[5]/div[3]/div[2]/div/div[1]/main/div/div/div[2]/div/ul/li/div/div/div/div[2]/div/div[1]/div/span/span/a', by = By.XPATH)

                listaLinkEmpresa = []

                for elem in listaEmpresas:
                    listaLinkEmpresa.append(elem.get_attribute("href"))

                while isAvancarActivated:
                    botaoAvancar = self.find_element(selector='//button[@aria-label="Avançar"]', by = By.XPATH)
                    botaoAvancar.click()
                    self.wait(3000)
                    self.scroll_down(clicks=10)
                    isAvancarActivated = isElementEnabled(self,'//button[@aria-label="Avançar"]')
                    listaEmpresas = self.find_elements(selector='/html/body/div[5]/div[3]/div[2]/div/div[1]/main/div/div/div[2]/div/ul/li/div/div/div/div[2]/div/div[1]/div/span/span/a', by = By.XPATH)
                    for elem in listaEmpresas:
                        listaLinkEmpresa.append(elem.get_attribute("href"))
                                        
                if df.empty:
                        df = pd.DataFrame({'links': listaLinkEmpresa, 'Status': "Não processado"})
                        df.to_excel(caminho_arquivo_queue, index=False)
                else:
                        df2 = pd.DataFrame({'links': listaLinkEmpresa, 'Status': "Não processado"})
                        df_concatenado = pd.concat([df, df2], ignore_index=True)
                        df_concatenado.to_excel(caminho_arquivo_queue, index=False) 
        except Exception as e:
            # Enviar Email de erro System Excpetion indicando a palavra-chave que não foi possível extrair as informações
            print(e)
            print("\nNão foi possível realizar a extração dos links das empresas para palavra-chave: " + palavraChave)


def extrairInfoEmpresas(self, caminho_arquivo_empresas_queue, caminho_arquivo_empresas):

    df_queue = pd.read_excel(caminho_arquivo_empresas_queue)

    if df_queue.empty:
        #Enviar email notificando
        raise Exception("A fila de execução da extração de empresas está vazia. Processo encerrado")
    
    df = pd.DataFrame()
    
    for index, row in df_queue.iterrows():
        try:

            linkedinEmpresa = row['links']
            
            self.browse(linkedinEmpresa)
            self.wait(5000)
            nomeEmpresaElement = self.find_element(selector='h1', by = By.TAG_NAME)                                                                  
            nomeEmpresa = nomeEmpresaElement.text
            botaoSobre = self.find_element(selector="//a[contains(.,'Sobre')]", by = By.XPATH)
            botaoSobre.click()
            self.wait(5000)
            
            fullTextSectionSobre = self.find_element(selector="dl", by = By.TAG_NAME)
            listDl = fullTextSectionSobre.text.split("\n")

            site = 'Não informado'
            telefone = 'Não informado'
            setor = 'Não informado'
            tamanhoEmpresaFuncionarios = 'Não informado'
            tamanhoEmpresaUsuarios = 'Não informado'
            sede = 'Não informado'
            anoFundacao = 'Não informado'
            especializacoes = 'Não informado'

          

            for idx, elem in enumerate(listDl):
                if elem == 'Site':
                    site = listDl[idx + 1]
                if elem == 'Número de telefone':
                    telefone = listDl[idx + 1]
                if elem == 'Tamanho da empresa':
                    tamanhoEmpresaFuncionarios = listDl[idx + 1]
                    if 'usuários associados' in listDl[idx + 2]:
                        tamanhoEmpresaUsuarios = int(listDl[idx + 2].split(" ")[0])

                if elem == 'Sede':
                    sede = listDl[idx + 1]
                if elem == 'Fundada em':
                    anoFundacao = listDl[idx + 1]
                if elem == 'Especializações':
                    especializacoes = listDl[idx + 1]
                if elem == 'Setor':
                    setor = listDl[idx + 1]
                
            if df.empty:
                df = pd.DataFrame(columns=['Empresa', 'linkedinEmpresa', 'Setor', 'TamanhoEmpresaFuncionarios', 'TamanhoEmpresaUsuarios',  'Sede','Site', 'Telefone', 'AnoFundacao', 'Especializacoes', 'Extrair', 'NumeroExtracao'])

            nova_linha = {
            'Empresa': nomeEmpresa,
            'linkedinEmpresa': linkedinEmpresa,
            'Setor': setor,
            'TamanhoEmpresaFuncionarios': tamanhoEmpresaFuncionarios,
            'TamanhoEmpresaUsuarios' : tamanhoEmpresaUsuarios,
            'Sede': sede,
            'Site': site,
            'Telefone': telefone,
            'AnoFundacao': anoFundacao,
            'Especializacoes': especializacoes,
            'Extrair': 0,
            'NumeroExtracao': 0

            }

            df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
            df_queue.at[index, 'Status'] = "Sucesso"
        except Exception as e:
            df_queue.at[index, 'Status'] = "Falha"
            print(e)
            print("\nNão foi possível realizar a extração das informações da empresa de link: " + linkedinEmpresa)

    # Escrever o DataFrame no arquivo Excel
    indices_deletar = df[df['TamanhoEmpresaUsuarios'] == "Não informado"].index
    df_empresas_filtrado = df.drop(indices_deletar)
    df_empresas_filtrado = df_empresas_filtrado[df_empresas_filtrado['TamanhoEmpresaUsuarios'] > 200]
    df_empresas = df.drop(columns=['Extrair', 'NumeroExtracao'])

   
    # Criar um writer usando openpyxl e associar o workbook carregado
    with pd.ExcelWriter(caminho_arquivo_empresas, engine='xlsxwriter') as writer:
        
        # Escrever os DataFrames nas abas especificadas
        df_empresas_filtrado.to_excel(writer, sheet_name='Processamento', index=False)
        df_empresas.to_excel(writer, sheet_name='Base de dados', index=False)
        
   

     # Atualizar o relatório de execução
    relatorio = df_queue[(df_queue['Status'] == "Sucesso") | (df_queue['Status'] == "Falha")]
    relatorio.to_excel(caminho_arquivo_empresas_queue, index=False)

def isElementEnabled(self, seletor):
    element = self.find_element(selector=seletor, by = By.XPATH)

    if element:
        if element.is_enabled():
            return True
        else:
            print("O elemento não está ativo")
            return False
    else:
        print("O elemento não foi encontrado")
        return False
    
def prepararArquivo(caminho_arquivo_empresas, caminho_arquivo_empresas_queue):
    listOfFiles = [caminho_arquivo_empresas, caminho_arquivo_empresas_queue]
    wb = openpyxl.Workbook()
    for file in listOfFiles:
        wb.save(file)


def integrarBase(caminho_arquivo_empresas, caminho_arquivo_pessoas, caminho_arquivo_base):
    df_pessoas = pd.read_excel(caminho_arquivo_pessoas)
    df_empresas = pd.read_excel(caminho_arquivo_empresas, sheet_name="Base de dados")
    df_join = pd.merge(df_pessoas, df_empresas, on='linkedinEmpresa', how='inner')
    base_path = caminho_arquivo_base + "baseLinkedin-" + datetime.now().strftime("%m-%d-%Y%H-%M-%S") + ".xlsx"
    df_join.to_excel(base_path, index=False) 
        
def finalizar_contagem_tempo(inicio):
    fim = time.time()
    duracao_segundos = fim - inicio
    duracao_formatada = time.strftime("%H:%M:%S", time.gmtime(duracao_segundos))
    return duracao_segundos, duracao_formatada

def enviarEmail():
    corpo_email = """<p>hello</p>"""
    msg = email.message.Message()
    msg['Subject'] = "Teste"
    msg['From'] = "vinicius.souza@dkr.tec.br"
    msg['To'] = "vkskayo@gmail.com"

    password = 'vIni2020#'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.office365.com', 587)
    s.starttls()

    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print("Email enviado")


if __name__ == '__main__':
    Bot.main()
