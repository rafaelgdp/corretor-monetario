from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from io import BytesIO
from PIL import Image
import math
from PyPDF2 import PdfFileWriter, PdfFileReader
import io, os
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def get_valor_str(valor_original):
    f, i = math.modf(valor_original)
    valor_original_str = str(int(i)) + str("%.2f" % f)[2:]
    return valor_original_str

def get_mes_ano_by_str(data):
    mes = int(data[0:2])
    ano = int(data[2:])
    return mes, ano

def get_str_by_mes_ano(mes, ano):
    r = "%02d" % mes
    r += "%04d"% ano
    return r

def buscar_correcao(inicio, fim, valor):
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)
    fp.set_preference("browser.download.dir", "./")
    fp.set_preference("plugin.disable_full_page_plugin_for_types", "application/pdf")
    fp.set_preference("pdfjs.disabled", True)
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
    driver = webdriver.Firefox(fp)
    driver.implicitly_wait(10)
    driver.maximize_window()
    driver.get("https://www3.bcb.gov.br/CALCIDADAO/publico/exibirFormCorrecaoValores.do?method=exibirFormCorrecaoValores")
    Select(driver.find_element_by_id("selIndice")).select_by_value("00188INPC")
    driver.find_element_by_name("dataInicial").send_keys(inicio)
    driver.find_element_by_name("dataFinal").send_keys(fim)
    valor_correcao_element = driver.find_element_by_name("valorCorrecao")
    valor_correcao_element.clear()
    valor_correcao_element.send_keys(valor)
    driver.find_element(By.XPATH, '/html/body/div[7]/table/tbody/tr[3]/td/div/form/div[2]/input[1]').click()
    # driver.find_element(By.XPATH, '/html/body/div[7]/table/tbody/tr/td/div[2]/table[2]/tbody/tr/td[2]/input').click()
    driver.execute_script("document.body.style.scale='1'")
    return driver

meses = [
        "None", "janeiro", "fevereiro", "marco", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]

def get_nome_arquivo(mes, ano):
    nome_arquivo = meses[mes] + " " + str(ano) + ".pdf"
    return nome_arquivo

def get_nome_arquivo_temp(mes, ano):
    nome_arquivo = meses[mes] + " " + str(ano) + " temp.pdf"
    return nome_arquivo
    
def salvar_pdf(driver, mes, ano):
    img = Image.open(BytesIO(driver.find_element_by_tag_name('body').screenshot_as_png))
    
    # Margens
    top = 50
    bottom = 100
    left = 50
    right = 50
    
    tamanho_maior = (img.size[0] + left + right, img.size[1] + top + bottom)
    rgb = Image.new('RGB', tamanho_maior, (255, 255, 255))
    rgb.paste(img, box=(left, top),  mask=img.split()[3])
    rgb.save(get_nome_arquivo_temp(mes, ano), "PDF", quality=100, resolution=100.0)
    
    # Adicionando rodapé
    recuo_esquerda = left
    recuo_baixo = bottom // 3
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    data_atual = datetime.now()
    mensagem_informativa = "Cálculo realizado no dia "
    mensagem_informativa += "%d de %s de %04d" % (data_atual.day, meses[data_atual.month], data_atual.year)
    mensagem_informativa += ", %02d:%02d" % (data_atual.hour, data_atual.minute)
    mensagem_informativa += ", em João Pessoa, com base no site do Banco Central do Brasil."
    can.drawString(recuo_esquerda, recuo_baixo, mensagem_informativa)
    mensagem_informativa = "Disponível em "
    mensagem_informativa += "https://www3.bcb.gov.br/CALCIDADAO/publico/exibirFormCorrecaoValores.do?method=exibirFormCorrecaoValores."
    can.drawString(recuo_esquerda, recuo_baixo - (recuo_baixo // 2), mensagem_informativa)
    can.save()

    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    temp_file = open(get_nome_arquivo_temp(mes, ano), "rb")
    existing_pdf = PdfFileReader(temp_file)
    output = PdfFileWriter()
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    
    # Escrevendo PDF final
    outputStream = open(get_nome_arquivo(mes, ano), "wb")
    output.write(outputStream)
    outputStream.close()
    temp_file.close()
    os.remove(get_nome_arquivo_temp(mes, ano))
    driver.quit()

def calcular_periodo(inicio, fim, periodo, valor):
    if periodo == "mensal":
        mes, ano = get_mes_ano_by_str(inicio)
        mes_fim, ano_fim = get_mes_ano_by_str(fim)
        while(True):
            if (mes == mes_fim and ano == ano_fim):
                break
            proximo_mes = mes + 1
            proximo_ano = ano
            if proximo_mes == 13:
                proximo_mes = 1
                proximo_ano = ano + 1
            print("Criando correção referente a %s de %04d." % (meses[mes], ano))
            driver = buscar_correcao(get_str_by_mes_ano(mes, ano), get_str_by_mes_ano(proximo_mes, proximo_ano), get_valor_str(valor))
            salvar_pdf(driver, mes, ano)
            mes = proximo_mes
            ano = proximo_ano
    else:
        driver = buscar_correcao(inicio, fim, get_valor_str(valor))
        mes, ano = get_mes_ano_by_str(inicio)
        salvar_pdf(driver, mes, ano)

# data_inicial = "112018"
# periodicidade = "mensal" # "initerrupto"
# data_final = "122018"
# valor_original = 100.00
# calcular_periodo(data_inicial, data_final, periodicidade, valor_original)
# print("Finalizado!")

import kivy
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox

class ConversorMonetarioApp(App):

    def build(self):

        self.nao_errou = True

        main_box = BoxLayout(orientation='vertical')
        main_box.padding = 50
        main_box.size_hint.y = None
        main_box.size.y = 100
        title = Label(text="Automatizador de Conversão Monetária Mensal :P")
        main_box.add_widget(title)

        input_padding = 5

        valor_inicial_box = BoxLayout(orientation='horizontal')
        valor_inicial_box.padding = input_padding
        valor_inicial_label = Label(text="Valor Original Fixo")
        valor_inicial_input = TextInput(hint_text="0.00")
        valor_inicial_box.add_widget(valor_inicial_label)
        valor_inicial_box.add_widget(valor_inicial_input)
        main_box.add_widget(valor_inicial_box)

        prazo_inicial_box = BoxLayout(orientation='horizontal')
        prazo_inicial_box.padding = input_padding
        prazo_inicial_label = Label(text="Prazo Inicial (Formato MMAAAA)")
        prazo_inicial_input = TextInput(hint_text="MMAAAA")
        prazo_inicial_box.add_widget(prazo_inicial_label)
        prazo_inicial_box.add_widget(prazo_inicial_input)
        main_box.add_widget(prazo_inicial_box)

        prazo_final_box = BoxLayout(orientation='horizontal')
        prazo_final_box.padding = input_padding
        prazo_final_label = Label(text="Prazo Final (Formato MMAAAA)")
        prazo_final_input = TextInput(hint_text="MMAAAA")
        prazo_final_box.add_widget(prazo_final_label)
        prazo_final_box.add_widget(prazo_final_input)
        main_box.add_widget(prazo_final_box)

        periodicidade_box = BoxLayout(orientation='horizontal')
        periodicidade_box.padding = input_padding
        periodicidade_label = Label(text="Periodicidade do Cálculo")
        mensal_cb = CheckBox(color=[1, 1, 1, 1])
        mensal_label = Label(text="Mês a Mês")
        mensal_cb.active = True
        mensal_cb.group = "periodo"
        mensal_cb.color = (1, 1, 1, 1)
        mensal_cb.bind(active=self.on_mensal_active)
        corrido_cb = CheckBox()
        corrido_label = Label(text="Meses Corridos")
        corrido_cb.color = (1, 1, 1, 1)
        corrido_cb.active = False
        corrido_cb.group = "periodo"
        corrido_cb.bind(active=self.on_corrido_active)
        periodicidade_box.add_widget(periodicidade_label)
        periodicidade_box.add_widget(mensal_cb)
        periodicidade_box.add_widget(mensal_label)
        periodicidade_box.add_widget(corrido_cb)
        periodicidade_box.add_widget(corrido_label)
        main_box.add_widget(periodicidade_box)

        self.inicio_input = prazo_inicial_input
        self.fim_input = prazo_final_input
        self.valor_input = valor_inicial_input
        self.main_box = main_box
        self.periodicidade = "mensal"

        botao = Button(text="Iniciar Busca", on_release=self.buscar_informacoes)
        botao.margin = (80, 80)
        main_box.add_widget(botao)

        limpar_botao = Button(text="Resetar Valores", on_release=self.resetar_valores)
        limpar_botao.margin = (80, 80)
        main_box.add_widget(limpar_botao)

        self.erro_label = Label(text="Você digitou dados em formatos errados. Confira novamente!")

        main_box.add_widget(Label(text="Software ainda em desenvolvimento. Conferir informações antes de usar."))
        main_box.add_widget(Label(text="Feito por Rafael :P"))

        return main_box

    def on_mensal_active(self, cb, v):
        self.periodicidade = "mensal" if v else "corrido"

    def on_corrido_active(self, cb, v):
        self.periodicidade = "mensal" if not v else "corrido"

    def resetar_valores(self, botao):
        self.inicio_input.text = ""
        self.fim_input.text = ""
        self.valor_input.text = ""
        self.mensal_check.color = (1, 1, 1, 1)

    def validar(i, f, v):
        try: 
            int(i)
            int(f)
            float(v)
        except ValueError:
            return False
        if (len(i) != 6 and len(f) != 6):
            return False
        return True

    def buscar_informacoes(self, botao):
        inicio = self.inicio_input.text
        fim = self.fim_input.text
        valor = self.valor_input.text
        if (ConversorMonetarioApp.validar(inicio, fim, valor)):
            self.main_box.remove_widget(self.erro_label)
            calcular_periodo(inicio, fim, self.periodicidade, float(valor))
        else:
            if (self.nao_errou):
                self.main_box.add_widget(self.erro_label)
                self.nao_errou = False


if __name__ == '__main__':
    ConversorMonetarioApp().run()