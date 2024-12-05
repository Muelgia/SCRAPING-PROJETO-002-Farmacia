from funcoes import navegador
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains

driver = navegador(8888, "WebScraper")

categorais = ('medicamentos', 'saude-bem-estar', 'mamae-e-bebe',  'beleza', 'cuidados-para-cabelos', 'higiene-pessoal', 'busca/?fq=H:13253&utmi_p=_marca_propria&utmi_pc=Html&utmi_cp=produtos')

for categoria in categorais:

    driver.get(f'https://www.drogariasaopaulo.com.br/{categoria}')

    while True:

        wait = WebDriverWait(driver,60)

        #botao = driver.find_element(By.CSS_SELECTOR, ".rnk-comp-carregar-mais button.btn.btn-outline-dark")
        botao = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rnk-comp-carregar-mais button.btn.btn-outline-dark")))

        try:

            action = ActionChains(driver)
            action.move_to_element(botao).perform()
            botao.click()

        except:
            
            break
        
        print(f'Ok {categoria}')
        
        

        

     
    