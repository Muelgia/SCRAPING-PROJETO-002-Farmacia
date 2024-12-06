from funcoes import navegador, encontrar_ocorrencia
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
    countP = 0
    while True:

        wait = WebDriverWait(driver,10)
        
        lista = '.prateleira.vitrine.default.n48colunas.view_item_list_success'
        listaXpath = driver.find_elements(By.CSS_SELECTOR, lista)

        for item in listaXpath:

            li_elements = item.find_elements(By.TAG_NAME, 'li')

            for item in li_elements:

                count = 0 
                # Itere sobre cada 'li'
                for index, li in enumerate(li_elements, start=1):
                    
                    try:
                        
                        infos = li.text
                        
                        substituicoes = [
                            'ATIVAR DESCONTO\n',
                            'ADICIONAR AOS FAVORITOS',
                            'COMPRAR 3',
                            '60% OFF NA 3º UNIDADE\n',
                            '50% OFF NA 2º UNIDADE\n',
                            'Patrocinado\n',
                            'COMPRAR\n',
                            'Medicamento Refrigerado\n',
                            ';DESC. LABORATÓRIO\n',
                            'LEVE 3 PAGUE 2\n',
                            'LEVE 3 COM 25% OFF\n',
                        ]

                        # Substituindo cada texto indesejado em 'infos'
                        for texto in substituicoes:
                            infos = infos.replace(texto, '')
                        
                        infos = infos.strip()

                        infos = infos.replace('\n', ';') 
                        
                        achar = encontrar_ocorrencia(infos, ';', 3)
                        if achar == infos.find(';DESC. LABORATÓRIO'):
                            infos = infos.replace(';DESC. LABORATÓRIO', '')

                        infos = infos[:encontrar_ocorrencia(infos, ';', 5)] + '\n'

                        print(infos)

                        driver.execute_script("""
                        var element = arguments[0];
                        if (element) {
                            element.parentNode.removeChild(element);
                        }
                    """, li)
                        
                        countP += 1

                        with open('infos', 'a', encoding='utf-8') as ultima:
                            ultima.write(infos)
                        ultima.close()
                    
                    except:
                        pass


                #botao = driver.find_element(By.CSS_SELECTOR, ".rnk-comp-carregar-mais button.btn.btn-outline-dark")
                botao = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".rnk-comp-carregar-mais button.btn.btn-outline-dark")))

                try:

                    action = ActionChains(driver)
                    action.move_to_element(botao).perform()
                    botao.click()

                except:
                    
                    break
                
                print(f'Ok {categoria}')
                print(countP)
                break


