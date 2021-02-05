from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from python_anticaptcha import AnticaptchaClient, ImageToTextTask
from selenium.common.exceptions import NoSuchElementException, TimeoutException

import pandas as pd
import os
import time

class ScrapingCNOM:
    
    def __init__(self):
        
        self.url = 'https://www.conseil-national.medecin.fr/annuaire'
        self.df = pd.DataFrame()
            
    def solve_captcha(self, driver):
        driver.find_element_by_xpath("//img[@title = 'Image CAPTCHA']").screenshot('captcha_0.png')
        try:
            driver.find_element_by_id('edit-accept-tos')
            driver.execute_script("document.getElementById('edit-accept-tos').click()")
        finally:
            api_key = 'mykey'
            captcha_fp = open('captcha_0.png', 'rb')
            client = AnticaptchaClient(api_key)
            task = ImageToTextTask(captcha_fp)
            job = client.createTask(task)
            job.join()
            captcha = job.get_captcha_text()
            driver.find_element_by_id('edit-captcha-response').send_keys(captcha)
            driver.find_element_by_id('edit-op').click()
            return driver
        
    def access(self):
        options = webdriver.FirefoxOptions()
        options.add_argument('--headless')
        driver = webdriver.Firefox(options=options)
        driver.get(self.url)
        try:
            driver.find_element_by_xpath('//a[text() = "Sauvegarder mes préférences"]').click()
            pass
        except NoSuchElementException:
            pass
        regions = driver.find_element_by_id('region').get_attribute('innerHTML')
        soup = BeautifulSoup(regions, 'lxml')
        regions = soup.find_all('option')
        for reg in regions[1:4]:
            driver.get(self.url)
            driver.find_element_by_xpath('//option[contains(text(), "OPHTALMOLOGI")]').click()
            driver.execute_script("document.getElementById('edit-statut-1').setAttribute('checked','checked')")
            driver.execute_script("document.getElementById('edit-statut-3').removeAttribute('checked')")
            driver.find_element_by_xpath(f'//option[text() = "{reg.text}"]').click()
            driver.find_element_by_id('submit_adv_search').click()
            try:
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//img[@title = 'Image CAPTCHA']")))
                driver = self.solve_captcha(driver)
            except TimeoutException:
                pass
            liste_medecin = []
            verificateur =  True
            while verificateur:
                liste_medecin.extend([x.get_attribute('innerHTML') for x in driver.find_elements_by_class_name('search-med-result-item-wrapper')])
                try:
                    driver.find_element_by_class_name("pager-next").click()
                    try:
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//img[@title = 'Image CAPTCHA']")))
                        driver = self.solve_captcha(driver)
                    except TimeoutException:
                        pass
                except NoSuchElementException:
                    verificateur = False
                    print(len(liste_medecin))
                    self.organisateur(liste_medecin) 
        driver.quit()
        self.excel_writter()

       
    def organisateur(self, liste_medecin):
        for medecin in liste_medecin:
            departement_liste = []
            rpps_liste = []
            discipline_liste = []
            discipline_comp_liste = []
            autres_liste = []
            adresse_liste = []
            tel_liste = []
            fax_liste = []
            dicitionaire = {}
            soup = BeautifulSoup(medecin, 'lxml')
            medecin = soup.text
            medecin = medecin.split('\n')
            medecin = [x for x in medecin if x]
            nom = medecin[0]
            prenom = nom.split(' ')[0]
            departement_liste.extend([x.split(' : ')[1] for x in medecin if "Département d'inscription" in x])
            rpps_liste.extend([x.split(' : ')[1] for x in medecin if "RPPS" in x])
            discipline_liste.extend([medecin[x+1].title() for x in range(len(medecin)) if 'Discipline exercée' in medecin[x]])
            discipline_comp = [medecin[x+1].title() for x in range(len(medecin)) if 'Disciplines complémentaires' in medecin[x]]
            if 'Autres Titres Et Orientations Autorisés' in discipline_comp[0]:
                discipline_comp = 'null'
                discipline_comp_liste.append(discipline_comp)
            else:
                discipline_comp_liste.extend(discipline_comp)
            autres = [medecin[x+1].title() for x in range(len(medecin)) if 'Autres titres et ' in medecin[x]]
            if 'Adresse :' in autres[0]:
                autres = 'null' 
                autres_liste.append(autres)
            else:
                autres_liste.extend(autres)
            index = medecin.index('Adresse : ')
            for adress in medecin[index+1:]:
                if 'Tél :' not in adress and 'Fax :' not in adress:
                    adresse_liste.append(adress)
                else:
                    break
            for tel in medecin[index+1:]:    
                if 'Tél :' in adress:
                    tel = adress.split(' : ')[1]
                    if tel:
                        tel_liste.append(tel)
                    else:
                        tel_liste.append('null')
                else:
                    break
            for fax in medecin[index+1:]:
                if 'Fax :' in adress:
                    fax = adress.split(' : ')[1]
                    if fax:
                        fax_liste.append(tel)
                    else:
                        fax_liste.append('null')
            if not fax_liste:
                fax_liste.append('null')
            j = 0
            for adress in adresse_liste:
                dicitionaire[f'Adress_{j}'] = adress
                j += 1
            dicitionaire['Nom'] = nom
            dicitionaire['Prenom'] = prenom
            dicitionaire["Département d'inscription"] = departement_liste[0]
            dicitionaire['RPPS'] = rpps_liste[0]
            dicitionaire['Discipline exercée'] = discipline_liste[0]
            dicitionaire['Disciplines complémentaires'] = discipline_comp_liste[0]
            dicitionaire['Autres titres'] = autres_liste[0]
            dicitionaire['Tél'] = tel_liste[0]
            dicitionaire['Fax'] = fax_liste[0]
            print(dicitionaire)
            self.df = self.df.append(dicitionaire, ignore_index=True)
        print(self.df)
        
    def excel_writter(self):
        writer = pd.ExcelWriter(os.path.join(os.getcwd(), 'medicin_table.xlsx'), engine='xlsxwriter')
        self.df.to_excel(writer, sheet_name='solicitacao')
        worksheet = writer.sheets['solicitacao']
        for idx, col in enumerate(self.df):
            tamanho = 0
            for linha in self.df[col]:
                linha = str(linha)
                if len(linha) > tamanho:
                    tamanho = len(linha)
            if len(col) > tamanho:
                tamanho = len(col)
            tamanho += 1
            worksheet.set_column((idx + 1), (idx + 1), tamanho)
        writer.save()
        return print('Procédure finalisée')

if __name__ == '__main__':
    init_time = time.time()
    ScrapingCNOM().access()
    end_time = time.time()
    print(f'time elapsed {end_time - init_time}')