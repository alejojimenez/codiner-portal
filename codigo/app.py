import os
import time
import shutil
import requests
import pandas as pd
import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains

class Scraper_Codiner():

    def __init__(self,url, email, password, driver_path):
        print(url, email, password, driver_path)
        self.url = url
        self.email = email
        self.password = password
        self.driver_path = driver_path

    def wait(self, seconds):
        return WebDriverWait(self.driver, seconds)

    def close(self):
        self.driver.close()
        self.driver = None

    def quit(self):
        self.driver.quit()
        self.driver = None    

    def login(self):
        
        driver_exe = 'C:\\roda\\codiner-portal\\domain\\chromedriver.exe'
        credencials = 'C:\\roda\\codiner-portal\\config\\credenciales.xlsx'

        print('Entrando en la funcion login...')
        print('----------------------------------------------------------------------')
        
        #Seteo variables
        email = self.email
        url = self.url
        driver_path = self.driver_path
        password  = self.password
        
        options = webdriver.ChromeOptions()
            
        self.driver = webdriver.Chrome(driver_path, options=options)
        self.driver.get(url)
        self.driver.maximize_window()
        self.driver.implicitly_wait(40)

        # Controlar evento alert() Notificacion
        try:
            alert = WebDriverWait(self.driver, 35).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
                    
        except:
            print("No se encontró ninguna alerta.")  

        # Seleccionar codigo de servicio
        intentos = 0
        servicio = True
        while (servicio):
            try:
                print('Try en la funcion codigo de servicio...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_codigo= self.driver.find_element(By.ID, 'codigo')
                element_codigo.click()
                element_codigo.send_keys('200-7')
                servicio = False
            except:    
                print('Exception en la funcion codigo de servicio')
                print('----------------------------------------------------------------------')
                servicio = intentos <= 3                

        # Hacer click para ingresar
        intentos = 0
        ingreso = True
        while (ingreso):
            try:
                print('Try en la funcion click para ingresar..', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                time.sleep(5) 
                button_element = self.driver.find_element(By.CLASS_NAME, "btn-u")
                if not button_element.is_displayed():
                    action_chains = ActionChains(self.driver)
                    action_chains.move_to_element(button_element).perform()
                button_element.click()
                ingreso = False   
            except:
                print('Exception en la funcion click para ingresar')
                print('----------------------------------------------------------------------')
                ingreso = intentos <= 3      
         
    def scrapping_codiner(self):
        
        print('Entrando en la funcion Scrapping...')
        print('----------------------------------------------------------------------')
        
        folder_path_config = './config/'
        
        # Especifica la ruta de tu archivo Excel
        excel_file = folder_path_config + "clientes.xlsx"

        # Especifica el nombre de la hoja en la que se encuentran los datos
        hoja_excel = "Hoja1"

        # Carga los datos de Excel en un DataFrame
        df = pd.read_excel(excel_file, sheet_name=hoja_excel)
        print('Dataframe ', df)
        print('----------------------------------------------------------------------')
        
        #Esperamos detectar iframe para poder obtener los datos de la tabla
        wait = WebDriverWait(self.driver, 10)
        iframe_element = wait.until(EC.presence_of_element_located((By.NAME, "window")))
        
        #Hacemos el cambio de iframe    
        self.driver.switch_to.frame(iframe_element)
        
        #Hacemos diccionario para completar con url y fecha
        dicc = {}
        
        #Iteramos para obtener las url de descarga y la fecha
        i=1
        while i < 6:
        
            try:
                time.sleep(5)
                #Buscamos el elemento que contiene la url
                td_element = self.driver.find_element(By.ID, f"td{i}")
                a_element = td_element.find_element(By.TAG_NAME, "a")
                #Guardamos la url y la fecha
                url_a_hacer_clic = a_element.get_attribute("href")
                fecha = a_element.text
                dicc[f"{url_a_hacer_clic}"] = fecha
                time.sleep(5)
            except:
                print('No se han encontrado archivos de descarga')
                print('----------------------------------------------------------------------')
                    
            i += 2

        #Iteramos sobre diccionario para descargar pdf
        for url, date in dicc.items():
            
            #Separamos la fecha para tener mes y año
            mes,año = date.split()
            #Ejecutamos la url en otra ventana
            self.driver.execute_script("window.open(arguments[0]);", url)
        
            # Esperar a que se abra la ventana emergente
            intentos = 0
            reintentar_ventana = True
            while (reintentar_ventana):
                    try:
                        # Obtiene el identificador de la ventana actual
                        current_window = self.driver.current_window_handle
                        print('Ventana principal: ', current_window)
                        print('----------------------------------------------------------------------')
                        print('Try en la funcion manejo de ventanas abiertas...', intentos)
                        print('----------------------------------------------------------------------')
                        intentos += 1
                        window_handles_all = self.driver.window_handles
                        reintentar_ventana = False
                            
                    except:    
                        print('Exception en la funcion manejo de ventanas abiertas')
                        print('----------------------------------------------------------------------')
                        print('----------------------------------------------------------------------')
                        time.sleep(60) #espera para que cargue ventana emergente
                                    
                        reintentar_ventana = intentos <= 3
                                
            # Obtiene los identificadores de las ventanas abiertas
            self.driver.implicitly_wait(35)
            print('Ventanas abiertas: ', window_handles_all, len(window_handles_all))
            print('----------------------------------------------------------------------')            
                                
            # Cambiar al manejo de la ventana emergente
            for window_handle in window_handles_all:
                    if window_handle != current_window:
                        self.driver.switch_to.window(window_handle)
                        print('Ventana emergente: ', window_handle)
                        print('----------------------------------------------------------------------')
                        break

            # Esperar hasta que el elemento esté presente en la página
            self.driver.implicitly_wait(35)
                            
            try:
                # Obtener la URL de la ventana emergente
                ventana_emergente_url = self.driver.current_url
                print("URL de la ventana emergente:", ventana_emergente_url)
                print('----------------------------------------------------------------------')

                # Realizar una solicitud GET para obtener la data binaria del documento
                response = requests.get(ventana_emergente_url, stream=True)

                #Expresion regular para extraer 
                string_split = ventana_emergente_url
                regexp_split = re.split(pattern = r"[>: \: \ \_ \: \< ]", string = string_split)
                print('Resultante expresion regular con split: ', regexp_split[3][:-4])
                print('--------------------------------------------------------------------------')
                bill_number = regexp_split[3][:-4]

                # Cruce datos faltantes para ontener
                for index, row in df.iterrows():
                    
                    df_nro_cliente = df.loc[index, 'nro_cliente']
                    df_sucursal = df.loc[index, 'sucursal']
                    print('Nro. Cliente: ', df_nro_cliente, 'Sucursal: ', df_sucursal)
                    print('--------------------------------------------------------------------------')

                # Obtener el nombre del archivo a partir de los datos del proceso de descarga
                folder_path = './input/'
                file_name = folder_path + str(df_nro_cliente) +"_"+ str(bill_number)+"_"+ str(f'{año}')+".pdf" 

                # Guardar la data binaria en un archivo PDF
                with open(file_name, 'wb') as file:
                    response.raw.decode_content = True
                    shutil.copyfileobj(response.raw, file)
                    print("Guardando archivo:", file_name)
                    print('----------------------------------------------------------------------')
                                    
            except:    
                print("No se encontró el elemento con el id especificado...")
                print('----------------------------------------------------------------------')
                                
                count += 1
                print('Conteo de documentos: ', count)
                print('----------------------------------------------------------------------')                

            # Cerrar la ventana emergente
            time.sleep(5)
            self.driver.close()
            time.sleep(15)                            

            # Cambiar de nuevo al manejo de ventana principal
            self.driver.switch_to.window(current_window)
            print('Cual ventana es: ', current_window)
            print('----------------------------------------------------------------------')
               
            #Hacemos una pequeña pausa antes de pasar al siguiente archivo         
            time.sleep(5)
                
            print('pasamos al siguiente archivo')
