from unittest import skip
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import pandas as pd
from os.path import exists
import variables
from tools import sendEmail, startBrowser

import time
import os

VINCULACION = 5
ESTADO = 8
ESTADOS_EDICION = ['Inicial / Cargadas', 'Con suspensión', 'Pendiente para normalizar']

#open the source file
df = pd.read_excel(open('Actas\\datos.xlsx', 'rb'))
# start browser
path = "C:\\temp"
DELAY = 60  # seconds
browser = startBrowser(path)

file_path = '\\'.join(os.path.realpath(__file__).split('\\')[:-1])

 #<LOGIN>
DEBUG = variables.debug
url = 'https://intrachec.chec.com.co/index.php'
browser.get(url)

try:
    usuario = WebDriverWait(browser, DELAY).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="usuario"]')))
    password = WebDriverWait(browser, DELAY).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="contrasena"]')))
    print("Page is ready!")
except TimeoutException:
    print("Loading took too much time!")
# <AUTENTICACION>
usuario.clear()
usuario.send_keys(variables.usuario)
password.send_keys(variables.password)
password.send_keys(Keys.ENTER)
# </AUTENTICACION>
lastCuenta = 0
for index, row in df.iterrows():
    CUENTA = row['Cuenta']
    if CUENTA == lastCuenta:
        continue
    lastCuenta = CUENTA
    print(f"Procesando la cuenta ---|||  {CUENTA}  |||---")
    browser.switch_to.default_content()
    browser.switch_to.frame('menu')
    menuRecuperacion = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[text()='Recuperacion']")))

    menuRecuperacion.click()

    menuPresolicitudes = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[text()='Presolicitud']")))
    menuPresolicitudes.click()

    menuGestionPresolicitudes = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, "//*[contains(text(), 'Gestion Presolicitudes')]")))
    menuGestionPresolicitudes.click()
    # esperar hasta q el texto procesando... se desaparezca 
    WebDriverWait(browser, DELAY).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="progressbar"]')))
    browser.switch_to.default_content()
    browser.switch_to.frame('principal')
    fechaInicio = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="fecha_inicio"]')))
    
    # esperar a que la tabla se poble
    try:
        WebDriverWait(browser, DELAY/5).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="0"]')))
        time.sleep(1)
    except:
        print('no se pobló la tabla inicialmente')
    
    fechaInicio.clear()
    fechaInicio.send_keys('01-01-2022')

    # filtrar las que estan en estado inicial/cargada
    # WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
    #     (By.XPATH, '//*[@id="estado"]/option[6]')))
    # browser.find_element_by_xpath('//*[@id="estado"]/option[6]').click()

    inputCuenta = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="fil_cuenta"]')))
    inputCuenta.send_keys(CUENTA)

    botonFiltrar = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.CLASS_NAME, 'btn_buscar')))
    botonFiltrar.click()

    # esperar hasta q el texto procesando... se desaparezca 
    WebDriverWait(browser, DELAY).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="progressbar"]')))
    # esperar a que la tabla se poble
    try:
        WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="0"]')))
    except:
        browser.refresh()
        continue

    # traer la tabla de resultados 
    tabla = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.ID, 'jqGridRecepcion')))
    # convertir la tabla en dataframe
    tabla_df = pd.read_html(tabla.get_attribute('outerHTML'))[0]

    hallazgo_ya_gestionado = False
    sinRI = True
    tr_array = tabla.find_elements_by_tag_name('tr')
    if len(tr_array) > 2:
        for tr in tr_array:
            # se salta el ultimo pq ese debe ser procesado
            vinculacion = tr.find_element_by_xpath(f'.//td[{VINCULACION}]').text
            estado = tr.find_element_by_xpath(f'.//td[{ESTADO}]').text
            if vinculacion == 'RI' and sinRI and (estado in ESTADOS_EDICION):
                sinRI = False
                continue
            elif estado in ESTADOS_EDICION:
                try:
                    tr.find_element_by_xpath('.//td[2]').click() # seleccionar la fila de la orden
                    hallazgo_ya_gestionado = True
                except:
                    continue
        if hallazgo_ya_gestionado:    
            # seleccionar hallazgo ya gestionado
            WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
                (By.XPATH,'//*[@id="cestado"]/option[3]'))).click()
            # texto observacion
            sds_final = tr_array[-1].find_element_by_xpath('.//td[3]').text
            WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
                (By.ID,'obs_cambio_estado'))).send_keys(f'SE ENVIAN LOS DATOS CON LA SDS {sds_final}..')
            # click en el boton cambiar
            WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
                (By.CLASS_NAME,'btn_procesar'))).click()
            # Esta seguro de cambiar el estado a ...
            try:
                WebDriverWait(browser, 5).until(EC.alert_is_present(),
                                            'Timed out waiting for PA creation ' +
                                            'confirmation popup to appear.')
                browser.switch_to.alert.accept()
            except:
                print('no alert-151')
            

            try:
                WebDriverWait(browser, 5).until(EC.alert_is_present(),
                                            'Timed out waiting for PA creation ' +
                                            'confirmation popup to appear.')
                browser.switch_to.alert.accept()
            except:
                print('no alert-160')
            time.sleep(3.0)

    WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="0"]/td[2]/input')))
    firstRow = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="0"]')))
    # verificar si la orden coincide con la SDS
    # if(row['Orden'] == row['SDS']):
    #     print('iguales')
    # else:
    #     print('diferentes')

    if firstRow.find_element_by_xpath(f'.//td[{ESTADO}]').text not in ESTADOS_EDICION:
        browser.refresh()
        continue
    sds = int(firstRow.find_element_by_xpath('.//td[3]').text)
    # if sds == row['SDS']:
    while not firstRow.find_element_by_xpath('.//td[2]/input').is_selected():
        firstRow = WebDriverWait(browser, DELAY).until(EC.presence_of_element_located((By.XPATH, '//*[@id="0"]')))
        firstRow.find_element_by_xpath('.//td[2]/input').click() # seleccionar la fila de la orden
    WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, "btn_adjuntos")))
    btn_adjuntos = browser.find_element_by_class_name("btn_adjuntos") # boton generar
    # TODO: hacer click en el botón aceptar del cuadro de dialogo que sale a veces //*[@id="div_ventana"]/center[2]/center/center/center[2]/table/tbody/tr/td[1]/input
    btn_adjuntos.click()
    time.sleep(1.5)
    try:
        WebDriverWait(browser, 10).until(EC.alert_is_present(),
                                    'Timed out waiting for PA creation ' +
                                    'confirmation popup to appear.')
    except:
        if(browser.find_element_by_xpath('//*[@id="div_ventana"]').is_displayed()):
            browser.find_element_by_xpath('//*[@id="div_ventana"]/center[2]/center/center/center[2]/table/tbody/tr/td[1]/input').click()
            time.sleep(1.0)
    # Esta seguro de Generar una solicitud con ...
    browser.switch_to.alert.accept()
    try:
        WebDriverWait(browser, DELAY/4).until(EC.alert_is_present(),
                                    'Timed out waiting for PA creation ' +
                                    'confirmation popup to appear.')
        alert = browser.switch_to.alert
        alert.accept()
        print("alert accepted")
        time.sleep(0.5)
    except TimeoutException:
        print("no alert-185")
    # TODO: revisar esto
    

    origen = Select(browser.find_element(by= By.XPATH, value= '//*[@id="ORIGEN"]'))
    origen.select_by_visible_text('Perdidas')

    # TODO: completar para escoger el tipo de recuperacion y el municipio


    # cambiar envio a soporte clientes si la columna DIRE dice SC (Soporte Clientes)
    if row['DIRE'] == 'SC':
        enviar_a = Select(browser.find_element(by= By.XPATH, value= '//*[@id="DESTINATARIO"]'))
        enviar_a.select_by_visible_text('Soporte Clientes')
    
    WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="acepta"]')))
    # clic en aceptar
    browser.find_element_by_xpath('//*[@id="acepta"]/option[3]').click()
    # clic en adjuntos
    WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.CLASS_NAME, 'btn_adjuntos')))
    browser.find_element_by_class_name('btn_adjuntos').click()
    # agregar los archivos adjuntos
    WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="txt_archivo_acta"]')))
    sinPDFs = True
    contador = 0
    if exists(f'{file_path}\\Actas\\{CUENTA}.pdf'):
        contador = 1
        sinPDFs = False
        adjuntos = browser.find_element_by_xpath('//*[@id="txt_archivo_acta"]')
        adjuntos.send_keys(f'{file_path}\\Actas\\{CUENTA}.pdf')
        # seleccionar que si tiene firma el acta
        WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="verifico"]/option[3]')))
        browser.find_element_by_xpath('//*[@id="verifico"]/option[3]').click()
        # clic al boton Adjuntar acta
        WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="tabla_adjuntos_acta"]/tbody/tr[3]/td/input[2]')))
        browser.find_element_by_xpath('//*[@id="tabla_adjuntos_acta"]/tbody/tr[3]/td/input[2]').click()
        WebDriverWait(browser, 6*DELAY).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="0"]/td[2]')))
    else:
        time.sleep(1)
    # verificar si existe el archivo de instalacion
    if exists(f'{file_path}\\Actas\\{CUENTA}_INS.pdf'):
        sinPDFs = False
        # agregar los archivos adjuntos
        adjuntos = browser.find_element_by_xpath('//*[@id="txt_archivo_acta"]')
        adjuntos.send_keys(f'{file_path}\\Actas\\{CUENTA}_INS.pdf')
        # seleccionar que si tiene firma el acta
        WebDriverWait(browser, 6*DELAY).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="verifico"]/option[3]')))
        browser.find_element_by_xpath('//*[@id="verifico"]/option[3]').click()
        # clic al boton Adjuntar acta
        WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="tabla_adjuntos_acta"]/tbody/tr[3]/td/input[2]')))
        browser.find_element_by_xpath('//*[@id="tabla_adjuntos_acta"]/tbody/tr[3]/td/input[2]').click()
        WebDriverWait(browser, 3*DELAY).until(EC.presence_of_element_located(
            (By.XPATH, f'//*[@id="{contador}"]/td[2]')))
    else:
        time.sleep(1)
    # cerrar el dialogo de adjuntos
    WebDriverWait(browser, 3*DELAY).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'ui-dialog-titlebar-close')))
    browser.find_element_by_class_name('ui-dialog-titlebar-close').click()
    time.sleep(1.0)

    if (sinPDFs):
        browser.find_element_by_class_name("btn_eliminar").click()
        try:
            WebDriverWait(browser, 8).until(EC.alert_is_present(),
                                        'Timed out waiting for PA creation ' +
                                        'confirmation popup to appear.')
            alert = browser.switch_to.alert
            alertText = alert.text
            alert.accept()
            print(f"alert accepted: {alertText}")
            time.sleep(0.5)
            browser.refresh()
            continue
        except TimeoutException:
            print("no alert-291")

    # clic al boton editar
    WebDriverWait(browser, 3*DELAY).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'btn_editar')))
    browser.find_element_by_class_name('btn_editar').click()
    time.sleep(1.0)
    # ingresar el numero de la ot de normalizacion
    ot_normalizacion_inpt = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="orden_trabajo"]')))
    ot_normalizacion_inpt.send_keys(row['SDS'])
    # ingresar la serie del medidor
    serie_md_inpt = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="add_medidor_serie"]')))
    serie_md_inpt.send_keys(row['MED'])
    # ingresar la marca del medidor
    marca_md_inpt = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="add_medidor"]')))
    marca_md_inpt.send_keys(row['MAR'])

    # ingresar texto motivo novedad fecha normalizacion
    try:
        motivo_txt = WebDriverWait(browser, 1).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="mot_nov_normalizacion"]')))
        motivo_txt.send_keys('0')
    except:
        print('no estaba el campo motivo novedad')
    # ingresar la energia estimada
    WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="add_energia"]')))
    browser.find_element_by_xpath('//*[@id="add_energia"]').send_keys('0')
    # clic en el boton guardar
    WebDriverWait(browser, 5).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'btn_guardar')))
    browser.find_element_by_class_name('btn_guardar').click()

    # aceptar el mensaje
    try:
        WebDriverWait(browser, 8).until(EC.alert_is_present(),
                                    'Timed out waiting for PA creation ' +
                                    'confirmation popup to appear.')
        alert = browser.switch_to.alert
        alert.accept()
        print("alert accepted-272")
        time.sleep(0.5)
    except TimeoutException:
        print("no alert-275")
    time.sleep(1)
    try:
        if (marca_md_inpt.is_displayed()):
            time.sleep(3)
            if (browser.find_element_by_xpath("//*[contains(text(), 'no existe en la base de datos')]").is_displayed()):
                browser.find_element_by_xpath("//*[text()='Ok']").click()
                browser.find_element_by_class_name("ui-dialog-titlebar-close").click()
                time.sleep(1.5)
                browser.find_element_by_class_name("btn_eliminar").click()
                try:
                    WebDriverWait(browser, 8).until(EC.alert_is_present(),
                                                'Timed out waiting for PA creation ' +
                                                'confirmation popup to appear.')
                    alert = browser.switch_to.alert
                    alertText = alert.text
                    alert.accept()
                    print(f"alert accepted: {alertText}")
                    time.sleep(0.5)
                    browser.refresh()
                    continue
                except TimeoutException:
                    print("no alert-321")
    except:
        print('dialogo cerrado 324')
    btn_enviar = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.CLASS_NAME, 'btn_procesar')))
    btn_enviar.click()
    WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable(
        (By.CLASS_NAME, 'ui-button-text-only')))
    browser.find_element_by_class_name('ui-button-text-only').click()
    try:
        WebDriverWait(browser, 8).until(EC.alert_is_present(),
                                    'Timed out waiting for PA creation ' +
                                    'confirmation popup to appear.')
        alert = browser.switch_to.alert
        lote = alert.text
        lote = lote.split(':')[1]
        alert.accept()
        print(f"alert accepted- LOTE: {lote}")
        time.sleep(0.5)
    except TimeoutException:
        print("no alert-291")
        #</if>
    #</else>
    #</for>
    #recargar la pagina para volver a empezar
    browser.refresh()
    

salida_segura = WebDriverWait(browser, DELAY).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="master"]/div[6]')))
salida_segura.click()

WebDriverWait(browser, DELAY).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="usuario"]')))

browser.close()

print("!!!!FIN ROBOT!!!")


