# -*- coding: utf-8 -*-
from selenium.webdriver import Chrome
from selenium.webdriver.firefox.options import Options
from selenium.common import exceptions
from datetime import datetime
import time
import json
import logging

logger = logging.getLogger()
logger.setLevel(logging.INFO)

sh = logging.StreamHandler()
sh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - line %(lineno)s - %(message)s'))

logger.addHandler(sh)

logger.info("Iniciando Timecheat - O programa dos preguiçosos que não conseguem preencher a Timesheet da Claranet.")

# Read Data
logger.info("Lendo arquivo de JSON com as configurações.")
json_data = json.loads(open('data.json', encoding='utf-8').read())


transformed_tmsht = []

for timesheet in json_data['timesheet']:
    if 'dt_inicio' not in timesheet:
        timesheet_init = {
            "ent": timesheet['ent'],
            "dt_inicio": timesheet['dt'] + " 08:00",
            "dt_fim": timesheet['dt'] + " 12:00",
            "desc": timesheet['desc']
        }
        timesheet_end = {
            "ent": timesheet['ent'],
            "dt_inicio": timesheet['dt'] + " 13:00",
            "dt_fim": timesheet['dt'] + " 17:00",
            "desc": timesheet['desc']
        }
        transformed_tmsht.append(timesheet_init)
        transformed_tmsht.append(timesheet_end)
    else:
        transformed_tmsht.append(timesheet)


# Validate All Timesheet
logger.info("Validando dados.")

for idx, timesheet in enumerate(transformed_tmsht):
    try:
        dt_ini = datetime.strptime(timesheet['dt_inicio'], '%d/%m/%Y %H:%M')
        dt_fim = datetime.strptime(timesheet['dt_fim'], '%d/%m/%Y %H:%M')

        if (dt_fim - dt_ini).seconds > 14400:
            logger.error("Problema na data na linha %s. O periodo de tempo é maior que 4 horas." % str(idx + 1))
            exit(-1)
    except ValueError:
        logger.error("Problema na data na linha %s. O formato de data é dd/mm/yyyy hh:mm." % str(idx + 1))
        exit(-1)

logger.info("Fase de validação finalizada.")

logger.info("Iniciando o plugin do Selenium para preenchimento dos dados.")

opts = Options()
opts.headless = True

try:
    browser = Chrome(options=opts)

    browser.get('https://portalcredibilit.force.com/support/login?ec=302&inst=0B&startURL=%2Fsupport%2F04i')

    username_form = browser.find_element_by_id('username')
    username_form.send_keys(json_data['username'])

    username_form = browser.find_element_by_id('password')
    username_form.send_keys(json_data['password'])

    browser.find_element_by_id('Login').click()

    browser.find_element_by_css_selector("[title^='Guia Timesheets']").click()
    browser.find_element_by_css_selector("[title^='Criar Timesheet']").click()

    # ForEach Timesheet
    for timesheet in transformed_tmsht:
        entrega_form = browser.find_element_by_id('CF00NU0000003X1c5')
        entrega_form.send_keys(timesheet['ent'])

        dt_inicio_form = browser.find_element_by_id('00NU0000003VVEx')
        dt_inicio_form.send_keys(timesheet['dt_inicio'])

        dt_fim_form = browser.find_element_by_id('00NU0000003VVF2')
        dt_fim_form.send_keys(timesheet['dt_fim'])

        time.sleep(2)
        browser.execute_script("document.getElementsByClassName('cke_wysiwyg_frame cke_reset')[0]."
                               "contentWindow.document.body.innerText = '%s';" % timesheet['desc'])
        browser.find_element_by_css_selector("[title^='Salvar e novo']").click()

    browser.find_element_by_css_selector("[title^='Guia Timesheets']").click()

except exceptions.WebDriverException:
    logger.error("Problema de WebDriver, o Chrome Driver pode não estar instalado !!!")
    logger.info(" ----- Passo 1: Baixe o Chrome Driver em http://chromedriver.chromium.org/downloads.")
    logger.info(" ----- Passo 2: Coloque o executavel em um path do SO (Por exemplo, na pasta System32 do Windows)")

logger.info("Valeu e falou.")