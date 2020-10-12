from selenium import webdriver
from time import sleep
from pandas import DataFrame

def return_num(num):
    i=0
    for c in num:
        if c == '&':
            return int(num[:i])
        i += 1
    return num

def get_costs(driver):
    cost_flex, clost_fix, cost_ofice, cost_reunion, cost_event, cost_formacion = '', '', '', '', '', ''
    div = 0
    try:
        cost = driver.find_element_by_xpath('/html/body/div[2]/section[1]/div/div/article/div/section/div[4]/div/div/div[2]/div[1]/div[1]').get_attribute("innerHTML")
        div = 4
    except:
        div = 5
    for j in range(1,4):
        xpath_num, xpath_text, xpath_base = '', '', ''
        if div == 4:
            xpath_num = '/html/body/div[2]/section[1]/div/div/article/div/section/div[4]/div/div/div[2]/div['+str(j)+']/div[2]'
            xpath_text = '/html/body/div[2]/section[1]/div/div/article/div/section/div[4]/div/div/div[2]/div['+str(j)+']/div[1]'
            xpath_base = '/html/body/div[2]/section[1]/div/div/article/div/section/div[4]/div/div/div[2]/div['+str(j)+']/div[3]'
        else:
            xpath_num =   '/html/body/div[2]/section[1]/div/div/article/div/section/div[5]/div/div/div[2]/div['+str(j)+']/div[2]'
            xpath_text = '/html/body/div[2]/section[1]/div/div/article/div/section/div[5]/div/div/div[2]/div['+str(j)+']/div[1]'
            xpath_base = '/html/body/div[2]/section[1]/div/div/article/div/section/div[5]/div/div/div[2]/div['+str(j)+']/div[3]'

        try:
            cost = driver.find_element_by_xpath(xpath_num).get_attribute("innerHTML")
            nom_cost = driver.find_element_by_xpath(xpath_text).get_attribute("innerHTML")
            base = driver.find_element_by_xpath(xpath_base).get_attribute("innerHTML")
            if nom_cost == 'Mesa Flexible':
                cost_flex = str(return_num(cost)) + '€/' + base
            elif nom_cost == 'Mesa Fija':
                clost_fix = str(return_num(cost)) + '€/' + base
            elif nom_cost == 'Oficina':
                cost_ofice = str(return_num(cost)) + '€/' + base
            elif nom_cost == 'Sala de Reuniones':
                cost_reunion = str(return_num(cost)) + '€/' + base
            elif nom_cost == 'Sala de Eventos':
                cost_event = str(return_num(cost)) + '€/' + base
            elif nom_cost == 'Sala de Formación':
                cost_formacion = str(return_num(cost)) + '€/' + base
        except:
            break
    return clost_fix, cost_flex, cost_ofice, cost_reunion, cost_event, cost_formacion

URL = "https://coworkingspain.es/espacios"

options = webdriver.ChromeOptions()

options.add_argument("headless")

print('Cargando todas las oficinas de la página...')

driver = webdriver.Chrome(executable_path=r"drivers/chromedriver.exe",options=options)
driver.get(URL)

finished = False

while not finished:
    try:
        submit = driver.find_element_by_xpath('//*[@id="block-system-main"]/div/div/div[2]/ul/li/a')
        submit.click()
        sleep(0.7)
    except:
        finished = True

elements = driver.find_elements_by_class_name('views-row')

list_elements = []

print('Extrayendo links para cada oficina...')

for i in range(1,len(elements)+1):
    xpath = '//*[@id="block-system-main"]/div/div/div[1]/div['+str(i)+']/div[2]/div/div[1]/div[1]/a/h2'
    title = driver.find_element_by_xpath(xpath)
    xpath = '//*[@id="block-system-main"]/div/div/div[1]/div['+str(i)+']/div[2]/div/div[1]/div[1]/a'
    link = driver.find_element_by_xpath(xpath).get_attribute('href')
    list_elements.append([title.text, '', '', '', '', '', '', '', '', '', link])
    
print('Extrayendo la información de contacto de cada oficina y precios...')

for i in range(len(list_elements)):
    driver.get(list_elements[i][10])
    list_elements[i][1], list_elements[i][2], list_elements[i][3], list_elements[i][4], list_elements[i][5], list_elements[i][6] = get_costs(driver)
    try:
        list_elements[i][7] = driver.find_element_by_xpath('/html/body/div[2]/section[1]/div/div/article/section[2]/div[2]/div[2]/div[2]/div[1]/div/div').get_attribute("innerHTML")
        list_elements[i][8] = driver.find_element_by_xpath('/html/body/div[2]/section[1]/div/div/article/section[2]/div[1]/div/ol/li[3]/a/span').get_attribute("innerHTML")
        list_elements[i][9] = driver.find_element_by_xpath('/html/body/div[2]/section[1]/div/div/article/section[2]/div[2]/div[2]/div[2]/div[3]/div/div').get_attribute("innerHTML")
        if '_blank' in list_elements[i][9]:
            list_elements[i][9] = driver.find_element_by_xpath('/html/body/div[2]/section[1]/div/div/article/section[2]/div[2]/div[2]/div[2]/div[2]/div/div').get_attribute("innerHTML")
    except:
        pass
        
driver.close()

print('Exportando a excel...')
        
excel = DataFrame(list_elements, columns=['Título','Mesa Fija','Mesa Flexible','Oficina','Sala de Reuniones','Sala de Eventos','Sala de Formación', 'Dirección', 'Ciudad', 'Teléfono','Link'])

excel.to_excel('base_datos_coworkingspain.xlsx', index=False)

print("Se han exportado a excel la informacion y precios de "+str(len(list_elements))+" oficinas.")