import time
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random
import os
from selenium.webdriver.common.by import By
import speech_recognition as sr
import ffmpy
import requests
import urllib
import pydub
import pandas as pd
import psycopg2
from psycopg2 import Error
import json
from selenium_stealth import stealth
from fake_useragent import UserAgent
from fp.fp import FreeProxy


def get_data_excel():
    filename = 'Копия Фосфоглив ОФД скидки 3 плюс 1_ (003).xlsx'
    df = pd.read_excel(filename, usecols=['Дата чека', 'ФП', 'РН ККТ'])
    return df


# def test_add_to_db():
#     try:
#         # Подключиться к существующей базе данных
#         connection = psycopg2.connect(user="postgres",
#                                       # пароль, который указали при установке PostgreSQL
#                                       password="DtnthDgjkt2002",
#                                       host="127.0.0.1",
#                                       port="5432",
#                                       database="initpro")
#         set_data = ('ООО Социальная Аптека 8',
#                     '6166074939',
#                     '20.04.2022 15:33:00',
#                     '[{"number_product": 1, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 2, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 3, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 4, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 5, "name_product": "Пакет ТК-433 Социальная Аптека ед.", "price_product": 1.0, "quantity_product": 1.0, "sum_price_product": 1.0}, {"number_product": 6, "name_product": "Вода мин Смирновская 0,5 л ед.", "price_product": 44.0, "quantity_product": 1.0, "sum_price_product": 44.0}]',
#                     1938.0,
#                     172.08,
#                     7.5,
#                     '00106720018661',
#                     '0005801569021922',
#                     '17401',
#                     '9960440301067936',
#                     '2762389667',
#                     '414004, Астрахань г, Бориса Алексеева ул, 34а д., 03 помещение',
#                     'Аптека №170')
#         print(set_data)
#         # Создайте курсор для выполнения операций с базой данных
#         cursor = connection.cursor()
#         # insert_script = """INSERT INTO public.checks_info(name_organization, inn, date_check, products_info, price, vat_10, vat_20, zn_kkt, rn_kkt, fd, fn, fp, address, place_settlement) VALUES ('ООО Социальная Аптека 8', '6166074939', '20.04.2022 15:33:00', '[{"number_product": 1, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 2, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 3, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 4, "name_product": "Фосфоглив капс N50 ед.", "price_product": 473.25, "quantity_product": 1.0, "sum_price_product": 473.25}, {"number_product": 5, "name_product": "Пакет ТК-433 Социальная Аптека ед.", "price_product": 1.0, "quantity_product": 1.0, "sum_price_product": 1.0}, {"number_product": 6, "name_product": "Вода мин Смирновская 0,5 л ед.", "price_product": 44.0, "quantity_product": 1.0, "sum_price_product": 44.0}]', 1938.0, 172.08, 7.5, '00106720018661', '0005801569021922', '17401', '9960440301067936', '2762389667', '414004, Астрахань г, Бориса Алексеева ул, 34а д., 03 помещение', 'Аптека №170');"""
#         # insert_script_test = f"INSERT INTO public.checks_info(name_organization, inn, date_check, products_info, price, vat_10, vat_20, zn_kkt, rn_kkt, fd, fn, fp, address, place_settlement) VALUES ({});"
#         insert_script = """INSERT INTO public.checks_info(
# 	name_organization, inn, date_check, products_info, price, vat_10, vat_20, zn_kkt, rn_kkt, fd, fn, fp, address, place_settlement)
# 	VALUES (%s, %s, %s, %s, %s,%s, %s, %s, %s, %s,%s, %s, %s, %s);"""
#         cursor.execute(insert_script, set_data)
#         connection.commit()
#         cursor.execute('SELECT * from public.checks_info')
#         print(f"элемент успешно добавлен")
#     except (Exception, Error) as error:
#         print("Ошибка при работе с PostgreSQL", error)
#     finally:
#         if connection:
#             cursor.close()
#             connection.close()
#             print("Соединение с PostgreSQL закрыто")


def add_to_database(insert_data):
    try:
        # Подключиться к существующей базе данных
        connection = psycopg2.connect(user="postgres",
                                      # пароль, который указали при установке PostgreSQL
                                      password="DtnthDgjkt2002",
                                      host="127.0.0.1",
                                      port="5432",
                                      database="initpro")
        # Создайте курсор для выполнения операций с базой данных
        cursor = connection.cursor()
        k = 0
        for dict_data in insert_data:
            k += 1
            set_data = (dict_data['name_organization'],
                        dict_data['inn'],
                        dict_data['date_check'],
                        dict_data['products_info'],
                        dict_data['price'],
                        dict_data['vat_10'],
                        dict_data['vat_20'],
                        dict_data['zn_kkt'],
                        dict_data['rn_kkt'],
                        dict_data['fd'],
                        dict_data['fn'],
                        dict_data['fp'],
                        dict_data['address'],
                        dict_data['place_settlement'])
            print(set_data)

            insert_script = """INSERT INTO public.checks_info(
            	name_organization, inn, date_check, products_info, price, vat_10, vat_20, zn_kkt, rn_kkt, fd, fn, fp, address, place_settlement)
            	VALUES (%s, %s, %s, %s, %s,%s, %s, %s, %s, %s,%s, %s, %s, %s);"""
            cursor.execute(insert_script, set_data)
            connection.commit()
            print(f"{k} элемент успешно добавлен")
    except (Exception, Error) as error:
        print("Ошибка при работе с PostgreSQL", error)
    finally:
        if connection:
            cursor.close()
            connection.close()
            print("Соединение с PostgreSQL закрыто")


def delay():
    time.sleep(random.randint(5, 6))


def long_delay():
    time.sleep(random.randint(30, 45))


def super_long_delay():
    time.sleep(random.randint(1000, 1800))


def settings_and_start():
    global browser
    option = webdriver.ChromeOptions()
    # options.add_argument("start-maximized")

    # options.add_argument("--headless")
    # option.add_argument("--no-sandbox")
    option.add_experimental_option("excludeSwitches", ["enable-automation"])
    option.add_experimental_option('useAutomationExtension', False)
    ser = Service("C:\\Users\\mrtik\\PycharmProjects\\testOpencv\\chromedriver\\chromedriver.exe")
    option.add_argument('--disable-blink-features=AutomationControlled')
    # option.add_argument("/home/tellowit/home/initpro/chromedriver")
    # option.add_argument(r'--user-data-dir=C:\Users\mrtik\AppData\Local\Google\Chrome\User Data\Default')
    browser = webdriver.Chrome(options=option, service=ser)

    stealth(browser,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
            )
    # proxy = '178.176.74.1'
    url = "https://ofd-initpro.ru/check-bill/"
    # ser = Service(r"C:\Users\mrtik\PycharmProjects\Initpro\chromedriver\chromedriver.exe")
    # option = webdriver.ChromeOptions()
    # ua = UserAgent()
    # user_agent = ua.random
    # print(user_agent)
    # try:
    #     proxy = FreeProxy(country_id=['EU'], timeout=5, rand=True).get()
    #     proxy.replace('http://', '')
    #     print(proxy)
    #     option.add_argument(f"--proxy-server={proxy}")
    # except Exception as e:
    #     print("Proxy not find")
    # option.add_argument(f'user-agent={user_agent}')
    # option.add_argument('headless')
    # option.add_argument('--disable-blink-features=AutomationControlled')
    # browser = webdriver.Chrome(service=ser, options=option)
    browser.implicitly_wait(5)
    browser.get(url)
    delay()


def solving_captcha(data_input, fpd_input, kkt_input):
    global browser
    try:
        iframe = browser.find_elements(by=By.TAG_NAME, value='iframe')[0]
        browser.switch_to.frame(iframe)
        act = browser.find_element(by=By.CSS_SELECTOR, value='.recaptcha-checkbox-border')
        act.click()  # click to field
        delay()
        browser.switch_to.default_content()
        all_frames = browser.find_elements(by=By.TAG_NAME, value='iframe')
        print(all_frames)
        print(f'Len: {len(all_frames)}')
        # browser.switch_to.frame(all_frames[3])
        k = 0
        for i in range(0, len(all_frames), 1):
            try:
                browser.switch_to.default_content()
                browser.switch_to.frame(all_frames[i])
                print(f'Iteration: {i}')
                browser.find_element(by=By.ID, value='recaptcha-audio-button').click()
                k = i
            except:
                print(f'Incorrect frame click to audio: ID-{i}')
        print(f'CORRECT FRAME click to audio: ID-{k}')

        delay()

        # switch do audio frame
        browser.switch_to.default_content()
        all_frames = browser.find_elements(by=By.TAG_NAME, value='iframe')
        browser.switch_to.frame(all_frames[k])
        delay()

        try:
            act = browser.find_element(by=By.XPATH, value='//*[@id=":2"]')
            act.click()
        except Exception as e:
            print('Captcha block, because think program robot')
            print('Problems with listen audio')
            delay()
            browser.close()
            # browser.quit()
            super_long_delay()
            print('RESTART')
            settings_and_start()
            delay()
            filling_out_form(data_input, fpd_input, kkt_input)

        src = browser.find_element(by=By.ID, value='audio-source').get_attribute('src')
        print(f'Aurdio src: {src}')

        path_to_mp3 = os.path.normpath(os.path.join(os.getcwd(), "sample.mp3"))
        path_to_wav = os.path.normpath(os.path.join(os.getcwd(), "sample.wav"))
        # download the mp3 audio file from the source
        urllib.request.urlretrieve(src, path_to_mp3)
        sound = pydub.AudioSegment.from_mp3(path_to_mp3)
        sound.export(path_to_wav, format="wav")
        sample_audio = sr.AudioFile(path_to_wav)
        delay()
        r = sr.Recognizer()
        with sample_audio as source:
            audio = r.record(source)
        key = r.recognize_google(audio)
        print(f"Recaptcha Passcode: {key}")
        browser.find_element(by=By.ID, value='audio-response').send_keys(key.lower())
        browser.find_element(by=By.ID, value='audio-response').send_keys(Keys.ENTER)
        delay()
        browser.switch_to.default_content()
        browser.find_element(by=By.XPATH, value='//*[@id="js-check"]/div/div[1]/form/button').click()
        print('Click4 to button')
        delay()
    except Exception as e:
        print('Problems with solving captcha')
        delay()
        browser.close()
        # browser.quit()
        long_delay()
        print('RESTART')
        settings_and_start()
        delay()
        filling_out_form(data_input, fpd_input, kkt_input)


def filling_out_form(data_input, fpd_input, kkt_input):
    """Found field where need indicate that you not robot"""
    try:
        browser.find_element(by=By.ID, value='datepicker').send_keys(data_input)
        browser.find_element(by=By.ID, value='datepicker').send_keys(Keys.ENTER)
        delay()
        browser.find_element(by=By.ID, value='fpd').send_keys(fpd_input)  # frp
        delay()
        browser.find_element(by=By.ID, value='rnm').send_keys(kkt_input)
        delay()

        solving_captcha(data_input, fpd_input, kkt_input)
    except Exception as e:
        print('Problems with filling out form or solving captcha')


def collect_information(data_input, fpd_input, kkt_input):
    global browser
    try:
        browser.switch_to.default_content()
        iframe = browser.find_element(by=By.XPATH, value='//*[@id="iframe_elem"]')
        browser.switch_to.frame(iframe)
        delay()
        info_of_check = dict()

        name_organization = browser.find_element(by=By.XPATH,
                                                 value='/html/body/table/tbody/tr/td/div/table[1]/tbody/tr[1]/td').text
        info_of_check['name_organization'] = name_organization
        print('name_organization : ', name_organization)

        inn = browser.find_element(by=By.XPATH, value='/html/body/table/tbody/tr/td/div/table[1]/tbody/tr[2]/td').text
        inn = inn.replace('ИНН ', '')
        info_of_check['inn'] = inn
        print('inn : ', inn)

        date_check = browser.find_element(by=By.XPATH,
                                          value='/html/body/table/tbody/tr/td/div/table[2]/tbody/tr[2]/td').text
        info_of_check['date_check'] = date_check
        print('date_check : ', date_check)

        table_products = browser.find_element(by=By.CLASS_NAME, value='itemsTable')
        table_body_products = table_products.find_element(by=By.TAG_NAME, value='tbody')
        products = table_body_products.find_elements(by=By.CLASS_NAME, value='item')
        products_arr = []
        for i in range(len(products)):
            number_product = products[i].find_element(by=By.CLASS_NAME, value='num').text
            name_product = products[i].find_element(by=By.CLASS_NAME, value='name').text
            price_product = products[i].find_element(by=By.CLASS_NAME, value='price').text
            quantity_product = products[i].find_element(by=By.CLASS_NAME, value='quantity').text
            sum_price_product = products[i].find_element(by=By.CLASS_NAME, value='sum').text

            prodict_dict = dict()
            prodict_dict['number_product'] = int(number_product)
            prodict_dict['name_product'] = name_product
            prodict_dict['price_product'] = float(price_product)
            prodict_dict['quantity_product'] = float(quantity_product)
            prodict_dict['sum_price_product'] = float(sum_price_product)

            products_arr.append(prodict_dict)

        info_of_check['products_info'] = json.dumps(products_arr, ensure_ascii=False)
        print('products_info : ', info_of_check['products_info'])

        price = browser.find_element(by=By.XPATH,
                                     value='/html/body/table/tbody/tr/td/div/table[7]/tbody/tr[1]/td[2]').text
        info_of_check['price'] = float(price)
        print('price : ', price)

        vat_10 = 0
        try:
            vat_10 = browser.find_element(by=By.XPATH,
                                          value='/html/body/table/tbody/tr/td/div/table[7]/tbody/tr[5]/td[2]').text
            info_of_check['vat_10'] = float(vat_10)
        except Exception as e:
            info_of_check['vat_10'] = float(vat_10)
        print('vat_10 : ', vat_10)

        vat_20 = 0
        try:
            vat_20 = browser.find_element(by=By.XPATH,
                                          value='/html/body/table/tbody/tr/td/div/table[7]/tbody/tr[6]/td[2]').text
            info_of_check['vat_20'] = float(vat_20)
        except Exception as e:
            info_of_check['vat_20'] = float(vat_20)
        print('vat_20 : ', vat_20)

        zn_kkt = browser.find_element(by=By.XPATH,
                                      value='/html/body/table/tbody/tr/td/div/table[8]/tbody/tr[1]/td[2]').text
        info_of_check['zn_kkt'] = zn_kkt
        print('zn_kkt : ', zn_kkt)

        rn_kkt = browser.find_element(by=By.XPATH,
                                      value='/html/body/table/tbody/tr/td/div/table[8]/tbody/tr[2]/td[2]').text
        info_of_check['rn_kkt'] = rn_kkt
        print('rn_kkt : ', rn_kkt)

        fd = browser.find_element(by=By.XPATH, value='/html/body/table/tbody/tr/td/div/table[8]/tbody/tr[3]/td[2]').text
        info_of_check['fd'] = fd
        print('fd : ', fd)

        fn = browser.find_element(by=By.XPATH, value='/html/body/table/tbody/tr/td/div/table[8]/tbody/tr[4]/td[2]').text
        info_of_check['fn'] = fn
        print('fn : ', fn)

        fp = browser.find_element(by=By.XPATH, value='/html/body/table/tbody/tr/td/div/table[8]/tbody/tr[5]/td[2]').text
        info_of_check['fp'] = fp
        print('fp : ', fp)

        address = browser.find_element(by=By.XPATH,
                                       value='/html/body/table/tbody/tr/td/div/table[9]/tbody/tr[1]/td[2]').text
        info_of_check['address'] = address
        print('address : ', address)

        place_settlement = browser.find_element(by=By.XPATH,
                                                value='/html/body/table/tbody/tr/td/div/table[9]/tbody/tr[2]/td[2]').text
        info_of_check['place_settlement'] = place_settlement
        print('place_settlement : ', place_settlement)

        print(info_of_check)

        return info_of_check
    except Exception as e:
        print('Problems with collecting data')
        # print(e)
        delay()
        browser.close()
        # browser.quit()
        long_delay()
        print('RESTART')
        settings_and_start()
        delay()
        filling_out_form(data_input, fpd_input, kkt_input)


def collect_into_table(query_result):
    """Формирование таблицы с данными и генерация xlsx файла"""
    # создание Dataframe колоннами
    df_data = pd.DataFrame()
    try:
        df_data = df_data.from_records(query_result)
        # создание xlsx файла

        # engine = create_engine("postgresql+psycopg2://postgres:DtnthDgjkt2002@localhost:5432/initpro")
        df_data.to_excel(f'result.xlsx', index=False)
        # df_data.to_sql(name='checks_info', con=engine, if_exists='append', index=False)
        print('Successful generate xlsl-file')
    except Exception as e:
        print('Error in collect_into_table(): ', e)


def main(date_arr, fpd_arr, kpp_arr):
    total_info = []
    success_counter = 0
    for i in range(len(date_arr)):
        settings_and_start()
        filling_out_form(date_arr[i], fpd_arr[i], kpp_arr[i])
        total_info.append(collect_information(date_arr[i], fpd_arr[i], kpp_arr[i]))
        success_counter += 1
        delay()
        browser.close()
        browser.quit()
        print(f'SUCCESS CHECK: {success_counter}')
        long_delay()
    collect_into_table(total_info)
    add_to_database(total_info)


def super_main(date_arr, fpd_arr, kpp_arr):
    total_info_super = []
    success_counter = 0
    for i in range(0, len(date_arr), 4):
        settings_and_start()
        list_elements_date = str(date_arr[i])[:10].split('-')
        date_i = str(list_elements_date[2] + '.' + list_elements_date[1] + '.' + list_elements_date[0])
        fpd_i = int(fpd_arr[i])
        kpp_i = '000' + str(int(kpp_arr[i]))
        filling_out_form(date_i, fpd_i, kpp_i)
        total_info_super.append(collect_information(date_i, fpd_i, kpp_i))
        success_counter += 1
        delay()
        browser.close()
        browser.quit()
        print(f'SUCCESS CHECK: {success_counter}')
        long_delay()
    collect_into_table(total_info_super)
    add_to_database(total_info_super)


def start_2():
    d_arr = ['20.04.2022', '08.12.2021']
    f_arr = ['2762389667', '3386209579']
    k_arr = ['0005801569021922', '0005801569021922']

    main(d_arr, f_arr, k_arr)


def start_all():
    data_exl = get_data_excel()
    d_all_arr = data_exl._get_column_array(0)
    f_all_arr = data_exl._get_column_array(2)
    k_all_arr = data_exl._get_column_array(1)

    super_main(d_all_arr, f_all_arr, k_all_arr)


if __name__ == '__main__':
    try:
        start_all()
    except Exception as e:
        print('Undefined problems, restart all program')
        start_all()
