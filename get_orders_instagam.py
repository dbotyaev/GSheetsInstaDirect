"""
Скрипт №3
ссылка на задание: https://disk.yandex.ru/i/WJZIe2Zbg3rsfw

Краткое описание: необходимо в инстаграме на странице Сохраненное получить список сохраненных постов,
у каждого поста получить его текст, а также комментарии.
Текст поста необходимо распарсить и получить название аромата, позиции для прайса, которые загрузить
в Google-таблицу на лист Прайс, существующие записи обновить, новые добавить.
Комментарии необходимо сгруппировать по ароматам и добавить в таблицу на лист Заказы.
"""
import csv
import os
import pandas as pd
import pickle
import pygsheets
import random
import re
import requests
import sys
import time

from datetime import datetime
from instabot import Bot
from loguru import logger
from selenium.webdriver.common.keys import Keys

from instagrambot import InstagramBot
from gsheetsbot import GSheetsBot
from settings.settings import INSTA_SERVICE_LOGIN, INSTA_SERVICE_PASSWORD, URL_SAVE_POSTS
from settings.settings import INSTALOGIN, INSTAPASSWORD
from settings.settings import BROWSER_MODE, AUTH_MODE, HEADLESS
from settings.headers import HEADERS
from settings.settings import SERVICE_ACCOUNT_FILE, URL_GSHEET, PRICE_SHEET


def _get_headers():
    logger.debug(f'Добавляем в HEADERS данные из сохраненных cookies')
    cookie = ''
    headers = HEADERS
    with open(os.getcwd() + '\\cookies\\cookies.pickle', 'rb') as file:
        cookies = pickle.load(file)
        for elem in cookies:
            cookie += f"{elem['name']}={elem['value']}; "
    headers['cookie'] = cookie[:-2]  # обрезаем пробел и точку с запятой в конце

    return headers


def _parsing_post_json(data_json):
    """
    Паринг json данных о постах
    :param data_json: словарь постов с одной страницы в формате json
    :return: список словарей данных поста
    """
    result_list = []
    for data in data_json:
        result_dict = {
            'media_id': data['media']['caption']['media_id'],
            'text': data['media']['caption']['text'],
            'code_url': data['media']['code'],
        }
        result_list.append(result_dict)
    logger.info(f'Получено и обработано {len(result_list)} постов')
    return result_list


def _parsing_text_post(text):
    """
    Получение из текста поста позиции прайса, например, 'NORAN PERFUMES - KADOR PRIVAT 2,5	300'
    :param text: текст поста в instagram
    :return: result_price = [[]]  список списков
    """
    result_price = []
    ch = '*'  # ключевой символ поиска
    index_ch = [i for i in range(0, len(text)) if text[i] == ch]  # получаем все позиции символа ch в тексте
    name_parfum = text[index_ch[0] + 1:index_ch[1]].strip()  # получаем название парфюма заключенного в "*"

    # перебираем все остальные найденные позиции символа ch для поиска объема и цены
    for i in index_ch[2:]:
        ind_n = text.find('\n', i)
        str_price = text[i:ind_n]  # ожидаем получение вот такой строки, например, '*2,5 мл - 300₽ + 50₽ атомайзер'
        # получаем объем, например, '2,5'
        volume = re.search('\*\d+[.,]?\d?', str_price).group(0)[1:].strip().replace('.', ',')
        price = re.search('\d+₽', str_price).group(0)[:-1]  # получаем цену, например, '300'
        result_price.append([f'{name_parfum} {volume}', price])
    # name_parfum возвращаем, чтобы добавить данные в result_parsing_posts для парсинга комментариев
    return result_price, name_parfum


def _get_orders_from_comments():
    result_orders = []
    try:
        for post in instagram.result_parsing_posts:
            logger.debug(f'Пауза между запросами')
            time.sleep(random.randrange(15, 30))  # задержка между получением комментариев поста

            media_id = post['media_id']  # идентификатор поста
            aromat_parfum = post['name_parfum'].upper()
            url_post = f'https://www.instagram.com/p/{post["code_url"]}'
            logger.info(f'Получаем комментарии "{aromat_parfum}" {url_post}')
            comments_post = comments_bot.get_media_comments_all(media_id)  # получение всех комментарием поста

            # распарсивание поста
            count_comments = 0  # счетчик комментариев
            for comment in comments_post:
                date_post = datetime.fromtimestamp(comment['created_at']).strftime('%d.%m.%Y %H:%M:%S')
                account = comment['user']['username'].upper()
                # пропускаем комментарии основного аккаунта
                if account == INSTALOGIN:
                    continue
                text_comment = comment['text']  # комментарий
                try:
                    # получение из комментария объема в числовом выражении
                    text_volume = re.search('\d+[.,]?\d?', text_comment).group(0).replace('.', ',')
                except Exception:
                    text_volume = ''
                result_orders.append([account, aromat_parfum, text_volume, text_comment, date_post, url_post])
                count_comments += 1
            logger.info(f'Получено {count_comments} комментариев')
        logger.success(f'Общее кол-во комментариев всех постов {len(result_orders)}')
        return result_orders

    except Exception as ex:
        logger.error(f'Возникла ошибка при получении комментариев поста {ex}')
        if result_orders:
            orders_error_file_csv = f'{os.getcwd()}\\price_orders\\orders_error.csv'
            logger.info(f'Записываем часть комментариев во временный файл {orders_error_file_csv}')
            with open(orders_error_file_csv, 'w', encoding='utf-8') as file_csv:
                file_writer = csv.writer(file_csv, delimiter=";", lineterminator="\r")
                file_writer.writerows(result_orders)
            logger.info(f'Часть комментариев успешно сохранена во временный файл')
        logger.error(f'Попробуйте позже или обратитесь к разработчику')
        raise


def _set_new_price():
    """
    Обновляем существующие позиции в прайс-листе и добавляем новые
    :return: обновленный gsheets_price.df_values
    """
    # удаляем строки, где НАИМЕНОВАНИЕ = NaN и ЦЕНА = NaN
    gsheets_price.df_values.dropna(subset=['НАИМЕНОВАНИЕ', 'ЦЕНА'], how='all', inplace=True)

    price_for_append = []  # список новых позиций для прайс-листа

    for price in instagram.price_list:
        # проверяем наименование товара в прайс-листе из Google-таблицы
        if price[0].upper() in gsheets_price.df_values['НАИМЕНОВАНИЕ'].values:
            # обновляем цены существующих позиций в прайс-листе
            gsheets_price.df_values.loc[gsheets_price.df_values['НАИМЕНОВАНИЕ'] == price[0].upper(), 'ЦЕНА'] = \
                float(price[1])
            gsheets_price.df_values.loc[gsheets_price.df_values['НАИМЕНОВАНИЕ'] == price[0].upper(), 'ДАТА'] = \
                datetime.now().strftime('%d.%m.%Y')
        else:
            # добавляем в список новые позиции для последующего добавления в прайс-лист
            price_for_append.append({
                'НАИМЕНОВАНИЕ': price[0].upper(),
                'ЦЕНА': float(price[1]),
                'ДАТА': datetime.now().strftime('%d.%m.%Y')
            })

    # добавляем новые позиции в прайс-лист
    gsheets_price.df_values = gsheets_price.df_values.append(price_for_append, ignore_index=True)


def _set_price_gsheets():
    """
    Формирование структурированного прайс-листа для последующей выгрузки в Google-таблтцу.
     После каждого вида аромата вставляется пустая строка
    :return: список списков result_price
    """

    def get_name_aromat(s):
        # получаем наименование аромата без объема - ключеввой символ - последний пробел в названии
        index = s.rfind(' ')
        return s[:index]

    result_price = []

    price_list = gsheets_price.df_values.values.tolist()
    aromat_prev = get_name_aromat(price_list[0][0])
    for price in price_list:
        if get_name_aromat(price[0]) != aromat_prev:
            result_price.append(['', ''])
        result_price.append(price)
        aromat_prev = get_name_aromat(price[0])

    return result_price


class InstagramParsing(InstagramBot):
    """
    Инициализация и функция авторизации необходимы только для того, чтобы получить свежие cookies
    Функионал парсинга может использовать cookies сессий рассылки сообщений. В этом случае придется
    не наследовать __init__ и закомментировать функцию авторизации
    """

    def __init__(self, username, password):
        super().__init__(username, password)
        self.price_list = []
        self.result_parsing_posts = []  # список словарей резльтата парсинга постов
        self.headers = ''

    # Отключаем всплывающее окно "Включить уведомления"
    def _enable_notifications(self, button_not_now):
        logger.info(f'Проверка появления всплывающего окна Включить уведомления и попытка его отключить')
        # в режиме HEADLESS окно "Включить уведомления" не появлется
        if not HEADLESS:
            # Отключаем всплывающее окно "Включить уведомления"
            if self.xpath_exists(button_not_now):
                self.browser.find_element_by_xpath(button_not_now).click()
                time.sleep(random.randrange(2, 4))

    def authorization(self):
        self.browser.get('https://www.instagram.com')
        self.browser.set_window_size(768, 704)
        # self.browser.maximize_window()

        time.sleep(random.randrange(3, 5))

        # режим использования логина и пароля на странице авторизации
        if AUTH_MODE:
            username_input = self.browser.find_element_by_name('username')
            username_input.clear()
            username_input.send_keys(self.username)

            time.sleep(random.randrange(1, 3))

            password_input = self.browser.find_element_by_name('password')
            password_input.clear()
            password_input.send_keys(self.password)

            time.sleep(random.randrange(1, 3))

            password_input.send_keys(Keys.ENTER)

            time.sleep(60)  # время задержки для ввода кода подтверждения при двухфакторной авторизации

            # Попытка Отключить всплывающее окно "Включить уведомления", если оно появится
            self._enable_notifications('/html/body/div[4]/div/div/div/div[3]/button[2]')

            time.sleep(random.randrange(1, 2))

            # сохранение cookies в файл после успешной авторизации
            if BROWSER_MODE == 'COOKIES':
                self.save_cookies()
                logger.info(f'Сохранение cookies после успешной авторизации во временный файл')

        # читаем cookies предыдущей сессии из файла cookies.pickle и авторизуемся без логина и пароля
        elif BROWSER_MODE == 'COOKIES':  # AUTH_MODE == 0
            logger.info(f'Чтение cookies для авторизации из файла')
            self.load_cookies()
            logger.info(f'Cookies из файла добавлены в браузер')
            # обновляем страницу браузера и появляется всплывающее окнов Включить уведомления
            time.sleep(1)
            self.browser.refresh()

            time.sleep(random.randrange(3, 5))

            # Попытка отключить всплывающее окно "Включить уведомления", если оно появится
            self._enable_notifications('/html/body/div[4]/div/div/div/div[3]/button[2]')

            time.sleep(1)

            # сохранение cookies в файл после успешной авторизации
            self.save_cookies()
            logger.info(f'Сохранение cookies после успешной авторизации во временный файл')

    def get_post_instagram(self):
        """
        Функция делает get-запросы и получает словарь с данными постов (id, текст поста, код ссылки для поста)
        :return: формирует список словарей с данными постов self.result_parsing_posts
        """
        id_collection = re.search('\/\d+\/', URL_SAVE_POSTS).group(0)[1:-1]  # получаем идентификатор папки из ссылки
        self.headers = _get_headers()  # добавляем в headers параметр cookie
        # Параметры запроса и ответа
        url_api = f'https://i.instagram.com/api/v1/feed/collection/{id_collection}/posts/?max_id='
        more_available = True  # параметр последней страницы списка постов
        next_max_id = ''  # параметр запроса, в котором передается идентификатор страницы запроса
        try:
            while more_available:
                response = requests.get(url_api + next_max_id, headers=self.headers)
                if response.status_code == 200:
                    post_json = response.json()
                    if post_json['status'] == 'ok':
                        # распарсиваем информацию о постах со страницы и добавляем в список словарей
                        self.result_parsing_posts += (_parsing_post_json(post_json['items']))
                        more_available = post_json['more_available']
                        if more_available:
                            next_max_id = post_json['next_max_id']

                        logger.debug(f'Пауза между запросами')
                        time.sleep(random.randrange(7, 12))

                    else:
                        logger.error(f'Ошибка получения списка постов. Статус ответа {post_json["status"]}')
                        logger.error(f'Попробуйте позже или обратитесь к разработчику')
                        raise
                else:
                    logger.error(f'Сервер Instagram вернул неверный ответ {response.status_code}')
                    logger.error(f'Попробуйте позже или обратитесь к разработчику')
                    raise
        except Exception as ex:
            logger.error(f'Произошла ошибка при парсинге информации постов {ex}')
            logger.error(f'Возможно изменилась структура данных. Обратитесь к разработчику')
            raise

    def get_price_from_posts(self):
        """
        Формирует прайс-лист из текста поста
        :return: заполняет список self.price_list
        """
        for post in self.result_parsing_posts:
            text_post = post['text']
            price, parfum = _parsing_text_post(text_post)
            self.price_list += price
            post['name_parfum'] = parfum


class GSheetsOrdersPrice(GSheetsBot):
    """
    Унаследовали класс GSheetsBot, который используется в скрипте для отправки сообщений
    """

    def __init__(self, url, start_cell=None, end_cell=None, name_worksheet=None):
        """
        :param url: ссылка на Google-таблицу
        :param name_worksheet: имя листа с диапазоном для загрузки
        :param start_cell: адрес первой ячейки диапазона для загрузки
        :param end_cell: адрес последней ячейки диапазона для загрузки
        """

        if start_cell is None:
            self.start_cell = 'A1'

        try:
            logger.info(f'Подключаемся к Google-таблице и получаем данные')
            path_api = os.getcwd() + SERVICE_ACCOUNT_FILE
            self.client = pygsheets.authorize(service_account_file=path_api)
            self.google_sheet = self.client.open_by_url(url)
            self.worksheet = self.google_sheet.worksheet_by_title(name_worksheet)
            self.protect_wsheet_id = 0

            # дата-фрейм ЗНАЧЕНИЙ таблицы заказов диапазона ячеек start_cell:end_cell
            self.df_values = self.worksheet.get_as_df(has_header=True,
                                                      index_column=None,
                                                      start=start_cell,
                                                      end=end_cell,
                                                      numerize=True,
                                                      empty_value='',
                                                      value_render='FORMATTED_VALUE')
            # дата-фрейм ФОРМУЛ таблицы заказов диапазона ячеек start_cell:end_cell
            # numerize=False чтобы сохранить исходное форматирование при чтении
            # self.df_formulas = self.worksheet.get_as_df(has_header=True,
            #                                             index_column=None,
            #                                             start=start_cell,
            #                                             end=end_cell,
            #                                             numerize=False,
            #                                             empty_value='',
            #                                             value_render='FORMULA')

            # устанавливаем защиту на лист
            try:
                self.protect_wsheet_id = self.set_protect_worksheet(self.worksheet.id)
                logger.info(f'Установлена защита на лист {self.worksheet.title} на период работы скрипта')
            except Exception:
                logger.warning(f'Невозможно установить защиту на лист {self.worksheet.title} на период работы скрипта '
                               f'Проверьте установлена ли защита на листе')

        except Exception as ex:
            logger.error(f'Возникла ошибка при получении данных их Google таблицы')
            logger.debug(f'{ex}')
            raise

    def set_protect_worksheet(self, worksheet_id):
        """
        Устанавливает защиту на лист
        :param worksheet_id: идентификатор листа, на который необходимо установить защиту
        :return: возвращает идентификатор листа на который была установлена защита, чтобы в конце снять ее
        """
        request = {
            "addProtectedRange": {
                "protectedRange": {
                    "range": {
                        "sheetId": worksheet_id
                    },
                    "description": 'Защита при работе скрипта',
                    "warningOnly": 'true'
                }
            }
        }
        res = self.google_sheet.custom_request(request, fields='replies')
        protect_id = res['replies'][0]['addProtectedRange']['protectedRange']['protectedRangeId']
        return protect_id

    def save_df_gsheet(self):
        """
        итоговое сохрание данных в Google-таблицу
        :return: 
        """
        try:
            # локализация значений для Google-таблицы, в столбце 'ЦЕНА' меняем плавающую точку на запятую
            if self.google_sheet.to_json()['properties']['locale'] == 'ru_RU':
                self.df_values['ЦЕНА'] = self.df_values['ЦЕНА'].astype(str).str.replace('.', ','). \
                    fillna(self.df_values['ЦЕНА']).astype(str)

            self.worksheet.copy_to(self.google_sheet.id)
            logger.info(f'Перед записью измененений создана копия листа {self.worksheet.title}')

            if self.protect_wsheet_id:
                self.worksheet.remove_protected_range(self.protect_wsheet_id)
                logger.success(f'Снята ранее установленная защита на лист для записи изменений в Google-таблицу')
            else:
                logger.warning(f'Невозможно снять защиту на листе, установленную до запуска скрипта')

            # очищаем данные на листе
            self.worksheet.clear(start=self.start_cell, fields='userEnteredValue')
            # запись данных поверх существующих значений
            self.worksheet.set_dataframe(self.df_values,
                                         start=self.start_cell,
                                         copy_index=False, copy_head=True, extend=False, fit=False,
                                         escape_formulae=False,
                                         nan='')
            logger.success(f'Зафиксированные изменения успешно записаны в Google-таблицу')

        except Exception as ex:
            logger.error(f'Произошла ошибка. Новые данные в Google-таблицу не записаны {ex}')
            path_xls = os.getcwd() + '\\copy_dataframe\\' + datetime.now().strftime("%Y%m%d %H%M%S") + '.xlsx'
            self.df_values.to_excel(excel_writer=path_xls,
                                    sheet_name=datetime.now().strftime("%Y%m%d %H%M%S"),
                                    engine='xlsxwriter')
            logger.info(f'Копия данных сохранена в Excel-файл {path_xls}')
            raise

    def save_new_worksheet(self, df_data_values, title_worksheet):
        try:
            # создаем новый лист
            new_worksheet = self.google_sheet.add_worksheet(title_worksheet)
            new_worksheet.set_dataframe(df_data_values,
                                        start=self.start_cell,
                                        copy_index=False, copy_head=True, extend=False, fit=False,
                                        escape_formulae=False,
                                        nan='')
            # устанавливаем ширину столбцов
            new_worksheet.adjust_column_width(start=1, end=1, pixel_size=190)
            new_worksheet.adjust_column_width(start=2, end=2, pixel_size=240)
            new_worksheet.adjust_column_width(start=3, end=3, pixel_size=60)
            new_worksheet.adjust_column_width(start=4, end=4, pixel_size=270)
            new_worksheet.adjust_column_width(start=5, end=5, pixel_size=140)
            new_worksheet.adjust_column_width(start=6, end=6, pixel_size=270)
            logger.success(f'Зафиксированные изменения успешно записаны в Google-таблицу')
        except Exception as ex:
            # TODO можно выделить в отдельную функцию
            logger.error(f'Произошла ошибка. Новые данные в Google-таблицу не записаны {ex}')
            path_xls = os.getcwd() + '\\copy_dataframe\\' + datetime.now().strftime("%Y%m%d %H%M%S") + '.xlsx'
            df_data_values.to_excel(excel_writer=path_xls,
                                    sheet_name=datetime.now().strftime("%Y%m%d %H%M%S"),
                                    engine='xlsxwriter')
            logger.info(f'Копия данных сохранена в Excel-файл {path_xls}')
            raise


if __name__ == '__main__':
    path_log = os.getcwd() + f'\\logs\\debug_orders.log'
    logger.add(path_log, level='DEBUG', compression="zip", rotation="9:00", retention="10 days")
    logger.info(f'Запуск скрипта')

    try:
        logger.info(f'Открываем браузер и авторизуемся')
        instagram = InstagramParsing(INSTALOGIN, INSTAPASSWORD)
        instagram.authorization()
    except Exception as ex:
        logger.critical(f'Возникла ошибка на этапе авторизации. Завершаем работу')
        try:
            instagram.close_browser()
        except Exception:
            pass
        sys.exit(1)

    try:
        logger.info(f'Переходим в раздел Сохраненное на страницу {URL_SAVE_POSTS}')
        instagram.browser.get(URL_SAVE_POSTS)
        logger.info(f'Начинаем получение основной информации о постах в разделе Сохраненное')
        instagram.get_post_instagram()
        logger.success(f'Информация успешно получена. Общее кол-во постов {len(instagram.result_parsing_posts)}')

    except Exception as ex:
        logger.info(f'Завершаем работу из-за ошибки {ex}')
        instagram.close_browser()
        sys.exit(1)

    logger.info(f'На основании полученных постов формируем прайс')
    instagram.get_price_from_posts()
    logger.success(f'Прайс сформирован. Общее количество позиций {len(instagram.price_list)}')
    instagram.close_browser()  # закрываем браузер

    price_file_csv = f'{os.getcwd()}\\price_orders\\price_from_instagram.csv'
    logger.info(f'Записываем полученный прайс из инстаграм во временный файл {price_file_csv}')
    with open(price_file_csv, 'w', encoding='utf-8') as file_csv:
        file_writer = csv.writer(file_csv, delimiter=";", lineterminator="\r")
        file_writer.writerows(instagram.price_list)
    logger.success(f'Прайс из инстаграм успешно сохранен во временный файл')

    logger.info(f'Обновляем прайс-лист')
    gsheets_price = GSheetsOrdersPrice(URL_GSHEET, name_worksheet=PRICE_SHEET)
    _set_new_price()
    price_file_csv = f'{os.getcwd()}\\price_orders\\new_price.csv'
    logger.info(f'Записываем обновленный прайс во временный файл {price_file_csv}')
    gsheets_price.df_values.to_csv(price_file_csv, sep=';', encoding='utf-8', index=False)
    logger.success(f'Полный обновленный прайс успешно сохранен во временный файл')
    logger.info(f'Форматируем прайс для загрузки в Google-таблицу')
    price_list_gsheets = _set_price_gsheets()
    gsheets_price.df_values = pd.DataFrame(price_list_gsheets, columns=['НАИМЕНОВАНИЕ', 'ЦЕНА', 'ДАТА'])

    logger.info(f'Записываем обновленный прайс в Google-таблицу')
    try:
        gsheets_price.save_df_gsheet()
    except Exception as ex:
        logger.warning(f'Обновления прайс-листа не записаны в Google-таблицу')
        logger.warning(f'Завершаем работу скрипта. Запустите скрипт повторно или обратитесть с разработчику')
        sys.exit(1)

    logger.info(f'Подключаем служебный аккаунт и начинаем получение комментариев')
    new_orders = []  # список заказов из комменатариев
    try:
        comments_bot = Bot()  # бот для  получения комментариев
        comments_bot.login(username=INSTA_SERVICE_LOGIN, password=INSTA_SERVICE_PASSWORD)
        new_orders = _get_orders_from_comments()  # получаем список комментариев
        # проверяем наличие комментариев
        if new_orders:
            new_orders_file_csv = f'{os.getcwd()}\\price_orders\\new_orders.csv'
            logger.info(f'Записываем все комментарии во временный файл {new_orders_file_csv}')
            with open(new_orders_file_csv, 'w', encoding='utf-8') as file_csv:
                file_writer = csv.writer(file_csv, delimiter=";", lineterminator="\r")
                file_writer.writerows(new_orders)
            logger.success(f'Все комментарии успешно сохранены во временный файл')
        else:
            logger.warning(f'Новые комментарии не найдены. Завершаем работу скрипта')
            sys.exit(1)
    except Exception as ex:
        logger.error(f'Завершаем работу скрипта из-за ошибки {ex}')
        sys.exit(1)

    logger.info(f'Записываем комментарии и новые заказы в Google-таблицу')
    # создаем датафрейм новых заказов из списка списков комментариев
    df_new_orders = pd.DataFrame(new_orders, columns=['АККАУНТ', 'АРОМАТ', 'ОБЪЕМ',
                                                      'КОММЕНТАРИЙ', 'ДАТА', 'ССЫЛКА'])
    # сортируем датафрейм на месте по значениям в указанных столбцах
    df_new_orders.sort_values(by=['АРОМАТ', 'АККАУНТ'], inplace=True)
    # генерируем имя листа
    title = f'comments{datetime.now().strftime("%Y%m%d%H%M")}'
    try:
        gsheets_price.save_new_worksheet(df_data_values=df_new_orders, title_worksheet=title)
        logger.success(f'Комментарии и новые заказы успешно сохранены на лист {title}')
    except Exception as ex:
        logger.error(f'Завершаем работу скрипта из-за ошибки {ex}')
        sys.exit(1)

    logger.success(f'Завершаем работу скрипта УСПЕШНО')
