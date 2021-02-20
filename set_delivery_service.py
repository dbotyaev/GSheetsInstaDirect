import os
import re
import sys
import pandas as pd
import pygsheets

from loguru import logger
from datetime import datetime
from gsheetsbot import GSheetsBot
from fuzzywuzzy import process
from settings.settings import FILE_YANDEX_FORM, LEVEL_MATCH
from settings.settings import SERVICE_ACCOUNT_FILE, URL_GSHEET, START_ADDR, END_ADDR, NAME_WORKSHEET
from settings.settings import STATUS_NEW, STATUS_SEND, STATUS_PAYMENT, STATUS_DELIVERY


class GSheetsDelivery(GSheetsBot):
    """
    Унаследовали класс GSheetsBot, который используется в скрипте для отправки сообщений
    """
    def __init__(self, url, start_cell, end_cell, name_worksheet):
        """
        :param url: ссылка на Google-таблицу
        :param start_cell: адрес первой ячейки диапазона для загрузки
        :param end_cell: адрес последней ячейки диапазона для загрузки
        :param name_worksheet: имя листа с диапазоном для загрузки
        """
        try:
            logger.info(f'Подключаем к Google-таблице и получаем данные')
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
            self.df_formulas = self.worksheet.get_as_df(has_header=True,
                                                        index_column=None,
                                                        start=start_cell,
                                                        end=end_cell,
                                                        numerize=False,
                                                        empty_value='',
                                                        value_render='FORMULA')

            # получаем список уникальных аккаунтов со статусом НОВЫЙ
            self.usernames = pd.unique(self.df_values[
                                           (self.df_values['СТАТУС'] == STATUS_NEW) |
                                           (self.df_values['СТАТУС'] == STATUS_SEND) |
                                           (self.df_values['СТАТУС'] == STATUS_DELIVERY) |
                                           (self.df_values['СТАТУС'] == STATUS_PAYMENT)]['АККАУНТ']).tolist()

            if not self.usernames:
                logger.info(f'Заказы для работы скрипта со статусами '
                            f'{STATUS_NEW}, {STATUS_SEND}, {STATUS_DELIVERY}, {STATUS_PAYMENT} отсутствуют')
            else:
                logger.success(f'Сформирован список аккаунтов для сравнения. Общее кол-во {len(self.usernames)-1}')
                try:
                    self.protect_wsheet_id = self.set_protect_worksheet(self.worksheet.id)
                    logger.info(f'Установлена защита на лист на период работы скрипта')
                except Exception:
                    logger.warning(f'Невозможно установить защиту на лист на период работы скрипта '
                                   f'Проверьте установлена ли защита на листе')
                    self.protect_wsheet_id = 0
        except Exception as ex:
            logger.error(f'Возникла ошибка при получении данных их Google таблицы')
            logger.debug(f'{ex}')
            logger.error(f'Завершаем работу скрипта. Попробуйте позже или обратитесь к разработчику')
            raise

    # установка защиты на лист от редактирование на период работы скрипта
    def set_protect_worksheet(self, worksheet_id):
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

    def change_status_delivery(self, user, delivery):
        """
        Изменение статуса позиций в заказе по фильтру АККАУНТа и СТАТУСов отбора
        Загрузка данных и все расчеты делаются из датафрейма df_values,
        а изменения и обратная загрузка из df_formulas, чтобы сохранить формулы и исходное форматирование
        :param user: АККАУНТ
        :param delivery: служба доставки
        :return: изменяет датафрейм с заказами из Google-таблицы
        """
        self.df_formulas.loc[((self.df_formulas['СТАТУС'] == STATUS_NEW) |
                              (self.df_formulas['СТАТУС'] == STATUS_SEND) |
                              (self.df_formulas['СТАТУС'] == STATUS_DELIVERY) |
                              (self.df_formulas['СТАТУС'] == STATUS_PAYMENT)) &
                              (self.df_formulas['АККАУНТ'] == user), 'НУЖНА \nОТПРАВКА'] = 'ДА'
        self.df_formulas.loc[((self.df_formulas['СТАТУС'] == STATUS_NEW) |
                              (self.df_formulas['СТАТУС'] == STATUS_SEND) |
                              (self.df_formulas['СТАТУС'] == STATUS_DELIVERY) |
                              (self.df_formulas['СТАТУС'] == STATUS_PAYMENT)) &
                              (self.df_formulas['АККАУНТ'] == user), 'ВИД \nОТПРАВКИ'] = delivery

        # здесь внесение изменений необходимо для выгрузки в Excel (если была ошибка выгрузки в Google-таблицу)
        self.df_values.loc[((self.df_values['СТАТУС'] == STATUS_NEW) |
                            (self.df_values['СТАТУС'] == STATUS_SEND) |
                            (self.df_values['СТАТУС'] == STATUS_DELIVERY) |
                            (self.df_values['СТАТУС'] == STATUS_PAYMENT)) &
                            (self.df_values['АККАУНТ'] == user), 'НУЖНА \nОТПРАВКА'] = 'ДА'
        self.df_values.loc[((self.df_values['СТАТУС'] == STATUS_NEW) |
                            (self.df_values['СТАТУС'] == STATUS_SEND) |
                            (self.df_values['СТАТУС'] == STATUS_DELIVERY) |
                            (self.df_values['СТАТУС'] == STATUS_PAYMENT)) &
                            (self.df_values['АККАУНТ'] == user), 'ВИД \nОТПРАВКИ'] = delivery

        # сохраняем копии объектов dataframe self.df_formulas и self.df_values в объекты pickle в папку copy_dataframe
        # через встроенный функционал pandas
        path_pkl = os.getcwd() + '\\copy_dataframe\\'
        self.df_values.to_pickle(path_pkl+'df_values_delivery.pkl')
        self.df_formulas.to_pickle(path_pkl + 'df_formulas_delivery.pkl')
        # logger.success(f'Успешные изменения сохранены в файлы *.pkl')


def get_excel_yandex_forms():
    try:
        logger.info(f'Считываем данные из Excel-файла Яндекс-форм')
        df_delivery = pd.read_excel(f'{path_delivery}{FILE_YANDEX_FORM}', header=0, engine='openpyxl')
        df_delivery = df_delivery.sort_values(['Время создания'])  # сортируем по времения создания
        # группируем по нику и выбираем последний элемент в группе, это будет последний пришедший ответ от аккаунта
        df_last_answer = df_delivery[['ID', 'Ваш ник в Instagram', 'Время создания', 'Доставка:']]. \
            groupby(['Ваш ник в Instagram']).last().reset_index()
        # print(df_last_answer[['ID', 'Ваш ник в Instagram', 'Время создания', 'Доставка:']])
        # print(df_last_answer.columns)
        # uniq_nickname = pd.unique(df_username_yandex_forms['Ваш ник в Instagram'].to_list())
        # добавляем новые столбы для формирования отчета
        df_delivery['Нечеткий поиск'] = ''
        df_delivery['Уровень'] = ''
        logger.success(f'Аккаунты из Excel-файла Яндекс-форм успешно прочитаны. Общее кол-во {df_last_answer.shape[0]}')
    except Exception as ex:
        logger.error(f'При считывании данных из Excel-файла возникла ошибка {ex}')
        logger.critical(f'Из-за ошибки завершаем скрипт')
        raise
    return df_delivery, df_last_answer


def set_delivery_service():
    """
    Функция делает необходимые преобразования названий аккаунтов и
    далее делает перебор аккаунтов, полученных их Excel-файл Яндекс формы и сравнивает их
    с данными датафрейма Google-таблицы. При успешном совпадении вносит изменения в датафрейм данных
    для последующей записи в Google-таблицу
    :return: измененяет данные в датафрейм self.df_values и df_formulas для выгрузки в Google-таблицу,
             а также df_excel_yandex_forms для формирования отчета
    """

    # создаем копию списка аккаунтов из Google_таблицы для поиска индекса
    gsheet_accounts = [user.lower() for user in gsheet.usernames]
    # итерируемся по датафрейму по строками выбираем столбец с ником инстаграмма
    for i, row in df_unique_nicknames.iterrows():
        # print(f'Index: {i}')
        # print(f'{row["Ваш ник в Instagram"]}')
        nickname_orig = row["Ваш ник в Instagram"]  # ник из Excel без преобразования
        nickname = re.sub('[\$@?& \\\/]', '', row["Ваш ник в Instagram"]).lower()  # поиск и удаление символов

        # получаем службу доставка из датафрейма файла Excel
        delivery_service = df_unique_nicknames.loc[
            df_unique_nicknames["Ваш ник в Instagram"] == nickname_orig, 'Доставка:'].values[0]
        if delivery_service == 'Почта России':
            delivery_service = 'ПОЧТА'
        try:
            # ищем nickname из Excel в списке аккаунтов Google-таблицы и получаем индекс элемента,
            # чтобы получить аккаунт без преобразований и делать дальше поиск по нему в датафрейме из Google-таблицы
            index = gsheet_accounts.index(nickname)  # если ошибка, переход на обработку исключения
            username = gsheet.usernames[index]
            # записываем изменения в датафрейм
            logger.success(f'Найдено точное совпадение аккаунтов {nickname_orig} - {username}')
            logger.info(f'Записываем изменения для Google-таблицы на аккаунт {username}')
            gsheet.change_status_delivery(user=username, delivery=delivery_service)
        # обрабатываем ошибку поиска, когда nickname из Excel-таблицы не найден в Google-таблице
        except ValueError:
            logger.warning(f'Значение в столбце "Ваш Ник Instagram" {nickname_orig} не найден в Google-таблице')
            logger.info(f'Делаем нечеткий поиск похожих аккаунтов')
            # result_fuzz_list = process.extract(nickname, gsheet_accounts, limit=3)
            result_fuzz_one = process.extractOne(nickname, gsheet_accounts)

            # НАЧАЛО БЛОКА - ИСПОЛЬЗОВАНИЕ МОЖЕТ ПРИВЕСТИ К НЕПРАВИЛЬНЫМ РЕЗУЛЬТАТАМ
            # если уровень совпадения соответствует параметру LEVEL_MATCH, записываем изменения в датафрейм
            # if result_fuzz_one[1] >= LEVEL_MATCH:
            #     # получаем индекс в списке аккаунтов Google-таблицы приведенном к lower()
            #     # и по индеку находим неизменный объект для передачи в функцию записи данных
            #     index = gsheet_accounts.index(result_fuzz_one[0])
            #     username = gsheet.usernames[index]
            #     logger.warning(f'Уровень похожести составил {result_fuzz_one[1]} с аккаунтом {username}')
            #     logger.warning(f'Записываем изменения для Google-таблицы на аккаунт {username}')
            #     # записываем изменения в датафрейм для выгрузки в Google-таблицу
            #     gsheet.change_status_delivery(user=username, delivery=delivery_service)
            # else:
            #     logger.critical(f'Уровень похожести ниже установленного уровня. Данные не будут записаны')
            # КОНЕЦ БЛОКА

            # получаем ID оригинального nickname_orig в датафрейме уникальных аккаунтов из Excel-файла
            column_id = df_unique_nicknames.loc[df_unique_nicknames['Ваш ник в Instagram'] == nickname_orig, 'ID']
            # записываем все несовпадения и уровень похожести в датафрейм для выгрузки отчета в Excel-файл
            df_excel_yandex_forms.loc[df_excel_yandex_forms['ID'] == int(column_id), 'Нечеткий поиск'] =\
                result_fuzz_one[0]
            df_excel_yandex_forms.loc[(df_excel_yandex_forms['ID'] == int(column_id)), 'Уровень'] =\
                result_fuzz_one[1]

    # сортируем на месте по времения создания, приводим к изначальному порядку, как в Excel0-файле
    df_excel_yandex_forms.sort_values(['Время создания'], ascending=False, inplace=True)
    # print(df_excel_yandex_forms[df_excel_yandex_forms['Уровень'] != '']
    #       [['Ваш ник в Instagram', 'ID', 'Нечеткий поиск', 'Уровень']])


if __name__ == '__main__':

    path_log_delivery = os.getcwd() + f'\\logs\\debug_delivery.log'
    logger.add(path_log_delivery, level='DEBUG', compression="zip", rotation="9:00", retention="10 days")

    logger.info(f'Запуск скрипта')

    # получаем датафрейм полной таблицы Excel Яндекс формы, в которую потом при необходимости сформируется отчет,
    # а такаже датафрейм уникальных аккаунтов инстаграм из этой таблицы и список этих аккаунтов
    path_delivery = os.getcwd() + f'\\delivery\\'
    try:
        df_excel_yandex_forms, df_unique_nicknames = get_excel_yandex_forms()
    except Exception:
        sys.exit(1)

    # получаем данные из Google-таблицы
    try:
        gsheet = GSheetsDelivery(URL_GSHEET, START_ADDR, END_ADDR, NAME_WORKSHEET)
    except Exception:
        sys.exit(1)

    logger.info(f'Начинаем перебор и поиск аккаунтов в Яндекс-форме и Google-таблице')
    set_delivery_service()
    logger.info(f'Завершили перебор и поиск аккаунтов в Яндекс-форме и Google-таблице')

    try:
        logger.info(f'Начинаем запись результатов в Google-таблицу')
        gsheet.save_df_gsheet()
    except Exception as ex:
        logger.critical(f'Ошибка! Не удалось записать данные в Google_таблицу. {ex}')

    try:
        logger.info(f'Начинаем запись отчета в Excel-файл')
        file_report = f'{path_delivery}{datetime.now().strftime("%Y%m%d%H%M%S")}_Report.xlsx'
        df_excel_yandex_forms.to_excel(excel_writer=file_report, sheet_name='Report', engine='xlsxwriter')
        logger.success(f'Отчет успешно записан в файл {file_report}')
    except Exception as ex:
        logger.error(f'Ошибка записи отчета в Excel-файл. Обратитесь к разработчику {ex}')
        path_pkl_report = os.getcwd() + '\\copy_dataframe\\'
        df_excel_yandex_forms.to_pickle(path_pkl_report + 'df_excel_yandex_forms.pkl')
