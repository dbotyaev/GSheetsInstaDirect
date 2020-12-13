import os
import pandas as pd
import pygsheets

from datetime import datetime

from settings.settings import NAME_WORKSHEET, SERVICE_ACCOUNT_FILE, START_ADDR, END_ADDR,\
    STATUS_NEW, RECOVERY_DATAFRAME


class GSheetsBot:
    """
    Класс для работы с google-таблицей, содержащей аккаунты инстраграм, данные заказов.
    Данные загружаются в дата-фрейм, формируются сообщения в зависимости от статуса заказа.
    После отправки изменяются статусы и обновленный дата фрейм выгружается в google-таблицу.

    Используются 2 датафрейма из-за невозможности считать одновременно формулы
    и сохранить значения ячеек с плавающе точкой. С одного датафрейма данные
    используются для расчетов, в другой записываются изменения для обратной выгрузки в Google-таблицу
    """

    def __init__(self, url):
        try:
            path_api = os.getcwd() + SERVICE_ACCOUNT_FILE
            self.client = pygsheets.authorize(service_account_file=path_api)
            self.google_sheet = self.client.open_by_url(url)
            self.worksheet = self.google_sheet.worksheet_by_title(NAME_WORKSHEET)
            # дата-фрейм ЗНАЧЕНИЙ таблицы заказов диапазона ячеек START_ADDR:END_ADDR
            if not RECOVERY_DATAFRAME:
                self.df_values = self.worksheet.get_as_df(has_header=True,
                                                          index_column=None,
                                                          start=START_ADDR,
                                                          end=END_ADDR,
                                                          numerize=True,
                                                          empty_value='',
                                                          value_render='FORMATTED_VALUE')
                # дата-фрейм ФОРМУЛ таблицы заказов диапазона ячеек START_ADDR:END_ADDR
                # numerize=False чтобы сохранить исходное форматирование при чтении
                self.df_formulas = self.worksheet.get_as_df(has_header=True,
                                                            index_column=None,
                                                            start=START_ADDR,
                                                            end=END_ADDR,
                                                            numerize=False,
                                                            empty_value='',
                                                            value_render='FORMULA')
            else:
                print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                      f'Запущен режим восстановления. Считывание данных из сохраненных файлов *.pkl.')
                path_pkl = os.getcwd() + '\\copy_dataframe\\'
                self.df_values = pd.read_pickle(path_pkl + 'df_values.pkl')
                self.df_formulas = pd.read_pickle(path_pkl + 'df_formulas.pkl')

            # установка защиты на лист от редактирование на период работы скрипта
            def set_protect_worksheet(worksheet_id):
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

            # локализаиця дата-фрейма для работы в python, если локализация гугл-таблицы 'Россия'
            # при обратной загрузке требуется обратная замена
            # if self.google_sheet.to_json()['properties']['locale'] == 'ru_RU':
            #     self.df_formulas['ОБЪЕМ'] = self.df_formulas['ОБЪЕМ'].astype(str).str.replace(',', '.'). \
            #         fillna(self.df_formulas['ОБЪЕМ']).astype(str)

            # получаем список уникальных аккаунтов со статусом НОВЫЙ
            self.usernames = pd.unique(self.df_values[self.df_values['СТАТУС'] == STATUS_NEW]['АККАУНТ']).tolist()
            if not self.usernames:
                print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                      f'Заказы для рассылки со статусом {STATUS_NEW} отсутствуют.')
            else:
                print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                      f'Сформирован список аккаунтов для отправки. Общее кол-во {len(self.usernames)}.')
                try:
                    self.protect_wsheet_id = set_protect_worksheet(self.worksheet.id)
                    print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                          f'Установлена защита на лист на период работы скрипта.')
                except Exception:
                    print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                          f'NoError. Невозможно установить защиту на лист на период работы скрипта. '
                          f'Проверьте установлена ли защита на листе.')
                    self.protect_wsheet_id = 0
        except:
            raise

    def orders_by_status(self, username, status):
        """
        возвращает список позиций в датафрейме и общую сумму по ним
        :param username: условие отбора по АККАУНТу
        :param status: условие отбора по СТАТУСу
        :return: список позиций в датафрейме по статусу и общая сумма по ним
        """
        products = self.df_values[(self.df_values['СТАТУС'] == status) & \
                                  (self.df_values['АККАУНТ'] == username)] \
            [['АРОМАТ', 'ОБЪЕМ', 'ЦЕНА']].values.tolist()
        total_sum = self.df_values[(self.df_values['СТАТУС'] == status) & \
                                   (self.df_values['АККАУНТ'] == username)]['ЦЕНА'].sum()
        return products, total_sum

    def change_status_order(self, username, old_status, new_status):
        """
        Изменение статуса позиций в заказе по фильтру АККАУНТа и текущего СТАТУСа
        Загрузка данных и все расчеты делаются из датафрейма df_values,
        а изменения и обратная загрузка из df_formulas, чтобы сохранить формулы и исходное форматирование
        :param username: АККАУНТ
        :param old_status: текущий СТАТУС позиций в заказе
        :param new_status: новый СТАТУС позиций в заказе
        :return: изменяет датафрейм с заказами
        """

        self.df_formulas.loc[(self.df_formulas['СТАТУС'] == old_status) &
                             (self.df_formulas['АККАУНТ'] == username), 'СТАТУС'] = new_status
        # здесь внесение изменений необходимо для выгрузки в Excel (если была ошибка выгрузки в Google-таблицу
        self.df_values.loc[(self.df_values['СТАТУС'] == old_status) &
                           (self.df_values['АККАУНТ'] == username), 'СТАТУС'] = new_status
        # сохраняем копии объектов dataframe self.df_formulas и self.df_values в объекты pickle в папку copy_dataframe
        # через встроенный функционал pandas
        path_pkl = os.getcwd() + '\\copy_dataframe\\'
        self.df_values.to_pickle(path_pkl+'df_values.pkl')
        self.df_formulas.to_pickle(path_pkl + 'df_formulas.pkl')
        print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
              f'Изменения сохранены в файлы *.pkl')

    def save_df_gsheet(self):  # итоговое сохрание данных в Google-таблицу
        try:
            # локализация значений для Google-таблицы, в столбце 'ОБЪЕМ' меняем плавающую точку на запятую
            if self.google_sheet.to_json()['properties']['locale'] == 'ru_RU':
                self.df_formulas['ОБЪЕМ'] = self.df_formulas['ОБЪЕМ'].astype(str).str.replace('.', ','). \
                    fillna(self.df_formulas['ОБЪЕМ']).astype(str)

            try:
                self.worksheet.copy_to(self.google_sheet.id)
                print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                      f'Перед записью измененений создана копия листа с заказами в Google-таблице.')
            except Exception:
                print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                      f'Не удалось создать копию листа перед загрузкой измененений в Google-таблицу.'
                      f'Продолжаем работу скрипта.')

            if self.protect_wsheet_id:
                self.worksheet.remove_protected_range(self.protect_wsheet_id)
                print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                      f'Снята ранее установленная защита на лист для записи изменений в Google-таблицу.')
            else:
                print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                      f'NoError. Невозможно снять защиту на листе, установленную до запуска скрипта.')

            # запись данных повверх существующих значений
            self.worksheet.set_dataframe(self.df_formulas,
                                         START_ADDR,
                                         copy_index=False, copy_head=True, extend=False, fit=False,
                                         escape_formulae=False,
                                         nan='')
            print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                  f'Зафиксированные изменения успешно записаны в Google-таблицу.')
        except:
            path_xls = os.getcwd() + '\\copy_dataframe\\' + datetime.now().strftime("%Y%m%d %H%M%S") + '.xlsx'
            self.df_values.to_excel(excel_writer=path_xls,
                                    sheet_name=datetime.now().strftime("%Y%m%d %H%M%S"),
                                    engine='xlsxwriter')
            print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                  f'Error. Произошла ошибка. Новые статусы в Google-таблицу не записаны.'
                  f'Копия данных сохранена в Excel-файл в папку copy_dataframe.')
            raise


if __name__ == '__main__':
    # gsheet = GSheetsBot(URL_GSHEET)
    pass