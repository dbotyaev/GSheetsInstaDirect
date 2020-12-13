"""
Задача:
Из Google-таблицы необходимо выбирать записи по статусом "НОВЫЙ", группировать по каждому аккаунту Instagram
и суммировать значения по столбцу "ЦЕНА".
Далее используя иммитацию работы в браузере отправлять сообщения каждому аккаунту в Instagram Direct.
Для работы в браузере используется библиотека Selenium, а работа с Google-таблицей ведется через API,
используя библиотеку pygsheets.

В коде созданы два класса: InstagramBot для работы в браузере. GSheetsBot - для работы с Google-таблицей.
"""

import time
from datetime import datetime

from instagrambot import InstagramBot
from gsheetsbot import GSheetsBot
from settings.settings import INSTALOGIN, INSTAPASSWORD, URL_GSHEET, STATUS_NEW, \
    TEMPLATE_MESSAGE_START, STATUS_SEND, STATUS_USERNOTFOUND, DELAY_SEND, LIMIT_USER, RECOVERY_DATAFRAME

if __name__ == '__main__':
    try:
        if not RECOVERY_DATAFRAME:
            print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                  f'Начало работы. Считывание данные из Google-таблицы.')
        gsheet = GSheetsBot(URL_GSHEET)
        if not gsheet.usernames:  # Если нет заказов для рассылки, заврешаем работу скрипта
            raise SystemExit(1)
    except Exception as ex:
        print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
              f'Error. Возникла ошибка на этапе получения данных из Google-таблицы. {ex}')
        raise SystemExit(1)

    print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
          f'Открытие браузера.')
    instagram = InstagramBot(INSTALOGIN, INSTAPASSWORD)

    try:
        print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
              f'Переход на страницу instagram.')
        instagram.login_and_direct()
        time.sleep(2)
    except Exception as ex:
        print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
              f'Error. Возникла ошибка на этапе авторизации и перехода на страницу direct. {ex}')
        instagram.close_browser()
        raise SystemExit(1)

    counter = 0  # счетчик лимита кол-ва аккаунтов для отправки в рамках запуска одного скрипта LIMIT_USER
    for username in gsheet.usernames:
        counter += 1
        messages = [TEMPLATE_MESSAGE_START]
        if username != '':
            print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                  f'{username} - начата подготовка и отправка сообщений')
            # по каждому аккаунту формируем список позиций товаров для сообщения и общую сумму за все позиции
            products, total_sum = gsheet.orders_by_status(username, STATUS_NEW)
            message_order = ''
            for number, product in enumerate(products):
                message_order += f'{number + 1}. {product[0]}\n' \
                                 f'Объем, мл: {product[1]}\n' \
                                 f'Цена: {product[2]} руб\n'
            message_order += f'\n' \
                             f'Итого к оплате без учёта доставки: {total_sum} руб.'
            messages.append(message_order)

            # отправка полученного списка сообщений каждому аккаунту
            status_delivered = instagram.send_direct_message(username.lower(), messages)

            if status_delivered == 'message delivered':  # сообщение успешно доставлено
                # записать изменения в датафрейм, поменять на STATUS_SEND
                gsheet.change_status_order(username, STATUS_NEW, STATUS_SEND)
            elif status_delivered == 'username not found':  # аккаунт с заданным именем не найден
                # записать изменения в датафрейм, поменять на STATUS_USERNOTFOUND
                gsheet.change_status_order(username, STATUS_NEW, STATUS_USERNOTFOUND)
            elif not status_delivered:  # ошибка отправки сообщения
                break  # выходим из цикла отправки сообщений по списку аккаунтов, т.к. браузер закрыт

        # выводим позиции заказа
        print(gsheet.df_values[(gsheet.df_values['СТАТУС'] == STATUS_SEND) &
                               (gsheet.df_values['АККАУНТ'] == username)])
        time.sleep(DELAY_SEND)  # пауза между отправками каждому аккаунту
        if counter >= LIMIT_USER:
            print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} '
                  f'Достигнут установленный лимит отправки сообщений. {LIMIT_USER}')
            break  # достигнут лимит отправки, выходим из цикла

    # записываем изменения статусов в Google-таблицу
    try:
        gsheet.save_df_gsheet()
    except Exception as ex:
        print(ex)
        instagram.close_browser()
        raise SystemExit(1)

    print(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")} Завершение работы. Закрытие браузера.')

    try:
        instagram.close_browser()
    except Exception:
        print(instagram.browser.session_id)