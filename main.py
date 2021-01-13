"""
Задача:
Из Google-таблицы необходимо выбирать записи по статусом "НОВЫЙ", группировать по каждому аккаунту Instagram
и суммировать значения по столбцу "ЦЕНА".
Далее используя иммитацию работы в браузере отправлять сообщения каждому аккаунту в Instagram Direct.
Для работы в браузере используется библиотека Selenium, а работа с Google-таблицей ведется через API,
используя библиотеку pygsheets.

В коде созданы два класса: InstagramBot для работы в браузере. GSheetsBot - для работы с Google-таблицей.
"""

import os
import time

from loguru import logger
from instagrambot import InstagramBot
from gsheetsbot import GSheetsBot
from settings.settings import INSTALOGIN, INSTAPASSWORD, URL_GSHEET, STATUS_NEW, \
    TEMPLATE_MESSAGE_START, STATUS_SEND, STATUS_USERNOTFOUND, DELAY_SEND, LIMIT_USER, RECOVERY_DATAFRAME

if __name__ == '__main__':
    # настраиваем логирование
    path_log = os.getcwd() + f'\\logs\\debug.log'
    logger.add(path_log, level='DEBUG', compression="zip", rotation="9:00", retention="10 days")

    try:
        if not RECOVERY_DATAFRAME:
            logger.info(f'Начало работы. Считывание данных из Google-таблицы.')
        gsheet = GSheetsBot(URL_GSHEET)
        if not gsheet.usernames:  # Если нет заказов для рассылки, заврешаем работу скрипта
            raise SystemExit(1)
    except Exception as ex:
        logger.error(f'Error. Возникла ошибка на этапе получения данных из Google-таблицы. {ex}')
        raise SystemExit(1)

    logger.info(f'Открытие браузера.')
    instagram = InstagramBot(INSTALOGIN, INSTAPASSWORD)

    try:
        logger.info(f'Переход на страницу instagram.')
        instagram.login_and_direct()
        time.sleep(2)
    except Exception as ex:
        logger.error(f'Error. Возникла ошибка на этапе авторизации и перехода на страницу direct. {ex}')
        instagram.browser.save_screenshot(os.getcwd() + f'\\logs\\exception.png')
        try:
            instagram.close_browser()
        except Exception:
            pass
        raise SystemExit(1)

    counter = 0  # счетчик лимита кол-ва аккаунтов для отправки в рамках запуска одного скрипта LIMIT_USER
    for username in gsheet.usernames:
        counter += 1
        messages = [TEMPLATE_MESSAGE_START]
        if username != '':
            logger.info(f'{username} - начата подготовка и отправка сообщений')
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
            try:
                status_delivered = instagram.send_direct_message(username.lower(), messages)
            except Exception as ex:
                instagram.browser.save_screenshot(os.getcwd() + f'\\logs\\send_direct_message.png')
                logger.critical(f'Error. Возникла необработанная ошибка на этапе отправки сообщения. {ex}')
                break

            if status_delivered == 'message delivered':  # сообщение успешно доставлено
                # записать изменения в датафрейм, поменять на STATUS_SEND
                gsheet.change_status_order(username, STATUS_NEW, STATUS_SEND)
            elif status_delivered == 'username not found':  # аккаунт с заданным именем не найден
                # записать изменения в датафрейм, поменять на STATUS_USERNOTFOUND
                gsheet.change_status_order(username, STATUS_NEW, STATUS_USERNOTFOUND)
            elif not status_delivered:  # ошибка отправки сообщения
                break  # выходим из цикла отправки сообщений по списку аккаунтов

        # выводим позиции заказа
        df_order_log = gsheet.df_values[(gsheet.df_values['СТАТУС'] == STATUS_SEND) &
                                        (gsheet.df_values['АККАУНТ'] == username)]
        logger.debug(f'\n'
                     f'{df_order_log}')

        time.sleep(DELAY_SEND)  # пауза между отправками каждому аккаунту
        if counter >= LIMIT_USER:
            logger.info(f'Достигнут установленный лимит отправки сообщений. {LIMIT_USER}')
            break  # достигнут лимит отправки, выходим из цикла

    # записываем изменения статусов в Google-таблицу
    try:
        gsheet.save_df_gsheet()
    except Exception as ex:
        logger.critical(f'Error. Не удалось записать данные в Google_таблицу. {ex}'
                        f'Следующий запуск скрипта рекомендуется с параметром RECOVERY_DATAFRAME = 1')

    logger.info(f'Завершение работы. Закрытие браузера.')

    try:
        instagram.close_browser()
    except Exception:
        instagram.browser.save_screenshot(os.getcwd() + f'\\logs\\exception.png')
        logger.info(instagram.browser.session_id)
