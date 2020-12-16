import os
import pickle
import time
import random

from loguru import logger
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

from settings.settings import HEADLESS, BROWSER_MODE, AUTH_MODE


class InstagramBot:
    """Instagram Bot"""

    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.options = webdriver.ChromeOptions()
        if BROWSER_MODE == 'PROFILE':  # режим авторизации из профиля браузера
            self.options.add_argument("--user-data-dir="
                                      "C:\\Users\\dboty\\AppData\\Local\\Yandex\\YandexBrowser\\User Data\\Instagram")
            # self.options.add_argument("--profile-directory=Default")  # никак не влияет на yandexdriver

        # проверка запуска браузера в скрытом режиме и установка параметров
        if HEADLESS:
            self.options.add_argument(HEADLESS)  # режим невидимого запуска браузера
            self.options.add_argument('--disable-gpu')
            self.options.add_argument('--enable-javascript')
            self.options.add_argument(
                f'--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                f' Chrome/87.0.4280.88 Safari/537.36')
            # self.options.add_argument('--no-sandbox')
            # self.options.add_argument('--ignore-certificate-errors')
            # self.options.add_argument('--allow-insecure-localhost')
        # self.driver_file = os.getcwd() + '/chromedriver/chromedriver.exe'  # path to ChromeDriver
        self.driver_file = os.getcwd() + '/yandexdriver/yandexdriver.exe'  # path to ChromeDriver
        self.browser = webdriver.Chrome(self.driver_file, options=self.options)

        time.sleep(5)  # после открытия браузера

    # метод для закрытия браузера
    def close_browser(self):
        self.browser.close()
        self.browser.quit()

    # метод проверяет по xpath существует ли элемент на странице
    def xpath_exists(self, url):
        try:
            self.browser.find_element_by_xpath(url)
            exist = True
        except NoSuchElementException:
            exist = False
        return exist

    # метод чтения cookies из файла cookies.pickle и добавление в бразуер для авторизации
    def load_cookies(self):
        with open(os.getcwd() + '\\cookies\\cookies.pickle', 'rb') as file:
            cookies = pickle.load(file)
            for cookie in cookies:
                self.browser.add_cookie(cookie)

    # метод сохранения cookies в файл cookies.pickle
    def save_cookies(self):
        with open(os.getcwd()+'\\cookies\\cookies.pickle', 'wb') as file:
            pickle.dump(self.browser.get_cookies(), file)

    # метод авторизации и перехода на страницу отправки сообщений
    def login_and_direct(self):

        # Проверяем наличие кнопки Direct и переходим на страницу отправки сообщения
        def button_direct():
            direct_message_button = '/html/body/div[1]/section/nav/div[2]/div/div/div[3]/div/div[2]/a'
            if self.xpath_exists(direct_message_button):
                self.browser.find_element_by_xpath(direct_message_button).click()
                time.sleep(random.randrange(3, 5))
                return True
            else:
                logger.error(f'Error. Кнопка direct для отправки сообщений не найдена!')
                return False

        # Отключаем всплывающее окно "Включить уведомления"
        def enable_notifications(button_not_now):
            if not HEADLESS:  # в режиме HEADLESS окно "Включить уведомления" не появлется
                # Отключаем всплывающее окно "Включить уведомления"
                #                  '/html/body/div[4]/div/div/div/div[3]/button[2]'
                # button_not_now = '/html/body/div[5]/div/div/div/div[3]/button[2]'
                if self.xpath_exists(button_not_now):
                    self.browser.find_element_by_xpath(button_not_now).click()
                    time.sleep(random.randrange(2, 4))
                    return True
                else:
                    logger.error(f'Error. Не получилось отключить всплывающее окно Включить уведомления')
                    return False

        self.browser.get('https://www.instagram.com')
        self.browser.set_window_size(768, 704)

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

            # Проверяем наличие кнопки Direct и переходим на страницу отправки сообщения
            if not button_direct():
                raise NoSuchElementException

            # Отключаем всплывающее окно "Включить уведомления"
            if not enable_notifications('/html/body/div[5]/div/div/div/div[3]/button[2]'):
                raise NoSuchElementException

            # сохранение cookies в файл после успешной авторизации
            if BROWSER_MODE == 'COOKIES':
                self.save_cookies()
                logger.info(f'Сохранение cookies после успешной авторизации в файл')

        # читаем cookies предыдущей сессии из файла cookies.pickle и авторизуемся без логина и пароля
        elif BROWSER_MODE == 'COOKIES':  # AUTH_MODE == 0
            logger.info(f'Чтение cookies для авторизации из файла')
            self.load_cookies()
            logger.info(f'Cookies из файла добавлены в браузер')
            # обновляем страницу браузера и появляется всплывающее окнов Включить уведомления
            time.sleep(1)
            self.browser.refresh()

            time.sleep(random.randrange(3, 5))

            # Отключаем всплывающее окно "Включить уведомления"
            if not enable_notifications('/html/body/div[4]/div/div/div/div[3]/button[2]'):
                raise NoSuchElementException

            # Проверяем наличие кнопки Direct и переходим на страницу отправки сообщения
            if not button_direct():
                raise NoSuchElementException

            # сохранение cookies в файл после успешной авторизации
            self.save_cookies()
            logger.info(f'Сохранение cookies после успешной авторизации в файл')

            # должны быть авторизованы из ранее сохраненного профиля в браузере
        elif BROWSER_MODE == 'PROFILE':  # AUTH_MODE == 0
            # Проверяем наличие кнопки Direct и переходим на страницу отправки сообщения
            # всплывающего окна Включить уведомления быть не должно
            if not button_direct():
                raise NoSuchElementException

    def send_direct_message(self, username="", messages=None):
        if messages is None:
            messages = []
        time.sleep(random.randrange(2, 4))

        # Нажимаем на кнопку новое сообщение
        button_new_message = '/html/body/div[1]/section/div/div[2]/div/div/div[1]/div[1]/div/div[3]/button'
        if self.xpath_exists(button_new_message):
            self.browser.find_element_by_xpath(button_new_message).click()
        else:
            logger.critical(f'Error. Отправка сообщений остановлена. Кнопка новое сообщение не найдена!')
            self.close_browser()
            return False

        time.sleep(random.randrange(1, 3))

        # В поле Кому вводим имя получателя
        to_name_input = self.browser.find_element_by_name('queryBox')
        to_name_input.clear()
        to_name_input.send_keys(username)

        time.sleep(random.randrange(3, 5))

        # выбираем получателя из найденного списка и жмем кнопку "Далее"
        button_select_user = '/html/body/div[5]/div/div/div[2]/div[2]/div[1]/div/div[3]/button'
        if self.xpath_exists(button_select_user):
            self.browser.find_element_by_xpath(button_select_user).click()
            time.sleep(random.randrange(3, 5))

            # проверка получателя для отправки с найденным и выбранным
            select_user = self.browser.find_element_by_xpath('/html/body/div[5]/div/div/div[2]/div[1]/div/div[2]'). \
                find_element_by_tag_name('button').text
            if select_user != username:
                logger.warning(f'Error. Получатель для отправки {username} '
                               f'не совпал с найденным и выбранным {select_user}')
                # закрываем окно поиска и выбора получателя
                self.browser.find_element_by_xpath('/html/body/div[5]/div/div/div[1]/div/div[1]/div/button').click()
                time.sleep(random.randrange(1, 3))
                return 'username not found'  # для изменения статуса на Аккаунт не найден

            # получатель найден и проверен, жмем кнопку Далее
            self.browser.find_element_by_xpath('/html/body/div[5]/div/div/div[1]/div/div[2]/div/button').click()
        else:
            logger.warning(f'NoError. Получатель {username} не найден')
            # закрываем окно поиска и выбора получателя
            self.browser.find_element_by_xpath('/html/body/div[5]/div/div/div[1]/div/div[1]/div/button').click()
            time.sleep(random.randrange(1, 3))
            return 'username not found'  # для изменения статуса на Аккаунт не найден

        time.sleep(random.randrange(3, 5))

        # вводим сообщение и отправляем получателю
        try:
            text_message_area = self.browser.find_element_by_xpath(
                "/html/body/div[1]/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[2]/textarea")
            text_message_area.clear()

            # решение проблемы невозможности отправки целиком сообщений с абзацами
            for message in messages:
                if '\n' in message:  # check if exists \n tag in text
                    textsplit = message.split('\n')  # explode
                    textsplit_len = len(textsplit) - 1  # get last element
                    for text in textsplit:
                        text_message_area.send_keys(text)
                        # do what you need each time, if not the last element
                        if textsplit.index(text) != textsplit_len:
                            text_message_area.send_keys(Keys.SHIFT + Keys.ENTER)
                else:
                    text_message_area.send_keys(message)
                if text_message_area.text:  # проверяем поле ввода перед отправкой
                    time.sleep(random.randrange(2, 4))
                    text_message_area.send_keys(Keys.ENTER)
                    if not text_message_area.text:  # проверяем поле ввода после отправки
                        logger.success(f'Сообщение для {username} успешно отправлено!')
                    time.sleep(random.randrange(2, 4))
                else:
                    logger.error(f'Error. Ошибка при вводе текста в поле для отправки сообщения для {username}.')
                    self.close_browser()
                    return False
        except Exception as ex:
            logger.critical(f'Error. Возникла ошибка при отправке сообщения. Отправка сообщений приостановлена.', ex)
            self.close_browser()
            return False
        # сохранение cookies в файл после успешной отправки сообщений аккаунту
        if BROWSER_MODE == 'COOKIES':
            self.save_cookies()
            logger.info(f'Сохранение cookies в файл после успешной отправки сообщений')
        return 'message delivered'  # успешная отправка сообщения


if __name__ == '__main__':
    # instagram = InstagramBot('username', 'password')
    pass
