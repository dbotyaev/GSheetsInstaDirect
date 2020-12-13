# Использовать ОСТОРОЖНО.
# параметр считывания данных в датафреймы df_values и df_formulas
# из сохраненного объекта pickle в папке copy_dataframe.
# можно использовать при сбое, когда была сделана отправка данных в instagram,
# а изменения не были сохранены в Google-таблицу.
# 0 - по умолчанию, датафреймы формируются из Google-таблицы
# 1 - датафреймы формируются из Google-таблицы из объектов df_formulas.pkl и df_values.pkl

RECOVERY_DATAFRAME = 0

# логин и пароль instagram-аккаунта, с которого будут отправляться сообшения
INSTALOGIN = '*'
INSTAPASSWORD = '*'

# данные для работы с google-таблицей
URL_GSHEET = 'https://docs.google.com/spreadsheets/d/*/'
NAME_WORKSHEET = 'ЗАКАЗЫ'  # имя листа с заказами
START_ADDR = 'A1'
END_ADDR = 'E2000'
SERVICE_ACCOUNT_FILE = '\settings\*.json'

STATUS_NEW = 'НОВЫЙ'  # статус, по которому будет производиться поиск и отправка, можно изменять при необходимости
STATUS_SEND = 'СЧЕТ'
STATUS_USERNOTFOUND = 'АККАУНТ НЕ НАЙДЕН'  # статус, на который менять, когда АККАУНТ не найден в Instagram

DELAY_SEND = 15  # Пауза в секундах между отправками каждому аккаунту

LIMIT_USER = 2  # Лимит количества пользователей для рассылки при одном запуске скрипта

# Режим работы браузера
# HEADLESS = '--headless'  # скрытый режим в chromedriver не работает
HEADLESS = ''  # видимый режим

# Шаблон сообщения получателю
TEMPLATE_MESSAGE_START = '*'