#!/usr/bin/env python3
# coding = utf-8

import argparse
import copy
import datetime
import email
import imaplib
import os
import re
import sys
import tempfile
import time
import uuid
from email.header import decode_header

import bs4
from python_rucaptcha import ImageCaptcha
from selenium import webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from termcolor import colored

arguments_parser = argparse.ArgumentParser()
arguments_parser.add_argument('--email', help='Почта', action='store')
arguments_parser.add_argument('--password', help='Пароль для почты', action='store')
arguments_parser.add_argument('--only_read', help='Обрабатываем только прочитанные письма', action='store_true')
arguments_parser.add_argument('--only_unread',
                              help=f'''Обрабатываем только непрочитанные письма. {colored('Важно!', color='red')} 
Если ни один из параметров --only-read или --only-unread не указан, то будет обрабатываться вся почта''',
                              action='store_true')
arguments_parser.add_argument('--from_date', help='начальная дата письма')
arguments_parser.add_argument('--number', help='номера заявок росреестра')
arguments_parser.add_argument('--to_date', help=f'''конечная дата письма. {colored('Важно!', color='red')}
Если даты не будут указаны, обрабатываться будет вся почта''')
arguments_parser.add_argument('--out_dir', help='Каталог, в который будут складываться скачанные файлы', action='store')
arguments = arguments_parser.parse_args()

out_dir = os.path.abspath(arguments.out_dir)

# От кого будем искать письма, захардкодим значение
email_from = 'portal@rosreestr.ru'
# ключик рукапчи
rucaptcha_key = '8433ba8b9b3f97f8b017eff242df2531'


#
# sys.path.append('./bin/')
# sys.path.append('./bin/firefox/')

# if not os.path.isfile('./bin/geckodriver'):
#     print(colored('Не найден geckodriver!', color='red'))


# ! Используем системный firefox
# if not os.path.isfile('./bin/firefox/firefox'):
#    print(colored('Не найден исполняемый файл firefox!', color='red'))


class ImapSession:
    def __init__(self, email, password,
                 imap_host, imap_port, domain=None):
        self.email, self.password = email, password
        self.imap_host, self.imap_port = imap_host, imap_port
        #
        self.connection = None
        #
        # сюда складываем Id непрочитанных сообщений
        self.ids_messages = []
        #
        self.folder_list = []
        self.folder_code = None
        #
        self.message = None
        self.content = None
        self.message_headers = None
        self.html_message = None
        self.subject = None
        self.domain = domain

    def _search_folder(self, folder_name):
        try:
            folder_status, folder_list = self.connection.list()
        except Exception as error:
            message = f'{self.email}: не удалось получить список каталогов, подробная информация: {error}'
            print(colored(message, color='red'))
            return False
        else:
            if folder_status == 'OK':
                _tmp_folder = []
                for folder in folder_list:
                    _tmp_folder.append(folder.decode('utf-8'))
                for folder in _tmp_folder:
                    if folder_name in folder:
                        self.folder_code = folder.split(' "|" ')[1]
                        print(self.folder_code)
                        return True
                else:
                    message = f'{self.email}: не удалось найти каталог {folder_name}'
                    print(colored(message, color='red'))
                    return False
            else:
                message = f'{self.email}: не удалось получить имена каталогов'
                print(colored(message, color='red'))
                return False

    def _get_folder_list(self):
        self.folder_list.clear()
        try:
            folder_status, folder_list = self.connection.list()
        except Exception as error:
            message = f'{self.email}: не удалось получить список каталогов, подробная информация: {error}'
            print(colored(message, color='red'))
            return False
        else:
            if folder_status == 'OK':
                _folder_list = [folder.decode('utf-8').split(' "|" ')[1] for folder in folder_list]
                self.folder_list.extend(_folder_list)
                return True
            else:
                message = f'{self.email}: не удалось получить имена каталогов'
                print(colored(message, color='red'))
                return False

    def check_folder_exist(self, folder_name):
        message = f'{self.email}: проверяем существование каталога "{folder_name}"'
        print(colored(message, color='cyan'))
        if self._get_folder_list():
            if folder_name in self.folder_list:
                message = f'{self.email}: каталог "{folder_name}" существует'
                print(colored(message, color='cyan'))
                return True
            else:
                message = f'{self.email}: каталога "{folder_name}" не существует'
                print(colored(message, color='red'))
                return False
        else:
            return False

    def create_folder(self, folder_name):
        '''
        Создаём каталог folder_name в текущем местоположении
        :param folder_name:
        :return:
        '''
        message = f'{self.email}: создаём каталог "{folder_name}"'
        print(colored(message, color='cyan'))
        try:
            create_status, create_info = self.connection.create(folder_name)
        except Exception as error:
            message = f'{self.email}: при создании каталога {folder_name} произошла ошибка, подробная информация: {error}'
            print(colored(message, color='red'))
            return False
        else:
            if create_status == 'OK':
                message = f'{self.email}: каталог {folder_name} успешно создан'
                print(colored(message, color='cyan'))
                return True
            else:
                message = f'{self.email}: не удалось создать каталог {folder_name}'
                print(colored(message, color='red'))
                return False

    def load_messages(self, unread=False, read=False,
                      from_date=None, to_date=None, from_addr=None):
        '''
        Загружаем id непрочитанных писем в текущем каталоге
        :return:
        '''
        commands = []
        if unread:
            message = f'{self.email}: загружаем непрочитанные письма'
            commands.append('(UNSEEN)')
        else:
            if read:
                message = f'{self.email}: загружаем прочитанные письма'
                commands.append('(SEEN)')
            else:
                message = f'{self.email}: загружаем все письма'
                commands.append('ALL')

        time_format = '%d-%b-%Y'
        if from_date is not None and to_date is not None:
            commands.append(f'(SINCE "{from_date.strftime(time_format)}" BEFORE "{to_date.strftime(time_format)}")')

        else:
            if from_date is not None:
                commands.append(f'(SINCE "{from_date.strftime(time_format)}")')

            if to_date is not None:
                commands.append(f'(BEFORE "{to_date.strftime(time_format)}")')

        '''
        if to_date is not None:
            commands.append(f'(BEFORE "{str(to_date)}")')
        '''
        if from_addr is not None:
            commands.append(f'(FROM "{from_addr}")')

        self.ids_messages.clear()

        print(colored(message, color='cyan'))

        try:
            if self.domain and self.domain == 'mail':
                code, messages = self.connection.search(None, 'ALL')
            else:
                code, messages = self.connection.search(None, ' '.join(commands))
        except Exception as error:
            message = f'{self.email}: произошла ошибка при поиске непрочитанных писем, подробная информация: {error}'
            print(colored(message, color='red'))
            return False
        else:
            if code == 'OK':
                messages = messages[0].decode('utf-8').split(' ')
                for piece in messages:
                    self.ids_messages.append(bytearray(piece.encode('utf-8')))

                message = f'{self.email}: загружено "{len(self.ids_messages)}" непрочитанных сообщений'
                print(colored(message, color='cyan'))
                return True

            else:
                message = f'{self.email}: произошла ошибка при загрузке непрочитанных сообщений'
                print(colored(message, color='red'))
                return False

    def move_message_to_folder(self, message_id, folder_name):
        '''

        :param message_id:
        :param folder_name:
        :return:
        '''
        message = f'{self.email}: перемещаем письмо "{message_id.decode("utf-8")}" в каталог {folder_name}'
        print(colored(message, color='cyan'))
        try:
            status, info = self.connection.copy(message_id, folder_name)
        except Exception as error:
            message = f'{self.email}: произошла ошибка при перемещении письма, подробная информация: {error}'
            print(colored(message, color='red'))
            return False
        else:
            if status == 'OK':
                message = f'{self.email}: сообщение перемещено'
                print(colored(message, color='cyan'))
                return True
            else:
                message = f'{self.email}: произошла ошибка при перемещении письма'
                print(colored(message, color='red'))
                return False

    def delete_message(self, message_id):
        print(colored(f'{self.email}: удаляем сообщение {message_id.decode("utf-8")}', color='cyan'))
        try:
            status, info = self.connection.store(message_id, '+FLAGS', '\\Deleted')
        except Exception as error:
            message = f'{self.email}: произошла ошибка при удалении письма, подробная информация: {error}'
            print(colored(message, color='red'))
            return False
        else:
            if status == 'OK':
                return True
            else:
                message = f'{self.email}: произошла ошибка при удалении письма'
                print(colored(message, color='red'))
                return False

    def move_to_folder(self, folder_name):
        '''
        Переходим в каталог folder_name
        :param folder_name: название каталога
        :return:
        '''
        message = f'{self.email}: переходим в каталог "{folder_name}"'
        print(colored(message, color='cyan'))
        try:
            code, messages = self.connection.select(folder_name)
        except Exception as error:
            message = f'{self.email}: при переходе в каталог "{folder_name}" произошла ошибка, подробная информация: {error}'
            print(colored(message, color='red'))
            return False
        else:
            if code == 'OK':
                message = f'{self.email}: перешли в каталог "{folder_name}"'
                print(colored(message, color='cyan'))
                return True
            else:
                message = f'{self.email}: не удалось перейти в каталог "{folder_name}"'
                print(colored(message, color='red'))
                return False

    def loading_message_headers(self, message_id):
        '''
        Загружаем хидеры _одного_ сообщения
        :return:
        '''
        message = f'{self.email}: загружаем сообщение с id "{message_id.decode("utf-8")}"'
        print(colored(message, color='cyan'))
        try:
            status, message_headers = self.connection.fetch(message_id, '(BODY.PEEK[HEADER])')
        except Exception as error:
            message = f'{self.email}: загрузка сообщения не удалась, подробная информация: "{error}"'
            print(colored(message, color='red'))
            return False
        else:
            if status == 'OK':
                message = f'{self.email}: сообщение загружено'
                print(colored(message, color='cyan'))
                try:
                    self.message_headers = email.message_from_bytes(message_headers[0][1])
                except TypeError:
                    message = f'{self.email}: ошибка почтового сервера, отсутствует тело письма, письмо обработается при следующем проходе'
                    print(colored(message, color='red'))
                    return False
                else:
                    return True

            else:
                message = f'{self.email}: загрузка сообщения не удалась'
                print(colored(message, color='red'))
                return False

    def connect(self):
        message = f'{self.email}: подключаемся к {self.imap_host}'
        print(colored(message, color='cyan'))

        try:
            self.connection = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
        except Exception as error:
            message = f'{self.email}: не удалось подключиться к {self.imap_host}:{self.imap_port}, причина: {error}'
            print(colored(message, color='red'))
            return False
        else:
            connection_status = self.connection.login(self.email, self.password)
            try:
                auth_status = connection_status[0]
            except IndexError:
                message = f'{self.email}: не удалось найти код успешной авторизации, подробная информация: {connection_status}'
                print(colored(message, color='red'))
                return False

            else:
                if auth_status == 'OK':
                    message = f'{self.email}: успешно подключились'
                    print(colored(message, color='cyan'))
                    return True
                else:
                    message = f'{self.email}: подключение неудалось, причина: {auth_status}'
                    print(colored(message, color='red'))
                    return False

    def reconnect(self):
        print(colored(f'{self.email}: переподключаемся к серверу', color='red'))
        if self.connect():
            if self.move_to_folder('inbox'):
                return True
        return False

    def load_message(self, message_id):
        '''
        Загружаем хидеры _одного_ сообщения
        :return:
        '''
        message = f'{self.email}: загружаем сообщение с id "{message_id.decode("utf-8")}"'
        print(colored(message, color='cyan'))
        try:
            status, message = self.connection.fetch(message_id, '(RFC822)')
        except Exception as error:
            message = f'{self.email}: загрузка сообщения не удалась, подробная информация: "{error}"'
            print(colored(message, color='red'))
            # переподключаемся в случае ошибки
            if self.reconnect():
                return self.load_message(message_id)

            return False

        else:
            if status == 'OK':
                try:
                    self.message = email.message_from_bytes(copy.copy(message[0][1]))
                except TypeError:
                    message = f'{self.email}: ошибка почтового сервера, отсутствует тело письма, письмо обработается при следующем проходе'
                    print(colored(message, color='red'))
                    return False
                else:
                    self.content = self.message.get_payload(decode=True)
                    message = f'{self.email}: сообщение загружено'
                    try:
                        self.html_message = self.message.get_payload(0).get_payload(1).get_payload(decode=True).decode(
                            'utf-8')
                    except:
                        print(f'{self.email}: не удалось распарсить тело сообщения')
                        return False

                    else:
                        self.html_message = bs4.BeautifulSoup(self.html_message, features="html.parser")

                    try:
                        self.subject = decode_header(self.message.get('Subject'))[0][0].decode('utf-8')
                    except:
                        print(f'{self.email}: не удалось распарсить тему сообщения')
                        return False

                    print(colored(message, color='cyan'))
                    return True

            else:
                message = f'{self.email}: загрузка сообщения не удалась'
                print(colored(message, color='red'))
                return False

    def logout(self):
        print(colored(f'{self.email}: отключаемся', color='white'))
        self.connection.close()


def start_browser(download_path):
    print(colored('Запускаем Firefox', color='green'))

    options = Options()
    options.headless = True
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", download_path)
    if sys.platform != 'win32':
        options.set_preference("browser.helperApps.neverAsk.saveToDisk",
                           "application/octet-stream,application/vnd.ms-excel,application/zip,application/x-zip-compressed")
        browser = webdriver.Firefox(options=options, executable_path='c:\\GECKoDRIVER\\geckodriver.exe')
    browser.set_page_load_timeout(120)
    # устанавливаем размер экрана чуть больше чем обычно - иногда элементы на которые нужно кликнуть - не показываются
    # на странице и происходит ошибка
    browser.set_window_size(1920, 1080)
    return browser


def parse_link(browser, result, email):
    count, max_count = 0, 4
    while count < max_count:
        # загружаем страницу
        count += 1
        print(colored(f'{email}: загружаем {result["download_url"]}', color='green'))
        try:
            browser.get(result['download_url'])
        except:
            print(colored('Произошла ошибка при загрузке страницы! Вероятно браузер умер',
                          color='red'))
            continue

        # Ищем элемент с капчей
        try:
            image_elements = browser.find_elements_by_tag_name('img')
        except:
            print(colored('Не удалось найти элемент капчи', color='red'))
            continue

        else:
            for image_element in image_elements:
                try:
                    src_attr = image_element.get_attribute('src')
                except:
                    print(colored('Не удалось получить информацию о капче', color='red'))
                else:
                    if 'captcha' in src_attr.lower():
                        captcha_element = image_element
                        break
            else:
                print(colored('На странице не обнаружена капча', color='red'))
                continue

            # нашли элемент с капчей, сохраняем его куда-нибудь
            captcha_file_dst = os.path.join(tempfile.mkdtemp(), f'{str(uuid.uuid4())}.png')
            captcha_element.screenshot(captcha_file_dst)
            print(colored(f'{email}: отсылаем капчу на решение', color='cyan'))
            user_answer = ImageCaptcha.ImageCaptcha(rucaptcha_key=rucaptcha_key).captcha_handler(
                captcha_file=captcha_file_dst)

            if not user_answer['error']:
                captcha_code = user_answer['captchaSolve']
                print(colored(f'{email}: получен код капчи "{captcha_code}"', color='green'))
            else:
                print(colored(f'{email}: не удалось решит капчу, пробуем заново', color='yellow'))
                break

            print(colored(f'{email}: вводим капчу в форму', color='cyan'))
            try:
                captcha_form = browser.find_element_by_name('captchaText')
            except:
                print(colored(f'{email}: не удалось найти форму для ввода капчи'))
                continue
            else:
                captcha_form.send_keys(captcha_code)

            try:
                browser.find_element_by_class_name('terminal-button-light').click()
            except:
                print(colored(f'{email}: не удалось найти кнопку для отправки капчи', color='red'))
                continue
            else:
                time.sleep(2)

            try:
                a_elements = browser.find_elements_by_tag_name('a')
            except:
                print(colored(f'{email}: не удалось найти кнопку для ввода ключа'))
                continue
            else:
                for element in a_elements:
                    attribute = element.get_attribute('onclick')
                    if attribute is not None:
                        if 'setAccessType' in attribute:
                            access_key_button = element
                            try:
                                access_key_button.click()
                            except:
                                print(colored(f'{email}: не удалось обработать ключ'))
                            break
                else:
                    print(colored(f'{email}: не удалось найти кнопку для ввода ключа'))
                    continue

        try:
            access_key_form = browser.find_element_by_name('accessKey')
        except:
            print(colored(f'{email}: не удалось найти форму для ввода ключа', color='green'))
            continue

        print(colored(f'{email}: вводим ключ {result["key"]} в форму', color='cyan'))
        access_key_form.send_keys(result.get('key'))
        try:
            get_file_buttons = browser.find_elements_by_class_name('terminal-button-light')
        except:
            print(colored(f'{email}: не удалось найти кнопку для скачивания файла', color='green'))
            continue
        else:
            current_files = get_current_list_of_files(out_dir)
            for element in get_file_buttons:
                try:
                    element_text = element.text
                except:
                    continue

                if 'файл' in element_text:
                    print(colored(f'{email}: загружаем файл', color='cyan'))
                    print(colored(f'{email}: ожидаем окончание загрузки файла', color='cyan'))
                    i = 0
                    while i < 5:
                        try:
                            print(f'try = {i}')
                            element.click()
                        except:
                            print(colored(f'{email}: не удалось скачать файл'))
                            i += 1
                            continue
                        else:
                            break
                    time.sleep(15)
                    break
            else:
                print(colored(f'{email}: не удалось найти кнопку для скачивания файла', color='green'))
                continue

        return calculate_new_files_in_dir(current_files, out_dir)

    print(colored(f'{email}: количество попыток обработать письмо достигло максимума'))
    return None


def parse_message(message_body, message_subject):
    '''
    Парсим само html сообщение
    :param message_body: раскодированное html сообщение, объект bs4
    :param message_subject: раскодированный сабж
    :return:
    '''
    result = {
        # номер заявления
        'application_number': None,
        # Дата регистрации заявления
        'reg_date': None,
        # ссылка на скачивание
        'download_url': None,
        # ключ доступа
        'key': None
    }
    true_subject = ['документ по заявлению', 'заявление выполнено']
    if not true_subject[0] in message_subject and true_subject[1] in message_subject:
        print(colored('Не используем это сообщение, не соответствует запросу', color='yellow'))
        return result
    # for subj in true_subject:
    #     if message_subject == 'Портал Росреестра: заявление выполнено  (45-3819900)':
    #         print("ok")
    #     print(message_subject)
    #     if subj in message_subject:
    #         break
    #     else:

    # не очень красивые регулярки, увы
    application_number = re.search('\(\S+\)', str(message_subject))
    if application_number is not None:
        application_number = application_number.group().replace('(', '').replace(')', '')
    else:
        print('Не удалось распарсить номер заявления!')
        return

    result['application_number'] = application_number
    urls = message_body.find_all('a')
    for piece in urls:
        if 'requestNumber' in piece.get('href'):
            result['download_url'] = piece.get('href')
            break

    key = re.search('код <b>\S\S\S\S\S</b>', str(message_body))
    if key is not None:
        key = key.group().replace('код <b>', '').replace('</b>', '')
        result['key'] = key
    else:
        key = re.search('ключ <b>\S\S\S\S\S</b>', str(message_body))
        if key is not None:
            key = key.group().replace('ключ <b>', '').replace('</b>', '')
            result['key'] = key
        else:
            key = re.search('ключ \S\S\S\S\S', str(message_body))
            if key is not None:
                key = key.group().replace('ключ ', '').replace('</b>', '')
                result['key'] = key
            else:
                print('Не удалось распарсить ключ!')
                return
    try:
        reg_date = re.findall('\d\d.\d\d.\d\d\d\d', str(message_body))[1]
    except:
        print('Не удалось распарсить время регистрации!')
        return
    else:
        result['reg_date'] = reg_date

    return result


def get_current_list_of_files(path):
    '''Получаем список файлов в каталоге'''
    return os.listdir(path)


def calculate_new_files_in_dir(old_files_list, path):
    '''Ищем новые файлы в каталоге'''
    new_files_list = get_current_list_of_files(path)
    for file in new_files_list:
        if file in old_files_list:
            continue
        else:
            # нашли новый файл
            return file


def main():
    email, password = arguments.email, arguments.password
    log_dst = os.path.join(os.path.dirname(__file__), f'log\\{email}.txt')
    print(log_dst)
    if os.path.isfile(log_dst):
        log_file = open(log_dst, 'a')
    else:
        log_file = open(log_dst, 'w')
    log_file.write('=' * 40)
    log_file.write('\n')
    log_file.write(f'{datetime.datetime.now()}: {email}: начинаем работу\n')
    log_file.flush()

    rosreestr_number = arguments.number
    if rosreestr_number:
        rosreestr_number = rosreestr_number.split(',')

    IMAPS_EMAIL = {
        'yandex': 'imap.yandex.ru',
        'mail': 'imap.mail.ru',
        'ya': 'imap.ya.ru'
    }

    if email is not None and password is not None:
        time_format = '%Y-%m-%d'
        if arguments.from_date is not None:
            try:
                from_date = datetime.datetime.strptime(arguments.from_date, time_format)
            except:
                print(colored('Неверно указано значение from_date', color='red'))
                sys.exit(1)
        else:
            from_date = None

        if arguments.to_date is not None:
            try:
                to_date = datetime.datetime.strptime(arguments.to_date, time_format)
            except:
                print(colored('Неверно указано значение to_date', color='red'))
                sys.exit(1)

        else:
            to_date = None

        current_domin = [(k, v) for (k, v) in IMAPS_EMAIL.items() if k in email]

        if not current_domin:
            print("i dont know this is domain {}".format(email))

        current_domin, current_imap = current_domin[0]

        imap_session = ImapSession(email, password, current_imap, 993, current_domin)
        if out_dir is not None:
            if not os.path.isdir(out_dir):
                print(colored(f'Каталог для загрузки файлов не найден, создаём каталог {out_dir}', color='yellow'))
                os.mkdir(out_dir)
        else:
            print(colored('Не указан каталог для загрузки файлов!', color='red'))
            sys.exit(1)
        if imap_session.connect():
            try:
                if imap_session.move_to_folder('inbox'):
                    if imap_session.load_messages(unread=arguments.only_unread,
                                                  read=arguments.only_read,
                                                  to_date=to_date,
                                                  from_date=from_date,
                                                  from_addr=email_from):
                        if len(imap_session.ids_messages) > 1:
                            browser = start_browser(out_dir)
                            try:
                                for pos_i, message_id in enumerate(imap_session.ids_messages):
                                    log_file.write(f"----------------------- \n")
                                    log_file.write(f"{pos_i} \n")
                                    if imap_session.load_message(message_id):
                                        result = parse_message(imap_session.html_message,
                                                               imap_session.subject)
                                        for piece in result.keys():
                                            if result[piece] is None:
                                                print('Не удалось распарсить сообщение')
                                                break
                                        else:
                                            print(colored(f'{email}: сообщение обработано, данные: ', color='cyan'))
                                            print(f'''
              Номер заявления: {colored(result["application_number"], color="green")}, 
              дата регистрации: {colored(result["reg_date"], color="green")}, 
              ссылка для доступа: {colored(result["download_url"], color="green")}, 
              ключ для доступа: {colored(result["key"], color="green")}
                                                      ''')

                                            # processing  only exactly
                                            if rosreestr_number and not result[
                                                                            "application_number"] in rosreestr_number:
                                                continue
                                            new_file = parse_link(browser, result, email)
                                            if new_file is not None:
                                                print(colored(f'{email}: загружен файл {new_file}'))
                                                # log_file.write(
                                                #     f'{str(datetime.datetime.now())},{email},{result["application_number"]},{result["reg_date"]},{result["download_url"]},{result["key"]},{out_dir + "/" + new_file}\n')
                                                # log_file.flush()

                                            else:
                                                print(colored(f'{email}: не удалось обработать ссылку', color='red'))
                                                log_file.write(
                                                    f'не удалось загрузит файл {email}  {result["application_number"]} \n')
                                                log_file.flush()
                                else:
                                    print('===========END============')
                            except Exception as e:
                                print(e)
                            finally:
                                browser.quit()
            except Exception as e:
                print(e)
            finally:
                imap_session.logout()
                log_file.close()


if __name__ == '__main__':
    main()
