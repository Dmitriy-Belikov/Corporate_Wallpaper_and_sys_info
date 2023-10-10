import time
import winreg
import platform
import socket
import os
import subprocess
import datetime
import struct
import ctypes
import configparser
from PIL import Image, ImageDraw, ImageFont
from win32com.client import Dispatch
from win32api import GetSystemMetrics
from xml.dom import minidom
from urllib.request import urlopen, urlretrieve


'''Функция создания конфига'''
def new_config():
    with open(f'{os.getcwd()}\config.ini', 'w') as f:
        f.write('[DEFAULT]\n')
        f.write(f"local = {os.getcwd()}\local_wallp.jpg\n") #Директория хранения фото на ПК
        f.write('auto = True\n') #автосмена обоев вкл или выкл
        f.write("logo = OFS.JPG\n") # Диретория хранения логотипа
        f.write(f"new_wallp = {os.getcwd()}\corp_wallpaper.jpg\n") #Директория хранения нового файла рабочего стола
        f.write("config = config.ini")

'''Проверка автоматической смены изображений'''
def check_auto_wallpaper():
    #print('Проверка автосмены')
    if cauto == 'True':
        #print('Автосмена включена')
        #print('Скачиваем изображение')
        download_wallpaper()
        if os.path.exists(clocal):
            create_image(clocal)
        else:
            #print('Не удалось скачать изображение')
            check_logo_weather()
    else:
        #print('Автосмена выключена')
        check_logo_weather()
def wallpaper():
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    aKey = winreg.OpenKey(aReg, r"Control Panel\Desktop")
    keyname = winreg.QueryValueEx(aKey, 'WallPaper')[0]
    return keyname


#Проверка логотипа
def check_logo_weather(): #проверка логотипа посредством нахождения фото в папке
    #print('Проверка логотипа')
    keyname = wallpaper()
    if keyname == cnew_wallp:
        pass
        #print('Лого есть, ждем')
        #Здесь ссылка на функцию ожидания 2 часа
    else:
        #print('Лого нет, создаем лого')
        create_image(keyname)

#Копирование картринки с сервера
def join_path(*args):
    # Takes an list of values or multiple values and returns an valid path.
    if isinstance(args[0], list):
        path_list = args[0]
    else:
        path_list = args
    val = [str(v).strip(' ') for v in path_list]
    return os.path.normpath('/'.join(val))
'''
dir_path = os.path.dirname(os.path.realpath(__file__))
save_dir = join_path(dir_path, 'images')
if not os.path.exists(save_dir):
    os.makedirs(save_dir)'''
def download_wallpaper(idx=0, use_wallpaper=True):
    # Getting the XML File
    try:
        usock = urlopen(''.join(['http://www.bing.com/HPImageArchive.aspx?format=xml&idx=',
                                 str(idx), '&n=1&mkt=ru-RU']))  # ru-RU, because they always have 1920x1200 resolution
    except Exception as e:
        #print('Error while downloading #', idx, e)
        return
    try:
        xmldoc = minidom.parse(usock)
    # This is raised when there is trouble finding the image url.
    except Exception as e:
        #print('Error while processing XML index #', idx, e)
        return
    # Parsing the XML File
    for element in xmldoc.getElementsByTagName('url'):
        url = 'http://www.bing.com' + element.firstChild.nodeValue
        # Get Current Date as fileName for the downloaded Picture
        now = datetime.datetime.now()
        date = now - datetime.timedelta(days=int(idx))
        pic_path = join_path(clocal)

        '''if os.path.isfile(pic_path):
            print('Image of', date.strftime('%d-%m-%Y'), 'already downloaded.')
            if use_wallpaper:
                set_wallpaper(pic_path)
            return
        print('Downloading: ', date.strftime('%d-%m-%Y'), 'index #', idx)'''

        # Download and Save the Picture
        # Get a higher resolution by replacing the file name
        urlretrieve(url.replace('_1366x768', '_1920x1200'), pic_path)
        # Set Wallpaper if wanted by user

'''Запрос информации о ПК'''
def get_pc_info():
    #print('Запрос информации о ПК')
    '''Список для параметров'''
    param = []
    '''Запросы из регистра, версии программ и скрипты запроса нужной информации'''
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
    '''Версия SCCM'''
    try:
        aSCCM_version = winreg.OpenKey(aReg, r'SOFTWARE\Microsoft\SMS\Mobile Client')
        SCCM_version = winreg.QueryValueEx(aSCCM_version, 'SmsClientVersion')[0]
    except:
        SCCM_version = 'None'
    '''#Номер сборки'''
    Build_number = winreg.QueryValueEx(aKey, 'CurrentBuildNumber')[0] #Build Number
    '''#Версия системы'''
    OS_Build = winreg.QueryValueEx(aKey, 'DisplayVersion')[0] #OS Build
    '''#Имя ПК'''
    Host_Name = platform.node()
    '''#Сервис таг'''
    Mother_name = subprocess.check_output("wmic bios get serialnumber").decode().split()[1]
    '''#Узнаем Uptime PC'''
    time_start = subprocess.check_output('powershell Get-CimInstance Win32_OperatingSystem | select LastBootUpTime').decode().split('--------------')[1]
    time_start = time_start.replace("\r", '').replace("\n",'').replace(' ', '')
    time_start = datetime.datetime.strptime(str(time_start), '%d.%m.%Y%H:%M:%S')
    Uptime = datetime.datetime.now() - time_start
    Uptime = str(Uptime)[:str(Uptime).find('.')]
    ver_parser = Dispatch('Scripting.FileSystemObject')
    '''#Версия антивируса Crowdstrike
    try:
        Crowdstrike_version = ver_parser.GetFileVersion('C:\Program Files\CrowdStrike\CSFalconContainer.exe')
    except:
        Crowdstrike_version = 'None'
        '''
    '''#Версия MS Edge '''
    try:
        Edge_version = ver_parser.GetFileVersion('C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    except:
        Edge_version = 'None'
    '''#Версия ОС'''
    OS_Version = platform.platform(terse=True)
    '''#Ip адрес'''
    try:
        IP_adress = socket.gethostbyname(Host_Name)
    except:
        IP_adress = 'None'
    '''#Домен в котором состоит машина'''
    try:
        Machine_Domain = socket.getfqdn().split('.', 1)[1]
        if Machine_Domain == 'ent.techofs.com':
            Machine_Domain = 'TECHOFS'
        Logon_Domain = os.environ['userdomain']
    except:
        Logon_Domain = 'None'
        Machine_Domain = 'None'
    '''#Домен к которому подключен ПК'''
    try:
        Logon_Server = subprocess.check_output('nltest /dsgetdc:ent.techofs.com').decode('cp866').split()[2].split('.')[0].replace('\\', '')
    except:
        Logon_Server = 'None'
    '''#имя пользователя'''
    User_name = subprocess.check_output('whoami').decode('cp866').replace('\n', '')
    symb = User_name.find('\\')
    User_name = User_name[symb+1:]
    '''#Собираем список параметров полученных выше и передаем в функцию создания watermark'''
    param.extend([Host_Name, Mother_name, IP_adress, Machine_Domain,Logon_Domain, Logon_Server, User_name, OS_Version, Build_number, OS_Build, SCCM_version,Edge_version, Uptime])
    return param
'''Создание картинки с Watermark'''
def create_image(dir_walpp):
    if os.path.exists(dir_walpp):
        dir_walpp = dir_walpp
    else:
        dir_walpp = "c:\windows\web\wallpaper\windows\img0.jpg"
    #print('Создаем изображение')
    monitor_width = GetSystemMetrics(0)
    general_watermark = Image.new('RGBA', (300, 370))
    draw_text = ImageDraw.Draw(general_watermark)
    name_param = ['Host Name', 'Serial number', 'IP Adress', 'Machine Domain', 'Logon Domain', 'Logon Server', 'User Name', 'OS Version', 'Build Number', 'OS Build', 'SCCM Client Version', 'Edge Version', 'Uptime']
    value_param = get_pc_info()
    font = ImageFont.truetype('arial.ttf', size=15)
    coordy = 170
    i = 0
    for item_param in name_param:
        draw_text.text((10, coordy), item_param, font=font, stroke_width=1, stroke_fill="black")
        draw_text.text((295, coordy + 13), value_param[i], font=font, anchor='rs', stroke_width=1, stroke_fill="black")
        coordy += 15
        i += 1
    #подгоняем логотип к размеру watermark
    new_width = 290
    try:
        watermark = Image.open(clogo)
        #определение соотношения сторон
        width, height = watermark.size
        new_height = int(new_width * height / width)
        watermark = watermark.resize((new_width, new_height))
    except:
        pass

    #Вставляем лого в вотермарк
    #print(dir_walpp)
    try:
        general_watermark.paste(watermark, (8, 80), mask=watermark.convert('RGBA'))
    except:
        pass
    new_wallpapper = Image.open(dir_walpp)
    #Меняем размер изображения под размер монитора
    width_percent = (monitor_width / float(new_wallpapper.size[0]))
    height_size = int((float(new_wallpapper.size[1]) * float(width_percent)))
    new_wallpapper = new_wallpapper.resize((monitor_width, height_size))
    new_wallpapper.convert('RGBA')
    #Вычисление установки логотипа
    position = (new_wallpapper.width - general_watermark.width - 50,
                new_wallpapper.height - general_watermark.height - 50)
    #Вставляем лого
    new_wallpapper.paste(general_watermark, position, mask=general_watermark.convert('RGBA'))
    #Сохраняем изображение
    new_wallpapper.save(cnew_wallp)
    #Удаляем временный файл
    if os.access(clocal, os.W_OK):
        os.remove(clocal)
    #new_wallpapper.show()

    changeBG()

def changeBG():
    #print("Устанавливаем на рабочий стол")
    """Проверка разрядности системы"""
    bit64 = struct.calcsize('P') * 8 == 64
    SPI_SETDESKWALLPAPER = 20
    corp_wllp = os.path.abspath(cnew_wallp)
    if bit64 is True:
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, corp_wllp, 3)
    else:
        ctypes.windll.user32.SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, corp_wllp, 3)
    #print('Обои установлены')

def read_conf():
    homepath = os.getenv('USERPROFILE')
    config = configparser.ConfigParser()
    conf_file = f'{os.getcwd()}\config.ini'
    config.read(conf_file)
    clocal = config['DEFAULT']['local']
    cauto = config['DEFAULT']['auto']
    clogo = config['DEFAULT']['logo']
    cnew_wallp = config['DEFAULT']['new_wallp']
    cconfig = config['DEFAULT']['config']
    return clocal, cauto, clogo, cnew_wallp, cconfig


if __name__ == "__main__":
    homepath = os.getenv('USERPROFILE')
    '''Проверка наличия конфигурационного файла'''
    if os.path.exists(f'{os.getcwd()}\config.ini'):
        pass
        #print('Файл конфигурации успешно загружен')
    else:
        new_config()
        time.sleep(2)
    clocal, cauto, clogo, cnew_wallp, cconfig = read_conf()
    while True:
        check_auto_wallpaper()
        time.sleep(3600)

