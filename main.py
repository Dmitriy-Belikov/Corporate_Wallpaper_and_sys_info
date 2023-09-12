import time
import winreg
import platform
import socket
import os
import subprocess
import datetime
import struct
import ctypes
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from win32com.client import Dispatch
from win32api import GetSystemMetrics
from tqdm import tqdm
import bing_wallpaper


'''Функция создания конфига'''
def new_config():
    '''
    config.server #Директория хранения фото на сервере
    config.local # Директория хранения фото на ПК
    config.auto #автосмена обоев вкл или выкл
    config.logo # Диретория хранения логотипа
    config.new_wallp #Директория хранения нового файла рабочего стола
    '''

    with open('config.py', 'w') as f:
        f.write("server = 'C:/PS/server_wallp.jpg'\n") #Директория хранения фото на сервере
        f.write("local = 'C:/PS/local_wallp.jpg'\n") #Директория хранения фото на ПК
        f.write('auto = False\n') #автосмена обоев вкл или выкл
        f.write("logo = 'OFS.JPG'\n") # Диретория хранения логотипа
        f.write("new_wallp = 'C:/PS/corp_wallpaper.jpg'") #Директория хранения нового файла рабочего стола

'''Проверка автоматической смены изображений'''
def check_auto_wallpaper():
    print('Проверка автосмены')
    if config.auto is True:
        print('Автосмена включена')
        print('Скачиваем изображение')
        copy_server_to_pc_wallpaper()
        if os.path.exists(config.local):
            create_image(config.local)
        else:
            print('Не удалось скачать изображение')
            check_logo_weather()
    else:
        print('Автосмена выключена')
        check_logo_weather()
def wallpaper():
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    aKey = winreg.OpenKey(aReg, r"Control Panel\Desktop")
    keyname = winreg.QueryValueEx(aKey, 'WallPaper')[0]
    return keyname
#Проверка логотипа
def check_logo_weather(): #проверка логотипа посредством нахождения фото в папке
    print('Проверка логотипа')
    keyname = wallpaper()
    if keyname == config.new_wallp:
        print('Лого есть, ждем')
        #Здесь ссылка на функцию ожидания 2 часа
    else:
        print('Лого нет, создаем лого')
        create_image(keyname)
#Копирование картринки с сервера
def copy_server_to_pc_wallpaper():
    bing_wallpaper.download_wallpaper()

    '''
    print('Проверяю доступность сервера и папки на пк')
    if os.access(config.server, os.R_OK) is True and os.access(config.local, os.W_OK) is True:
        print('Копирую с сервера')
        shutil.copy(config.server, config.local)
    else:
        print('Доступ не получен. Вставляю лого в имеющуюся картинку')
        wallREG = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        awallREG = winreg.OpenKey(wallREG, r"Control Panel\Desktop")
        wallpaper_user = winreg.QueryValueEx(awallREG, 'WallPaper')[0].replace('\\\\', '\\')
        try:
            shutil.copy(wallpaper_user, config.local)
        except:
            print('файл не найден, копирую новую картинку с сервера')
            shutil.copy(config.server, config.local)'''
'''Запрос информации о ПК'''
def get_pc_info():
    print('Запрос информации о ПК')
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
    '''#Версия антивируса Crowdstrike'''
    try:
        ver_parser = Dispatch('Scripting.FileSystemObject')
        Crowdstrike_version = ver_parser.GetFileVersion('C:\Program Files\CrowdStrike\CSFalconContainer.exe')
    except:
        Crowdstrike_version = 'None'
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
    print('Создаем изображение')
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
    watermark = Image.open(config.logo)
    # определение соотношения сторон
    width, height = watermark.size
    new_height = int(new_width * height / width)
    watermark = watermark.resize((new_width, new_height))
    #Вставляем лого в вотермарк
    general_watermark.paste(watermark, (8, 80), mask=watermark.convert('RGBA'))
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
    new_wallpapper.save(config.new_wallp)
    #Удаляем временный файл
    if os.access(config.local, os.W_OK):
        os.remove(config.local)
    #new_wallpapper.show()

    changeBG()

def changeBG():
    print("Устанавливаем на рабочий стол")
    """Проверка разрядности системы"""
    bit64 = struct.calcsize('P') * 8 == 64
    SPI_SETDESKWALLPAPER = 20
    if bit64 is True:
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, config.new_wallp, 3)
    else:
        ctypes.windll.user32.SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, config.new_wallp, 3)
    print('Обои установлены')




if __name__ == "__main__":
    '''Проверка наличия конфигурационного файла'''
    if os.path.exists('config.py'):
        print('Файл конфигурации успешно загружен')
        import config
    else:
        new_config()
        time.sleep(2)
    import config

    while True:
        check_auto_wallpaper()
        print('Время до следующей смены обоев')
        for i in tqdm(range(3600)):
            time.sleep(1)