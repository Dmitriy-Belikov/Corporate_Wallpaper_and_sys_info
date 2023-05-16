import winreg
import shutil
import platform
import socket
import os
import subprocess

import filecmp
import struct
import ctypes
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from win32com.client import Dispatch

'''config.local # Директория хранения фото на ПК
config.auto #автосмена обоев вкл или выкл
config.logo # Диретория хранения логотипа
config.server #Директория хранения фото на сервере
config.new_wallp #Директория хранения нового файла рабочего стола'''

'''Создание конфигурационного файла'''
if os.path.exists('config.py'):
    print('Файл конфигурации успешно загружен')
    pass
else:
    with open('config.py', 'a') as f:
        f.write("server = 'C:/CSV/wallpaper.jpg'\n")
        f.write("local = 'C:/CSV/new/wallpaper.jpg'\n")
        f.write('auto = True\n')
        f.write("logo = 'C:/CSV/logo.png'\n")
        f.write("new_wallp = 'C:/CSV/new/corp_wallpaper.jpg'")

import config


'''Проверка автоматической смены изображений'''
def check_auto_wallpaper():
    print('Проверка автосмены')
    if config.auto is True:
        if os.path.isfile(config.server) is True:
            try:
                print('Сравнение фото на сервере и на пк')
                filecmp.cmp(config.server, config.local)
                print('Файлы одинаковые, замена не требуется')
            except:
                print('Копирую с сервера')
                copy_server_to_pc_wallpaper()
                print('Создаю обои')
                create_image(config.local)
                pass
        else:
            pass
        #Установка обоев
        #Запуск функции ожидания 2 часа
    else:
        print('Проверка лого')
        check_logo_weather()
#Проверка логотипа
def check_logo_weather(): #проверка логотипа посредством нахождения фото в папке
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    aKey = winreg.OpenKey(aReg, r"Control Panel\Desktop")
    keyname = winreg.QueryValueEx(aKey, 'WallPaper')[0]
    if keyname == config.new_wallp.replace('/', '\\'):
        print('Лого есть, ждем')
        pass
        #Здесь ссылка на функцию ожидания 2 часа
    else:
        print('Лого нет, создаем лого')
        create_image(keyname)
#Копирование картринки с сервера
def copy_server_to_pc_wallpaper():
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
            shutil.copy(config.server, config.local)
'''Запрос информации о ПК'''
def get_pc_info():
    print('Запрос информации о ПК')
    '''Список для параметров'''
    param = []
    '''Запросы из регистра, версии программ и скрипты запроса нужной информации'''
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
    try:
        aSCCM_version = winreg.OpenKey(aReg, r'SOFTWARE\Microsoft\SMS\Mobile Client')
        SCCM_version = winreg.QueryValueEx(aSCCM_version, 'SmsClientVersion')[0]
    except:
        SCCM_version = 'None'
    Build_number = winreg.QueryValueEx(aKey, 'CurrentBuildNumber')[0] #Build Number
    OS_Build = winreg.QueryValueEx(aKey, 'DisplayVersion')[0] #OS Build
    Host_Name = platform.node()
    Mother_name = subprocess.check_output("wmic bios get serialnumber").decode().split()[1]
    try:
        ver_parser = Dispatch('Scripting.FileSystemObject')
        Crowdstrike_version = ver_parser.GetFileVersion('C:\Program Files\CrowdStrike\CSFalconContainer.exe')
    except:
        Crowdstrike_version = 'None'
    try:
        Edge_version = ver_parser.GetFileVersion('C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    except:
        Edge_version = 'None'
    OS_Version = platform.platform(terse=True)
    try:
        IP_adress = socket.gethostbyname(Host_Name)
    except:
        IP_adress = 'None'
    try:
        Machine_Domain = socket.getfqdn().split('.', 1)[1]
        if Machine_Domain == 'wft.root.loc':
            Machine_Domain = 'WFT'
        Logon_Domain = os.environ['userdomain']
    except:
        Machine_Domain = 'None'
    try:
        Logon_Server = subprocess.check_output('nltest /dsgetdc:wft.root.loc').decode('cp866').split()[2].split('.')[0].replace('\\', '')
    except:
        Logon_Server = 'None'
    User_name = subprocess.check_output('whoami').decode('cp866').replace('\n', '')
    if User_name[0:3] == 'wft':
        User_name = User_name[4:]
        if User_name[0] == 'e':
            User_name = User_name[0:7]
    param.extend([Host_Name, Mother_name, IP_adress, Machine_Domain,Logon_Domain, Logon_Server, User_name, OS_Version, Build_number, OS_Build, SCCM_version,Edge_version, Crowdstrike_version])
    return param
'''Создание картинки с Watermark'''
def create_image(dir_walpp):
    print('Создаем изображение')
    general_watermark = Image.new('RGBA', (300, 370))
    draw_text = ImageDraw.Draw(general_watermark)
    name_param = ['Host Name', 'Serial number', 'IP Adress', 'Machine Domain', 'Logon Domain', 'Logon Server', 'User Name', 'OS Version', 'Build Number', 'OS Build', 'SCCM Client Version', 'Edge Version', 'Crowdstrike Version']
    value_param = get_pc_info()
    font = ImageFont.truetype('arial.ttf', size=15)
    coordy = 170
    i = 0
    for item_param in name_param:
        draw_text.text((10, coordy), item_param, font=font, stroke_width=1, stroke_fill="black")
        draw_text.text((290, coordy + 13), value_param[i], font=font, anchor='rs', stroke_width=1, stroke_fill="black")
        coordy += 15
        i += 1
    watermark = Image.open(config.logo)
    general_watermark.paste(watermark, (8, 8), mask=watermark.convert('RGBA'))
    copy_server_to_pc_wallpaper()
    new_wallpapper = Image.open(dir_walpp)
    new_wallpapper.convert('RGBA')
    position = (new_wallpapper.width - general_watermark.width,
                new_wallpapper.height - general_watermark.height)
    new_wallpapper.paste(general_watermark, position, mask=general_watermark.convert('RGBA'))
    new_wallpapper.save(config.new_wallp)
    os.remove(config.local)
    new_wallpapper.show()
    #changeBG()

'''Функция установки обоев ломает Bing Wallpaper
Нужно исправить'''
def changeBG():
    """Change background depending on bit size"""
    bit64 = struct.calcsize('P') * 8 == 64
    SPI_SETDESKWALLPAPER = 20
    if bit64 is True:
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, config.new_wallp, 3)
    else:
        ctypes.windll.user32.SystemParametersInfoA(SPI_SETDESKWALLPAPER, 0, config.new_wallp, 3)
    print('Обои установлены')

check_auto_wallpaper()