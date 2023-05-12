import winreg
import shutil
from win32com.client import Dispatch
import platform
import socket
import os
import subprocess
from PIL import Image, ImageDraw, ImageFont
dir = []
with open('config.txt') as file:
    dir = file.read().split('\n')
new_dir = dir[1] # Директория хранения фото на ПК
auto = dir[2] #автосмена обоев вкл или выкл
dir = dir[0] #Директория хранения фото на сервере


#Проверка автоматической смены изображений
def check_auto_wallpaper():
    if auto is True:
        pass
        #Проверка наличия нового изображения на сервере
        #ЕСЛИ ЕСТЬ НОВОЕ
        #Затем копирование нового фото на комп
        #Отрисовка логотипа
        #Установка обоев

        #ЕСЛИ НЕТ НОВОГО
        #проверяем логотип
        #check_logo_weatherford
    else:
        check_logo_weather()
        pass
#Проверка логотипа
def check_logo_weather(): #проверка логотипа посредством нахождения фото в папке
    aReg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    aKey = winreg.OpenKey(aReg, r"Control Panel\Desktop")
    keyname = winreg.QueryValueEx(aKey, 'WallPaper')[0]
    if keyname == new_dir:
        pass
        #Здесь ссылка на функцию ожидания 2 часа
    else:
        pass
        #Здесь функция установки логотипа
    print(keyname)
#Копирование картринки с сервера
def copy_server_to_pc_wallpaper():
    shutil.copy(dir, new_dir)
def get_pc_info():
    param = []

    aReg = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
    aSCCM_version = winreg.OpenKey(aReg, r'SOFTWARE\Microsoft\SMS\Mobile Client')
    SCCM_version = winreg.QueryValueEx(aSCCM_version, 'SmsClientVersion')[0]
    Build_number = winreg.QueryValueEx(aKey, 'CurrentBuildNumber')[0] #Build Number
    OS_Build = winreg.QueryValueEx(aKey, 'DisplayVersion')[0] #OS Build
    Host_Name = platform.node()
    Mother_name = subprocess.check_output("wmic bios get serialnumber").decode().split()[1]

    ver_parser = Dispatch('Scripting.FileSystemObject')
    Crowdstrike_version = ver_parser.GetFileVersion('C:\Program Files\CrowdStrike\CSFalconContainer.exe')
    Edge_version = ver_parser.GetFileVersion('C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    OS_Version = platform.platform(terse=True)
    IP_adress = socket.gethostbyname(Host_Name)
    Machine_Domain = socket.getfqdn().split('.', 1)[1]
    if Machine_Domain == 'wft.root.loc':
        Machine_Domain = 'WFT'
    Logon_Domain = os.environ['userdomain']
    Logon_Server = None
    User_name = None

    print(Host_Name)
    print(Mother_name)
    print(IP_adress)
    print(Machine_Domain)
    print(Logon_Domain)
    print(Logon_Server)
    print(User_name)
    print(OS_Version)
    print(Build_number)
    print(OS_Build)
    print(SCCM_version)
    print(Edge_version)
    print(Crowdstrike_version)


def create_image():
    im = Image.new('RGB', (300, 400))
    draw_text = ImageDraw.Draw(im)
    name_param = ['Host Name', 'IP Adress', 'Machine Domain', 'Logon Domain', 'Logon Server', 'User Name', 'OS Version', 'Build Number', 'OS Build', 'SCCM Client Version', 'Edge Version', 'Crowdstrike Version']
    font = ImageFont.truetype('arial.ttf', size=13)
    coordy = 100
    for item_param in name_param:
        draw_text.text((10, coordy), item_param, font=font)
        draw_text.text((290, coordy+13), item_param, font=font, anchor='rs')
        coordy += 20

    im.show()


get_pc_info()