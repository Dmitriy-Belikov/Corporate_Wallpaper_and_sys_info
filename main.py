import winreg
import shutil
from win32com.client import Dispatch
import platform
import socket
import os
import subprocess
import csv
from PIL import Image, ImageDraw, ImageFont, ImageFilter

dir = []
with open('config.txt', 'r', newline='') as file:
    reader = csv.reader(file)
    for row in reader:
        dir.append(row[0])

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
    #Logon_Server = subprocess.run("gpresult /r", encoding='utf-8')
    Logon_Server = subprocess.check_output('nltest /dsgetdc:wft.root.loc').decode('cp866').split()[2].split('.')[0].replace('\\', '')
    User_name = subprocess.check_output('whoami').decode('cp866').replace('\n','')
    if User_name[0:3] == 'wft':
        User_name = User_name[4:]

    param.extend([Host_Name, Mother_name, IP_adress, Machine_Domain,Logon_Domain, Logon_Server, User_name[:-1], OS_Version, Build_number, OS_Build, SCCM_version,Edge_version, Crowdstrike_version])
    return param

def create_image():
    general_watermark = Image.new('RGB', (300, 440))
    draw_text = ImageDraw.Draw(general_watermark)
    name_param = ['Host Name', 'Serial number', 'IP Adress', 'Machine Domain', 'Logon Domain', 'Logon Server', 'User Name', 'OS Version', 'Build Number', 'OS Build', 'SCCM Client Version', 'Edge Version', 'Crowdstrike Version']
    value_param = get_pc_info()
    font = ImageFont.truetype('arial.ttf', size=13)
    coordy = 170
    i = 0
    for item_param in name_param:
        draw_text.text((10, coordy), item_param, font=font)
        draw_text.text((290, coordy + 13), value_param[i], font=font, anchor='rs')
        coordy += 20
        i += 1

    watermark = Image.open('C:/Intel/logo.jpg')

    general_watermark.paste(watermark, (10, 10))
    #general_watermark.show()

    general_watermark.putalpha(128)
    new_wallpapper = Image.open(new_dir)
    position = (new_wallpapper.width - general_watermark.width,
                new_wallpapper.height - general_watermark.height)

    new_wallpapper.alpha_composite(general_watermark, position)
    new_wallpapper.show()


create_image()