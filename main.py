import winreg
import shutil
from PIL import Image, ImageDraw
dir = []
with open('config.txt') as file:
    dir = file.read().split('\n')
new_dir = dir[1]
auto = dir[2]
dir = dir[0]

print(new_dir)
print(auto)
print(dir)

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

def create_image():
    im = Image.new('RGB', (300, 400))
    draw_text = ImageDraw.Draw(im)
    param = ['Host Name', 'IP Adress', 'Machine Domain', 'Logon Domain', 'Logon Server', 'User Name', 'OS Version', 'Build Number', 'OS Build', 'SCCM Client Version', 'Edge Version', 'Crowdstrike Version']
    coordy = 100
    for i in param:
        draw_text.text((10, coordy), i)
        coordy += 20

    im.show()
