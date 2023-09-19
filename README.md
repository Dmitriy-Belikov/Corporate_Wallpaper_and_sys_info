# Corporate_Wallpaper_and_sys_info
Скрипт для установки watermark на изображение рабочего стола.
При первом запуске будет создан файл конфигурации config.ini
    [DEFAULT]
    local = local.jpg # Директория хранения фото на ПК
    auto = True #автосмена обоев вкл или выкл
    logo = OFS.JPG # Диретория хранения логотипа
    new_wallp = corp_wallpaper.jpg #Директория хранения нового файла рабочего стола
    config = config.ini
    
Для правильной идеального отображения лого, лого должо быть размером 804х193
В конфиге есть функция автоматической загрузки фото дня bing.
Для отключения необходимо изменить True на False.

При создании watermark указываются следующие данные:
Host Name
Serial Number
IP Adress
Machine Domain
Logon Domain
Logon Server
User Name
OS Version
Build Number
OS Build
SCCM Client Version
Edge Version
