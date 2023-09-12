import datetime
from urllib.request import urlopen, urlretrieve
from xml.dom import minidom
import os
import config
def join_path(*args):
    # Takes an list of values or multiple values and returns an valid path.
    if isinstance(args[0], list):
        path_list = args[0]
    else:
        path_list = args
    val = [str(v).strip(' ') for v in path_list]
    return os.path.normpath('/'.join(val))

dir_path = os.path.dirname(os.path.realpath(__file__))
save_dir = join_path(dir_path, 'images')
if not os.path.exists(save_dir):
    os.makedirs(save_dir)

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
        pic_path = join_path(config.local)

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

if __name__ == "__main__":
    download_wallpaper()

