# System env
# OS:Win10 21H1 x64
# Python: 3.8 form conda

# Init Packages
import os
import sys
from string import ascii_lowercase
from zipfile import ZipFile

import requests
import xmltodict
from win32com import client as wincom_client


# find chrome install path
working_os = sys.platform
chrome_path = None

# find Chrome Path
if working_os == 'win32':
    root_paths = list(item+":" for item in ascii_lowercase)
    for root_path in root_paths:
        for path, _, files in os.walk(root_path):
            if 'chrome.exe' in files:
                chrome_path = path + r'\chrome.exe'
                break
    if chrome_path:
        wincom_obj = wincom_client.Dispatch('Scripting.FileSystemObject')
        chrome_version = wincom_obj.GetFileVersion(chrome_path)
        print(f'Chrome Path: {chrome_path} with Version {chrome_version}')
    else:
        print('Please check Chrome installed!!')

elif working_os == 'darwin':
    pass

#
# Get chromedriver list:
chromedriver_api_url = 'https://chromedriver.storage.googleapis.com/'
api_res = requests.get(url=chromedriver_api_url)
api_res.encoding = 'utf-8'


if api_res.status_code == 200:
    api_content = xmltodict.parse(api_res.text)
    for key in api_content.keys():
        api_content = api_content[key]['Contents']
    chrome_version = chrome_version[0:8]
    all_files = list()
    for item in api_content:
        if chrome_version in item['Key']:
            all_files.append(item['Key'])
    if working_os == 'win32':
        for item in all_files:
            if 'win' in item:
                file_url = chromedriver_api_url+f'{item}'
                file_res = requests.get(url=file_url)
                project_path = os.path.abspath(os.path.dirname(__name__))
                project_path = os.path.join(project_path, 'chromedriver.zip')
                with open(project_path, 'wb') as f:
                    f.write(file_res.content)
    else:
        pass
else:
    'Chromedriver api Error!!'


# unzip
with ZipFile(project_path, 'r') as zf:
    zf.extractall()