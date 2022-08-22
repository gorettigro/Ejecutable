import pywinauto
import win32clipboard
import json
import pywinauto
import win32clipboard
import json
import codecs
import psutil
import ctypes
import sys
import os
import shutil

from time import sleep, time
from pywinauto import keyboard
from pywinauto.application import Application
from concurrent.futures import process
from importlib.resources import path
from socket import timeout
from turtle import tilt, title
from unicodedata import name
from pathlib import Path

with codecs.open("data.json", "r", encoding="utf-8-sig") as file: 
    data = json.load(file)
print(data['repositorio'])

for (dirpath,names,filename) in os.walk(data['repositorio']):
    for f in filename:
        if f[-3:] == "txt":
            os.remove(f"{data['repositorio']}/{f}")

sleep(8)

PROCNAME = "SAEWIN80.exe"

pid = None

for proc in psutil.process_iter():
    if proc.name() == PROCNAME:
        pid=proc.pid

if not pid:
    print("No se ha encontrado el proceso")
    ctypes.windll.user32.MessageBoxW(0,"No se ha encontrado el proceso","Error",1)
    sys.exit(0)

app = pywinauto.Application(backend="win32").connect(process=pid)
main=app.window(title_re='.*Aspel.*')
lg=app.window(title='Abrir empresa')
    
if lg.exists(timeout=5):
    app['Abrir empresa']['Edit5'].type_keys(data['user_name'])
    app['Abrir empresa']['Edit7'].type_keys(data['password'])
    app['Abrir empresa']['Button3'].click()

if main.exists(timeout=5) and not lg.exists(timeout=5):
    main.set_focus()
else:
    for i in range(15):
        try:
            app.set_focus()
            keyboard.send_keys('%x')
            keyboard.send_keys('%a')

            if lg.exists(timeout=5):
                app['Abrir empresa']['Edit5'].type_keys(data['user_name'])
                app['Abrir empresa']['Edit7'].type_keys(data['password'])
                app['Abrir empresa']['Button3'].click()
            
            main.set_focus()
            break

        except Exception as e:
            pass
sleep(10)
keyboard.send_keys('%e')
sleep(10)
keyboard.send_keys('^%e')

main2=main.child_window(title_re=".*Administrador de Estadísticas.*")

if main2.exists(timeout=2):
    main2.set_focus
else:
    for i in range(30):
        try:
            app.set_focus()
            keyboard.send_keys('^%e')
            main2.set_focus()
            break
        except Exception as e:
            pass

sleep(10)
keyboard.send_keys('^c')

win32clipboard.OpenClipboard()
got = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()
grid = got.splitlines()
grid.pop(0)

for estad in data['name_stad']:
    index = 0
    flag = False
    for i in grid:
        if estad in i:
            flag = True
            break
        index += 1
    if not flag:
        print(f"Estadistic Not Found {estad}")
        continue

    print(f"Estadistic Found {estad}")

    keyboard.send_keys('{HOME}')

    for i in range(index):
        keyboard.send_keys('{DOWN}')

    keyboard.send_keys('{ENTER}')

    sleep(8)

    error = app.window(title="Error")
    
    if error.exists(timeout=3):
        app['Error']['Button'].click()
        continue


    sleep(10)
    
    ven=app.window(title="Estadísticas de ventas")
    if ven.exists:
        pass
    else:
        ven=app.window(title="Estadísticas de facturas")
        if ven.exists:
            pass
        else:
            ven=app.window(title="Estadísticas de productos")
            if ven.exists:
                pass
            else:
                print("Estadística no accesible")
                continue

    keyboard.send_keys('^e')

    export_info=app.window(title="Exportar información")

    if export_info.exists(timeout=2):
        app['Exportar información']['ComboBox2'].type_keys("%{DOWN}")
        app['Exportar información']['ComboBox2'].type_keys(data['formato'])
        app['Exportar información']['ComboBox2'].click()
        app['Exportar información']['Button6'].click()
        
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'], with_spaces=True)
        app['Exportar información']['Button3'].click()

    error2 = app.window(title="Error")
    if error2.exists(timeout=5):
        app['Error']['Button'].click()
        app['Exportar Información']['Edit7'].type_keys(data['repositorio_pre'], with_spaces=True)
        app['Exportar información']['Button3'].click()

    confi=app.window(title="Confirmación")
    if confi.exists(timeout=2):
        app['Confirmación']['Button1'].click()

    sleep(8)

    info=app.window(title="Información")
    if info.exists(timeout=2):
        app['Información']['Button'].click()
    
    if main2.exists(timeout=2):
        main2.set_focus
    else:
        for i in range(30):
            try:
                main.set_focus()
                keyboard.send_keys('^%e')
                main2.set_focus()
                break
            except Exception as e:
                pass

sleep(5)
   
keyboard.send_keys('%{F4}')

sleep(5)

confi2=app.window(title="Confirmación")
if confi2.exists(timeout=2):
        app['Confirmación']['Button1'].click()

sleep(3)

os.system("TASKKILL/F /IM notepad.exe")
