from importlib.resources import path
from socket import timeout
from turtle import tilt, title
from unicodedata import name
import pywinauto
import win32clipboard
import json
import os
from time import sleep, time
from pywinauto import application, keyboard

app = application.Application()
app.connect(path=r"C:\Program Files (x86)\Aspel\Aspel-SAE 8.0\SAEWIN80.exe")
main=app.window(title_re='Aspel-SAE')
lg=app.window(title='Abrir empresa')

with open("data.json", "r", encoding="utf-8") as file: 
    data = json.load(file)
    
if lg.exists(timeout=5):
    app['Abrir empresa']['Edit5'].type_keys(data['user_name'])
    app['Abrir empresa']['Edit7'].type_keys(data['password'])
    app['Abrir empresa']['Button3'].click()

main.set_focus()
keyboard.send_keys('%e')
sleep(2)
keyboard.send_keys('^%e')

main=main.child_window(title_re=".*Administrador.*")
#main.set_focus()
#keyboard.send_keys('^%e')

sleep(3)
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

    sleep(3)

    error = app.window(title="Error")
    
    if error.exists(timeout=3):
        app['Error']['Button'].click()
        continue
    
    ven=app.window(title="Estadísticas de ventas")
    if ven.exists:
        pass
    else:
        ven=app.window(title="Estadísticas de compras")
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
        app['Exportar información']['Edit3'].type_keys(estad)
        app['Exportar información']['Button6'].click()
        #app['Exportar Información']['Edit5'].type_keys(data['repositorio'])
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'])
        app['Exportar información']['Button3'].click()

    error2 = app.window(title="Error")
    
    if error2.exists(timeout=2):
        app['Error']['Button'].click()
        app['Exportar Información']['Edit7'].type_keys("C:/Users/auditor/Desktop")
        app['Exportar información']['Button3'].click()

    confi=app.window(title="Confirmación")
    if confi.exists(timeout=2):
        app['Confirmación']['Button1'].click()

    info=app.window(title="Información")
    if info.exists(timeout=2):
        app['Información']['Button'].click()