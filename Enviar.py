import pandas
import pyautogui
import time
import webbrowser

Escribir = 621, 691

def Abrir(pos,click=1):
    pyautogui.moveTo(pos)
    pyautogui.click(clicks=click)

print("Bienvanide")
time.sleep(2)
#variables de exel
Exel = r'\Enviar.xlsx'
df = pandas.read_excel(Exel, sheet_name='Hoja1')
url = 'https://web.whatsapp.com/send?phone=+54'
print("Se cargaron los datos de Exel")
print ("Abriendo el Navegador")
for i in df.index:
    numero = str(df['Numero'][i])
    mensaje = df['Mensaje'][i]
    nombre = str(df['Nombre'][i])
    orden = str(df['Orden'][i])
    hora = str(df['Hora'][i])
    fecha = str(df['Fecha'][i])
    obser= df['Observ'][i]
    priv = ("Hola "+nombre+"! Orden N°: "+orden+". "+mensaje+ " "+fecha+" a las "+hora+". "+obser )
    webbrowser.open(url+numero+"&text=")
    time.sleep(20)
    Abrir(Escribir,click=1)
    time.sleep(2)
    pyautogui.typewrite(priv)
    time.sleep(2)
    pyautogui.press('enter')
    print("Enviado a "+numero+" "+priv)
    time.sleep (6)
    pyautogui.hotkey('Ctrl','w')
    print("Terminado")
    pass
