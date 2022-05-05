### Este es el archivo madre del proyecto 01
from modules.func_rick import *
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')
bienvenida()

ticket = input('Nombre del ticket: ').upper()
inicio = input('Fecha del periodo inicial (YYYY-MM-DD) : ')
fin = input('Fecha del periodo final (YYYY-MM-DD)  : ')

try :
    datetime.strptime(inicio, '%Y-%m-%d')
    datetime.strptime(fin, '%Y-%m-%d')
    df = data(ticket,s=inicio, f=fin,interval='1d')
except:
    print('Variables incorrectas')
    quit()

con_excel(df, ticket)

ajuste = input('Ajuste (N,M,A,UA): ')
bolling(df, ticket,ajuste)
