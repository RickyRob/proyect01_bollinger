### Este es el archivo madre del proyecto 01
from modules.func_rick import *
import pandas as pd

ticket = input('Nombre del ticket: ')
inicio = input('Fecha del periodo inicial : ')
fin = input('Fecha del periodo final : ')
intervalo= input('Intervalo de datos: ')

df = data(ticket,s=inicio, f=fin,interval=intervalo)
ln_rend(df)
con_excel(df, ticket) 

#print(df.head())