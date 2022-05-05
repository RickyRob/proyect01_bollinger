import pandas as pd
import yfinance as yf
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt
from datetime import timedelta

# Funcion de bienvenida

def bienvenida():
    print('\n')
    print('####################################################')
    print('################# RICKY INVESTING ##################')
    print('####################################################')
    print('\n')

### Esta funcion importa los datos de Yahoo Finance y retorna un dataframe limpio con solo los
#Datos de cierre ajustado
def data(ticket:str,s:str,f:str,interval:str):
    df = yf.download(ticket,start=s,end=f,interval = interval)
    df.drop(['Open','High','Low','Close','Volume'], axis = 1, inplace=True)
    df.dropna(inplace=True)
    return df

# Esta función regresa el rendimiento logaritmico
def ln_rend(df):
    df['R_LN'] =  np.log(df['Adj Close']/df['Adj Close'].shift(periods=1))

# Esta funcion abre la aplicacion Excel y crea un Libro nuevo
def con_excel(df, ticket:str):
    wb = xw.Book()
    sheet_data = wb.sheets[0]
    sheet_data
    xw.sheets.active
    xw.books.active
    sheet_data.name = ticket
    sheet_data.range('A1').value = df
    return wb

# Esta función es la más importante recibe dos datos de entrada:
"""
1. Recibe el dataframe generado en la funcion data
2. Recibe el nombre del ticket que se le pregunta al usuario por consola
3. Recibe la bondad de ajuste del calculo

"""
def bolling(df, ticket, ajuste):
    wb = xw.books.active # Trabaja sobre el Libro creado
    sheet_estr1 = wb.sheets.add() # Agrega una hoja nueva al Libro creado y activo
    sheet_estr2 = wb.sheets.add() # Agrega una hoja nueva al Libro creado y activo
    sheet_estr2 = wb.sheets[0] # Se le asigna la hoja por posición
    sheet_estr1 = wb.sheets[1] # Se le asigna la hoja por posición
    sheet_estr1.name = 'Bandas_Boll' # Se le asignamnombre a la hoja por posición aqui viviran los resultados de las bandas y medias moviles de 10
    sheet_estr2.name = 'Resultados' # Se le asigna nombre a la hoja por posición aqui viviran los resultados de compras y ventas

    """
    Creación de las bandas de volatilidad por encima y por debajo
    """
    df['2+std'] = df['Adj Close'].rolling(20).std()+df['Adj Close']+df['Adj Close'].rolling(10).std() 
    df['2-std'] = -df['Adj Close'].rolling(20).std()+df['Adj Close']-df['Adj Close'].rolling(10).std()

    """
    Calculo de la media movil a 10 datos
    """
    df['Rol_SWM10'] = df['Adj Close'].rolling(10).mean()

    """
    En el dataframe original se agregaran dos columnas:
        1. df['Venta'] tendrá la relación "voltailidad/media movil" un valor
        de 1 quiere decir que la media movil a tocado a la volatilidad, un valor por debajo de 1
        quiere decir que esta por tocar la linea de volatilidad

        2. df['Compra'] tendrá la relación "media movil/ Volatilidad por de bajo" un valor
        de 1 quiere decir que la media movil a tocado a la volatilidad, un valor por encima de 1
        quiere decir que esta por tocar la linea de volatilidad

        Todo el dataframe se guardará en la hoja del Libro llamada 'Bandas_Boll'
    """
    df['Venta'] = np.round(df['2+std']/df['Rol_SWM10'],2) 
    df['Compra'] = np.round(df['Rol_SWM10']/df['2-std'],2)
    sheet_estr1.range('A1').value = df


    fig = plt.figure(figsize=(10,6)) # Se instancia el objeto figura de matplotlib

    """
    Se crean dos dataframes nuevos:
        1. df_ventas : Tendrá los registros de venta con los criterios que se muestrán en el código
        todo valor contenido en la columna df['Venta'] que este entre 0.99 y 1.00 se alamcenará en el dataframe
        su registro. De manera similar para df_compra y la columna df['Compra'].
    """
    df_ventas = df[df['Venta'].isin([0.99,1.00])]
    #df_compra = df[df['Compra'].isin([0.99,1.00])]
    df_compra = df[df['Compra'].isin([1.00,1.01])]


    ### Agregando una columna bandera de color según es compra o venta a los dataframes nuevos
    df_ventas['Color']='g'
    df_compra['Color']='r'

    ### Haciendo negativos los valores de compra ya que es dinero desembolsado
    df_compra['Compra']=df_compra['Compra'].apply(lambda x : x * -1)
    df_compra['Adj Close']=df_compra['Adj Close'].apply(lambda x : x * -1)

    ### Nuevo dataframe de apoyo uniendo las compras y ventas. Ordenando los indices.
    df_union = pd.concat([df_compra,df_ventas], axis=0)
    df_union = df_union.sort_index()
    
    ## Lista para almacenar un campo de 1 y -1 las filas con -1 se eliminarán ya que quiere decir que no alterna la compra y venta
    lista = []
    x = 0
    while x < (len(df_union['Color'])-1): # Este ciclo durará un registro menos en el dataframe
        for i in range(len(df_union['Color'])-1): # Este ciclo durará un registro menos en el dataframe
             if df_union.iloc[i]['Color'] != df_union.iloc[i-1]['Color']: # Si en el registro 'i' con valor en la columna Color 
                # es diferente al color del registro anterior asignar un valor de 1 a la lista vacia 
                # llamada lista y -1 en el caso contrario
                 lista.append(1)
             else:
                 lista.append(-1)
             x += 1
    # Este append agrega el valor faltante para tener las mismas dimensiones que el df original union
    lista.append(-1)
    # Uniendo la columna bandera de -1 y 1
    df_union = df_union.assign(Validator=lista)
    # Este dataframe es el que se usará para alternar las compras y ventas todos los valores 1 alternan entre compras y ventas
    # De esta manera se garantiza la primer señal de compra o venta de manera alternada
    # El nuevo dataframe con esta condición es df_simple
    df_simple = df_union[df_union['Validator']==1]

    # El algoritmo debe iniciar con una compra y esta linea garantiza esta situación
    if df_simple.iloc[0]['Adj Close'] > 0:
        df_simple.drop(df_simple.index.tolist()[0],inplace=True)
    
    # Este bloque especifica el nivel de ajuste en días indice que se deberan restar a la señal original
    # Si el usuario teclea 'N' no se recorren los dias de ajuste, si el usuario teclea 'M' se restan 5 dias a la señal original
    # Esto hace más responsiva la señal con una aticipación
    rol = 0

    if ajuste == 'N':
        rol += 0
    elif ajuste == 'M':
        rol += 5
    elif ajuste == 'A':
        rol += 10
    elif ajuste == 'UA':
        rol += 15
    else:
        rol += 25

    # Con este bloque extraemos el indice de datos ajustados segun lo especifique el usuario
    # Ademas garantizamos que exista el indice 
    lista_aux1 = [] # En esta lista se agregaran los indices ajustados
    n = rol
    for r in df_simple.index.tolist():
        f = r-timedelta(n)
        if f in df.index.tolist():
            lista_aux1.append(f)
        elif (f - timedelta(1)) in df.index.tolist():
            f = f - timedelta(1)
            lista_aux1.append(f)
        elif (f - timedelta(2)) in df.index.tolist():
            f = f - timedelta(2)
            lista_aux1.append(f)
        else:
            f = f - timedelta(3)
            lista_aux1.append(f)
    
    df_opt = df.loc[lista_aux1] # Este será el dataframe optimizado
    df_opt['Colores'] = df_simple['Color'].tolist() # Se asignan los colores al dataframe optimizado extrayendolos del original
    # Los colores seran 'r' = rojo indica compra, 'v' = verde indica venta

    # Los valores ajustados de cierre se multiplicaran por -1 donde la columna color sea 'r' que indica compra
    df_opt['Adj Close'] = np.where(df_opt['Colores']=='r',df_opt['Adj Close']*-1,df_opt['Adj Close'])

    # print('############# Estrategía sin Ajuste ####################')
    # print(df_simple)
    # print('############# Estrategía con Ajuste ####################')
    # print(df_opt)

    # Se grafican los puntos de compra y venta con su color indicado
    plt.scatter(df_opt.index, df_opt['Rol_SWM10'],color=df_opt['Colores'],marker="v", s=10)

    # Si el ultimo registro es de compra no se toma en cuenta para el calculo del rendimietno pero si en el grafico
    if df_opt.iloc[-1]['Adj Close'] < 0:
        df_opt.drop(df_opt.index.tolist()[-1],inplace=True)

    # Se suman los valores de compra y venta y se asignan en estas variables
    salidas = -1*(df_opt[df_opt['Adj Close']< 0]['Adj Close'].sum())
    entradas = 1*(df_opt[df_opt['Adj Close']> 0]['Adj Close'].sum())
    
    print(f'Salidas de efectivo: {np.round(salidas,2)}')
    
    print(f'Entradas de efectivo: {np.round(entradas,2)}')

    # Este es el rendimiento en importe
    rend= np.round(df_opt['Adj Close'].sum(),2)
    print(f'Rendimiento de la estrategía: {rend}')

    # Este es el rendimiento en procentaje
    rend_por = np.round(rend/salidas,2)
    print(f'Rendimiento de la estrategía: {rend_por}')

    # En esta hoja del libro se pegan los valores del dataframe optimizado
    sheet_estr2.range('A1').value = df_opt

    # Se grafican las volatilidades, precios ajustados y media movil
    plt.plot(df.index,df['2+std'],'r--',label='2+std', linewidth=0.35)
    plt.plot(df.index,df['Adj Close'],'y--',label='Precio', linewidth=0.35)
    plt.plot(df.index,df['2-std'],'g--',label='2-std', linewidth=0.35)
    plt.plot(df.index,df['Rol_SWM10'],color='b', linewidth=0.6)
    plt.title(f'Bandas de Bollinger {ticket}. Rendimiento: {rend_por}')

    # Se pega la grafica en las hojas señaladas del libro
    plot_bolli = sheet_estr1.pictures.add(fig,name='Bolli')
    plot_bolli2 = sheet_estr2.pictures.add(fig,name='Bolli2')
    plot_bolli.left = sheet_estr1.range('L3').left
    plot_bolli2.left = sheet_estr1.range('K3').left

    plt.show()





