import pandas as pd
import yfinance as yf
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt
from datetime import timedelta

def data(ticket:str,s:str,f:str,interval:str):
    df = yf.download(ticket,start=s,end=f,interval = interval)
    df.drop(['Open','High','Low','Close','Volume'], axis = 1, inplace=True)
    df.dropna(inplace=True)
    return df

def ln_rend(df):
    df['R_LN'] =  np.log(df['Adj Close']/df['Adj Close'].shift(periods=1))

def con_excel(df, ticket:str):
    wb = xw.Book()
    sheet_data = wb.sheets[0]
    sheet_data
    xw.sheets.active
    xw.books.active
    sheet_data.name = ticket
    sheet_data.range('A1').value = df
    return wb

def bolling(df, ticket, ajuste):
    wb = xw.books.active
    sheet_estr1 = wb.sheets.add()
    sheet_estr2 = wb.sheets.add()
    sheet_estr2 = wb.sheets[0]
    sheet_estr1 = wb.sheets[1]
    sheet_estr1.name = 'Bandas_Boll'
    sheet_estr2.name = 'Resultados'
    df['2+std'] = df['Adj Close'].rolling(20).std()+df['Adj Close']+df['Adj Close'].rolling(10).std()
    df['2-std'] = -df['Adj Close'].rolling(20).std()+df['Adj Close']-df['Adj Close'].rolling(10).std()
    df['Rol_SWM10'] = df['Adj Close'].rolling(10).mean()
    df['Venta'] = np.round(df['2+std']/df['Rol_SWM10'],2) 
    df['Compra'] = np.round(df['Rol_SWM10']/df['2-std'],2)
    sheet_estr1.range('A1').value = df

    fig = plt.figure(figsize=(10,6))
    ### Creando dos slides del dataframe original para separar compras y ventas con los criterios definidos
    #df_ventas = df[df['Venta'].isin([1,1.01])]
    df_ventas = df[df['Venta'].isin([0.99,1.00])]
    df_compra = df[df['Compra'].isin([0.99,1.00])]
    ### Agregando una columna bandera de color según es compra o venta
    df_ventas['Color']='g'
    df_compra['Color']='r'
    ### Haciendo negativos los valores de compra ya que es dinero desembolsado
    df_compra['Compra']=df_compra['Compra'].apply(lambda x : x * -1)
    df_compra['Adj Close']=df_compra['Adj Close'].apply(lambda x : x * -1)

    ### Nuevo dataframe de apoyo uniendo las compras y ventas
    df_union = pd.concat([df_compra,df_ventas], axis=0)
    df_union = df_union.sort_index()
    
    ## Lista para almacenar un campo de 1 y -1 las filas con -1 se eliminarán ya que quiere decir que no alterna la compra y venta
    lista = []
    x = 0
    while x < (len(df_union['Color'])-1):
        for i in range(len(df_union['Color'])-1):
             if df_union.iloc[i]['Color'] != df_union.iloc[i-1]['Color']:
                 lista.append(1)
             else:
                 lista.append(-1)
             x += 1
    # Este append agrega el valor faltante para tener las mismas dimensiones que el df original union
    lista.append(-1)
    # Uniendo la columna bandera de -1 y 1
    df_union = df_union.assign(Validator=lista)
    # Este dataframe es el que se usará para alternar las compras y ventas
    df_simple = df_union[df_union['Validator']==1]

    if df_simple.iloc[0]['Adj Close'] > 0:
        df_simple.drop(df_simple.index.tolist()[0],inplace=True)

    #print(df_simple)

    lista_aux1 = []
    n = ajuste
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
        
    df_opt = df.loc[lista_aux1]
    df_opt['Colores'] = df_simple['Color'].tolist()

    #print(lista_aux1)
    df_opt['Adj Close'] = np.where(df_opt['Colores']=='r',df_opt['Adj Close']*-1,df_opt['Adj Close'])
    print(df_simple)
    print(df_opt)

    plt.scatter(df_opt.index, df_opt['Rol_SWM10'],color=df_opt['Colores'],marker="v", s=10)

    if df_opt.iloc[-1]['Adj Close'] < 0:
        df_opt.drop(df_opt.index.tolist()[-1],inplace=True)
    #print(df_simple)
    #rend = np.round(df_simple['Adj Close'].diff().sum(),2)
    salidas = -1*(df_opt[df_opt['Adj Close']< 0]['Adj Close'].sum())
    entradas = 1*(df_opt[df_opt['Adj Close']> 0]['Adj Close'].sum())
    print(salidas)
    print(entradas)
    rend= np.round(df_opt['Adj Close'].sum(),2)
    print(rend)
    rend_por = np.round(rend/salidas,2)
    print(rend_por)
    sheet_estr2.range('A1').value = df_opt
    plt.plot(df.index,df['2+std'],'r--',label='2+std', linewidth=0.35)
    plt.plot(df.index,df['Adj Close'],'y--',label='Precio', linewidth=0.35)
    plt.plot(df.index,df['2-std'],'g--',label='2-std', linewidth=0.35)
    plt.plot(df.index,df['Rol_SWM10'],color='b', linewidth=0.6)
    plt.title(f'Bandas de Bollinger {ticket}. Rendimiento: {rend_por}')
    plot_bolli = sheet_estr1.pictures.add(fig,name='Bolli')
    plot_bolli2 = sheet_estr2.pictures.add(fig,name='Bolli2')
    plot_bolli.left = sheet_estr1.range('L3').left
    plot_bolli2.left = sheet_estr1.range('K3').left





