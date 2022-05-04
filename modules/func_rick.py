import pandas as pd
import yfinance as yf
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt

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

def bolling(df, ticket):
    wb = xw.books.active
    sheet_estr1 = wb.sheets.add()
    sheet_estr1 = wb.sheets[0]
    sheet_estr1.name = 'Bandas_Boll'
    df['2+std'] = df['Adj Close'].rolling(20).std()+df['Adj Close']+df['Adj Close'].rolling(10).std()
    df['2-std'] = -df['Adj Close'].rolling(20).std()+df['Adj Close']-df['Adj Close'].rolling(10).std()
    df['Rol_SWM10'] = df['Adj Close'].rolling(10).mean()
    df['Venta'] = np.round(df['2+std']/df['Rol_SWM10'],2) 
    df['Compra'] = np.round(df['Rol_SWM10']/df['2-std'],2)
    sheet_estr1.range('A1').value = df

    fig = plt.figure(figsize=(10,6))
    ### Creando dos slides del dataframe original para separar compras y ventas con los criterios definidos
    df_ventas = df[df['Venta'].isin([1,1.01])]
    df_compra = df[df['Compra'].isin([0.99,1])]
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
    plt.scatter(df_simple.index, df_simple['Rol_SWM10'],color=df_simple['Color'],marker="v", s=10)

    if df_simple.iloc[-1]['Adj Close'] < 0:
        df_simple.drop(df_simple.index.tolist()[-1],inplace=True)
    #print(df_simple)
    rend = np.round(df_simple['Rol_SWM10'].diff().sum(),2)
    print(rend)

    plt.plot(df.index,df['2+std'],'r--',label='2+std', linewidth=0.35)
    plt.plot(df.index,df['2-std'],'g--',label='2-std', linewidth=0.35)
    plt.plot(df.index,df['Rol_SWM10'],color='b', linewidth=0.6)
    plt.title(f'Bandas de Bollinger {ticket}. Rendimiento: {rend}')
    plot_bolli = sheet_estr1.pictures.add(fig,name='Bolli')
    plot_bolli.left = sheet_estr1.range('L3').left





