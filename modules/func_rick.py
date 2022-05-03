#import pandas as pd
from unicodedata import name
from pandas import Period
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

def bolling(df):
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

    df_ventas = df[df['Venta'].isin([1,1.01])]
    df_compra = df[df['Compra'].isin([0.99,1])]

    
    plt.scatter(df_ventas.index, df_ventas['Rol_SWM10'],color='g',marker="v", s=10)
    plt.scatter(df_compra.index, df_compra['Rol_SWM10'],color='r',marker="v", s=10)

    plt.plot(df.index,df['2+std'],'r--',label='2+std', linewidth=0.35)
    plt.plot(df.index,df['2-std'],'g--',label='2-std', linewidth=0.35)
    plt.plot(df.index,df['Rol_SWM10'],color='b', linewidth=0.6)
    plt.title('Bandas de Bollinger')
    plot_bolli = sheet_estr1.pictures.add(fig,name='Bolli')
    plot_bolli.left = sheet_estr1.range('L3').left





