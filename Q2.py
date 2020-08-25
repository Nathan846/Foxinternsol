import pandas as pd
df = pd.read_excel('NIFTY25JUN2010000PE.xlsx')
df.loc[:,'DateTime'] = pd.to_datetime(df.Date.astype(str)+' '+df.Time.astype(str))
df['DateTime'] = df['DateTime'].apply(lambda x: x.to_pydatetime())
df.set_index('DateTime',inplace=True)
df.drop(['Ticker','Date','Time'],inplace=True,axis=1)
odf = df.resample('15Min')
index = odf.sum().index
pdf = pd.DataFrame(index=index)
pdf['Open'] = odf.first()['Open ']
pdf['High'] = odf.max()['High ']
pdf['Low'] = odf.min()['Low ']
pdf['Close'] = odf.last()['Close ']
pdf['Volume'] = odf.sum()['Volume']
print('\n\n\n\n\n')
print('15 Min Granularity data')
print(pdf.head())
print('\n\n\n\n\n')
exitstarted=0
entryprice = pdf.iloc[0]['Volume']
low = pdf.iloc[0]['Low']
for i in range(1,len(pdf)):
    lownow = pdf.iloc[i]['Low']
    if(pdf.iloc[i-1]['Low']<lownow):
        if(pdf.iloc[i-1]['High']<pdf.iloc[i]['High']):
            exitprice = pdf.iloc[i]['Volume']
            break
        exitstarted=1
    if(exitstarted and pdf.index[i-1].day!=pdf.index[i].day):
        exitprice = pdf.iloc[i-1]['Volume']
        break
    if(i==len(pdf)):
        exitprice = pdf.iloc[i]['Volume']
cash = entryprice - exitprice
if(cash<0):
    print('Profit = ',cash)
else:
    print('Loss =',cash)
        