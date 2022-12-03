import os
import pandas as pd
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM
from tensorflow.keras.layers import Dense
from tensorflow.keras.layers import Dropout
import numpy as np
from sklearn import preprocessing
import tensorflow as tf
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

if os.path.exists("SalesForTrain&Test.csv"):
    pass
else:    
    wget.download("https://drive.google.com/file/d/1EBe213TESkHvEwOJu_LpVpt3736MsWQh/view?usp=share_link")



df=pd.read_csv(f"SalesForTrain&Test.csv")

df.columns=["Date","Holiday","Avg. % ad spend of gross revenue","Avg. % discount","Sales"]

train_length=len(df["Sales"])-df["Sales"].isnull().sum()

test_lenth=df["Sales"].isnull().sum()

le=preprocessing.LabelEncoder()
encoded=le.fit_transform(df["Holiday"])

from sklearn.preprocessing import MinMaxScaler
scaler=MinMaxScaler()

scaled_ad_spend=scaler.fit_transform(df["Avg. % ad spend of gross revenue"][:,np.newaxis])

scaled_discount=scaler.fit_transform(df["Avg. % discount"][:,np.newaxis])
scaled_sales= scaler.fit_transform(df["Sales"][:,np.newaxis])

scaled_discount2=np.array(scaled_discount)
scaled_sales2=np.array(scaled_sales)
encoded=np.array(encoded)

encoded=encoded.reshape((len(df["Sales"]),1))

final3=np.array((encoded[:train_length],scaled_ad_spend[:train_length],scaled_discount[:train_length]))

test=np.array((encoded[train_length:],scaled_ad_spend[train_length:],scaled_discount[train_length:]))
# print(test.shape)

test=test.reshape((test_lenth,3))

final3=final3.reshape((train_length,3))

from keras.preprocessing.sequence import TimeseriesGenerator
n_input=3
n_features=3
generator=TimeseriesGenerator(final3,scaled_sales[:train_length],length=n_input,batch_size=1)

model = Sequential()
model.add(LSTM(100, activation='relu', input_shape=(n_input,n_features)))
model.add(Dense(1))
model.summary()

model.compile(loss=tf.keras.losses.mse,optimizer=tf.keras.optimizers.Adam(lr=0.001))
model.fit(generator,epochs=150,batch_size=30)

first_eval_batch=final3[-n_input:,:,np.newaxis]
current_batch = first_eval_batch.reshape((1, n_input, n_features))


final3=final3.astype(np.uint8)

test=test.astype(np.uint8)

predictions=[]
for i in range(0,test_lenth):
    if i==0:
        predictions.append(model.predict(current_batch)[0])
        # print("hi")
    elif i==1:
        # print("pass")
        # first_eval_batch=final3[-n_input+1:,:,np.newaxis]
        first_eval_batch=final3[-n_input+1:,:,np.newaxis]
        first_eval_batch=list(first_eval_batch)
        # print(first_eval_batch)
        first_eval_batch.append(test[0,:,np.newaxis])
        # print(first_eval_batch)
        first_eval_batch=np.array(first_eval_batch)
        current_batch = first_eval_batch.reshape((1, n_input, n_features))
        # print(current_batch)
        predictions.append(model.predict(current_batch)[0])
    elif i==2:
        first_eval_batch=final3[-n_input+2:,:,np.newaxis]
        first_eval_batch=list(first_eval_batch)
        first_eval_batch.append(test[0,:,np.newaxis])
        first_eval_batch.append(test[1,:,np.newaxis])
        # print(first_eval_batch)
        first_eval_batch=np.array(first_eval_batch)

        current_batch = first_eval_batch.reshape((1, n_input, n_features))
        # print(current_batch)
        predictions.append(model.predict(current_batch)[0])
    else:
        first_eval_batch=test[i-n_input:i,:,np.newaxis]
        current_batch = first_eval_batch.reshape((1, n_input, n_features))
        # print(current_batch)
        predictions.append(model.predict(current_batch)[0])

scaled_predictions = scaler.inverse_transform(predictions)
print(scaled_predictions)
true_predictions=[]

for y in range(0,len(scaled_predictions)):
    
    for i in scaled_predictions[y]:

        if i<=0:
            i=-i
            true_predictions.append(i)
   
        else:
            true_predictions.append(i)

Workbookname=np.random.randint(0,1000)
workbook = xlsxwriter.Workbook(f'{Workbookname}.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Date')
worksheet.write('B1', 'Sales')

row = 1
column = 0
for i in df["Date"][train_length:]:

    worksheet.write(row, column, i)
    row += 1

row = 1
column +=1
for i in true_predictions :

    worksheet.write(row, column, i)
    row += 1
    
workbook.close()

print("Done!")

