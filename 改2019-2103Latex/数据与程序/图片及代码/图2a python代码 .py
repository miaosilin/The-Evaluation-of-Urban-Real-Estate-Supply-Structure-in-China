# -*- coding: utf-8 -*-
"""
Created on Wed May 29 15:59:28 2019

@author: qw
"""

import numpy as np
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签


#读取数据
datafile = u'E:\\房地产耦合协调度论文\\bubble data.xlsx'
data = pd.read_excel(datafile)
 
plt.figure(figsize=(10,5))#设置画布的尺寸

plt.axvline(0.1,color = 'black',linestyle = ':',alpha=0.5)
plt.axhline(0.974,color = 'black',linestyle = ':',alpha=0.5)

plt.xlabel(u'协调度',fontsize=14)#设置x轴，并设定字号大小
plt.ylabel(u'耦合度',fontsize=14)#设置y轴，并设定字号大小

colors = ['white','cornflowerblue','red','pink','lightblue']


for index in range(5):
    C = data.loc[data['loc'] == index]['耦合度']
    T = data.loc[data['loc'] == index]['协调度']
    D = data.loc[data['loc'] == index]['耦合协调度']*data.loc[data['loc'] == index]['耦合协调度']*5000
    plt.scatter(T, C, c=colors[index],  s=D,  marker='o')  

for i in range(35):
    plt.annotate(data['城市'][i], xy = (data['协调度'][i], data['耦合度'][i]), \
                 xytext = (data['协调度'][i]-0.008, data['耦合度'][i]-0.003))
    # 这里xy是需要标记的坐标，xytext是对应的标签坐标


plt.savefig('2a.pdf',dpi=600,format='pdf')
plt.show()#显示图像
