import pandas as pd
import os
import json
import datetime
import numpy as np
from itertools import combinations, permutations
t=datetime.datetime.today()
x=pd.read_excel('charac_data.xlsx','INV',header=None).dropna(how='all')
x1=x.fillna('')
l1=[i for i,j in enumerate(x1[0]) if j=='ARC']
l2 = []
for i in range(len(l1)):
  if i == len(l1)-1:
    p=(l1[i],x1.shape[0])
    l2.append(p)
  else:
    p = (l1[i], l1[i+1])
    l2.append(p)
# print(l2)
n1=['A_R_Z_F_cell_delay','A_R_Z_F_fall_transition','A_F_Z_R_cell_delay','A_F_Z_R_rise_transition']
def INV(y1, y2):
    dict1={}
    x2=list(x.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(x.columns)][1].dropna())
    x3=x.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(x.columns)][1:2].dropna(axis=1).values
    x5 = x.reset_index().drop('index', axis = 1).iloc[y1+2:y2,0:len(x.columns)]
    x4=x3.tolist()
    dict1['index1']=x2
    for data in x4:
        dict1['index2']=data
    x5=x.iloc[y1+1:y2,0+1:len(x.columns)].dropna().values
    x6=x5.tolist()
    for data in x6:
        data.remove(data[0])
    x7=np.array(x6)
    dict1['values']=x7
    return dict1
def ND2():
    print('Working on it........')
def AN2():
    print('will be do soon........')
def main():
    while True:
        print('1.INV')
        print('2.ND2')
        print('3.AN2')
        opt=input('Enter your option for which sheet?:')
        if opt=='1':
            for i in range(len(l2)):
                k,l=l2[i]
                result=INV(k,l)
                with open('Inv.json','w')as f:
                    n2=0
                    for j in range(len(n1)):
                        f.write(n1[n2])
                        f.write('\n')
                        for i,j in result.items():
                            print(f'{i}:{j}',end='/\n',file=f)
                        n2+=1
                
            os.startfile('Inv.json')
        elif opt=='2':
            ND2()
        elif opt=='3':
            AN2()
        else:
            print('oops! you have choose wrong option please try again!')
        q=input("Do you want to continue?:")
        if q=='yes':
            continue
        else:
            break
     
main()
