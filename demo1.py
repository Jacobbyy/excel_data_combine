import pandas as pd
import glob
files = glob.glob(r'DODECANE/*.xlsx')
OM = []
OM1 = []
T = []
T1 = []
for file in files:
    OM.append(eval(file[-13:-5]))
    T.append(eval(file[-17:-14]))

#去重
for m in T:
    if m not in T1:
        T1.append(m)
for n in OM:
    if n not in OM1:
        OM1.append(n)

df1 = pd.DataFrame(index=OM1,columns=T1)  # O2C
df2 = pd.DataFrame(index=OM1,columns=T1)  # 产率
print(df1.shape)
for file in files:
    for i in df1.index:
        for j in df1.columns:
            if (eval(file[-17:-14])==j) and (eval(file[-13:-5])==i):
                data_temp = pd.read_excel(file)
                df1.loc[[i],[j]] = data_temp.iloc[-2,-2]
                df2.loc[[i],[j]] = data_temp.iloc[-2,-1]
df1.to_excel('finally_O2C.xlsx')
df2.to_excel('finally_Yiled.xlsx')