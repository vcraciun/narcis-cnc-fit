import pandas as pd
import os

for fl in os.listdir('.'):
    _, ext = os.path.splitext(fl)
    if ext != '.xlsx':
        continue
    df = pd.DataFrame(pd.read_excel(fl))
    c1 = list(df.iloc[:,1])
    c2 = list(df.iloc[:,2])
    c3 = list(df.iloc[:,3])
    c4 = list(df.iloc[:,4])
    data = [(c1[i], c2[i], c3[i], c4[i]) for i in range(len(c1)) if type(c1[i]) == int and type(c2[i]) == int and type(c3[i]) == int]    
    print(fl, len(data))

    