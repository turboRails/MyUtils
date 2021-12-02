import pandas as pd
 
filename = "P.xls"
 
rows = []
columns = ['OP.No.', 'OP.NAME', 'Model', 'Engine', 'Body', 'FRT']
 
data = pd.read_excel(
    filename, ["EN", "BO", "OTHERS"], index_col=None, header=None)
 
eng = data["ENG"]
opname = eng[1].fillna('')
remark = eng[2].fillna('')
 
for val in [1,2,3]:
    eng.drop(val, inplace=True, axis=1)
 
for r, s in eng.iterrows():
    if (r < 3):
        continue
    if (len(opname[r]) > 0):
        firstname = opname[r]
        fullname = opname[r]
        if(len(remark[r]) > 0):
            fullname += " -- " + remark[r]
    else:
        fullname = firstname + " -- " + remark[r]
    s = s.dropna()
    for c, value in s.items():
        if (c < 4):
            continue
        row = []
        row.append(eng[0][r] + str(eng[c][2]))
        row.append(fullname)
        row.append(eng[c][0])
        row.append(eng[c][1])
        row.append("")
        row.append(eng[c][r])
        rows.append(row)
 
eng = data["BODY"]
opname = eng[1].fillna('')
remark = eng[2].fillna('')
 
for val in [1,2,3]:
    eng.drop(val, inplace=True, axis=1)
 
for r, s in eng.iterrows():
    if (r < 3):
        continue
    if (len(opname[r]) > 0):
        firstname = opname[r]
        fullname = opname[r]
        if(len(remark[r]) > 0):
            fullname += " -- " + remark[r]
    else:
        fullname = firstname + " -- " + remark[r]
    s = s.dropna()
    for c, value in s.items():
        if (c < 4):
            continue
        row = []
        row.append(eng[0][r] + str(eng[c][2]))
        row.append(fullname)
        row.append(eng[c][0])
        row.append(eng[c][1])
        row.append("")
        row.append(eng[c][r])
        rows.append(row)
 
eng = data["OTHERS"]
opname = eng[1].fillna('')
remark = eng[2].fillna('')
 
for val in [1,2,3]:
    eng.drop(val, inplace=True, axis=1)
 
for r, s in eng.iterrows():
    if (r < 3):
        continue
    if (len(opname[r]) > 0):
        firstname = opname[r]
        fullname = opname[r]
        if(len(remark[r]) > 0):
            fullname += " -- " + remark[r]
    else:
        fullname = firstname + " -- " + remark[r]
    s = s.dropna()
    for c, value in s.items():
        if (c < 4):
            continue
        row = []
        row.append(eng[0][r] + str(eng[c][2]))
        row.append(fullname)
        row.append(eng[c][0])
        row.append(eng[c][1])
        row.append("")
        row.append(eng[c][r])
        rows.append(row)
 
df = pd.DataFrame(rows, columns=columns)
filename = filename + "x"
df.to_excel(filename,  index=False)
