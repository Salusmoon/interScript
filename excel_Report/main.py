import sys
import os
import pandas as pd
import xlsxwriter   

currentPath=os.getcwd()
print(currentPath)

first_excel_path = sys.argv[1]
second_excel_path= sys.argv[2]


first_excel = pd.ExcelFile(currentPath+"/"+first_excel_path)
second_excel = pd.ExcelFile(currentPath+"/"+second_excel_path)


df1 = pd.read_excel(first_excel, "TestiniumCLOUDResults")
df2 = pd.read_excel(second_excel, "TestiniumCLOUDResults")

df1= df1.loc[df1["Durum"] == "SUCCESS"]
df2= df2.loc[df2["Durum"] == "SUCCESS"]

passCasedf1 = []
passCasedf2 = []
df1_test=[]
df2_test=[]

for index, row in df1.iterrows():
    df1_test.append(row["Test Senaryosu"])


for index, row in df2.iterrows():
    df2_test.append(row["Test Senaryosu"])

count_dup=0
for i in range(len(df1_test)):
    if (df1_test[i] in df2_test):
        count_dup=count_dup+1
    else:
        passCasedf1.append(df1_test[i])


count_dup=0
for i in range(len(df2_test)):
    if (df2_test[i] in df1_test):
        count_dup=count_dup+1
    else:
        passCasedf2.append(df2_test[i])


print("ilk testten " , len(passCasedf1), " adet farklı case geçmiş")
print("ilk testten " ,len(passCasedf2) , " adet farklı case geçmiş")


workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()
A="A"
worksheet.write("A1", 'İlk excelin farkı')
for i in range(len(passCasedf1)):

    index= str(i+2)
    worksheet.write(A+index, passCasedf1[i])

index2 = str(len(passCasedf1)+5)
worksheet.write(A+index2, 'İkinci excelin farkı')
for j in range(len(passCasedf2)):
    index=str(len(passCasedf1)+5+j+1)
    worksheet.write(A+index, passCasedf2[j])

workbook.close()
print("finish")


