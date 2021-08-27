import sys
import os
import pandas as pd
import xlsxwriter
import platform
import xlrd

os_platform = platform.system()
first_excel_path = sys.argv[1]
second_excel_path= sys.argv[2]
durum = sys.argv[3]

if os_platform == "Linux":
    currentPath=os.getcwd()
    first_excel = pd.ExcelFile(currentPath+"/"+first_excel_path)
    second_excel = pd.ExcelFile(currentPath+"/"+second_excel_path)
elif os_platform == "Windows" :
    currentPath=os.getcwd()
    first_excel = pd.ExcelFile(currentPath+"\\"+first_excel_path)
    second_excel = pd.ExcelFile(currentPath+"\\"+second_excel_path)

first_excel_sheet_names = xlrd.open_workbook(first_excel, on_demand=True).sheet_names()
second_excel_sheet_names = xlrd.open_workbook(first_excel, on_demand=True).sheet_names()


if ("TestiniumCLOUDResults" in first_excel_sheet_names and "TestiniumCLOUDResults" in second_excel_sheet_names):

    df1 = pd.read_excel(first_excel, "TestiniumCLOUDResults")
    df2 = pd.read_excel(second_excel, "TestiniumCLOUDResults")

    df1= df1.loc[df1["Durum"] == durum]
    df2= df2.loc[df2["Durum"] == durum]

    passCasedf1 = []
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

    print("ilk testten " , len(passCasedf1), " adet farklı case geçmiş")

    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    A="A"
    column = first_excel_path+ " de olup " + second_excel_path + "de olmayan, durumu " + durum + " olan caseler "
    worksheet.write("A1", column)
    for i in range(len(passCasedf1)):

        index= str(i+2)
        worksheet.write(A+index, passCasedf1[i])


    workbook.close()

else:

    df1 = pd.read_excel(first_excel, "TestScenarioResults")
    df2 = pd.read_excel(second_excel, "TestScenarioResults")

    df1_test = []
    df1_durum = []

    for index, row in df1.iterrows():
        if row["Başarılı"]== 1: 
            df1_durum.append("SUCCESS")
            df1_test.append(row["Test Senaryosu"])
        elif row["Uyarı"]==1:
            df1_durum.append("WARNING")
            df1_test.append(row["Test Senaryosu"])
        elif row["Başarısız"]==1:
            df1_durum.append("FAILURE")
            df1_test.append(row["Test Senaryosu"])
        elif row["Hatalı"]==1:
            df1_durum.append("ERROR")
            df1_test.append(row["Test Senaryosu"])
        else:
            df1_durum.append("BLOCKED")
            df1_test.append(row["Test Senaryosu"])

    dict = {"Test Senaryosu": df1_test, "Durum": df1_durum}
    df1= pd.DataFrame(data=dict)

    df2_test=[]
    df2_durum=[]
    for index, row in df2.iterrows():
        if row["Başarılı"]== 1:
            df2_durum.append("SUCCESS")
            df2_test.append(row["Test Senaryosu"])
        elif row["Uyarı"]==1:
            df2_durum.append("WARNING")
            df2_test.append(row["Test Senaryosu"])
        elif row["Başarısız"]==1:
            df2_durum.append("FAILURE")
            df2_test.append(row["Test Senaryosu"])
        elif row["Hatalı"]==1:
            df2_durum.append("ERROR")
            df2_test.append(row["Test Senaryosu"])
        else:
            df2_durum.append("BLOCKED")
            df2_test.append(row["Test Senaryosu"])

    dict = {"Test Senaryosu": df2_test, "Durum": df2_durum}
    df2= pd.DataFrame(data=dict)

    df1= df1.loc[df1["Durum"] == durum]
    df2= df2.loc[df2["Durum"] == durum]

    df1_test=[]
    df2_test=[]

    for index, row in df1.iterrows():
        df1_test.append(row["Test Senaryosu"])
    for index, row in df2.iterrows():
        df2_test.append(row["Test Senaryosu"])

    passCasedf1=[]
    count_dup=0
    for i in range(len(df1_test)):
        if (df1_test[i] in df2_test):
            count_dup=count_dup+1
        else:
            passCasedf1.append(df1_test[i])


    print("ilk testten " , len(passCasedf1), " adet farklı case geçmiş")

    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    A="A"
    column = first_excel_path+ " de olup " + second_excel_path + "de olmayan, durumu " + durum + " olan caseler "
    worksheet.write("A1", column)
    for i in range(len(passCasedf1)):

        index= str(i+2)
        worksheet.write(A+index, passCasedf1[i])


    workbook.close()

print("finish")


