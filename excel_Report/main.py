import sys
import os
import pandas as pd
import xlsxwriter
import platform
import xlrd

# İşletim sistemi ve path işlemleri
os_platform = platform.system()
first_excel_path = sys.argv[1]
second_excel_path= sys.argv[2]
durum = sys.argv[3]

if os_platform == "Linux" or os_platform == "Darwin":
    currentPath=os.getcwd()
    first_excel = pd.ExcelFile(currentPath+"/"+first_excel_path)
    second_excel = pd.ExcelFile(currentPath+"/"+second_excel_path)
elif os_platform == "Windows" :
    currentPath=os.getcwd()
    first_excel = pd.ExcelFile(currentPath+"\\"+first_excel_path)
    second_excel = pd.ExcelFile(currentPath+"\\"+second_excel_path)

first_excel_sheet_names = xlrd.open_workbook(first_excel, on_demand=True).sheet_names()
second_excel_sheet_names = xlrd.open_workbook(first_excel, on_demand=True).sheet_names()

#  
# excel: excel dosyası      sheetName: İstenilen excel sayfası      durum: istenilen test durumu
# return dataFrame
def excelRead(excel, sheetName, Testdurum):
    df=pd.read_excel(excel, sheetName)
    df=df.loc[df["Durum"] == Testdurum]
    return df
#  
# data: filtrelenmiş data
# return sadece test isimleri 
def testCases(data):
    testCase=[]
    for index, row in data.iterrows():
        testCase.append(row["Test Senaryosu"])
    return testCase
#  
# data1: Birinci excelden istelilen durumun test isimleri       data2 : ikinci excelden istelilen durumun test isimleri
# return: birinci excelde olup ikinci excelde olamayan test isimleri
def findTestCaseDiff(data1, data2):
    uniqueTestCases = []
    count_dup=0
    for i in range(len(data1)):
        if (data1[i] in data2):
            count_dup=count_dup+1
        else:
            uniqueTestCases.append(data1[i])
    return uniqueTestCases
#  
# data: testScenarioResult verisi
# return düzenlenmiş TestCloudResult tablosuna çevrilmiş Data
def dataTestScenarioResılt(data):
    data_test = []
    data_durum = []
    for index, row in data.iterrows():
        if row["Başarılı"]== 1: 
            data_durum.append("SUCCESS")
            data_test.append(row["Test Senaryosu"])
        elif row["Uyarı"]==1:
            data_durum.append("WARNING")
            data_test.append(row["Test Senaryosu"])
        elif row["Başarısız"]==1:
            data_durum.append("FAILURE")
            data_test.append(row["Test Senaryosu"])
        elif row["Hatalı"]==1:
            data_durum.append("ERROR")
            data_test.append(row["Test Senaryosu"])
        else:
            data_durum.append("BLOCKED")
            data_test.append(row["Test Senaryosu"])
    dict = {"Test Senaryosu": data_test, "Durum": data_durum}
    newData= pd.DataFrame(data=dict)
    return newData



if ("TestiniumCLOUDResults" in first_excel_sheet_names and "TestiniumCLOUDResults" in second_excel_sheet_names):

    df1 = excelRead(first_excel, "TestiniumCLOUDResults", durum)
    df2 = excelRead(second_excel, "TestiniumCLOUDResults", durum)

    df1_test = testCases(df1)
    df2_test = testCases(df2)

    testCaseDiff= findTestCaseDiff(df1_test,df2_test)
    
    #print("ilk testten " , len(passCasedf1), " adet farklı case geçmiş")

    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    A="A"
    column = first_excel_path+ " de olup " + second_excel_path + "de olmayan, durumu " + durum + " olan caseler "
    worksheet.write("A1", column)
    for i in range(len(testCaseDiff)):

        index= str(i+2)
        worksheet.write(A+index, testCaseDiff[i])
    workbook.close()

else:

    df1 = pd.read_excel(first_excel, "TestScenarioResults")
    df2 = pd.read_excel(second_excel, "TestScenarioResults")

    df1=dataTestScenarioResılt(df1)
    df2=dataTestScenarioResılt(df2)

    df1= df1.loc[df1["Durum"] == durum]
    df2= df2.loc[df2["Durum"] == durum]

    df1_test=testCases(df1)
    df2_test=testCases(df2)

    testCaseDiff= findTestCaseDiff(df1_test,df2_test)

    #print("ilk testten " , len(testCaseDiff), " adet farklı case geçmiş")

    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    A="A"
    column = first_excel_path+ " de olup " + second_excel_path + "de olmayan, durumu " + durum + " olan caseler "
    worksheet.write("A1", column)
    for i in range(len(testCaseDiff)):

        index= str(i+2)
        worksheet.write(A+index, testCaseDiff[i])
    workbook.close()

print("finish")


