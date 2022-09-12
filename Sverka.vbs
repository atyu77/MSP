Dim FilePath
FilePath = "C:\Scripts\СверкаМСП"
'файл из налоговой
sFileMSPN = FilePath&"\rsmp.xlsx"

Set objConnection1 = CreateObject("ADODB.Connection")
Set objRecordSet1  = CreateObject("ADODB.Recordset")
Set oADOStream = CreateObject("ADODB.Stream")


Set objExcel1 = CreateObject("Excel.Application")
objExcel1.Workbooks.Open(sFileMSPN)
Set objWorkbook1 = objExcel1.ActiveWorkbook.Worksheets("Лист1")

sServ="vabs2"
sDB="MSP"
sLogin = "loader"
sPasswd = "123456"

dt = CStr("31.08.2022 09:02")
objConnection1.Open("provider=SQLOLEDB.1;data source="&sServ&";database="&sDB&";uid="&sLogin&";Password="&sPasswd)
SQL = "delete from pSverkaMSP "

   RSN = 4
      Do Until objWorkbook1.Cells(RSN, 1).Value =""                 
        
         SQL = SQL+" insert pSverkaMSP select "+"'"+objWorkbook1.Cells(RSN, 6).Value+"'"+","+"'"+objWorkbook1.Cells(RSN, 4).Value+"'"+","+"'"+objWorkbook1.Cells(RSN, 13).Value+"'"+","+"'"+objWorkbook1.Cells(RSN, 14).Value+"'"+","+"'"+dt+"'"
         'MsgBox(SQL)      
                  
         RSN = RSN+1
      Loop  
        

objRecordSet1.Open SQL, objConnection1
objConnection1.Close

objExcel1.ActiveWorkbook.Close
objExcel1.Workbooks.Close
objExcel1.Application.Quit
