dim objShellApp

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objShellApp = CreateObject("Shell.Application")
set objBL = CreateObject("SQLXMLBulkLoad.SQLXMLBulkload.4.0")  
'Set objWSH = CreateObject("WScript.Shell")
Set objConnection1 = CreateObject("ADODB.Connection")
Set objRecordSet1  = CreateObject("ADODB.Recordset")
Set oADOStream = CreateObject("ADODB.Stream")

'для начала определим дату для имени файла

dt = Date()
dt1 = "10"&Right(0 & Month(dt),2)&CStr(Year(dt))
dt2 = CStr(Year(dt))&Right(0 & Month(dt),2)&"10"



FilePath = "C:\Scripts\MSP\Files"
sPathBase = "C:\Scripts\MSP\XSD"
Set objPath = FSO.GetFolder(FilePath) 

'почистим файлы от предыдущей загрузки
If FSO.FolderExists(FilePath) Then
	FSO.DeleteFile FSO.BuildPath(FilePath, "*.*"), True
End If


'Скачивание файла
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP.Open "GET", "https://file.nalog.ru/opendata/7707329152-rsmp/data-"&dt1&"-structure-10032022.zip", 0
oXMLHTTP.Send

oADOStream.Mode = 3
oADOStream.Type = 1
oADOStream.Open
oADOStream.Write oXMLHTTP.responseBody
oADOStream.SaveToFile FilePath+"\data-"&dt1&".zip", 2

For Each objItem In objPath.Files 'objFolder.Files
        If StrComp(FSO.GetExtensionName(objItem.Name), "zip", vbTextCompare) = 0 Then
        set FilesInZip=objShellApp.NameSpace(objItem.Path).items
        objShellApp.NameSpace(objPath.Path).CopyHere(FilesInZip)
        FSO.DeleteFile objItem
        End If
    Next

'разбираем файл и копируем данные в таблицу
      
      sServ="vabs2"
      sDB="MSP"
      sLogin = "loader"
      sPasswd = "123456"
      
      objConnection1.Open("provider=SQLOLEDB.1;data source="&sServ&";database="&sDB&";uid="&sLogin&";Password="&sPasswd)
      'objConnection1.ConnectionTimeout
      SQL = "delete from tMSP"
      
      objRecordSet1.Open SQL, objConnection1
      
      objConnection1.Close
      
      objBL.ConnectionString = "provider=SQLOLEDB.1;data source="&sServ&";database="&sDB&";uid="&sLogin&";Password="&sPasswd

      Set objFolderItems = CreateObject("Shell.Application").NameSpace(FilePath).Items
          objFolderItems.Filter 64 + 128, "VO_RRMSPSV*.xml"


if objFolderItems.Count > 0 then
Set Files = FSO.GetFolder(FilePath).Files


      sFileXSD=sPathBase & "\alor_msp.xsd"
       For each File In objFolderItems
                strFilePath = File.path
                objBL.Execute sFileXSD, strFilePath


next
end if
Set objBL=Nothing
WScript.Quit (0)

'    sServ="abs3"
'    sDB="AlorBank"
'    sLogin = "loader"
'    sPasswd = "123456"


'objConnection1.Open("provider=SQLOLEDB.1;data source="&sServ&";database="&sDB&";uid="&sLogin&";Password="&sPasswd)

'MsgBox("Запуск MSPInst")

'Set resultSet = objConnection1.Execute("execute MSPInst '"+dt2+"'")
'objConnection1.Close


