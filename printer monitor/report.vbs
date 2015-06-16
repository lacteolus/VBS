'On Error Resume Next

'===== Задается пользователем =========================

'Имя файла для записи результатов
Const strResultFile = "C:\otchet\result.txt"

'Задаем период выборки - от strFrom до strTo
'const strFrom = "2009.11.18 00:00:00"
'const strTo = "2010.02.15 23:59:59"

'======================================================

'=====Параметры подключения к базе данных==============
Set objConn = CreateObject("ADODB.Connection")
'Set objComm = CreateObject("ADODB.Command")
Set objRecordset = CreateObject("ADODB.Recordset")
'Определяем параметры подключения к базе данных
ServerName = "xi.mdi.ru" 'имя или IP-адрес сервера
DSN = "print" 'имя базы данных
UID = "print" 'логин пользователя SQL-сервера
PWD = "print" 'пароль пользователя SQL-сервера
ConnectString = "Provider=SQLOLEDB;" & _
                "Data Source=" & ServerName & _
                ";Initial Catalog=" & DSN & _
                ";UID=" & UID & ";PWD=" & PWD
objConn.ConnectionString = ConnectString
objConn.ConnectionTimeOut = 15
objConn.CommandTimeout = 30
'objConn.Mode = adModeReadWrite
'======================================================

'======Подключаемся к базе данных для получения времени последней записи==========
objConn.Open
strSQLString = "SELECT Date FROM printing ORDER BY Date Desc"
objRecordset.Open strSQLString, objConn, adOpenStatic, adLockOptimistic
If objRecordset.BOF = FALSE Then
	objRecordset.MoveFirst
	strFrom = objRecordset.Fields("Date").Value + 0.0007
Else strFrom = Now() - 31
End If
'WScript.Echo strFrom
objConn.Close

strTo = Now()

'Wscript.Echo strFrom
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

Set Fso = CreateObject("Scripting.FileSystemObject")
'Set FolderName = Fso.GetFolder(strRepFolder)

'Создаем файл с результатами или очищаем существующий
Set Destination = Fso.CreateTextFile(strResultFile)
Destination.Close


'Set objConn = CreateObject("ADODB.Connection")
'Set objComm = CreateObject("ADODB.Command")
'Set objRecordset = CreateObject("ADODB.Recordset")
'Определяем параметры подключения к базе данных
'ServerName = "win-2008.mdi.ru" 'имя или IP-адрес сервера
'DSN = "print" 'имя базы данных
'UID = "print" 'логин пользователя SQL-сервера
'PWD = "print" 'пароль пользователя SQL-сервера
'ConnectString = "Provider=SQLOLEDB;" & _
'                "Data Source=" & ServerName & _
'                ";Initial Catalog=" & DSN & _
'                ";UID=" & UID & ";PWD=" & PWD
'objConn.ConnectionString = ConnectString
'objConn.ConnectionTimeOut = 15
'objConn.CommandTimeout = 30

Function WMIDateStringToDate(dtmDate)
	WMIDateStringToDate = Left(dtmDate, 4) & Mid(dtmDate, 5, 2) & Mid(dtmDate, 7, 2) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2)
End Function

Function WMIDateStringToDate2(dtmDate)
	WMIDateStringToDate2 = CDate(Mid(dtmDate, 7, 2) & "." & _
	Mid(dtmDate, 5, 2) & "." & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

dateTo = CDate(strTo)
dateFrom = CDate(strFrom)
arrComputers = Array("VM-PRINT")
For Each strComputer In arrComputers
   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'System'", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
'	WScript.Echo WMIDateStringToDate2(objItem.TimeGenerated)
     If objItem.EventCode = 10 and (WMIDateStringToDate2(objItem.TimeGenerated) < dateTo) and (WMIDateStringToDate2(objItem.TimeGenerated) > dateFrom) then
	EndPoint = 0
	startPoint = 0
	strLenght = 0

	'Поиск принтера
	EndPoint = Instr(objItem.Message, "via port")
	startPoint = Instr(objItem.Message, "printed on ") + 11
	strLenght = EndPoint - startPoint
	strPrinter = Trim(Mid(objItem.Message, startPoint, strLenght))

	'Поиск страниц
	EndPoint = Instr(objItem.Message, chr(13))
	startPoint = Instr(objItem.Message, "pages printed: ") + 15
	strLenght = EndPoint - startPoint
	strPages = Replace(Mid(objItem.Message, startPoint, strLenght), "%", "")
	
	'Поиск юзера
	EndPoint = Instr(objItem.Message, "was printed")
	startPoint = Instr(objItem.Message, "owned by ") + 9
	strLenght = EndPoint - startPoint
	strUser = Trim(Mid(objItem.Message, startPoint, strLenght))
	
	'Поиск документа
	EndPoint = Instr(objItem.Message, "owned by")
	startPoint = Instr(objItem.Message, ", ") + 2
	strLenght = EndPoint - startPoint
	strDocument = Trim(Replace(Mid(objItem.Message, startPoint, strLenght), "'", "%2019"))

	'Время события
	strTime = WMIDateStringToDate(objItem.TimeGenerated)
		'WScript.Echo strTime & " " & Now() - 20

'	objConn.Open
'	objComm.ActiveConnection = objConn
'	objComm.CommandText = "INSERT INTO dbprint (datetime,username,printer,documentname,numpages) VALUES ('" & strTime & "','" & strUser & "','" & strPrinter & "','" & strDocument & "','" & strPages & "')"
'	'objComm.CommandText = "INSERT INTO dbprint (datetime,username,printer,documentname,numpages) VALUES ('" & Now() & "','5','6','7','8')"
'	objRecordset = objComm.Execute
'	objConn.Close

'If Fso.FileExists(strResultFile) Then
'			Set Destination = Fso.OpenTextFile(strResultFile, 8)
'			Destination.WriteLine strPrinter & ";" & strUser & ";" & strPages & ";" & strTime
'			Destination.Close
'End If

'============Пишем  в базу==================
objConn.Open
strSQLString = "INSERT INTO printing (Date,Printer,PUser,Pages,Document) VALUES ('" & strTime & "','" & strPrinter & "','" & strUser & "','" & strPages & "','" & strDocument & "')"
objRecordset = objConn.Execute(strSQLString)
objConn.Close





   End If    

'WScript.Echo "Document: " & strDocument & " Pages: " & strPages & " User: " & strUser & " Printer: " & strPrinter
   Next
Next
'WScript.Echo "The end"
Set objConn = Nothing
Set objRecordset = Nothing
