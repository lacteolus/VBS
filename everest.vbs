'************************************************************************** 
'* Файл: everest.vbs                                                      * 
'* Язык: VBScript                                                         * 
'* Назначение: Cкрипт для выборки данных из отчетов Everest               *
'**************************************************************************

'Option Explicit
'On Error Resume Next

Dim strFileName, strMACAdress, strIPAdress, strCompName, strCPUType, strSysMemory, strDisk1,  strDisk2, strDisk3,  strDisk4, strDisk5, strDisk6,_
	 strNameOS, strOSSP, strOffice, strSymantec, intCount, strOf

'Пути для директории с отчетами и выходного файла
Const strRepFolder = "\\delta\log$\Reports\mgsn\reports-2006-06-18\Брянская 9"
Const strResultFile = "\\delta\log$\Reports\mgsn\reports-2006-06-18\Брянская 9\result.txt"

Set Fso = CreateObject("Scripting.FileSystemObject")
Set FolderName = Fso.GetFolder(strRepFolder)

'Создаем файл с результатами или очищаем существующий
Set Destination = Fso.CreateTextFile(strResultFile)
Destination.Close

'Счетчик записей
intCount = 0

'Просмотр всех файлов *.ini в директории
For Each Member in FolderName.Files
	strFileName = Member.Name
	If LCase(Right(strFileName, 3)) = "ini" Then 

'Поиск необходимых данных
		Set txtStream = fso.OpenTextFile(Member, 1)
		While Not txtStream.AtEndOfStream
    			Str = txtStream.ReadLine()
    			If Left(Str, 27) = "Network|Primary IP Address=" Then
				strIPAdress = Mid(Str, 28, 15)   
    			End If
    			If Left(Str, 28) = "Network|Primary MAC Address=" Then
				strMACAdress = Mid(Str, 29, 17)
			End If
			If Left(Str, 27) = "NetBIOS Name|Computer Name=" Then
        			strCompName = Mid(Str, 28, 16)
    			End If
    			If Left(Str, 21) = "Motherboard|CPU Type=" Then
       				strCPUType = Mid(Str, 22, 50)
    			End If
    			If Left(Str, 26) = "Motherboard|System Memory=" Then
        			strSysMemory = Mid(Str, 27, 8)
    			End If
    			If Left(Str, 20) = "Storage|Disk Drive1=" Then
        			strDisk1 = Mid(Str, 21, 70)
    			End If
    			If Left(Str, 20) = "Storage|Disk Drive2=" Then
        			strDisk2 = ", " + Mid(Str, 21, 70)
    			End If
    			If Left(Str, 20) = "Storage|Disk Drive3=" Then
        			strDisk3 = ", " + Mid(Str, 21, 70)
    			End If
    			If Left(Str, 20) = "Storage|Disk Drive4=" Then
        			strDisk4 = ", " + Mid(Str, 21, 70)
    			End If
    			If Left(Str, 20) = "Storage|Disk Drive5=" Then
        			strDisk5 = ", " + Mid(Str, 21, 70)
    			End If
    			If Left(Str, 20) = "Storage|Disk Drive6=" Then
        			strDisk6 = ", " + Mid(Str, 21, 70)
    			End If
    			If Left(Str, 36) = "Operating System Properties|OS Name=" Then
        			strNameOS = Mid(Str, 37, 50)
    			End If
    			If Left(Str, 44) = "Operating System Properties|OS Service Pack=" Then
        			strOSSP = Mid(Str, 45, 50)
    			End If
				If InStr(2, Str, "crosoft") and InStr(Str, "Office") and InStr(Str, "Version=") and (not Left(Str, 3) = "Sun") and InStr(Str, "Proof") = 0 and InStr(Str, "MUI") = 0 and InStr(Str, "FrontPage") = 0  and InStr(Str, "Visio") = 0 Then
        			strOffice = Mid(Str, InStr(Str, "=")+1, 9)
				End If
    			If ((Left(Str, 18) = "Symantec AntiVirus" or Left(Str, 17) = "Symantec Endpoint") and InStr(Str, "Version=")) or Left(Str, 17) = "Symantec|Version=" Then
        			If Right(Left(Str, 32), 3) <> "Mgr" Then
						strSymantec = Mid(Str, InStr(Str, "=")+1, 20)
					End If
    			End If
		Wend
		txtStream.Close

'Запись в файл с результатами
 		If Fso.FileExists(strResultFile) Then
			Set Destination = Fso.OpenTextFile(strResultFile, 8)
			Destination.WriteLine strMACAdress & ";" & strIPAdress & ";" & strCompName & ";" & strCPUType & ";" & strSysMemory & _
				";" & strDisk1 & strDisk2 & strDisk3 & strDisk4 & strDisk5 & ";" & strNameOS & _
				";" & strOSSP & ";" & strOffice & ";" & strSymantec
			Destination.Close	
 		End If
	intCount = intCount + 1
	End If

'Сбрасываем все переменные
strMACAdress = null
strIPAdress = null
strCompName = null
strCPUType = null
strSysMemory = null
strDisk1 = null
strDisk2 = null
strDisk3 = null
strDisk4 = null
strDisk5 = null
strNameOS = null
strOSSP = null
strOffice = "No"
strSymantec = "No"

Next

MsgBox "В файл " & strResultFile & " записано строк: " & intCount, 0, "Результат"

