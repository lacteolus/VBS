'************************************************************************** 
'* Файл: ping.vbs                                                         * 
'* Язык: VBScript                                                         * 
'* Назначение: Скрипт для определения доступности хостов в сети           *
'* В скрипте используется запрос к классу Win32_PingStatus, который       *
'* доступен только в версиях Windows старше XP!!!                         *
'**************************************************************************

'Option Explicit

Dim objPing, objRetStatus, fso, fnPing, i, objWMIService, oShell, oTF, strAddress, LogFile
Const NETWORK = "192.168.4." 'Network
Const LogPath = "\\delta\log$\night_ping.txt" 'Path to log-file
Const sTempFile = "C:\temp_file.txt" 'Temporary file
Set WshShell = CreateObject("WScript.Shell") 'Command prompt
Set fso = CreateObject("Scripting.FileSystemObject") 'Working with files
Set osf = CreateObject("Scripting.FileSystemObject")
If osf.FileExists (LogPath) then
	Set LogFile = osf.OpenTextFile(LogPath,2,True) 'Open log-file
End If	
			
For i = 2 To 254 'Range of ip addresses
	strAddress = NETWORK & i
	Set objPing = GetObject( "winmgmts:{impersonationLevel=impersonate}" ).ExecQuery _
	( "select * from Win32_PingStatus where address = '" & strAddress & "'" ) 'Ping hosts 
	For Each objRetStatus in objPing
		If objRetStatus.StatusCode = 0 Then
			WshShell.Run "%comspec% /c nslookup.exe " & strAddress & " > " & sTempFile, 0, True ' 'Run NSLookup via Command Prompt. Dump results into a temp text file 
			Set oTF = fso.OpenTextFile(sTempFile) 'Open the temp Text File and Read out the Data 
			Do While Not oTF.AtEndOfStream 'Parse the text file
				sLine = Trim(oTF.Readline)
				If LCase(Left(sLine, 5)) = "name:" Then 'Looking for string contained 'name'
					sData = Trim(Mid(sLine, 6))
					Exit Do
				End If
			Loop 
			oTF.Close 'Close file	
			LogFile.WriteLine (now() & " " & strAddress & " " & sData) 'Write in log-file
		End If
	Next
Next

LogFile.Close 'Close log-file

'strCommandString = "postie -host:mdi.mdi.ru -to:yusupov@mdi.ru -from:nimda@mdi.ru -s:""Computers at night"" -import -file:" & LogPath
'WshShell.Exec(strCommandString)
strCommandString = "postie -host:mdi.mdi.ru -to:kosmachevda@mdi.ru -from:nimda@mdi.ru -s:""Computers at night"" -import -file:" & LogPath
WshShell.Exec(strCommandString)
'strCommandString = "postie -host:mdi.mdi.ru -to:maksim@mdi.ru -from:nimda@mdi.ru -s:""Computers at night"" -import -file:" & LogPath
'WshShell.Exec(strCommandString)