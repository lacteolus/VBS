'**************************************************************************
'* File: event_log.vbs                                                    * 
'* Language: VBScript                                                     * 
'* Function: Looking for errors and warnings in Windows event log         *
'**************************************************************************
'Option Explicit
On Error Resume Next
Dim objFso, objFolder, objWMI, objEvent ' Objects
Dim strFile, strComputer, strFolder, strFileName, strPath ' Strings
Dim intEvent, intNumberID, intRecordNum, colLoggedEvents
'arrEventLogs = Array() 


'Reading parameters from config file
ExecuteGlobal _
    CreateObject("Scripting.FileSystemObject"). _
    OpenTextFile("event_log.config").ReadAll

' Set your variables
nowHour = (DatePart("h",Now))


'String to date
Function WMIDateStringToDate(dtmDate)
	WMIDateStringToDate = CDate(Mid(dtmDate, 7, 2) & "/" & _
	Mid(dtmDate, 5, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

'Main section
If nowHour = strRunTime1 Then
		chPeriod = 0.0417 * strPeriod1 
Elseif nowHour = strRunTime2 Then
		chPeriod = 0.0417 * strPeriod2 
Else
		WScript.Quit
End If

'Connecting to WMI and getting computer name
Set objWMI = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\.\root\cimv2")
	strQuery = "SELECT * FROM Win32_OperatingSystem"
Set colItems = objWMI.ExecQuery (strQuery)
For Each objItem in colItems
	strComputer = LCase(objItem.CSName)
Next

For Each strEventLog In arrEventLogs
	
	'Creating folder and file
	Set objFso = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(strFolder) Then
		Set objFolder = objFSO.GetFolder(strFolder)
	Else
		Set objFolder = objFSO.CreateFolder(strFolder)
	End If
	strPath = strFolder & "\" & strEventLog & ".html"
	Set strFile = objFso.CreateTextFile(strPath, True)
	strFile.WriteLine ("<p align=""center""><FONT Size=""+1"">Error and Warning events in " & strEventLog & " log on <b>" & UCase(strComputer) & "</b> between <b>" & (Now() - chPeriod) & "</b> and <b>" & Now() & "</b></FONT></p>")
	strFile.WriteLine ("<hr>")
	strFile.WriteLine ("<br><br>")
	
	'Getting Error and Warning events
	strQuery = "Select * from Win32_NTLogEvent Where Logfile = '" & strEventLog & "' And (Type = 'Error' Or Type = 'Warning')"
	Set colLoggedEvents = objWMI.ExecQuery (strQuery)

	' Looking for events and writing into the file
	strNotImportantSymantec = 0
	strNotImportantUserenv = 0
	intEvent = 0
	For Each objEvent in colLoggedEvents
		result = 0
		If (WMIDateStringToDate(objEvent.TimeGenerated) > (Now() - chPeriod)) Then
			For i = 0 To Ubound(arrExclusions) - 1
				If (objEvent.SourceName = arrExclusions(i, 0) And objEvent.Category = arrExclusions(i, 1) And objEvent.EventCode = arrExclusions(i, 2)) Then
					arrExclusions(i, 3) = arrExclusions(i, 3) + 1
					result = 1
					'Exit For
				End If
			Next
			If result = 0 Then
				strFile.WriteLine ("<b>Record No:</b>" & vbTab & (intEvent + 1) & "<br>")
				strType = UCase(objEvent.Type)
				Select Case strType
					Case "ERROR"
						strFile.WriteLine ("<table BORDER=""0"" WIDTH=""100%""><tr BGCOLOR=""dcdcdc""><td WIDTH=""100px""><b>Type:</b></td><td><FONT COLOR=""cc0000""><b>" & vbTab & strType & "</FONT></b></td></tr>")
					Case "WARNING"
						strFile.WriteLine ("<table BORDER=""0"" WIDTH=""100%""><tr BGCOLOR=""dcdcdc""><td WIDTH=""100px""><b>Type:</b></td><td><FONT COLOR=""ff9900""><b>" & vbTab & strType & "</FONT></b></td></tr>")
				End Select
				strFile.WriteLine ("<tr><td><b>Time: </b></td><td>" & vbTab & WMIDateStringToDate(objEvent.TimeGenerated) & "</td></tr>")
				strFile.WriteLine ("<tr BGCOLOR=""dcdcdc""><td><b>Source:</b></td><td>" & vbTab & objEvent.SourceName & "</td></tr>")
				strFile.WriteLine ("<tr><td><b>Category: </b></td><td>" & vbTab & objEvent.Category & "</td></tr>")
				strFile.WriteLine ("<tr BGCOLOR=""dcdcdc""><td><b>Event ID:</b></td><td>" & vbTab & objEvent.EventCode & "</td></tr>")
				strFile.WriteLine ("<tr><td><b>User:</b></td><td>" & vbTab & objEvent.User & "</td></tr>")
				'strFile.WriteLine ("Computer: " & vbTab & objEvent.ComputerName & "</BR>")
				' strFile.WriteLine ("Message: " & objEvent.Message)
				' strFile.WriteLine ("Record Number: " & objEvent.RecordNumber)
				' strFile.WriteLine ("Time Written: " & objEvent.TimeWritten)
				strFile.WriteLine ("<tr BGCOLOR=""dcdcdc""><td><b>Description:</b></td><td>")
				strDescription = Replace(objEvent.Message, "<", "{")
				strDescription = Replace(strDescription, ">", "}")
				strFile.WriteLine (strDescription & "</td></tr>")
				strLink = Replace(objEvent.SourceName, " ", "+")
				strFile.WriteLine ("<tr><td><b>Link 1:</b></td><td><A HREF=""http://www.eventid.net/display.asp?eventid=" & objEvent.EventCode & "&source=" & strLink & """>EventID.Net</A></td></tr>")
				strFile.WriteLine ("<tr BGCOLOR=""dcdcdc""><td><b>Link 2:</b></td><td><A HREF=""http://www.microsoft.com/technet/support/ee/SearchResults.aspx?Type=1&Source=" & objEvent.SourceName & "&ID=" & objEvent.EventCode & "&Language=1033"">Microsoft Events and Errors Message Center</A></td></tr></table>")
				strFile.WriteLine ("<br><br>")
			End If
			IntEvent = intEvent + 1
		End If
	Next
	If intEvent = 0 Then
		strSubject = "CLEAR"
		strFile.WriteLine ("There are no Error or Warning events in " & strEventLog & " log for the specified period")
	Else
		strSubject = "ERRORS"
		For i = 0 To Ubound(arrExclusions)
			If arrExclusions(i, 3) > 0 Then
				strFile.WriteLine ("There are <b>" & arrExclusions(i, 3) & "</b> errors/warnings from <b>" & arrExclusions(i, 0) & "</b>  (Category: <b>" & arrExclusions(i, 1) & "</b>, Code: <b>" & arrExclusions(i, 2) & "</b>) in " & strEventLog & " log for the specified period<br>")
				arrExclusions(i, 3) = 0
			End If
		Next
	End If
	strFile.WriteLine ("</BODY>")
	strFile.WriteLine ("</HTML>")
	
	'Posting with CDO
	If strSubject = "ERRORS" Then
		'Reading from html file
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(strPath, 1)
		BodyText = f.ReadAll
		f.Close
		Set f = Nothing
		Set fso = Nothing
	
		'Posting in HTML format
		Set objEmail = CreateObject("CDO.Message")
		objEmail.From = strComputer & strFromSuffix
		objEmail.To = strToAdresses
		objEmail.Subject = strEventLog & " Event Log " & strSubject
		objEMail.CreateMHTMLBody "file://" & strPath
		objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strPostServer
		objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objEmail.Configuration.Fields.Update
		objEmail.Send
	End If

Next
WScript.Quit