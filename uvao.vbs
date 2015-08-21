'**************************************************************************
'* ����: UVAO.vbs                                                        *
'* ����: VBScript                                                         *
'* ����������: ������ ��� ������ �� ������                                *
'* �����: �������� �. �.                                                  * 
'**************************************************************************

'Option Explicit
On Error Resume Next

Dim Destination, i, intCount, strCommandString
Dim strArray()
Dim strPrevDate

'���� ��� ����� � ������� � ��������� �����
Const strFileName = "C:\Program Files\Tivoli\TSM\domino\domssch_uvao.log"
Const strResultFile = "C:\scripts\lastbackupuvao.txt"

intCount = 0
strPrevDate = CStr(Date - 1) '��������� ���������� ����

Set Fso = CreateObject("Scripting.FileSystemObject") '����� ����������� ������
Set txtStream = fso.OpenTextFile(strFileName, 1)
While Not txtStream.AtEndOfStream
    Str = txtStream.ReadLine()
    If Left(Str, 10) = strPrevDate Then
		ReDim Preserve strArray(intCount)
		strArray(intCount) = Str
		intCount = intCount + 1
	End If
    If Left(Str, 10) = CStr(Date) Then
		ReDim Preserve strArray(intCount)
		strArray(intCount) = Str
		intCount = intCount + 1
	End If
Wend
txtStream.Close

If intCount = 0 Then
	intCount = 1
	ReDim Preserve strArray(intCount)
	strArray(0) = "There are no backup records on " & strPrevDate & " in " & strFileName 
End If

If Fso.FileExists(strResultFile) Then '������ � ���� � ������������
	Set Destination = Fso.OpenTextFile(strResultFile, 2)
	For i = 0 to intCount - 1 
		Destination.WriteLine strArray(i)
	Next
	Destination.Close
 End If

Set WshShell = CreateObject("WScript.Shell") '�������� �����
strCommandString = "postie -host:mdi.mdi.ru -to:yusupov@mdi.ru -from:uvao_tsm@uvao.ru -s:""Last backup in UVAO"" -import -file:""C:\scripts\lastbackupuvao.txt"
WshShell.Exec(strCommandString)
strCommandString = "postie -host:mdi.mdi.ru -to:sofiich@mdi.ru -from:uvao_tsm@uvao.ru -s:""Last backup in UVAO"" -import -file:""C:\scripts\lastbackupuvao.txt"
WshShell.Exec(strCommandString)