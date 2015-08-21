'Скрипт для получения списка NTFS разрешений на папки общего доступа  
Dim objFolder, strFolder
Dim objSD, objAce
strComputer = "."
Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile ("C:\acl.txt", ForAppending, True)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
For each objShare in colShares
    If objShare.Path <> "" Then 
		strPath = Replace (objShare.Path, "\", "\\")
    Set objFolder = GetObject("winmgmts:Win32_LogicalFileSecuritySetting.path=""" & strPath & """" )
	If objFolder.GetSecurityDescriptor(objSD) = 0 Then
		objTextFile.Writeline objShare.Name
		objTextFile.Writeline objShare.Path
		objTextFile.Writeline ("Owner: " & objSD.Owner.Name)
	For Each objAce in objSD.DACL
		objTextFile.Writeline ("Trustee: " & objAce.Trustee.Domain & "\" & objAce.Trustee.Name)
		If objAce.AceType = 0 Then
			strAceType = "Allowed"
		Else
			strAceType = "Denied"
		End If
		objTextFile.Writeline ("Ace Type: " & strAceType)
		objTextFile.Writeline ("Ace Flags:")
		If objAce.AceFlags AND 1 Then
			objTextFile.Writeline (vbTab & "Child objects that are not containers inherit permissions.")
		End If
		If objAce.AceFlags AND 2 Then
			objTextFile.Writeline (vbTab & "Child objects inherit and pass on permissions.")
		End If
		If objAce.AceFlags AND 4 Then
			objTextFile.Writeline (vbTab & "Child objects inherit but do not pass on permissions.")
		End If
		If objAce.AceFlags AND 8 Then
			objTextFile.Writeline (vbTab & "Object is not affected by but passes on permissions.")
		End If
		If objAce.AceFlags AND 16 Then
			objTextFile.Writeline (vbTab & "Permissions have been inherited.")
		End If
		objTextFile.Writeline ("Access Masks:")
		If objAce.AccessMask AND 1048576 Then
			objTextFile.Writeline (vbtab & "Synchronize")
		End If
		If objAce.AccessMask AND 524288 Then
			objTextFile.Writeline (vbtab & "Write owner")
		End If
		If objAce.AccessMask AND 262144 Then
			objTextFile.Writeline (vbtab & "Write ACL")
		End If
		If objAce.AccessMask AND 131072 Then
			objTextFile.Writeline (vbtab & "Read security")
		End If
		If objAce.AccessMask AND 65536 Then
			objTextFile.Writeline (vbtab & "Delete")
		End If
		If objAce.AccessMask AND 256 Then
			objTextFile.Writeline (vbtab & "Write attributes")
		End If
		If objAce.AccessMask AND 128 Then
			objTextFile.Writeline (vbtab & "Read attributes")
		End If
		If objAce.AccessMask AND 64 Then
			objTextFile.Writeline (vbtab & "Delete dir")
		End If
		If objAce.AccessMask AND 32 Then
			objTextFile.Writeline (vbtab & "Execute")
		End If
		If objAce.AccessMask AND 16 Then
			objTextFile.Writeline (vbtab & "Write extended attributes")
		End If
		If objAce.AccessMask AND 8 Then
			objTextFile.Writeline (vbtab & "Read extended attributes")
		End If
		If objAce.AccessMask AND 4 Then
			objTextFile.Writeline (vbtab & "Append")
		End If
		If objAce.AccessMask AND 2 Then
			objTextFile.Writeline (vbtab & "Write")
		End If
		If objAce.AccessMask AND 1 Then
			objTextFile.Writeline (vbtab & "Read")
		End If
		objTextFile.Writeline
'Wscript.Echo
    'Wscript.Echo
	Next
	End If
	End If
        objTextFile.Writeline
		objTextFile.Writeline ("==================================================================================")
		objTextFile.Writeline
Next




