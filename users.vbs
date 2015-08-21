'************************************************************************** 
'* ����: users.vbs                                                        * 
'* ����: VBScript                                                         * 
'* ����������: C����� ��� ����������� ������������� � AD.                 *
'**************************************************************************
Dim objRoot, objDomain, objOU, objContainer
Dim strName
Dim intUser

'���� �� ������� �������������
'������ ������ � ����� - ��������� ������������ � �������: ������� ��� ��������;�����;������;���������������_�������������
Const strFile = ".\users.txt"

Set Fso = CreateObject("Scripting.FileSystemObject")
Set txtStream = fso.OpenTextFile(strFile, 1)
'������ �� �����
While Not txtStream.AtEndOfStream
	Str = txtStream.ReadLine()
	MyArray = Split(Str, ";", -1, 1)
	Set objRootDSE = GetObject("LDAP://rootDSE")
	Set objContainer = GetObject("LDAP://" & MyArray(3))
	'�������� ������������
	Set objLeaf = objContainer.Create("User", "cn=" & MyArray(0))
	'������� ������� �������-������������
	objLeaf.Put "sAMAccountName", MyArray(1)
	objLeaf.Put "userAccountControl", 544
	objLeaf.Put "displayName", MyArray(0)
	MyArrayName = Split(MyArray(0), " ", -1, 1)
	objLeaf.Put "sn", MyArrayName(0)
	objLeaf.Put "givenName", MyArrayName(1)
	objLeaf.Put "userPrincipalName", MyArray(1) & "@mgsn.loc"
 	'���������� ������
	objLeaf.SetInfo
	Set objUser = GetObject("LDAP://CN=" & MyArray(0) & "," & MyArray(3))
	'������� ������
	objLeaf.SetPassword MyArray(2)
Wend

WScript.Echo "Users created"
WScript.quit