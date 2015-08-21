'************************************************************************** 
'* Файл: users.vbs                                                        * 
'* Язык: VBScript                                                         * 
'* Назначение: Cкрипт для регистрации пользователей в AD.                 *
'**************************************************************************
Dim objRoot, objDomain, objOU, objContainer
Dim strName
Dim intUser

'Файл со списком пользователей
'Каждая строка в файле - отдельный пользователь в формате: Фамилия Имя Отчество;Логин;Пароль;Организационное_подразделение
Const strFile = ".\users.txt"

Set Fso = CreateObject("Scripting.FileSystemObject")
Set txtStream = fso.OpenTextFile(strFile, 1)
'Чтение из файла
While Not txtStream.AtEndOfStream
	Str = txtStream.ReadLine()
	MyArray = Split(Str, ";", -1, 1)
	Set objRootDSE = GetObject("LDAP://rootDSE")
	Set objContainer = GetObject("LDAP://" & MyArray(3))
	'Создание пользователя
	Set objLeaf = objContainer.Create("User", "cn=" & MyArray(0))
	'Задание свойств объекта-пользователя
	objLeaf.Put "sAMAccountName", MyArray(1)
	objLeaf.Put "userAccountControl", 544
	objLeaf.Put "displayName", MyArray(0)
	MyArrayName = Split(MyArray(0), " ", -1, 1)
	objLeaf.Put "sn", MyArrayName(0)
	objLeaf.Put "givenName", MyArrayName(1)
	objLeaf.Put "userPrincipalName", MyArray(1) & "@mgsn.loc"
 	'Сохранение данных
	objLeaf.SetInfo
	Set objUser = GetObject("LDAP://CN=" & MyArray(0) & "," & MyArray(3))
	'Задание пароля
	objLeaf.SetPassword MyArray(2)
Wend

WScript.Echo "Users created"
WScript.quit