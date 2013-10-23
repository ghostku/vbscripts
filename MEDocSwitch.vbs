' MEDoc Switcher v0.5 2013-10-21'
' Программа переключает настройки клиента MEDoc на другой сервер '
' в программе используется разбор XML файла '
Option Explicit
Dim strServer, strMEDocPath, objFSO
' Управляющие параметры не менять'
Const CONFIG_STATION		= "station.exe.config" ' Файлы конфигурации'
Const CONFIG_EZVIT			= "ezvit.exe.config" 
Const MEDOC_DISPLAY_NAME	= "m.e.doc.station" ' Имя МЕДка в списке установленных программ'
'Const 

Set objFSO = CreateObject("Scripting.FileSystemObject")

strServer = GetArgument()

If  objFSO.FolderExists(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%ProgramData%") & "\Medoc\Station") Then
	strMEDocPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%ProgramData%") & "\Medoc\Station"
Else
	strMEDocPath = GetInstalledSoftPath(MEDOC_DISPLAY_NAME, False)
	If strMEDocPath = "" Then 
		strMEDocPath=GetInstalledSoftPath(MEDOC_DISPLAY_NAME, True)
	End If
End If

If strMEDocPath = "" Then 
	WScript.Echo "ERROR. Cannot find M.E.Doc. Probably M.E.Doc is not installed on your system"
	WScript.Quit(1)
End If

ModifyConfig strMEDocPath & "\" & CONFIG_STATION,strServer
'ModifyConfig strMEDocPath & "\" & CONFIG_EZVIT,strServer
WScript.Echo "OK. Адрес сервера изменён на: " & strServer

WScript.Quit(0)

Function GetArgument()
	If WScript.Arguments.Count < 1 Then
		WScript.Echo "ERROR. Parameter missing. Example:"
		WScript.Echo "C:\MEDocSwitch.vbs SERVER:PORT"
		WScript.Quit (1)
	Else
		GetArgument = WScript.Arguments(0)
	End If
End Function

Function ModifyConfig(strPath, strServer)
	Dim objXMLDoc, objNodeList
	Set objXMLDoc = CreateObject("Msxml2.DOMDocument")
	objXMLDoc.Load(strPath)
	If(len(objXMLDoc.Text) = 0) Then
    	WScript.Echo "ERROR. Cannot read config file." 
    	WScript.Quit (1)
	End If
	Set objNodeList = objXMLDoc.SelectSingleNode("//setting[@name='RemoteServer']/value")
	objNodeList.Text = strServer
	objXMLDoc.Save(strPath)
End Function

Function GetInstalledSoftPath(strSoftName, x64)
	' Внутренни параметры'
	Const strCOMPUTER = "."
	Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
	Const REGKEY32 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
	Const REGKEY64 = "SOFTWARE\Wow6432node\Microsoft\Windows\CurrentVersion\Uninstall\"
	Const strNAME = "DisplayName"
	Const strPATH = "InstallLocation"
	
	Dim objReg, arrSubKeys, strSubKey, strValue1, intRet1, strRegPath
	If x64 then strRegPath = REGKEY64 else strRegPath = REGKEY32
	Set objReg = GetObject("winmgmts://" & strCOMPUTER & "/root/default:StdRegProv") 
	intRet1 = objReg.EnumKey (HKLM, REGKEY64, arrSubKeys)
	If intRet1 <> 0 Then
		WScript.Echo "Can not read data from Registry"
		WScript.Quit(1)
	End If
	'WScript.Echo UBound(arrSubkeys)
	For Each strSubKey In arrSubkeys 
		intRet1 = objReg.GetStringValue(HKLM, REGKEY64 & strSubKey, strNAME, strValue1) 
  		If strValue1 = strSoftName Then 
    		intRet1 = objReg.GetStringValue(HKLM, REGKEY64 & strSubKey, strPATH, strValue1)
    		GetInstalledSoftPath = strValue1
    		Exit Function
  		End If 
	Next 

	GetInstalledSoftPath = ""
End Function