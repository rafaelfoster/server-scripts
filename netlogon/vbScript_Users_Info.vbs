On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
Set WshShell     = CreateObject("WScript.Shell")
Set WshNetwork   = CreateObject("WScript.Network")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set SystemSet    = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const HKEY_LOCAL_MACHINE = &H80000002

Dim aOffID(4,1)
aOffID(0,0) = "XP"
aOffID(0,1) = "10.0"
aOffID(1,0) = "2003"
aOffID(1,1) = "11.0"
aOffID(2,0) = "2007"
aOffID(2,1) = "12.0"
aOffID(3,0) = "2010"
aOffID(3,1) = "14.0"
aOffID(4,0) = "2013"
aOffID(4,1) = "15.0"

Const Log_Anexar = 2 '( 1 = Read, 2 = Write, 8 = Append )

' --------------------------------------------------------------------------------------------
' Definição de Variaveis
strComputer = "."
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strLogonServer = wshShell.ExpandEnvironmentStrings( "%LOGONSERVER%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strProgFiles = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES%" )
strProgFilesx86 = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES(x86)%" )
Log_File = "\\rodrimar.com.br\Ti\AD-MGT\Logs\Log_usuarios_do_dominio\Log_UserInfo_" & strUserName & ".txt" 

Log_Header = "Data: " & date & " - " & time

if ( inStr(LCase(strSessionName),"rdp") <> 0 OR inStr(LCase(strComputerName),"ctx") <> 0 ) Then
	Wscript.Quit
End If

' Aguardar 3 minutos antes de iniciar
Wscript.Sleep 180000

If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, Log_Anexar, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If

For each System in SystemSet 
	SysVersion     = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
	dtmConvertedDate.Value = System.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
	SysInstallDate = dtmInstallDate
Next

' Get System TAG from Win32_BIOS.SerialNumber
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colBIOS = objWMIService.ExecQuery _
    ("Select * from Win32_BIOS")

For each objBIOS in colBIOS
    strServiceTag = objBIOS.SerialNumber
Next

objCriaLog.WriteLine "-------------------------------------------------------------------------------------------------------"
objCriaLog.WriteLine Log_Header
objCriaLog.WriteLine

' --------------------------------------------------------------------------------------------------------------------------
' Informações do Sistema

objCriaLog.WriteLine "-------[ System Information ]---------------------------------------------------"
objCriaLog.WriteLine
objCriaLog.WriteLine "Usuario               : " & strUserName
objCriaLog.WriteLine "Estacao de Trabalho   : " & strComputerName
objCriaLog.WriteLine "Servidor de Logon     : " & strLogonServer
objCriaLog.WriteLine "Sistema Operacional   : " & SysVersion
objCriaLog.WriteLine "Data de Instalacao    : " & SysInstallDate
objCriaLog.WriteLine "Service TAG           : " & strServiceTag
objCriaLog.WriteLine
objCriaLog.WriteLine

objCriaLog.WriteLine "-------[ Licence Information ]--------------------------------------------------"
objCriaLog.WriteLine 

Set oCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
oCtx.Add "__ProviderArchitecture", 64

Set oLocator = CreateObject("Wbemscripting.SWbemLocator")
Set oReg = oLocator.ConnectServer("", "root\default", "", "", , , , oCtx).Get("StdRegProv")

osType = 32
oReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "PROCESSOR_ARCHITECTURE", osProc
If osProc = "AMD64" Then osType = 64

wow = ""
If osType = "64" Then wow = "WOW6432Node\"
schKey97 "SOFTWARE\" & wow & "Microsoft\"
schKey2K "Office", "SOFTWARE\" & wow & "Microsoft\Office\9.0\", Array("0000","0001","0002","0003","0004","0010","0011","0012","0013","0014","0016","0017","0018","001A","004F"), "78E1-11D2-B60F-006097C998E7"
schKey2K "Visio", "SOFTWARE\" & wow & "Microsoft\Visio\6.0\", Array("B66F45DC"), "853B-11D3-83DE-00C04F3223C8"

For a = LBound(aOffID, 1) To UBound(aOffID, 1)
  schKey "SOFTWARE\Wow6432Node\Microsoft\Office\" & aOffID(a,1) & "\Registration", false
  schKey "SOFTWARE\Microsoft\Office\" & aOffID(a,1) & "\Registration", true
Next

Sub writeXML(oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote)
  objCriaLog.WriteLine _
  oProd & vbTab & vbTab & oKey
End Sub

objCriaLog.WriteLine
objCriaLog.WriteLine "-------[ Network Information ]--------------------------------------------------"

NetCardInformations = split(GetNetworkInformation(),";")
For Each NetInfo In NetCardInformations
	objCriaLog.WriteLine NetInfo
Next

' --------------------------------------------------------------------------------------------------------------------------
' Unidades Mapeadas

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_MappedLogicalDisk")

objCriaLog.WriteLine
objCriaLog.WriteLine
objCriaLog.WriteLine "-------[ Unidades Mapeadas neste computador ]-----------------------------------"
objCriaLog.WriteLine
For Each objItem in colItems
	objCriaLog.WriteLine "Unidade: " & objItem.Name & " - em: " & objItem.ProviderName
Next

objCriaLog.WriteLine
objCriaLog.WriteLine
objCriaLog.WriteLine "-------[ Softwares em Conformidade ]--------------------------------------------"
objCriaLog.WriteLine

If (objFSO.FolderExists(strProgFiles & "\OCS Inventory Agent") ) Then

	strOCSRootFolder = strProgFiles & "\OCS Inventory Agent"
	strBinOCSInventory = strOCSRootFolder & "\OCSInventory.exe"
	
	if ( objFSO.FileExists(strBinOCSInventory) ) Then
		objCriaLog.WriteLine "OCS Inventory NG: versao " & objFSO.GetFileVersion(strBinOCSInventory)
	End If

Elseif (objFSO.FolderExists(strProgFilesx86 & "\OCS Inventory Agent") ) Then

	strOCSRootFolder = strProgFilesx86 & "\OCS Inventory Agent"
	strBinOCSInventory = strOCSRootFolder & "\OCSInventory.exe"

	if ( objFSO.FileExists(strBinOCSInventory) ) Then
		objCriaLog.WriteLine "OCS Inventory NG: versao " & objFSO.GetFileVersion(strBinOCSInventory)
	End If

Else

	objCriaLog.WriteLine "OCS Inventory NG: ** NAO INSTALADO **"
	
End If


objCriaLog.WriteLine
objCriaLog.WriteLine
objCriaLog.WriteLine "-------[ Impressoras instaladas ]-----------------------------------------------"
GetPrintersInformation = Split(GetPrinters,";")
For Each Printer in GetPrintersInformation
	objCriaLog.WriteLine Printer
Next

objCriaLog.WriteLine

' ---------------------------------------------------------------------------------------------------------------------------
' Coleta Shortcuts e URLs do usuário

objCriaLog.WriteLine
objCriaLog.WriteLine "-------[ Atalhos de Area de Trabalho ]------------------------------------------"
objCriaLog.WriteLine
objStartFolder = strUserPath & "\Desktop\"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

For Each objFile in colFiles
	
	strFileExt = objFSO.GetExtensionName(objFile.Name)

	Select Case strFileExt
	
		Case "lnk"
			Set objShortcut = wshShell.CreateShortcut(objFile.Path)
			objCriaLog.WriteLine "Atalho:  " & objFile.Name

			if ( InStr(Lcase(SysOperation), "xp") <> 0 ) Then
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			Else
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			End If
			
			objCriaLog.WriteLine
		
		Case "url"
			Set objShortcut = wshShell.CreateShortcut(objFile.Path)
			objCriaLog.WriteLine "URL:     " & objFile.Name
			
			if ( InStr(LCase(SysOperation), "xp") <> 0 ) Then
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			Else
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			End If
			
			objCriaLog.WriteLine 
	
	End Select
	
	Set objShortcut = Nothing

	'-----------------------------

Next

objCriaLog.Close

Function GetNetworkInformation()

	' List IP Configuration Data

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colAdapters = objWMIService.ExecQuery _
		("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	 
	n = 1
	 
	For Each objAdapter in colAdapters
	   NetworkInformation = NetworkInformation & ";" &  "Adaptador de Rede " & n
	   NetworkInformation = NetworkInformation & ";" &  "==================="
	   NetworkInformation = NetworkInformation & ";" &  " Placa de Rede:            " & objAdapter.Description
	   NetworkInformation = NetworkInformation & ";" &  " MAC Address:              " & objAdapter.MACAddress
	 
	   If Not IsNull(objAdapter.IPAddress) Then
		  For i = 0 To UBound(objAdapter.IPAddress)
			 NetworkInformation = NetworkInformation & ";" & " Endereco IP:              " & objAdapter.IPAddress(i)
		  Next
	   End If
	 
	   If Not IsNull(objAdapter.IPSubnet) Then
		  For i = 0 To UBound(objAdapter.IPSubnet)
			NetworkInformation = NetworkInformation & ";" &  " Mascara de Rede:          " & objAdapter.IPSubnet(i)
		  Next
	   End If
	 
	   If Not IsNull(objAdapter.DefaultIPGateway) Then
		  For i = 0 To UBound(objAdapter.DefaultIPGateway)
			 NetworkInformation = NetworkInformation & ";" &  " Gateway Padrao:           " & _
				objAdapter.DefaultIPGateway(i)
		  Next
	   End If

	   NetworkInformation = NetworkInformation & ";" &  " Lista de Servidores DNS:"

	   If Not IsNull(objAdapter.DNSServerSearchOrder) Then
		  For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
			NetworkInformation = NetworkInformation & ";" &  vbTab & vbTab & vbTab & "   " & objAdapter.DNSServerSearchOrder(i)
		  Next
	   End If

	   If Not IsNull(objAdapter.DNSDomainSuffixSearchOrder) Then
		  For i = 0 To UBound(objAdapter.DNSDomainSuffixSearchOrder)
			 NetworkInformation = NetworkInformation & ";" &  " Sufixo DNS:               " & _
				objAdapter.DNSDomainSuffixSearchOrder(i)
		  Next
	   End If

	   n = n + 1

	Next

	GetNetworkInformation = NetworkInformation

End Function


Function GetPrinters()

	Set objWMIService = GetObject("winmgmts:" _ 
		& "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2") 

	Set colInstalledPrinters = objWMIService.ExecQuery _ 
		("Select * from Win32_Printer") 
		
	For Each objPrinter in colInstalledPrinters
		DefaultPrinter = ""
		If objPrinter.Default = True Then
			DefaultPrinter = " (Padrao)"
		End If
		GetPrinters = GetPrinters & ";" & "Impressora: " & objPrinter.Name & DefaultPrinter
		GetPrinters = GetPrinters & ";" & "Local     : " & objPrinter.Location
		GetPrinters = GetPrinters & ";" & "Porta/IP  : " & objPrinter.PortName
		GetPrinters = GetPrinters & "  ;"
	Next 

End Function


''--------------------------------------------------------------------------------------------------------------
'' msofficekey 2.2.3 (13/02/2013)
'' Plugin for OCS Inventory NG 2.x
'' Creative Commons BY-NC-SA 3.0
'' Nicolas DEROUET (nicolas.derouet[gmail]com)
'On Error Resume Next

' ------------------------------ 
' SUB XML

Sub schKey97(regKey)
  oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Office\8.0", "BinDirPath", oDir97
  If IsNull(oDir97) Then Exit Sub
  oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Microsoft Reference\BookshelfF\96L", "PID", oProdID
  oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Windows\CurrentVersion\Uninstall\Office8.0", "DisplayName", oProd
  oInstall = "1"
  If IsNull(oProd) Then
    oInstall = "0"
    oProd = "Microsoft Office 97"
  End If
  writeXML "97",oProd,oProdID,32,"",oInstall,"",""
End Sub

Sub schKey2K(name, regKey, guid1, guid2)
  oProd = Null
  oInstall = "0"
  oReg.GetBinaryValue HKEY_LOCAL_MACHINE, regKey & "Registration\DigitalProductID", "", aDPIDBytes
  oKey = ""
  If Not IsNull(aDPIDBytes) Then oKey = decodeKey(aDPIDBytes)

  oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Registration\ProductID", "", oProdID
  If IsNull(oProdID) Then Exit Sub

  oReg.EnumKey HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\", aKeys
  If Not IsNull(aKeys) Then
    For Each guid In aKeys
      If UCase(Right(guid,Len(guid)-InStr(guid,"-"))) = guid2 & "}" Then
        For i = LBound(guid1) To UBound(guid1)
          If UCase(Left(guid,Len(guid1(i)) + 1)) = "{" & guid1(i) Then
            oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\" & guid, "DisplayName", oProd
            oGUID = guid
            oInstall = "1"
          End If
        Next
      End If
    Next
  End If

  If IsNull(oProd) Then oProd = "Microsoft " & name & " 2000"
  writeXML "2000",oProd,oProdID,32,oGUID,oInstall,oKey,""
End Sub

Sub schKey(regKey, likeOS)
  oReg.GetBinaryValue HKEY_LOCAL_MACHINE, regKey, "DigitalProductID", aDPIDBytes
  If IsNull(aDPIDBytes) Then
    oReg.EnumKey HKEY_LOCAL_MACHINE, regKey, aGUIDKeys
    If Not IsNull(aGUIDKeys) Then
      For Each GUIDKey In aGUIDKeys
        schKey regKey & "\" & GUIDKey, likeOS
      Next
    End If
  Else
    oVer = aOffID(a,0)
    oProd = Null
    oKey = decodeKey(aDPIDBytes)
    oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ProductID", oProdID
    oBit = osType
    If Not likeOS Then oBit = 32
    oGUID = Right(regKey,InStr(StrReverse(regKey),"\")-1)
    oInstall = "1"
    wow = ""
    If Not likeOS Then wow = "WOW6432Node\"

    oEdit = ""
    If (oVer = "2010" Or oVer = "2013") Then
      For i = 280 to 320 Step 2
        If aDPIDBytes(i) <> 0 Then oEdit = oEdit & Chr(aDPIDBytes(i))
      Next
    End If
    oNote = oEdit

    If IsNull(oProd) And (oVer = "2010" Or oVer = "2013") Then
      kEdit = UCase(oEdit)
      If Mid(oGUID,11,4) = "003D" Then kEdit = "SingleImage"
      oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\Office" & Left(aOffID(a,1),2) & "." & kEdit, "DisplayName", oProd
    End If

    If IsNull(oProd) Then _
      oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\" & oGUID, "DisplayName", oProd

    If IsNull(oProd) Then
      oInstall = "0"
      oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ProductName", oProd
      If IsNull(oProd) Then oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ConvertToEdition", oProd

      ' Office Visio XP
      If IsNull(oProd) And (oVer = "XP") Then
        oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Office\XP\Common\ProductVersion", "LastProduct", pVer
        ' Original / SP1 / SP2
        If ((pVer = "10.0.525") Or (pVer = "10.1.2514") Or (pVer = "10.2.5110")) Then
          oProd = "Microsoft Office Visio XP"
        End If
      End If

      ' Office Visio Viewer 2003
      If IsNull(oProd) And (oVer = "2003") And (oKey = "MF4QD-3T4PM-26X66-4KH7R-QGTYT") Then
        oProd = "Microsoft Office Visio Viewer 2003"
      End If

      If IsNull(oProd) Then oProd = "Unidentifiable Office " & oVer
    End If
    writeXML oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote
  End If
End Sub

'Sub writeXML(oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote)
'  Wscript.Echo _
'  oProd & vbTab & vbTab & oKey
'End Sub

Function decodeKey(iValues)
  Dim arrDPID, foundKeys
  arrDPID = Array()
  foundKeys = Array()

  Select Case (UBound(iValues))
    Case 255:  ' 2000
      range = Array(52,66)
    Case 163:  ' XP, 2003, 2007
      range = Array(52,66)
    Case 1271: ' 2010, 2013
      range = Array(808,822)
    Case Else
      Exit Function
  End Select

  charset = "BCDFGHJKMPQRTVWXY2346789"

  For i = range(0) to range(1)
    ReDim Preserve arrDPID( UBound(arrDPID) + 1 )
    arrDPID( UBound(arrDPID) ) = iValues(i)
  Next

  withN = (arrDPID(UBound(arrDPID)) \ 6) And 1
  arrDPID(UBound(arrDPID)) = (arrDPID(UBound(arrDPID)) And &HF7) Or ((withN And 2) * 4)

  For i = 24 To 0 Step -1
    k = 0
    For j = 14 To 0 Step -1
      k = k * 256 Xor arrDPID(j)
      arrDPID(j) = k \ 24
      k = k Mod 24
    Next
    strProductKey = Mid(charset, k+1, 1) & strProductKey
  Next

  If (withN = 1) Then
    keypart = Mid(strProductKey,2,k)
    strProductKey = Replace(strProductKey, keypart, keypart & "N", 2, 1, 0)
    If k = 0 Then strProductKey = "N" & strProductKey
  End If

  decodeKey = ""
  For i = 1 To 25
    decodeKey = decodeKey & Mid(strProductKey,i,1)
    If i Mod 5 = 0 And i <> 25 Then decodeKey = decodeKey & "-"
  Next
End Function
