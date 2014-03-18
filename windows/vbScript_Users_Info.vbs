' This script generates a Log File With the Following informations:
'
' System information: username, Computername, LogonServer, Operational System, InstallDate
' Licence Information: Windows and Office Serial Keys
' Network Cards: NIC, MAC, IP, Mask, Gateway, DNS Servers, DNS Suffix
' Mapped Network Drivers: Letters and Paths of all Mapped Network Drivers
' Installed Printers: Printer Name, Location and Port of all printers
' Desktop Shortcuts: Name and Destination Path of All Desktop Shortcuts

On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
Set WshShell     = CreateObject("WScript.Shell")
Set WshNetwork   = CreateObject("WScript.Network")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set SystemSet    = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

' --------------------------------------------------------------------------------------------
' Definição de Variaveis
strComputer = "."
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strLogonServer = wshShell.ExpandEnvironmentStrings( "%LOGONSERVER%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Log_File = "\\SERVIDOR\COMPARTILHAMENTO\Logs\Log_userinfo_" & strUserName & ".txt" 

Log_Header = "Data: " & date & " - " & time

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSLogonSetting")
For Each objItem in colItems
	If ( Len(objItem.TerminalName) <> 0 ) Then
		Wscript.Quit
	End If
Next

' Aguardar 3 minutos antes de iniciar
'Wscript.Sleep 180000

If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, ForWriting, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If

For each System in SystemSet 
	SysVersion     = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
	dtmConvertedDate.Value = System.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
	SysInstallDate = dtmInstallDate
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
objCriaLog.WriteLine
objCriaLog.WriteLine

objCriaLog.WriteLine "-------[ Licence Information ]--------------------------------------------------"
objCriaLog.WriteLine 

WinKey = GetWinKey
OfficeKeys = GetOfficeKey("10.0") & GetOfficeKey("11.0") & GetOfficeKey("12.0") & GetOfficeKey("14.0") & GetOfficeKey("15.0")

objCriaLog.WriteLine WinKey
objCriaLog.WriteLine OfficeKeys
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

Function GetOfficeKey(sVer)
    On Error Resume Next
    Dim arrSubKeys
    Set wshShell = WScript.CreateObject( "WScript.Shell" )
    sBit = wshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
    if sBit <> "%ProgramFiles(x86)%" then
   sBit = "Software\wow6432node"
    else
   sBit = "Software"
    end if
    Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    objReg.EnumKey HKEY_LOCAL_MACHINE, sBit & "\Microsoft\Office\" & sVer & "\Registration", arrSubKeys
    Set objReg = Nothing
    if IsNull(arrSubKeys) = False then
        For Each Subkey in arrSubKeys
       if lenb(other) < 1 then other = wshshell.RegRead("HKLM\" & sBit & "\Microsoft\Office\" & sVer & "\Registration\" & SubKey & "\ProductName")
       if ucase(right(SubKey, 7)) = "0FF1CE}" then
                Set wshshell = CreateObject("WScript.Shell")
           key = ConvertToKey(wshshell.RegRead("HKLM\" & sBit & "\Microsoft\Office\" & sVer & "\Registration\" & SubKey & "\DigitalProductID"))
      oem = ucase(mid(wshshell.RegRead("HKLM\" & sBit & "\Microsoft\Office\" & sVer & "\Registration\" & SubKey & "\ProductID"), 7, 3))
        edition = wshshell.RegRead("HKLM\" & sBit & "\Microsoft\Office\" & sVer & "\Registration\" & SubKey & "\ProductName")
      if err.number <> 0 then 
          edition = other
            err.clear
      end if
           Set wshshell = Nothing
            'if oem <> "OEM" then oem = "Retail"
           if lenb(final) > 1 then
				final = final & vbnewline & final
           else
				final = edition & ":  " & vbTab & key 
			end if

       end if
        Next
   GetOfficeKey = final & vbnewline
    End If
End Function

Function GetWinKey()
    
	Set wshshell = CreateObject("WScript.Shell")
		edition = wshshell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
		oem = ucase(mid(wshshell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductID"), 7, 3))
		key = GetKey("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")
    
	set wshshell = Nothing
		'if oem <> "OEM" then oem = "Retail"
		GetWinKey = edition & ":  " & vbTab & vbTab & vbTab & key

End Function

Function GetKey(sReg)
    Set wshshell = CreateObject("WScript.Shell")
    GetKey = ConvertToKey(wshshell.RegRead(sReg))
    Set wshshell = Nothing
End Function

Function ConvertToKey(key)
    Const KeyOffset = 52
    i = 28
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    Do
        Cur = 0
        x = 14
        Do
            Cur = Cur * 256
            Cur = key(x + KeyOffset) + Cur
            key(x + KeyOffset) = (Cur \ 24) And 255
            Cur = Cur Mod 24
            x = x - 1
        Loop While x >= 0
        i = i - 1
        KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
        If (((29 - i) Mod 6) = 0) And (i <> -1) Then
            i = i - 1
            KeyOutput = "-" & KeyOutput
        End If
    Loop While i >= 0
    ConvertToKey = KeyOutput
End Function