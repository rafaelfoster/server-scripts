' -------------------------------------------------------------------------------
'
' Ocs Install script with plugins 
' ocs_agent_installer.vbs
' Script created by Rafael Foster (rafaelgfoster at gmail dot com)
'
' -------------------------------------------------------------------------------
On error resume Next
Set WshShell     = CreateObject("WScript.Shell")
Set objRegEx     = CreateObject("VBScript.RegExp")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set SystemSet    = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE
Const SubstituirPlugins = TRUE ' [ TRUE / FALSE ]

' -------------------------------------------------
' Definição do Regex 
objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = "\-\d{5,6}$"

' Variavel que define a versão minima requerida para checagem do sistema.
strMinVersionRequired="2.1.0.0"

' Variaveis de pastas de programas do Windows
strTempFolder   = wshShell.ExpandEnvironmentStrings( "%TMP%" )
strUserName     = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strProgFiles    = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES%" )
strProgFilesx86 = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES(x86)%" )

strTempFolder   = strTempFolder & "\"

For each System in SystemSet 
	SysVersion     = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
Next

if ( inStr(LCase(SysVersion),"server") <> 0 ) Then
	Wscript.Quit
End If

' Variaveis dos caminhos de instalação do Agente e seus plugins
strlatestOCSInstallFile="\\example.com.br\TI\Utils\Suporte\Programas\OCS\latest\ocspackage.exe"
strlatestOCSDefaultInstallFile="\\example.com.br\TI\Utils\Suporte\Programas\OCS\latest\ocspackage_default.exe"
strOCSInstallPluginsPath="\\example.com.br\TI\Utils\Suporte\Programas\OCS\latest\Plugins"
strOCSInstallLog="\\example.com.br\TI\AD-MGT\Logs\Log_Instalacao_OCS\"

' Define o nome do arquivo de Log
strCurrentInstallLog = strOCSInstallLog & "log_installOCS_" & "(" & strComputerName & ").log"

' Esta variavel passa os parametros para instalação do OCS silenciosamente (auto install)
' Caso esteja em branco, o programa de instalação iniciara sem nenhum parametro, a nao ser que 
' o executavel ja esteja compilado para ser executado com parametros.

' Esta é a linha de comando padrao para uma instalacao silenciosa
' strOCSInstallCMDArguments="/S /NOSPLASH /UPGRADE /NO_SYSTRAY /NOW /SERVER=http://servidor_ocs/ocsinventory /user=usuario_ocs /PWD=senha_do_ocs /TAG=123456789"
strOCSInstallCMDArguments="/S /NOSPLASH /UPGRADE /NO_SYSTRAY /NOW /SERVER=http://ocs.example.com.br/ocsinventory /user=ocs /PWD=r0dr!m@rocs"

' --------------------------------------------------------------------------------------------------------------
' Verificar se o diretorio de instalacao padrao OCS existe

' Verifica se o Client possui alguma TAG definida
If objRegEx.Test( strComputerName ) Then

	BP = Split(strComputerName,"-")
	For each BPPat in BP	
		strPatrimonio = BPPat
	Next
End If

If (objFSO.FileExists(strProgFiles & "\OCS Inventory Agent\OCSInventory.exe") ) Then
	strOCSRootFolder = strProgFiles & "\OCS Inventory Agent"
Elseif (objFSO.FileExists(strProgFilesx86 & "\OCS Inventory Agent\OCSInventory.exe") ) Then
	strOCSRootFolder = strProgFilesx86 & "\OCS Inventory Agent"
End If

	strBinOCSInventory = strOCSRootFolder & "\OCSInventory.exe"

	'---------------------------------------------------------------------------------------------------------------
	' Se o diretorio existir, verificar a versão do binario principal do sistema para detectar sua versão instalada.

	if ( objFSO.FileExists(strBinOCSInventory) ) Then

		' Determina a versão atual do executavel do OCS Inventory
		strCurOCSVersion = objFSO.GetFileVersion(strBinOCSInventory)

		' Caso o agent instalado seja maior ou igual a versão requerida, apenas os plugins são atualizados
		If ( strCurOCSVersion >= strMinVersionRequired ) Then
			cmd = strOCSInstallPluginsPath & "\*" & strOCSRootFolder & "\Plugins\"
			Result = WshShell.Run("cmd /c echo n | gpupdate /force",0,true)
			objFSO.CopyFile strOCSInstallPluginsPath & "\*", strOCSRootFolder & "\Plugins\", SubstituirPlugins
			
			' Adicionar TAG
			Wscript.Run strBinOCSInventory & " /NOW " & "/TAG=""" & strPatrimonio & """"

			Wscript.Quit
		End If

	End If

Wscript.Sleep 600000
	
' Executa a instalação do OCS Inventory Agent usando os parametros especificados e copia a pasta de plugins

' -------------------------------------------------------------------------------
' Tenta parar o servico do OCS para iniciar a instalacao
strComputerName = "."
 Set objWMIService = GetObject("winmgmts:" _ 
	& "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2") 

Set colServices = objWMIService.ExecQuery _ 
	("Select * from Win32_Service") 

strCurOCSVersion = objFSO.GetFileVersion(strBinOCSInventory)

For Each Service in colServices
	if instr(Service.Name,"OCS") > 0 Then

		strOCSServiceMSG = "Versao atual do OCS: " & strCurOCSVersion & vbCrlf & vbCr & _
						   "Servico " & Service.Name & " detectado." & vbCrlf & vbCr & _
						   "Tentando finalizar o servico...."

		if instr(Service.State,"Running") > 0 Then
			strServiceStopStatusCOD = Service.StopService()
		End If

		Select Case strServiceStopStatusCOD
			Case 0 
				strServiceStopStatusMSG = "SUCCESS!"
			Case 1 
				strServiceStopStatusMSG = "NOT SUPPORTED"
			Case 2 
				strServiceStopStatusMSG = "PERMISSION DENIED"
			Case 3 
				strServiceStopStatusMSG = "ERROR: DEPENDENTS SERVICES ARE RUNNING"
			Case 4 
				strServiceStopStatusMSG = "CANNONT SEND STOP CONTROL CODE TO SERVICE"
			Case 5 
				strServiceStopStatusMSG = "REQUEST CODE INVALID"
			Case 6 
				strServiceStopStatusMSG = "SERVICE NOT STARTED"
			Case 7 
				strServiceStopStatusMSG = "SERVICE NOT RESPONDING"
			Case 8 
				strServiceStopStatusMSG = "UNKNOWN ERROR ON START SERVICE"
			Case 9
				strServiceStopStatusMSG = "EXECUTABLE PATH NOT FOUNDED"
			Case 10 
				strServiceStopStatusMSG = "UNKNOWN ERROR ON START SERVICE"
			Case 11
				strServiceStopStatusMSG = "UNKNOWN ERROR ON START SERVICE"
		End Select

		strOCSServiceMSG = strOCSServiceMSG & " " & strServiceStopStatusMSG

	End if
Next

If objRegEx.Test( strComputerName ) Then
	BP = Split(strComputerName,"-")
	For each BPPat in BP	
		strPatrimonio = BPPat
	Next
	strOCSInstallBin = strlatestOCSDefaultInstallFile
	strOCSInstallArgs = strOCSInstallCMDArguments & "/TAG=" & strPatrimonio
Else
	strOCSInstallBin = strlatestOCSInstallFile
	strOCSInstallArgs = ""
End If

' Define o nome do executavel do OCS Install
strOCSInstallFileName = Split(strOCSInstallBin,"\",-1,1)
For Each arrName in strOCSInstallFileName
	strInstallFile = arrName
Next

' Copia o arquivo do OCS para a pasta %TEMP% do usuário e executa-o com os parametros definidos em 'strOCSInstallCMDArguments'
objFSO.CopyFile strOCSInstallBin, strTempFolder & "\", TRUE
WshShell.Run strTempFolder & "\" & strInstallFile & " " & strOCSInstallCMDArguments, 0, TRUE

' Caso o arquivo de log ja exista, substitui o mesmo
objFSO.CopyFile strTempFolder & "\" & "ocspackage.log", strCurrentInstallLog, TRUE

Set objOCSLogFile = objFSO.OpenTextFile(strTempFolder & "\" & "ocspackage.log", ForReading, True)
strFileContent = objOCSLogFile.ReadAll

' ------------------------------------------------------------------------------------------------------------------------------------------------
' Inicia a geração do Arquivo de Log
If objFSO.FileExists(strCurrentInstallLog) Then
	Set objLogFile = objFSO.OpenTextFile(strCurrentInstallLog, ForWriting, True)
Else
	Set objLogFile = objFSO.CreateTextFile(strCurrentInstallLog)
End If

objLogFile.Write strOCSServiceMSG
objLogFile.Writeline " " 
objLogFile.Write strFileContent

objFSO.DeleteFile strTempFolder & "\" & strInstallFile
objFSO.DeleteFile strTempFolder & "\" & "ocspackage.log"

Result = WshShell.Run("cmd /c echo n | gpupdate /force", 0, TRUE)

	If (objFSO.FolderExists(strProgFiles & "\OCS Inventory Agent") ) Then
		strOCSRootFolder = strProgFiles & "\OCS Inventory Agent"
	Elseif (objFSO.FolderExists(strProgFilesx86 & "\OCS Inventory Agent") ) Then
		strOCSRootFolder = strProgFilesx86 & "\OCS Inventory Agent"
	End If

objFSO.CopyFile strOCSInstallPluginsPath & "\*", strOCSRootFolder & "\Plugins\", SubstituirPlugins

Wscript.Quit
