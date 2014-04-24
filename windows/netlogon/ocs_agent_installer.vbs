' -------------------------------------------------------------------------------
'
' Ocs Install script with plugins 
' ocs_agent_installer.vbs
' Script created by Rafael Foster (rafaelgfoster at gmail dot com)
'
' -------------------------------------------------------------------------------
On error resume Next
Const SubstituirPlugins = TRUE ' [ TRUE / FALSE ]

Set WshShell     = CreateObject("WScript.Shell")
Set objFSO       = CreateObject("Scripting.FileSystemObject")


' Variavel que define a versão minima requerida para checagem do sistema.
strMinVersionRequired="2.1.0.0"

' Variaveis de pastas de programas do Windows
strTempFolder   = wshShell.ExpandEnvironmentStrings( "%TMP%" )
strUserName     = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strProgFiles    = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES%" )
strProgFilesx86 = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES(x86)%" )

strTempFolder   = strTempFolder & "\"

	if ( inStr(LCase(strSessionName),"rdp") <> 0 OR inStr(LCase(strComputerName),"ctx") <> 0 OR inStr(LCase(strComputerName),"rod-") <> 0  ) Then
		Wscript.Quit
	End If

' Variaveis dos caminhos de instalação do Agente e seus plugins
strlatestOCSInstallFile="\\rodrimar.com.br\TI\Utils\Suporte\Programas\OCS\latest\ocspackage.exe"
strOCSInstallPluginsPath="\\rodrimar.com.br\TI\Utils\Suporte\Programas\OCS\latest\Plugins"
strOCSInstallLog="\\rodrimar.com.br\TI\AD-MGT\Logs\Log_Instalacao_OCS\"

' Esta variavel passa os parametros para instalação do OCS silenciosamente (auto install)
' Caso esteja em branco, o programa de instalação iniciara sem nenhum parametro, a nao ser que 
' o executavel ja esteja compilado para ser executado com parametros.

' Esta é a linha de comando padrao para uma instalacao silenciosa
' strOCSInstallCMDArguments="/S /NOSPLASH /UPGRADE /NO_SYSTRAY /NOW /SERVER=http://servidor_ocs/ocsinventory /user=usuario_ocs /PWD=senha_do_ocs /TAG=123456789"
strOCSInstallCMDArguments=""

' --------------------------------------------------------------------------------------------------------------
' Verificar se o diretorio de instalacao padrao OCS existe

If (objFSO.FolderExists(strProgFiles & "\OCS Inventory Agent") ) Then
	strOCSRootFolder = strProgFiles & "\OCS Inventory Agent"
Elseif (objFSO.FolderExists(strProgFilesx86 & "\OCS Inventory Agent") ) Then
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
			Wscript.Quit
		End If

	End If

Wscript.Sleep 600000
	
' Executa a instalação do OCS Inventory Agent usando os parametros especificados e copia a pasta de plugins

strOCSInstallFileName = Split(strlatestOCSInstallFile,"\",-1,1)
For Each arrName in strOCSInstallFileName
	strInstallFile = arrName
Next

strCurrentInstallLog = strOCSInstallLog & "log_installOCS_" & strUserName & "-(" & strComputerName & ").log"

objFSO.CopyFile strlatestOCSInstallFile, strTempFolder & "\", TRUE
WshShell.Run strTempFolder & "\" & strInstallFile & " " & strOCSInstallCMDArguments, 0, TRUE

if objFSO.FileExists(strCurrentInstallLog) Then
	objFSO.DeleteFile strCurrentInstallLog
End If

objFSO.MoveFile strTempFolder & "\" & "ocspackage.log", strCurrentInstallLog
objFSO.DeleteFile strTempFolder & "\" & strInstallFile

Result = WshShell.Run("cmd /c echo n | gpupdate /force", 0, TRUE)

	If (objFSO.FolderExists(strProgFiles & "\OCS Inventory Agent") ) Then
		strOCSRootFolder = strProgFiles & "\OCS Inventory Agent"
	Elseif (objFSO.FolderExists(strProgFilesx86 & "\OCS Inventory Agent") ) Then
		strOCSRootFolder = strProgFilesx86 & "\OCS Inventory Agent"
	End If

objFSO.CopyFile strOCSInstallPluginsPath & "\*", strOCSRootFolder & "\Plugins\", SubstituirPlugins

Wscript.Quit
