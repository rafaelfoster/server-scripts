clear
$WarningPreference = "SilentlyContinue"
$ErrorActionPreference = "SilentlyContinue"

$PathBkpPST = "\\SERVERNAME\SHAREFOLDER\PST"

$mbxResult = "username"

if ( $mbxResult -eq $null ) {
	clear
	Write-Output "Nenhum usuario definido, ou usuário inexistente"
	Write-Output "Saindo..."
	break
}

foreach ($usuario in $mbxResult){
	Write-Output "Desabilitando as configs do usuario: $uName ($usuario) ..."

	Write-Output "Desabilitando OWA e ActiveSync....."
	Set-CASMailbox -identity $usuario -OWAEnabled $false –ActiveSyncEnabled $false | out-null 

	Write-Output "Removendo o usuario da lista de Contatos....."
	Set-Mailbox -identity $usuario -HiddenFromAddressListsEnabled $true | out-null 
  
	$UserPST = $PathBkpPST + "\" + $usuario + ".pst"
  
	If ( ! (Test-Path ( $UserPST ))){
		Write-Output "Exportando o PST do usuario para: $UserPST"
		New-MailboxExportRequest -Mailbox $usuario -FilePath $UserPST | out-null 
		Write-Host "Progresso -> [ " -nonewline; 

		while ((Get-MailboxExportRequest -Mailbox $usuario | Where {$_.Status -eq "Queued" -or $_.Status -eq "InProgress"})){
			Write-Host -nonewline .
			sleep 30
		}
		Write-Host -nonewline " ] - OK"; Write-Host ""
		Write-Output "ExportRequest - Finalizado"
	}
	Write-Output "Apagando o Export Job....."
	Get-MailboxExportRequest -Mailbox $usuario | Remove-MailboxExportRequest -Confirm:$false | out-null 

	$title = "Remover caixa"
	$message = "Remover completamente a caixa de entrada do usuario?"
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Sim", `
		"Remover completamente a caixa de entrada do usuario (apaga a mailbox)."

	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&Nao", `
		"Não remover caixa do usuário."
		
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($no, $yes)
	$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
	switch ($result)
		{
			0 { Write-Output "-> caixa do Usuario nao removida..." }
			1 { Disable-Mailbox -Identity $usuario -Confirm:$false; Write-Output "Removendo o usuario da Base do Exchange....." }
		}
	
	Write-Output "Processo de desabilitacao do usuario $usuario finalizada"
}