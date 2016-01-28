' -------------------------------------------------------------------------------
'
' Script: Nova Assinatura.vbs
' 
' Script that generates an signature based on a form 
' filled up by a user.
' The signature will be defined in user outlook profile.
'
' Script created by Marcos Sauda ( marcosauda at hotmail dot com )
'
' -------------------------------------------------------------------------------

On Error Resume Next
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
With objUser
	strId = MsgBox("Deseja sua assinatura em ingles?",4,"Idioma")
	strNome = InputBox("Digite seu nome completo:", "Nova assinatura de Email", .givenName & Chr(32) & .lastname)
	If strNome = "" Then
		Wscript.Quit
	End If
	strCargoDepartamento = InputBox("Digite seu cargo e seu departamento:", "Nova assinatura de Email", .Description & Chr(32) & Chr(124) & Chr(32) & .Department)
	If strCargoDepartamento = "" Then
		Wscript.Quit
	End If
	strPhone = InputBox("Digite o telefone completo do seu setor:" & Chr(13) & "(Se não tiver, apenas deixe em branco ou clique em cancelar)", "Nova assinatura de Email", "+55 (13) 32")
	If strPhone <> "" Then
		strPhone = Chr(11) & "Tel. " & strPhone
	Else
		strPhone = Chr(11)
	End If
	strFax = InputBox("Digite o Fax do seu setor" & Chr(13) & "(Se não tiver, apenas deixe em branco ou clique em cancelar)", "Nova assinatura de Email", "+55 (13) 32")
	If strFax <> "" Then
		If strPhone <> Chr(11) Then
			strFax = Chr(32) & Chr(124) & Chr(32) & "Fax " & strFax
		Else
			strFax = "Fax " & strFax
		End If
	End If
	strNextel = InputBox("Digite o seu Telefone Celular Empresarial e/ou Rádio" & Chr(13) & "(Se não tiver, apenas deixe em branco ou clique em cancelar)", "Nova assinatura de Email", "+55 (13) / 55*")
	If strNextel <> "" Then
		If strPhone <> Chr(11) Then
			If strFax <> "" Then
				strNextel = Chr(11) & "Rádio/Cel. " & strNextel
			Else
				strNextel = Chr(32) & Chr(124) & Chr(32) & "Rádio/Cel. " & strNextel
			End If
		ElseIf strFax <> "" Then
			strNextel = Chr(32) & Chr(124) & Chr(32) & "Rádio/Cel. " & strNextel
		Else
			strNextel = "Rádio/Cel. " & strNextel
		End If
	End If
	strMail = Chr(11) & .EmailAddress
End With

'strCompany = "GRUPO Exemplo - 70 ANOS - "
If strId=7 Then
'strlinha1 = Chr(34) & "Simplificando Processos. Ampliando Resultados." & Chr(34)
strlinha2 = "Preserve o meio ambiente praticando sua responsabilidade. Imprima apenas se for necessário."
Else
'strlinha1 = Chr(34) & "Simplifying Processes. Heightening Results." & Chr(34)
strlinha2 = "Preserve the environment practicing your responsibility. Print only what you need most."
End If

Set objword = CreateObject("Word.Application")
With objword
	Set objDoc = .Documents.Add()
	Set objSelection = .Selection
	Set objEmailOptions = .EmailOptions
End With

Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
With objSelection
	With .Font
		.Name = "Calibri"
		.Size = 10
		.Bold = True
		.Italic = False
		.Color = RGB(31, 73, 125)
	End With
		.TypeText strNome
	'With .Font
	'	.Name = "Calibri"
	'	.Size = 8
	'	.Bold = True
	'	.Italic = False
	'	.Color = RGB(31, 73, 125)
	'End With
		.TypeText Chr(11) & strCargoDepartamento
	'With .Font
	'	.Name = "Calibri"
	'	.Size = 8
	'	.Bold = True
	'	.Italic = False
	'	.Color = RGB(31, 73, 125)
	'End With
		.TypeText strPhone & strFax & strNextel
		.TypeText strMail

	'.TypeText Chr(11)
	'objSelection.Font.Size = "9" 
	'objSelection.Font.Name = "Arial"
	'objSelection.Font.Bold = True
	'objSelection.Font.Shadow = True
	'objSelection.Font.Color = RGB(0, 0, 128)
	'objSelection.TypeText strCompany
	'objDoc.Hyperlinks.Add objSelection.Range, "http://www.Exemplo.com.br/", , , "http://www.Exemplo.com.br/"
	'objSelection.Font.Bold = True
	
	.TypeText Chr(11)
	'objSelection.Font.Size = "8"
	'objSelection.Font.italic = True
	'objSelection.Font.Color = RGB(128, 0, 0)
	'objSelection.Font.Bold = True	
	'objSelection.TypeText strlinha1
	shape = objSelection.InlineShapes.AddPicture("\\Exemplo.com.br\Public\Exemplo\signature.jpg")
	shape.Width = 100
	shape.Height = 30

	.TypeText Chr(11)
	objSelection.Font.Size = "8"
	objSelection.Font.italic = False
	objSelection.Font.Color = RGB(0, 128, 0)
	objSelection.Font.Name = "Verdana"
	objSelection.Font.Bold = True
	objSelection.TypeText strlinha2
End With

Set objSelection = objDoc.Range()
objSignatureEntries.Add "Exemplo", objSelection
objSignatureObject.NewMessageSignature = "Exemplo"
objSignatureObject.ReplyMessageSignature = "Exemplo"
objDoc.Saved = True
objword.Quit
