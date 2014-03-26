Sub Save_DANF_ATTACHMENTS()

    Dim oNS As Outlook.NameSpace
    Dim oFld As Outlook.Folder
    Dim oMails As Outlook.Items
    Dim oMailItem As Outlook.MailItem
    Dim oProp As Outlook.PropertyPage
    
    Dim sSubject As String
    Dim sBody
    
    On Error GoTo Err_OL
    
    Set oNS = Application.GetNamespace("MAPI")
    'Set oFld = oNS.GetDefaultFolder(olFolderInbox)
    Set oFld = Application.ActiveExplorer.CurrentFolder
    Set oMails = oFld.Items
    
    For Each oMailItem In oMails
        sBody = oMailItem.Body
        sSubject = oMailItem.Subject
        MsgBox sSubject
        Exit Sub
    Next
    
    Exit Sub

Err_OL:
    If Err <> 0 Then
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
    Resume Next
    End If
    Exit Sub

End Sub