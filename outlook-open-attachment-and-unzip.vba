Dim strFolder As String
Dim strZipPassword As String
Dim str7ZipExe As String
Dim str7ZipParams As String
Dim strExec As String

Sub convertToPlain(MyMail As MailItem)
    Dim strID As String
    Dim objMail As Outlook.MailItem
    Dim objAttachments As Outlook.Attachments

    strFolder = "[PATH TO ATTACHMENT SAVE]"
        
    strZipPassword = "test"
    str7ZipExe = "[PATH TO 7ZIP EXE]"
    str7ZipParams = " x -y -p" & strZipPassword & " -o""" & strFolder & """ "
    
    strID = MyMail.EntryID
    Set objMail = Application.Session.GetItemFromID(strID)
    Set objAttachments = objMail.Attachments
    
    If objAttachments.Count = 1 Then
        atchmnt = objAttachments.Item(1)
        strFile = strFolder & atchmnt
        objAttachments.Item(1).SaveAsFile strFile
        strExec = str7ZipExe & str7ZipParams & strFile
        
        Shell strExec, vbHide
    End If
    
    Set objMail = Nothing
    Set objAttachments = Nothing
End Sub
