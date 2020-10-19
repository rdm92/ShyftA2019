Set Arg = WScript.Arguments
dim item1
dim objsubject
dim intcount
Dim i
dim savename
dim vTextFile
dim filename
dim extension
Dim t
Dim Itimestamp
dim savefolder
Dim vSenderEmailAddress
Dim vCcEmailAddress
Dim vFlagTextFileCreate
Dim RecipientObject
Dim objFso
Dim AtchmntCounter
Dim ExtnPos
Dim AtchFileName
vFlagTextFileCreate = True
savefolder = WScript.Arguments.Item(0)
'savefolder = "C:\Automation\AP 3Way Invoice Matching\Current\Email Attachments"
vTextFile = savefolder & "\Email Details Report.txt"
vFlagPDFAttachmentFound = False
Set fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set olApp = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then    'Could not get instance of Outlook, so create a new one
   Err.Clear
   Set olApp = CreateObject("Outlook.Application")
End If
on error goto 0
Set olns = olApp.GetNameSpace("MAPI")
olns.logon "Outlook",,False,True
Set objFolder = olns.GetDefaultFolder(6)
AtchmntCounter = 0
For each item1 in objFolder.Items

  if item1.Unread=true then
    objsubject = item1.subject
    if Instr(UCase(objsubject) ,"INVOICE") then
      intCount = item1.Attachments.Count
      For Each RecipientObject In item1.Recipients
         If RecipientObject.Type = 2 Then
            vCcEmailAddress = RecipientObject.Address
            if Instr(vCcEmailAddress ,"@") = 0 then
               vCcEmailAddress = RecipientObject.AddressEntry.GetExchangeUser.PrimarySmtpAddress
            End If
            Exit For
         End if
      Next
      If intcount > 0 Then
         'File format check BEGINS
         For i = 1 To intcount
            If InStr(LCase(item1.Attachments(i).filename), ".pdf") Then
               vFlagPDFAttachmentFound = True
            End if
         Next
         'File format check ENDS
         If vFlagPDFAttachmentFound = True Then
            For i = 1 To intcount
             AtchmntCounter = AtchmntCounter + 1
	     ExtnPos = InStr(1,item1.Attachments(i).filename,".pdf") - 1
	     AtchFileName = Mid(item1.Attachments(i).filename,1,ExtnPos) & "_" & AtchmntCounter & ".pdf"
              If InStr(LCase(item1.Attachments(i).filename), ".pdf") Then
                t = now()
                savename   = saveFolder & "\" & AtchFileName
                item1.Attachments(i).SaveAsFile savename
                WScript.Sleep 1000
                If item1.SenderEmailType = "SMTP" Then
                        vSenderEmailAddress = item1.SenderEmailAddress
                ElseIf item1.SenderEmailType = "EX" Then
                        vSenderEmailAddress = item1.Sender.GetExchangeUser.PrimarySmtpAddress
                End If 'If item1.SenderEmailType
                'Create InfoFile if does not exist
                If vFlagTextFileCreate = True Then
                   vFlagTextFileCreate = False
                   fso.CreateTextFile vTextFile
                End If
                Set ts = fso.OpenTextFile(vTextFile, 8, True, 0)
                ts.WriteLine AtchFileName & "," & vSenderEmailAddress & "," &  item1.Subject
                ts.Close
              end If 'If InStr(item1.Attachments(i).filename
            Next
            'Turning the unread mail to read
            item1.Unread = False
         Else
            If item1.SenderEmailType = "SMTP" Then
                    vSenderEmailAddress = item1.SenderEmailAddress
            ElseIf item1.SenderEmailType = "EX" Then
                    vSenderEmailAddress = item1.Sender.GetExchangeUser.PrimarySmtpAddress
            End If 'If item1.SenderEmailType
            'Create InfoFile if does not exist
            If vFlagTextFileCreate = True Then
               vFlagTextFileCreate = False
               fso.CreateTextFile vTextFile
            End If
            Set ts = fso.OpenTextFile(vTextFile, 8, True, 0)
            ts.WriteLine "No Attachment found" & "," & vSenderEmailAddress & "," &  item1.Subject
            ts.Close
            'Turning the unread mail to read
            item1.Unread = False
         End if
      ElseIf intcount = 0 Then
      'When no attachment is present in the mail
             'Create InfoFile if does not exist
             If item1.SenderEmailType = "SMTP" Then
                     vSenderEmailAddress = item1.SenderEmailAddress
             ElseIf item1.SenderEmailType = "EX" Then
                     vSenderEmailAddress = item1.Sender.GetExchangeUser.PrimarySmtpAddress
             End If 'If item1.SenderEmailType
             If vFlagTextFileCreate = True Then
                vFlagTextFileCreate = False
                fso.CreateTextFile vTextFile
             End If
             Set ts = fso.OpenTextFile(vTextFile, 8, True, 0)
             ts.WriteLine "No Attachment found" & "," & vSenderEmailAddress & "," &  item1.Subject
             ts.Close    
             'Turning the unread mail to read
             item1.Unread = False            
      end If 'If intcount > 0 Then
    'Below Exit as the bot is designed to just process one Retro file at a time
    'Exit For
    end If 'if Instr(objsubject ,
  end if 'if item1.Unread=true
Next
olns.logoff
Set olns  = Nothing
Set olApp = Nothing
WScript.Quit