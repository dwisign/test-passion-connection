

' Max's macros for Mass E-mail sending with attachments. Comments send to clustermass2@gmail.com

Sub SendEmail(what_address As String, subject_line As String, mail_body As String, path As String, year As String, month As String, maindir As String, fromfield As String)
Dim backslash As String
Dim filename As String
Dim fullpath As String
Dim oAccount As Outlook.account
backslash = "\"
filename = "attached_documents.zip"
fullpath = maindir & backslash & path & backslash & year & backslash & month & backslash & filename
Dim olApp As Outlook.application
Set olApp = CreateObject("Outlook.Application")
    ' Lets check if the account that user is trying to use exists in outlook session
    
    Dim checkifaccountexist As Boolean
    
    For Each oAccount In Outlook.application.Session.accounts
    If oAccount = fromfield Then
    checkifaccountexist = True
    End If
    Next
    
    If checkifaccountexist Then
    
    ' Lets check if the account that user is trying to use exists in outlook session

    
        Dim olMail As Outlook.MailItem
        For Each oAccount In Outlook.application.Session.accounts
        If oAccount = fromfield Then
        Set olMail = olApp.CreateItem(olMailItem)
        olMail.To = what_address
        olMail.Subject = subject_line
        olMail.Body = mail_body
        olMail.Attachments.Add fullpath
        olMail.SendUsingAccount = oAccount
        olMail.Send
        End If
        Next
    Else
        MsgBox fromfield & " was not found in current Outlook session! Please correct your Outlook account name in Excel. If you have just created new account, restart Outlook and then try again. "
        End
    End If

End Sub




Sub SendMassEmail()
minutes = Worksheets(1).Cells(6, 13)
seconds = Worksheets(1).Cells(7, 13)
hours = 0
waittime = TimeSerial(hours, minutes, seconds)
If Worksheets(1).Cells(5, 13) = 1 Then
logging = True
Else: logging = False
End If


If IsEmpty(Worksheets(1).Cells(4, 13)) Then
total_rows = Worksheets("Sheet1").Range("F65536").End(xlUp).row
Else
total_rows = Worksheets(1).Cells(4, 13)
End If
row_number = Worksheets(1).Cells(3, 13)
row_number = row_number - 1




Do
DoEvents
    row_number = row_number + 1

customer = Worksheets(1).Cells(row_number, 6)
msender = Worksheets(1).Cells(8, 13)
toaddress = Worksheets(1).Cells(row_number, 5)
mailsubject = Worksheets(1).Cells(row_number, 9)
attachedfile = Worksheets(1).Cells(2, 13) & "\" & Worksheets(1).Cells(row_number, 6) & "\" & Worksheets(1).Cells(row_number, 7) & "\" & Worksheets(1).Cells(row_number, 8) & "\attached_documents.zip"
mailbody = Worksheets(1).Cells(row_number, 10)
Set Skip = Worksheets(1).Cells(row_number, 1)
Dim foundincol As Range


If Skip.Value = vbNullString Then
'''''''''''''''''''''''''''''''''''''''
        Call SendEmail(Worksheets(1).Cells(row_number, 5), Worksheets(1).Cells(row_number, 9), Worksheets(1).Cells(row_number, 10), Worksheets(1).Cells(row_number, 6), Worksheets(1).Cells(row_number, 7), Worksheets(1).Cells(row_number, 8), Worksheets(1).Cells(2, 13), Worksheets(1).Cells(8, 13))
        Worksheets(1).Cells(row_number, 11) = "OK    " & "    " & Date & "   " & Time
        Worksheets(1).Cells(row_number, 11).Interior.Color = RGB(198, 239, 206)
        
        
        If logging Then
        ' logging starts here
       
        Set foundincol = Worksheets(2).Columns("A").Find(What:=customer, LookIn:=xlValues, LookAt:=xlWhole)
        ' looking for customer string in sheet2. if not found - create one.
        If foundincol Is Nothing Then
        Set foundincol = Worksheets(2).Range("A" & Rows.Count).End(xlUp).Offset(1)
        customer_row = foundincol.row
        Worksheets(2).Cells(customer_row, 1) = customer
        Else
        customer_row = foundincol.row
        End If
        ' looking for customer string end.

        Set myLastCell = Worksheets(2).Rows(customer_row).Find(What:="?", LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False, SearchFormat:=False)
        
        myLastCell.Offset(0, 1) = "OK    " & "   " & Date & "   " & Time & vbCrLf & "From: " & msender & vbCrLf & "To: " & toaddress & vbCrLf & "Subject: " & mailsubject & vbCrLf & mailbody & vbCrLf & "Attached file: " & attachedfile & vbCrLf & "Files that were attached:" & vbCrLf
        
        Set SHL = CreateObject("Shell.Application")
        For Each FileInZip In SHL.Namespace((attachedfile)).Items
          
          myLastCell.Offset(0, 1).Value = myLastCell.Offset(0, 1).Value & FileInZip & vbCrLf

        Next
        
        myLastCell.Offset(0, 1).Interior.Color = RGB(198, 239, 206)
        myLastCell.Offset(0, 1).ColumnWidth = 25.17
        Worksheets(2).Rows(customer_row).RowHeight = 15
        ' logging ends here
        End If
        

'''''''''''''''''''''''
Else

        Worksheets(1).Cells(row_number, 11) = "SKPD" & "   " & Date & "   " & Time
        Worksheets(1).Cells(row_number, 11).Interior.Color = RGB(149, 179, 215)
        If logging Then
        ' logging starts here

        Set foundincol = Worksheets(2).Columns("A").Find(What:=customer, LookIn:=xlValues, LookAt:=xlWhole)
        ' looking for customer string in sheet2. if not found - create one.
        If foundincol Is Nothing Then
        Set foundincol = Worksheets(2).Range("A" & Rows.Count).End(xlUp).Offset(1)
        customer_row = foundincol.row
        Worksheets(2).Cells(customer_row, 1) = customer
        Else
        customer_row = foundincol.row
        End If
        ' looking for customer string end.

        Set myLastCell = Worksheets(2).Rows(customer_row).Find(What:="?", LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False, SearchFormat:=False)
        
        myLastCell.Offset(0, 1) = "SKPD  " & Date & "   " & Time
        myLastCell.Offset(0, 1).Interior.Color = RGB(149, 179, 215)
        myLastCell.Offset(0, 1).ColumnWidth = 25.15
        ' logging ends here
        End If


End If


application.Wait (Now + waittime)
Loop Until row_number = total_rows
End Sub




