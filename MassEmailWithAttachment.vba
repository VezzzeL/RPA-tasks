Sub MassEmailWithAttachment()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim EmailSubject As String
    Dim EmailText As String
    Dim AttachmentPath As String
    Dim i As Long
    Dim ValidEmailFound As Boolean
    
    Dim FirstName As String
    Dim SecondName As String
    Dim TabValue As String
    Dim DateValue As String
    Dim ToValue As String
    Dim CcValue As String
    Dim BccValue As String
    
    EmailSubject = "Life Insurance 2021"
    'This pass was added cause cannot open attached PDF.
    AttachmentPath = "C:\Users\" & Environ("USERNAME") & "\Desktop\Attachment.pdf"
    
    Set ws = ThisWorkbook.Sheets("Data")
    Set OutApp = CreateObject("Outlook.Application")

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Set OutMail = OutApp.CreateItem(0)
        ValidEmailFound = False
        
        'Set table cells
        FirstName = ws.Cells(i, 1).Value
        SecondName = ws.Cells(i, 2).Value
        TabValue = ws.Cells(i, 3).Value
        DateValue = ws.Cells(i, 4).Value
        ToValue = ws.Cells(i, 5).Value
        CcValue = ws.Cells(i, 6).Value
        BccValue = ws.Cells(i, 7).Value
        
        'Email text form
        EmailText = "Dear " & FirstName & " " & SecondName & ",<br><br>" & _
        "We are glad to announce that in 2021 we continue cooperation with our current life insurance partner ""INGO"".<br><br>" & _
        "The insurance program remains the same. We remind you that the insurance risk includes establishment of disability I, II, and III groups; establishment of a diagnosis of critical diseases; <br>" & _
        "death. New insurance cards will be distributed after coming back to the office. You can get your personal card (" & TabValue & ") starting from " & DateValue & ". <br><br>" & _
        "As before, during the first two months from the beginning of the program, insured employees have the opportunity to insure close relatives with corporate rates.<br>" & _
        "All information on the life insurance program for employees and relatives in 2021 as well as contact numbers of insurance company INGO can be found on the site.<br><br>" & _
        "BR,<br><br>" & _
        "Nestle Family"
        
        'Check if email is valid
        If IsValidEmail(ToValue) Then
            OutMail.To = ToValue
            OutMail.CC = CcValue
            OutMail.Bcc = BccValue
            ValidEmailFound = True
        ElseIf IsValidEmail(CcValue) Then
            OutMail.To = CcValue
            OutMail.Bcc = BccValue
            ValidEmailFound = True
        ElseIf IsValidEmail(BccValue) Then
            OutMail.To = BccValue
            ValidEmailFound = True
        End If
        
        If ValidEmailFound Then
            With OutMail
                .Subject = EmailSubject
                .HTMLBody = EmailText
                .Attachments.Add AttachmentPath
                .Send
            End With
        End If
    Next i

    Set OutMail = Nothing
    Set OutApp = Nothing

    MsgBox "Done"
End Sub

Function IsValidEmail(email As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "^[\w\.-]+@[\w\.-]+\.\w+$"

    IsValidEmail = regex.Test(email)
End Function
