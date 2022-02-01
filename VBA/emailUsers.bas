Attribute VB_Name = "emailUsers"
Sub autoEmail()
'-------------------------------------------------------------------
' Send emails from Excel via sheet named ranges
' TO, CC, BCC, SUBJECT, BODY, ATCHMNT_PATH
'-------------------------------------------------------------------
    Dim emailApp As Outlook.Application
    Dim emailItem As Outlook.MailItem
    Dim source As String
    
    Set emailApp = New Outlook.Application
    Set emailItem = emailApp.CreateItem(olMailItem)
    
    If Not (Range("TO") = 0) Then
        emailItem.To = Range("TO")
    Else
        MsgBox ("Please provide an email address with a named ranged cell called 'TO'")
        Exit Sub
    End If
    
    If Not (Range("CC") = 0) Then
        emailItem.CC = Range("CC")
    End If
    
    If Not (Range("BCC") = 0) Then
        emailItem.BCC = Range("BCC")
    End If
    
    If Not (Range("SUBJECT") = 0) Then
        emailItem.Subject = Range("SUBJECT")
    Else
        MsgBox ("Please give your subject an email")
        Exit Sub
    End If
    
    If Not (Range("BODY") = 0) Then
        emailItem.HTMLBody = Range("BODY")
    End If
    
    If Not (Range("ATCHMNT_PATH") = 0) Then
        source = Range("ATCHMNT_PATH")
        emailItem.Attachments.Add source
    End If
    
    emailItem.Send
End Sub


Sub selectEmail()
'-------------------------------------------------------------------
' Send emails from Excel via cell selection
'-------------------------------------------------------------------
    Dim emailApp As Outlook.Application
    Dim emailItem As Outlook.MailItem
    Dim userInput As Variant
    
    Set emailApp = New Outlook.Application
    Set emailItem = emailApp.CreateItem(olMailItem)
    
    emailItem.To = Application.InputBox("To: ", "Select cell", " ", Type:=8)
    emailItem.CC = Application.InputBox("CC: ", "Select cell", " ", Type:=8)
    emailItem.Subject = Application.InputBox("Subject: ", "Select cell", " ", Type:=8)
    emailItem.Body = Application.InputBox("Email body: ", "Select cell", " ", Type:=8)
    ' currently no input validation
    
    'Debug.Print emailItem.CC
    
    emailItem.Send
End Sub


Sub manualEmail()
'-------------------------------------------------------------------
' Send emails from Excel via manual input
'-------------------------------------------------------------------
    Dim emailApp As Outlook.Application
    Dim emailItem As Outlook.MailItem
    Dim userInput As Variant
    
    Set emailApp = New Outlook.Application
    Set emailItem = emailApp.CreateItem(olMailItem)
    
    emailItem.To = InputBox("To: emails (seperated by semicolons)", "Email e-Manifest", "user email(s) here")
    emailItem.CC = InputBox("CC: emails (seperated by semicolons)", "Email e-Manifest", "CC email(s) here")
    emailItem.BCC = InputBox("BCC: emails (seperated by semicolons)", "Email e-Manifest", "BCC email(s) here")
    emailItem.Subject = InputBox("Subject", "Email e-Manifest", "Subject")
    emailItem.Send
End Sub

Sub pivotEmail()
'-------------------------------------------------------------------
' Loop through Pivot items to generate emails
'-------------------------------------------------------------------
    Dim emailApp As Outlook.Application
    Dim emailItem As Outlook.MailItem
    Dim basePath, source, idString As String
    Dim pt As PivotTable
    Dim pv As PivotField
    Dim epaId As PivotItem
    Dim i As Integer
    
    basePath = "C:\Users\dgraha01\OneDrive - Environmental Protection Agency (EPA)\e-Manifest - dg\DQ_issues\InvalidGenID\"
    Set pt = Worksheets("Summary").PivotTables(1)
    Set pv = pt.PivotFields("TSDF ID")
    
    For Each epaId In pv.PivotItems
        If epaId <> "" Then
            Worksheets("MTN suggestions").ListObjects("Table1").Range.AutoFilter _
            Field:=1, Criteria1:="=" & epaId
        Else
            MsgBox ("Value not found where expected")
            Exit Sub
        End If
        
        i = epaId.Position + 2
        If i = 3 Then
            i = i + 1
        End If
        Set emailApp = New Outlook.Application
        Set emailItem = emailApp.CreateItem(0)
        emailItem.To = Worksheets("Summary").Cells(i, 3).Value
        emailItem.Subject = "Test1: email with body " & epaId & " MTN with Invalid Generator IDs"
        emailItem.Body = Worksheets("Values").Cells(1, 1).Value & epaId & Worksheets("Values").Cells(2, 1).Value & epaId & Worksheets("Values").Cells(3, 1).Value
        source = basePath & epaId & ".csv"
        Debug.Print epaId, i
        emailItem.Attachments.Add source
        emailItem.Send
        'For testing purposes
        If i = 5 Then
            Exit Sub
        End If
    Next epaId
End Sub
