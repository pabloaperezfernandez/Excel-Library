Attribute VB_Name = "OutlookAutomation"
Option Explicit
Option Base 1

'***********************************************************************************
' FUNCTION NAME: SendNoAttachmentEmail(Recipient As String, Subject As String, Body As String)
'
' PURPOSE OF FUNCTION: To send a simple email to recipient with email address Recipient.
' The subject and body of the email are Subject and Email.
'
' PARAMETERS:
' ContactName : a string representing the name of the person whose ContactItem is
' desired.
'
'************************************************************************************
Public Sub SendNoAttachmentEmail(Recipient As String, Subject As String, Body As String)
    Dim olApp As Object 'Outlook.Application
    Dim olMail As Object 'MailItem

    Set olApp = CreateObject("Outlook.Application") 'New Outlook.Application
    Set olMail = olApp.CreateItem(0) 'olMailItem

    With olMail
        .to = Recipient
        .Subject = Subject
        .Body = Body
        .Send
    End With

    Set olMail = Nothing
    Set olApp = Nothing
End Sub


'***********************************************************************************
' FUNCTION NAME: SendAsAttachment(Recipient As String, Subject As String, Body As String)
'
' PURPOSE OF FUNCTION: Sends out the workbook to a single recipient. This routine
' hides the Outlook windows from the user and does not require the user to click on the
' send button.
'
' PARAMETERS:
' Recipient : A string with the email address of the recipient
' Subject : A string with the subject line
' Body : A string with the body of the email
'
'************************************************************************************
Public Sub SendAsAttachment(Recipient As String, FirstName As String, Subject As String, Body As String)
    Dim olApp As Object 'Outlook.Application
    Dim olMail As Object 'MailItem
    Dim CurrFile As String

    Set olApp = CreateObject("Outlook.Application") 'New Outlook.Application
    Set olMail = olApp.CreateItem(0) ' olMailItem

    ' Save the workbook using the recipient's email as the file name (plus the .xlsm extension)
    ActiveWorkbook.SaveCopyAs ActiveWorkbook.Path & "\" & FirstName & "-Distribution-" & ".xlsm"

    ' Get the complete path for the file name so may attach the workbook the email.
    CurrFile = ActiveWorkbook.Path & "\" & FirstName & "-Distribution-" & ".xlsm"

    With olMail
        .to = Recipient
        .Subject = Subject
        .Body = Body
        .Attachments.Add CurrFile
        ' Uncomment the next line to display the email dialogue
        '.Display '.Send
        ' Use the following line to send without displaying the dialogue or prompting the user to click send.
        .Send
    End With

    Set olMail = Nothing
    Set olApp = Nothing
End Sub
