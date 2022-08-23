Sub SendEmails()

' - DECLARATIONS AND INITIALISATION ------------------------------------------------------

' point to admin email
Dim adminEmail As String
adminEmail = Range("G7")

' initialise a counter to zero
Dim emailsSent As Integer
emailsSent = 0

' assign ms outlook to the EmailApp object
Dim EmailApp As Object
Set EmailApp = CreateObject("Outlook.Application")

' assign emails within outlook to the EmailItem object
Dim EmailItem As Object
Set EmailItem = EmailApp.CreateItem(0)

' point to list of names
Dim RangeOfList As Range
Set RangeOfList = Range("B3", Range("b3").End(xlDown))

' - MAIN ---------------------------------------------------------------------------------

' loop (for) through emails (which are 1 column offset) and send each an email
Dim R As Range
For Each R In RangeOfList
    If R.Offset(0, 2) = "Yes" Then
    Set EmailItem = EmailApp.CreateItem(0)
    EmailItem.To = R.Offset(0, 1)
    EmailItem.Subject = "x - Mailing List"
    EmailItem.Body = R.Offset(0, 0) & "," & vbNewLine & vbNewLine _
    & "Body of email." & vbNewLine _
    & vbNewLine & "Regards, x"
    EmailItem.Display 'EmailItem.Send
    emailsSent = emailsSent + 1
    End If
Next R
    
' send confirmation email to admin
Set EmailItem = EmailApp.CreateItem(0)
EmailItem.To = adminEmail
EmailItem.Subject = "Admin Notification"
EmailItem.Body = "Success! - " & CStr(emailsSent) & " customer(s) have been sent emails."
EmailItem.Display 'EmailItem.Send
    
' reset all objects and variables
Set EmailApp = Nothing
Set EmailItem = Nothing
emailsSent = 0

End Sub

