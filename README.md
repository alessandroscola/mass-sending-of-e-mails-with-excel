# Massive sending of emails with Excel

It may happen that you have to send your customers or suppliers some e-mails with customized text depending on the recipient, or with a different attachment for each recipient. If you don't have CRM software that allows you to automate all of this, don't worry! you can do it with Excel!

My [**blog article regarding this project**](https://www.alessandroscola.com/excel/invio-massivo-di-email-con-excel.html) (italian version)

## Technology:
* Excel and VBA
* Blat - Windows Command Line SMTP Mailer
* SMTP server


## Usage
Download [blat](https://sourceforge.net/projects/blat/files/)
Configure blat with you SMTP server
Customize the Excel Page 
Prepare your attachments files in the same folder 
Click on "Send Emails"


## Customize the VBA code

```
Option Explicit
Public conn, rs As Object

' Code by Alessandro Scola www.alessandroscola.com
'
' Rimbember for INSTALL BLAT:
' blat -install -f <sender-email> -server <email-server> -u <login> -pw <password>
'
' To SEND with BLAT:
'   blat -Subject "email Subject" -body "email body" -to recipient@domain.xxx -html
'
'   alternatively with the email body in a external file "email_body.txt", the command is:
'   blat -s "c:\path\to\email_body.txt" -Subject "email Subject" -to recipient@domain.xxx -attach "c:\path\to\attachment.pdf" -html

Sub send_emails()
Dim row As Long
Dim RetVal As Variant
Dim email As String
Dim command As String
Dim obj_fso As Object
Dim fileName As String
Dim subject As String
Dim body As String

RetVal = MsgBox("Are you sure you want to send all emails ?", vbQuestion + vbYesNo + vbDefaultButton2)
If (RetVal <> vbYes) Then
  Exit Sub
End If


subject = Trim(Cells(2, 2).Value)
subject = Replace(subject, Chr(34), "\" & Chr(34)) ' replaces any " character with a \" not to break the command string


row = 6

email = Trim(Range("B" & row).Value)

While (email <> "")
  body = Trim(Cells(4, 2).Value) ' the initial body text
  
  body = Replace(body, "%A%", Trim(Cells(row, 1))) ' replaces any %A% with the content of A column
  body = Replace(body, "%E%", Trim(Cells(row, 5))) ' replaces any %E%with the content of E column
  body = Replace(body, vbLf, "<br>") ' replaces any "new line" with the "<BR>" TAG. "<BR>" TAG work as "new line " in HTML e-mails
  body = Replace(body, Chr(34), "\" & Chr(34)) ' replaces any " character with \" to not to break the command string
  
  command = "blat.exe -Subject " & Chr(34) & subject & Chr(34) & " -body " & Chr(34) & body & Chr(34) & " -to " & email & " -html"
  fileName = ThisWorkbook.Path & "\" & Trim(Range("C" & row).Value) & Trim(Range("D" & row).Value)
  
  If (fileExists(fileName)) Then
     'Se il file allegato esiste aggiunge al command "BLAT" la parte di codice per allegarlo
     command = command & " -attach " & Chr(34) & fileName & Chr(34)
     Cells(row, 6).Value = "OK"
  Else
     'Altrimenti scrive alla colonna 6 un avviso !
     Cells(row, 6).Value = "ATTENTION: ATTACHMENT NOT FOUND!"
  End If
  RetVal = Shell(command, vbMinimizedFocus)
  
  Application.Wait (Now + TimeValue("0:00:1")) ' Pause for 1 second
  
  row = row + 1

  email = Trim(Range("B" & row).Value)
  fileName = ThisWorkbook.Path & "\" & Trim(Range("C" & row).Value) & Trim(Range("D" & row).Value)
Wend

row = row - 1
MsgBox "Program ended at row: " & row, vbInformation

End Sub

' Check if a file exists
Function fileExists(fileName As String) As Boolean

    Dim obj_fso As Object

    Set obj_fso = CreateObject("Scripting.FileSystemObject")
    fileExists = obj_fso.fileExists(fileName)
End Function

```


 
