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
' Ricorda che per INSTALLARE BLAT:
' blat -install -f <sender-email> -server <email-server> -u <login> -pw <password>
'
' Per INVIARE con BLAT:
'   blat -Subject "Oggetto mail" -body "corpo email" -to destinatario@dominio.it -html
'
'   oppure con corpo e-.mail in un file di testo "email_body.txt" il comando è:
'   blat -s "c:\path\to\email_body.txt" -Subject "oggetto" -to destinatario@dominio.it -attach "c:\path\to\attachment.pdf" -html

Sub invia_mails()
Dim riga As Long
Dim RetVal As Variant
Dim email As String
Dim comando As String
Dim obj_fso As Object
Dim fileName As String
Dim oggetto As String
Dim testo As String

RetVal = MsgBox("Sei Sicuro di voler inviare tutte le -mails ?", vbQuestion + vbYesNo + vbDefaultButton2)
If (RetVal <> vbYes) Then
  Exit Sub
End If


oggetto = Trim(Cells(2, 2).Value)
oggetto = Replace(oggetto, Chr(34), "\" & Chr(34)) ' sostituisce un eventuale carattere " con \" per non interrompere la stringa di comando


riga = 6

email = Trim(Range("B" & riga).Value)

While (email <> "")
  testo = Trim(Cells(4, 2).Value) ' preleva il testo on all'interno i TAGS %%
  
  testo = Replace(testo, "%A%", Trim(Cells(riga, 1))) ' sostituisce %A% con il contenuto della cella A (colonna 1) alla riga relativa
  testo = Replace(testo, "%E%", Trim(Cells(riga, 5))) ' sostituisce %E% con il contenuto della cella E (colonna 5) alla riga relativa
  testo = Replace(testo, vbLf, "<br>") ' sostituisce gli "a capo" con il tag "<BR>" per mandare a capo le righe nelle email HTML
  testo = Replace(testo, Chr(34), "\" & Chr(34)) ' sostituisce un eventuale carattere " con \" per non interrompere la stringa di comando
  
  comando = "blat.exe -Subject " & Chr(34) & oggetto & Chr(34) & " -body " & Chr(34) & testo & Chr(34) & " -to " & email & " -html"
  fileName = ThisWorkbook.Path & "\" & Trim(Range("C" & riga).Value) & Trim(Range("D" & riga).Value)
  
  If (fileExists(fileName)) Then
     'Se il file allegato esiste aggiunge al comando "BLAT" la parte di codice per allegarlo
     comando = comando & " -attach " & Chr(34) & fileName & Chr(34)
     Cells(riga, 6).Value = "OK."
  Else
     'Altrimenti scrive alla colonna 6 un avviso !
     Cells(riga, 6).Value = "ATTENZIONE: ALLEGATO NON TROVATO!"
  End If
  RetVal = Shell(comando, vbMinimizedFocus)
  
  Application.Wait (Now + TimeValue("0:00:1")) ' Pausa di 1 secondo
  
  riga = riga + 1

  email = Trim(Range("B" & riga).Value)
  fileName = ThisWorkbook.Path & "\" & Trim(Range("C" & riga).Value) & Trim(Range("D" & riga).Value)
Wend

riga = riga - 1
MsgBox "Finito alla riga " & riga, vbInformation

End Sub

' Check if a file exists
Function fileExists(fileName As String) As Boolean

    Dim obj_fso As Object

    Set obj_fso = CreateObject("Scripting.FileSystemObject")
    fileExists = obj_fso.fileExists(fileName)
End Function

```


 
