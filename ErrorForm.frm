VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorForm 
   Caption         =   "Microsoft Excel"
   ClientHeight    =   1320
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9396.001
   OleObjectBlob   =   "ErrorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()
    ErrorForm.Hide
    PasswordForm.Show
End Sub
