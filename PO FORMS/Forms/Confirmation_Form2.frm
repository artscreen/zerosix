VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Confirmation_Form2 
   Caption         =   "UserForm2"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   OleObjectBlob   =   "Confirmation_Form2.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Confirmation_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBttn_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()
 PrintDefault
End Sub

Private Sub ContinueBttn_Click()
    POLogEntry
    Unload Me
End Sub
