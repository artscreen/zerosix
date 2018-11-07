VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DescriptionForm 
   Caption         =   "Description Needed"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7875
   OleObjectBlob   =   "DescriptionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DescriptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If DescriptionBox.Value = "" Then
        MsgBox "Description Required"
    Exit Sub
    End If
        
    ActiveWorkbook.Sheets("POEntry").Range("Description").Value = DescriptionBox.Value
    Unload Me

End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

With DescriptionBox
    DescriptionBox.Value = ""
End With

End Sub
