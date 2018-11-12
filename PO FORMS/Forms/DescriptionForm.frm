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

Private Sub SaveButton_Click()
   
    ActiveWorkbook.Sheets("POEntry").Range("Description").Value = DescriptionBox.Value
    ActiveWorkbook.Sheets("POEntry").Range("vendor").Value = VendorBox.Value
    ActiveWorkbook.Sheets("POEntry").Range("jobnumber").Value = JobBox.Value
    ActiveWorkbook.Sheets("POEntry").Range("GLDesc").Value = GLBox.Value
    Unload Me
    Confirmation_Form2.Show
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim gllist As String
    gllist = ActiveWorkbook.Sheets("Dropdowns").Range("prefix").Value
With DescriptionBox
    VendorBox.Value = ActiveWorkbook.Sheets("POEntry").Range("vendor").Value
    JobBox.Value = ActiveWorkbook.Sheets("POEntry").Range("jobnumber").Value
    GLBox.List = ActiveWorkbook.Sheets("Dropdowns").Range(gllist).Value
    GLBox.Value = ActiveWorkbook.Sheets("POEntry").Range("GLDesc").Value
    DescriptionBox.Value = ActiveWorkbook.Sheets("POEntry").Range("Description").Value
    
End With

End Sub
