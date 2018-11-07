VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Address_Form 
   Caption         =   "Address Entry"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4515
   OleObjectBlob   =   "Address_Form.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Address_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If SiteNameBox.Value = "" Or Address1Box.Value = "" Or CityBox.Value = "" Or TaxRateBox.Value = "" Or StateComboBox.Value = "" Or ZipBox.Value = "" Then
        MsgBox "Please Complete All Fields"
    Exit Sub
    End If
    
    Dim LastAddRow As Long
    Dim WS As Worksheet
    
    Set WS = ActiveWorkbook.Sheets("Dropdowns")
    
    
    LastAddRow = WS.Range("a10000").End(xlUp).Row + 1
    
    WS.Range("A" & LastAddRow).Value = SiteNameBox.Value
    WS.Range("B" & LastAddRow).Value = Address1Box.Value
    WS.Range("C" & LastAddRow).Value = Address2Box.Value
    WS.Range("D" & LastAddRow).Value = CityBox.Value
    WS.Range("E" & LastAddRow).Value = StateComboBox.Value
    WS.Range("F" & LastAddRow).Value = ZipBox.Value
    WS.Range("G" & LastAddRow).Value = TaxRateBox.Value / 100
    
    ActiveWorkbook.Sheets("POEntry").Range("I42").Value = SiteNameBox.Value
    
    Unload Me
    
    Exit Sub
    
  
    

End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub



Private Sub UserForm_Initialize()

With SiteNameBox
    SiteNameBox.Value = ""
End With

With Address1Box
    Address1Box.Value = ""
End With

With Address2Box
    Address2Box.Value = ""
End With

With CityBox
    CityBox.Value = ""
End With

With TaxRateBox
    TaxRateBox.Value = ""
End With

With ZipBox
    ZipBox.Value = ""
End With

'Fill StateComboBox
With StateComboBox
    .AddItem "AL"
    .AddItem "AK"
    .AddItem "AZ"
    .AddItem "AR"
    .AddItem "CA"
    .AddItem "CO"
    .AddItem "CT"
    .AddItem "DE"
    .AddItem "FL"
    .AddItem "GA"
    .AddItem "HI"
    .AddItem "ID"
    .AddItem "IL"
    .AddItem "IN"
    .AddItem "IA"
    .AddItem "KS"
    .AddItem "KY"
    .AddItem "LA"
    .AddItem "ME"
    .AddItem "MD"
    .AddItem "MA"
    .AddItem "MI"
    .AddItem "MN"
    .AddItem "MS"
    .AddItem "MO"
    .AddItem "MT"
    .AddItem "NE"
    .AddItem "NV"
    .AddItem "NH"
    .AddItem "NJ"
    .AddItem "NM"
    .AddItem "NY"
    .AddItem "NC"
    .AddItem "ND"
    .AddItem "OH"
    .AddItem "OK"
    .AddItem "OR"
    .AddItem "PA"
    .AddItem "RI"
    .AddItem "SC"
    .AddItem "SD"
    .AddItem "TN"
    .AddItem "TX"
    .AddItem "UT"
    .AddItem "VT"
    .AddItem "VA"
    .AddItem "WA"
    .AddItem "WV"
    .AddItem "WI"
    .AddItem "WY"
End With



End Sub

