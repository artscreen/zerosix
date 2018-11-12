Dim LastRow As Double
Dim POFormWS As Worksheet
Dim POBookWB As Workbook
Dim POLogWS As Worksheet
Dim POSetupWS As Worksheet
Dim POPrefix As String
Dim POIncrement As Long
Dim PONextNumber As String
Dim PONumCell As Range
Dim PODateCell As Range
Dim POVendorCell As Range
Dim PODescCell As Range
Dim POJobCell As Range
Dim POGLCell As Range
Dim POSubCell As Range
Dim POTaxCell As Range
Dim POShipCell As Range
Dim POTotCell As Range
Dim PONoteCell As Range
Dim POItemsCells As Range
Dim PODateReqCell As Range
Dim POAttnCell As Range
Dim LogInitCell As Range
Dim LogNumCell As Range
Dim LogDateCell As Range
Dim LogVendorCell As Range
Dim LogDescCell As Range
Dim LogJobCell As Range
Dim LogGLCell As Range
Dim LogSubCell As Range
Dim LogTaxCell As Range
Dim LogShipCell As Range
Dim LogTotCell As Range
Dim reqmt_return As Integer

 
Sub POLogEntry()

Application.ScreenUpdating = False

    

'remove password protection
    sbUnProtectSheet

'define worksheets'
    Set POBookWB = ActiveWorkbook

    Set POFormWS = POBookWB.Sheets("POEntry")
    Set POLogWS = POBookWB.Sheets("POLog")
    Set POSetupWS = POBookWB.Sheets("Dropdowns")

'define PO info locations'
    
    Set PONumCell = POFormWS.Range("PONUMBER")
    Set PODateCell = POFormWS.Range("date")
    Set POVendorCell = POFormWS.Range("vendor")
    Set PODescCell = POFormWS.Range("Description")
    Set POJobCell = POFormWS.Range("jobnumber")
    Set POGLCell = POFormWS.Range("GL_CODE")
    Set POSubCell = POFormWS.Range("subtotal")
    Set POTaxCell = POFormWS.Range("tax")
    Set POShipCell = POFormWS.Range("freight")
    Set POTotCell = POFormWS.Range("total")

'find last row of po log'
    LastRow = POLogWS.Range("a10000").End(xlUp).Row + 1
   
'define PO log locations'
    Set LogInitCell = POLogWS.Range("A" & LastRow)
    Set LogNumCell = POLogWS.Range("B" & LastRow)
    Set LogDateCell = POLogWS.Range("C" & LastRow)
    Set LogVendorCell = POLogWS.Range("D" & LastRow)
    Set LogDescCell = POLogWS.Range("E" & LastRow)
    Set LogJobCell = POLogWS.Range("F" & LastRow)
    Set LogGLCell = POLogWS.Range("G" & LastRow)
    Set LogSubCell = POLogWS.Range("H" & LastRow)
    Set LogTaxCell = POLogWS.Range("I" & LastRow)
    Set LogShipCell = POLogWS.Range("J" & LastRow)
    Set LogTotCell = POLogWS.Range("K" & LastRow)
   
 'populate Log with data from PO form'
    LogInitCell = Left(UCase(Environ("Username")), 2)
    LogNumCell = PONumCell.Value
    LogDateCell = PODateCell.Value
    LogVendorCell = POVendorCell.Value
    LogDescCell = PODescCell.Value
    LogJobCell = POJobCell.Value
    LogGLCell = POGLCell.Value
    LogSubCell = POSubCell.Value
    LogTaxCell = POTaxCell.Value
    LogShipCell = POShipCell.Value
    LogTotCell = POTotCell.Value
    
'add hyperlink to po number in log
    POLogWS.Hyperlinks.Add Anchor:=LogNumCell, Address:="Purchase Orders\" & PONumCell.Value & " - " & RepIllegalChar(POVendorCell.Value, "_") & ".xlsx", TextToDisplay:=PONumCell.Value
 
'Save flat file to PO folder
    SavetoFile
    
'Generate a populate new PO number
    PO_Number
    
'Password protect Entry tab
    sbProtectSheet
    
 'Clear form and Save the Workbook
    ClearForm
    Application.ScreenUpdating = True
    POBookWB.Save
     
End Sub

Sub PO_Number()
    
    POPrefix = POSetupWS.Range("Prefix").Value
    POIncrement = Right(PONumCell.Value, 6)
    PONextNumber = POPrefix & (POIncrement + 1)
    PONumCell.Value = PONextNumber

End Sub

Sub ClearForm()
    sbUnProtectSheet
    
    Set POBookWB = ActiveWorkbook

    Set POFormWS = POBookWB.Sheets("POEntry")
    
    
   
    Set PODateReqCell = POFormWS.Range("date_req")
    Set PONoteCell = POFormWS.Range("notes")
    Set POAttnCell = POFormWS.Range("attn")
    
    Set PONumCell = POFormWS.Range("PONUMBER")
    Set PODateCell = POFormWS.Range("date")
    Set POVendorCell = POFormWS.Range("vendor")
    Set PODescCell = POFormWS.Range("Description")
    Set POJobCell = POFormWS.Range("jobnumber")
    Set POGLCell = POFormWS.Range("GL_CODE")
    Set POSubCell = POFormWS.Range("subtotal")
    Set POTaxCell = POFormWS.Range("tax")
    Set POShipCell = POFormWS.Range("freight")
    Set POTotCell = POFormWS.Range("total")


    Set POItemsCells = POFormWS.Range("lineitems")

'reset form values
    POVendorCell = ""
    POJobCell = ""
    PODescCell = ""
    POShipCell = ""
    POItemsCells = ""
    PONoteCell = ""
    PODateReqCell = Now
    
    POAttnCell = ""
    sbProtectSheet

End Sub

Sub PrintDefault()
    sbUnProtectSheet
    Set POBookWB = ActiveWorkbook
    Set POFormWS = POBookWB.Sheets("POEntry")
    
    'run sub for description check to see if it is empty
    ReqmtCheck
    If reqmt_return = 1 Then
        MsgBox "Please Fill in Required Fields"
        Exit Sub
    End If
    
    DescriptionCheck
    
    If POFormWS.Range("hide_check").Value = "x" Then
        POFormWS.Range("overflow").EntireRow.Hidden = False
    End If


    POFormWS.PrintOut
    
    If POFormWS.Range("hide_check").Value = "x" Then
        POFormWS.Range("overflow").EntireRow.Hidden = True
    End If

    sbProtectSheet
    'MsgBox ("Printed")

End Sub

Sub PrintPreview()
    sbUnProtectSheet
    Set POBookWB = ActiveWorkbook
    Set POFormWS = POBookWB.Sheets("POEntry")
    
    
    
    'run sub for description check to see if it is empty
    ReqmtCheck
    If reqmt_return = 1 Then
        MsgBox "Please Fill in Required Fields"
        Exit Sub
    End If
    
    DescriptionCheck
    
    If POFormWS.Range("hide_check").Value = "x" Then
        POFormWS.Range("overflow").EntireRow.Hidden = False
    End If


    POFormWS.PrintPreview
    
    If POFormWS.Range("hide_check").Value = "x" Then
        POFormWS.Range("overflow").EntireRow.Hidden = True
    End If

    sbProtectSheet
    'MsgBox ("Printed")

End Sub

Sub sbProtectSheet()

    ActiveSheet.Protect "password", True, True
    
End Sub

Sub sbUnProtectSheet()

    
    ActiveSheet.Unprotect "password"
    'ActiveWorkbook.Sheets("POLog").Unprotect "password"

End Sub

Sub AddressForm_Show()
    
    
    Address_Form.Show

End Sub

Sub DescriptionCheck()
    
    If ActiveWorkbook.Sheets("POEntry").Range("description").Value = "" Then
        DescriptionForm.Show
    End If

End Sub

Sub ReqmtCheck()
    Set POBookWB = ActiveWorkbook
    Set POFormWS = POBookWB.Sheets("POEntry")
    
    Set PODateReqCell = POFormWS.Range("date_req")
    Set PONoteCell = POFormWS.Range("notes")
    Set POAttnCell = POFormWS.Range("attn")
    
    Set PONumCell = POFormWS.Range("PONUMBER")
    Set PODateCell = POFormWS.Range("date")
    Set POVendorCell = POFormWS.Range("vendor")
    Set PODescCell = POFormWS.Range("Description")
    Set POJobCell = POFormWS.Range("jobnumber")
    Set POGLCell = POFormWS.Range("GL_CODE")
    Set POSubCell = POFormWS.Range("subtotal")
    Set POTaxCell = POFormWS.Range("tax")
    Set POShipCell = POFormWS.Range("freight")
    Set POTotCell = POFormWS.Range("total")
     
    
    If POVendorCell.Value = "" Or POJobCell.Value = "" Or POGLCell.Value = "Select" Then
        reqmt_return = 1
        Exit Sub
    End If

    reqmt_return = 0
    
End Sub

Sub ConfirmForm_Show()
    
    ReqmtCheck
    If reqmt_return = 1 Then
        DescriptionForm.Show
        Exit Sub
    End If
    
    DescriptionCheck
    
    If ActiveWorkbook.Sheets("POEntry").Range("description").Value = "" Then
        Exit Sub
    End If
    
    Confirmation_Form2.Show

End Sub
Sub HideRowToggle()
    sbUnProtectSheet
    Set POBookWB = ActiveWorkbook
    Set POFormWS = POBookWB.Sheets("POEntry")
    If POFormWS.Range("overflow").EntireRow.Hidden = True Then
        POFormWS.Range("overflow").EntireRow.Hidden = False
        POFormWS.Range("hide_check").Value = ""
        Else
    POFormWS.Range("overflow").EntireRow.Hidden = True
    POFormWS.Range("hide_check").Value = "x"
    End If
    sbProtectSheet
End Sub


Sub SavetoFile()
    
    Dim fname As String
    Dim fpath As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    fname = ActiveWorkbook.ActiveSheet.Range("ponumber").Value & " - " & RepIllegalChar(ActiveWorkbook.ActiveSheet.Range("vendor").Value, "_") & ".xlsx"
    fpath = ThisWorkbook.Path & "\Purchase Orders\"
    fsave = fpath & fname
    
    
    Sheets("POEntry").Copy
    Sheets("POEntry").Select
    Columns("O:O").Select
    Range(Selection, Selection.End(xlToRight)).Delete Shift:=xlToLeft
    Range("purchase_order").Copy
    Range("purchase_order").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("purchase_order").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveSheet.Shapes.Range(Array("Button 4")).Delete
    Range("A1").Select
    ActiveWorkbook.SaveAs fsave, FileFormat:=51
    ActiveWorkbook.Close
    Application.ScreenUpdating = True

 
End Sub

 Function RepIllegalChar(strIn As String, strChar As String) As String
    Dim strSpecialChars As String
    Dim i As Long
    strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next

    RepIllegalChar = strIn
End Function



