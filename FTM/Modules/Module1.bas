Option Compare Text

Sub importrecap()

    Dim recaptab As Worksheet
    Dim report As Workbook
    Dim data As Worksheet
    Dim lastrow As Double
    Dim descr As String
    Dim descrfix As String
    Dim deliverydate As String
    Dim fabricator As String
    Dim tonnage As Variant
    Dim matrixrow As Integer
    Dim matrixrowend As Integer
    Dim starttime As Date
    Dim endtime As Date
    Dim reportdate As Date
    Dim lookup As Worksheet
    Dim lookahead As Worksheet
    Dim selectdate As Date
    Dim toncol As String
    Dim rffcol As String
    Dim modtcol As String
    Dim delcol As String
    Dim fabcol As String
    Dim seqcol As String
    Dim i_count As Integer
    Dim username As String
    Dim dateinput As Date
    Dim dateoutput As Date
    
    sbUnProtectSheet
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    
    Set report = ActiveWorkbook
    Set data = report.Sheets("matrixData")
    Set lookahead = report.Sheets("Lookahead")
    Set lookup = report.Sheets("lookups")
   
    
    
    selectdate = Date$
    
    reportdate = Now()
    starttime = Now() - Weekday(Now(), 3)
    
    endtime = starttime + 20
    
    lastrow = 2
    i_count = 0
    

    data.Range("a2:d10000").ClearContents
    username = Application.username
    lookahead.Range("g3").Value = reportdate & Chr(10) & " by " & username
    lookup.Range("d5").Value = selectdate
        
        
    For Each recaptab In report.Sheets
            
        If recaptab.Name <> "TEMPLATE" And recaptab.Name <> "lookahead" And recaptab.Name <> "matrixdata" And recaptab.Name <> "lookups" And InStr(recaptab.Name, "CLOSED") = False And InStr(recaptab.Name, "closed") = False And InStr(recaptab.Name, "Closed") = False Then
            If recaptab.Range("q1") <> "x" Then
                modtcol = "d"
                rffcol = "n"
                delcol = "q"
                fabcol = "t"
                seqcol = "b"
                    
            Else
                    
                modtcol = "c"
                rffcol = "i"
                delcol = "l"
                fabcol = "o"
                seqcol = "a"
            End If
            
            matrixrowend = recaptab.Range("b10000").End(xlUp).Row
                
            For matrixrow = 29 To matrixrowend
                On Error GoTo ErrorMessage
                    
                dateinput = dateconvert(starttime, recaptab.Range(delcol & matrixrow).Value, endtime)

               
                If datecheck(dateinput) = True Then
                    descrfix = Left(Replace(recaptab.Range(seqcol & matrixrow).Value, "SEQUENCE", "SEQ"), 15)
                    descr = recaptab.Name & " " & descrfix
                    deliverydate = dateinput
                    If fabcheck(recaptab.Range(fabcol & matrixrow).Value) = False Then fabricator = "STEEL LLC" Else fabricator = recaptab.Range(fabcol & matrixrow).Value
                        
                    If IsError(recaptab.Range(rffcol & matrixrow)) = True Or IsError(recaptab.Range(modtcol & matrixrow)) = True Then
                        tonnage = 0
                    ElseIf toncheck((recaptab.Range(rffcol & matrixrow).Value), (recaptab.Range(modtcol & matrixrow).Value)) = True Then
                        toncol = modtcol
                    Else: toncol = rffcol
                            
                    End If
                        
                    If IsNumeric(recaptab.Range(toncol & matrixrow).Value) = True Then
                        tonnage = Round(recaptab.Range(toncol & matrixrow).Value, 2)
                    Else: tonnage = recaptab.Range(toncol & matrixrow).Value
                    End If
                        
            
            
                    If tonnage = 0 Or tonnage = "-" Then data.Range("a" & lastrow).Value = descr & " - " & fabricator Else data.Range("a" & lastrow).Value = descr & " - " & tonnage & " T" & " - " & fabricator
                    data.Range("b" & lastrow).Value = deliverydate
                    data.Range("c" & lastrow).Value = fabricator
                    data.Range("d" & lastrow).Value = tonnage
                    lastrow = lastrow + 1
                End If
            Next matrixrow
        End If
            
    Next recaptab

 
    sbProtectSheet
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
 
    
ErrorMessage:
    
    Resume Next
   
    
    
End Sub

Public Function dateconvert(starttime As Date, datein As Variant, endtime As Date) As Date
    
    'processes the info in the delivery date and checks to see if it falls with in the period.
    'If it does, pass date along in proper format.  If it does not, convert the date to 01/01/01 so date check function will skip it
    Debug.Print Format(starttime, "MM-DD-YYYY") & " - " & Format(datein, "MM-DD-YYYY") & " - " & Format(endtime, "MM-DD-YYYY")
    If datein > starttime - 2 And datein < endtime Then
        splitdate = Split(datein, "/")
        dateconvert = DateSerial(splitdate(2), splitdate(0), splitdate(1))
    Else: dateconvert = "01/01/01"
      
    End If
   
End Function

Public Function datecheck(cellvalue As Date) As Boolean

    'checks the value returned from the dateconvert function and allows sub to continue if the date is valid for the period
    
    If cellvalue = "01/01/01" Then
        datecheck = False
        Exit Function
    End If
    datecheck = True
End Function

Public Function fabcheck(cellvalue As Variant) As Boolean
    
    If cellvalue = "0" Or cellvalue = "" Then
        fabcheck = False
        Exit Function
    End If
    fabcheck = True
End Function

Public Function toncheck(rff As Variant, modt As Variant) As Boolean

    'checks released for fab weight field contains a weight
    
    If rff = "0" Or rff = 0 Then
        toncheck = True
        Exit Function
        
    ElseIf IsError(rff) = True Then
        If rff = CVErr(xlErrRef) Then
            toncheck = True
            Exit Function
        End If
    End If
    
    toncheck = False

End Function

Sub sbProtectSheet()
    ActiveSheet.Protect "PASSWORD", True, True
End Sub

Sub sbUnProtectSheet()
    ActiveSheet.Unprotect "PASSWORD"
End Sub

