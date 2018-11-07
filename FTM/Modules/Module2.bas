Attribute VB_Name = "Module2"
Option Compare Text
Sub SUBahead()
Attribute SUBahead.VB_ProcData.VB_Invoke_Func = "I\n14"

    Dim recap As Workbook
    Dim recaptab As Worksheet
    Dim report As Workbook
    Dim data As Worksheet
    Dim filename As Variant
    Dim lastrow As Double
    Dim output As Integer
    Dim errorout As Integer
    Dim descr As String
    Dim descrfix As String
    Dim deliverydate As String
    Dim fabricator As String
    Dim tonnage As Variant
    Dim matrixrow As Integer
    Dim matrixrowend As Integer
    Dim starttime As Date
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
    Dim i As Integer
    Dim username As String
    Dim ofacol As String
    Dim bfacol As String
    Dim rffdatecol As String
    Dim pename As String
    Dim penamerng As String
    Dim ofadate As Date
    Dim bfadate As Date
    Dim rffdate As Date
    Dim status As String
    
    sbUnProtectSheet
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    
    Set report = ActiveWorkbook
    Set data = report.Sheets("OFAData")
    Set lookahead = report.Sheets("Submittal Lookahead")
    Set lookup = report.Sheets("OFAlookups")
   
    
    
    selectdate = Date$
    
    reportdate = Now()
    starttime = Now()
    lastrow = 2
    i = 0
         
    
    'clear data table
        data.Range("a2:i10000").ClearContents
        
        
     'fill in updated box
        username = Application.username
        lookahead.Range("g3").Value = reportdate & Chr(10) & " by " & username
        lookup.Range("d5").Value = selectdate
        
      'loop through tabs, skipping unused and closed
      'select columns based on template
        For Each recaptab In report.Sheets
            
            If recaptab.Name <> "TEMPLATE 2" And recaptab.Name <> "OFAlookups" And recaptab.Name <> "matrixData" And recaptab.Name <> "Submittal Lookahead" And recaptab.Name <> "Lookahead" And recaptab.Name <> "OFAData" And recaptab.Name <> "lookups" And InStr(recaptab.Name, "CLOSED") = False And InStr(recaptab.Name, "closed") = False And InStr(recaptab.Name, "Closed") = False Then
                If recaptab.Range("q1") <> "x" Then
                    modtcol = "d"
                    rffcol = "n"
                    delcol = "q"
                    fabcol = "t"
                    seqcol = "b"
                    ofacol = "e"
                    bfacol = "k"
                    rffdatecol = "m"
                    penamerng = "c5"
                    
                    
                    Else
                    
                    modtcol = "c"
                    rffcol = "i"
                    delcol = "l"
                    fabcol = "o"
                    seqcol = "a"
                    ofacol = "d"
                    bfacol = "f"
                    rffdatecol = "h"
                    penamerng = "b5"
                End If
                
         'determine job tab range
            matrixrowend = recaptab.Range("b10000").End(xlUp).Row
            
          
          
                For matrixrow = 29 To matrixrowend
                On Error GoTo ErrorMessage
                'see if the dates are valid
                    If rffdatecheck(recaptab.Range(rffdatecol & matrixrow).Value) = True Then
                       
                        descrfix = Left(Replace(recaptab.Range(seqcol & matrixrow).Value, "SEQUENCE", "SEQ"), 15)
                        descr = recaptab.Name & " " & descrfix
                        
                        ofadate = recaptab.Range(ofacol & matrixrow).Value
                        bfadate = recaptab.Range(bfacol & matrixrow).Value
                        rffdate = recaptab.Range(rffdatecol & matrixrow).Value
                        pename = recaptab.Range(penamerng).Value
                        
                        
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
                        
            
            
            
                    
                        data.Range("b" & lastrow).Value = ofadate
                        data.Range("c" & lastrow).Value = bfadate
                        data.Range("d" & lastrow).Value = rffdate
                        If ofadate > selectdate Then
                            data.Range("f" & lastrow).Value = ofadate
                            status = "OFA"
                            Else
                                If bfadate > selectdate Then
                                    data.Range("f" & lastrow).Value = bfadate
                                    status = "BFA"
                                    Else
                                        data.Range("f" & lastrow).Value = rffdate
                                        status = "RFF"
                                End If
                            
                        End If
                        
                        If tonnage = 0 Or tonnage = "-" Then data.Range("a" & lastrow).Value = descr & " - " & status & " - " & pename Else data.Range("a" & lastrow).Value = descr & " - " & tonnage & " T" & " - " & status & " - " & pename
                        data.Range("e" & lastrow).Value = pename
                        data.Range("g" & lastrow).Value = tonnage
                        data.Range("H" & lastrow).Value = status
                        'data.Range("I" & lastrow).Value = Now()
                        'Debug.Print Now() & " - " & recaptab.Name & " -entry- " & matrixrow
                        lastrow = lastrow + 1
                    End If
                Next matrixrow
            End If
            'Debug.Print Now() & " - " & recaptab.Name & " -skip- " & matrixrow
        Next recaptab
    
  
    sbProtectSheet
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
 
    
ErrorMessage:
    'i = i + 1
    'Debug.Print "error - " & i
    Resume Next
    
    Application.Calculation = xlCalculationAutomatic
    
    
End Sub

Public Function rffdatecheck(cellvalue As Variant) As Boolean
    
     If cellvalue Like "*[a-zA-Z]*" Or cellvalue = "0" Or cellvalue = " " Or cellvalue = "" Or cellvalue = "-" Or cellvalue < Now() - 14 Then
        rffdatecheck = False
        'Debug.Print "RFF False"
        Exit Function
    End If
        rffdatecheck = True
        'Debug.Print "RFF TRUE"
End Function


