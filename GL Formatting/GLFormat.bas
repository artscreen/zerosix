Sub SplitGLtabs()
    
    Dim WB As Workbook
    Dim sltab As Worksheet
    Dim newtab As Worksheet
    Dim TabName As String
    Dim shape As String
    Dim GLCode As String
    Dim columnA As Range
    Dim lastrow As Long
    Dim startrow As Long
    Dim endrow As Long
    Dim NewtabEnd As Long
    Dim PRange As Range
    Dim PCache As PivotCache
    Dim Ptable As PivotTable

    Set WB = ActiveWorkbook
    Set sltab = WB.Sheets("Sheet1")
    lastrow = sltab.Range("a10000").End(xlUp).Row + 1
    Set columnA = sltab.Range("a1:a" & lastrow)

    If Application.Sheets.Count > 1 Or sltab.Range("A1").Value <> "Type" Then
        MsgBox "This ain't it, chief"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False

    For Each cell In columnA

        If cell.Value = "Acct:" Then
            
            startrow = cell.Row
            GLCode = cell.Offset(0, 1).Value

            If GLCode = "120002" Then TabName = "W 120002"
            If GLCode = "120003" Then TabName = "L 120004"
            If GLCode = "120004" Then TabName = "C MC 120005"
            If GLCode = "120005" Then TabName = "CP FB PL 120006"
            If GLCode = "120006" Then TabName = "RB REB SQ 120007"
            If GLCode = "120007" Then TabName = "HSS HSSR PI 120007"
            If GLCode = "120013" Then TabName = "COI 120013"
            If GLCode = "127000" Then TabName = "MIT 127000"
            If GLCode = "129900" Then TabName = "RES 129900"
        End If

        If cell.Value = "Acct" Then
            
            endrow = cell.Row

            NewtabEnd = endrow - startrow - 1
            if NewtabEnd > 2 then 
                    
                    Set newtab = Sheets.Add(After:=Sheets(Sheets.Count))
                    
                    newtab.Name = TabName
                    
                    Set newtab = Sheets(TabName)
                    
                    sltab.Range("A" & startrow & ":L" & endrow).Copy newtab.Range("a1")
                    newtab.Range("a1").Value = "A"
                    newtab.Range("b1").Value = "B"
                    newtab.Range("c1").Value = "C"
                    newtab.Range("d1").Value = "CODE"
                    newtab.Range("e1").Value = "PER"
                    newtab.Range("f1").Value = "NULL"
                    newtab.Range("g1").Value = "NBR"
                    newtab.Range("h1").Value = "DATE"
                    newtab.Range("i1").Value = "DESC"
                    newtab.Range("j1").Value = "INV $"
                    newtab.Range("k1").Value = "CREDIT $"
                    newtab.Range("l1").Value = "PO"
                    newtab.Range("l2:l" & NewtabEnd).FormulaR1C1 = "=mid(RC[-3],find(""-"",RC[-3],1)-5,9)"
                    
                    Set PRange = newtab.Range("A1:l" & NewtabEnd)

                    Set PCache = WB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange)

                    Set Ptable = PCache.CreatePivotTable(TableDestination:=newtab.Cells(2, 15), TableName:=GLCode & "Pivot")

                    With ActiveSheet.PivotTables(GLCode & "Pivot").PivotFields("PO")
                    .Orientation = xlRowField
                    .Position = 1
                    End With

                    With ActiveSheet.PivotTables(GLCode & "Pivot").PivotFields("INV $")
                    .Orientation = xlDataField
                    .Function = xlSum
                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"
                    End With

                    With ActiveSheet.PivotTables(GLCode & "Pivot").PivotFields("CREDIT $")
                    .Orientation = xlDataField
                    .Function = xlSum
                    .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* " & Chr(34) & "-" & Chr(34) & "??_);_(@_)"
                    End With
           End If
            newtab.UsedRange.EntireColumn.AutoFit
            
        End If
        
    Next cell
    
    sltab.Activate
    Application.ScreenUpdating = True

End Sub







