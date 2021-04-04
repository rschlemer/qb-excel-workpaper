Attribute VB_Name = "Online"
Sub QBO_Workpapers(control As IRibbonControl)
    Dim intLastIS As Integer, intLastRowIS As Integer, intLastRowBS As Integer

    Application.ScreenUpdating = False
    Application.StatusBar = "Hang on..."
        
    'left in, doesn't seem to be doing anything except maybe AJE sheet
    With ActiveWorkbook.Styles("Normal").Font
        .Name = "Arial"
        .Size = 10
    End With
    
    Call Rename_ISandBS_QBO
    
    Worksheets(1).Activate
    Call Flat_ISandBS_QBO
    
    Worksheets(2).Activate
    Call Flat_ISandBS_QBO
    
    Call AJE_Sheet_QBO
    
    intLastRowIS = Worksheets("income statement").Cells(7, 1).End(xlDown).Row
    intLastRowBS = Worksheets("balance sheet").Cells(7, 1).End(xlDown).Row
    
    Call Income_Sheet_QBO
    
    Call Balance_Sheet_QBO(intLastRowIS + 2)
    
    'deleting rows, adding title to adjusted column, centering title, changing font
    Worksheets("Income Statement").Activate
    Call Sum_Columns_QBO(intLastRowIS)
    Range("F6").Value = "Adjusted"
    'Rows(intLastRowIS - 2 & ":" & intLastRowIS).Delete
    Rows(5).Delete
    For i = 1 To 3
        Range("A" & i & ":G" & i).HorizontalAlignment = xlCenterAcrossSelection
    Next i
    With ActiveSheet.Cells.Font
        .Name = "Arial"
        .Size = 10
    End With
    Columns(1).AutoFit
    Worksheets("Balance Sheet").Activate
    Call Sum_Columns_QBO(intLastRowBS)
    Range("F6").Value = "Adjusted"
    Rows(5).Delete
    'Rows(intLastRowBS - 4 & ":" & intLastRowBS).Delete
    For i = 1 To 3
        Range("A" & i & ":G" & i).HorizontalAlignment = xlCenterAcrossSelection
    Next i
    With ActiveSheet.Cells.Font
        .Name = "Arial"
        .Size = 10
    End With
    Columns(1).AutoFit

    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
        Application.PrintCommunication = True
        Range("a1").Select
    Next ws
    
    Worksheets("Balance Sheet").Activate
    
    Application.CutCopyMode = False
    Application.StatusBar = "Done"
    Application.OnTime Now + TimeValue("00:00:05"), "ClearStatus"
    Application.ScreenUpdating = True
End Sub
Function Sum_Columns_QBO(intLastRowSheet As Integer)
    Columns("E:E").Insert
    Columns("E:E").ClearFormats
    
    With Range("E:E")
        .NumberFormat = "0"
        .ColumnWidth = 2.86
    End With
    
    For i = 1 To 3
        Range("E" & intLastRowSheet + i + 3).Value = i
        Range("F" & intLastRowSheet + i + 3).Formula = _
        "=sumif($E$1:$E$" & intLastRowSheet & ",$E" & intLastRowSheet + i + 3 & ",F$1:F$" & intLastRowSheet & _
        ")-sumif($E$1:$E$" & intLastRowSheet & ",-$E" & intLastRowSheet + i + 3 & ",F$1:F$" & intLastRowSheet & ")"
    Next i
End Function

Function Income_Sheet_QBO()
    Worksheets("Income Statement").Activate
    
    Dim intIncRow As Integer
    Dim intCogRow As Integer
    Dim intExpRow As Integer
    Dim intOthIncRow As Integer
    Dim intOthExpRow As Integer
    Dim intLastRow As Integer
    
    intLastRow = Cells(7, 1).End(xlDown).Row + 1
    
    Columns("C:E").Insert
    Columns("C:E").ClearFormats
    
    'finds intOthExpRow
    For i = 1 To intLastRow
        If LCase(Cells(i, 1).Value) = "other expenses" Then
            intOthExpRow = i
            Exit For
        End If
    Next i
    If intOthExpRow = 0 Then intOthExpRow = intLastRow
    
    'finds intOthIncRow
    For i = 1 To intLastRow
        If LCase(Cells(i, 1).Value) = "other income" Then
            intOthIncRow = i
            Exit For
        End If
    Next i
    If intOthIncRow = 0 Then intOthIncRow = intOthExpRow
    
    'finds inexprow number
    For i = 1 To intLastRow
        If LCase(Cells(i, 1).Value) = "expenses" Then
            intExpRow = i
            Exit For
        End If
    Next i
    If intExpRow = 0 Then intExpRow = intOthIncRow
    
    'finds intCogRow
    For i = 1 To intLastRow
        If LCase(Cells(i, 1).Value) = "cost of goods sold" Then
            intCogRow = i
            Exit For
        End If
    Next i
    If intCogRow = 0 Then intCogRow = intExpRow
    
    'finds intincrow number
    For i = 1 To 10
        If LCase(Cells(i, 1).Value) = "income" Then
            intIncRow = i
            Exit For
        End If
    Next i
    If intIncRow = 0 Then intIncRow = intCogRow
       
    'formats
    With Range("B" & intIncRow, "F" & intLastRow + 10)
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""_);_(@_)"
        .ColumnWidth = 13.57
    End With
    
    'creates aje formulas and adjusted column formulas
    For i = intIncRow To intLastRow
        If Cells(i, 2).Value <> "" And Cells(i, 2).Font.Bold = False Then
            Cells(i, 3).Formula = "=sumif(dname,a" & i & ",dval)"
            Cells(i, 4).Formula = "=sumif(cname,a" & i & ",cval)"
            If i < intCogRow Then
                Cells(i, 5).Formula = "=B" & i & "-" & "C" & i & "+" & "D" & i
            ElseIf i < intExpRow Then
                Cells(i, 5).Formula = "=B" & i & "+" & "C" & i & "-" & "D" & i
            ElseIf i < intOthIncRow Then
                Cells(i, 5).Formula = "=B" & i & "+" & "C" & i & "-" & "D" & i
            ElseIf i < intOthExpRow Then
                Cells(i, 5).Formula = "=B" & i & "-" & "C" & i & "+" & "D" & i
            Else
                Cells(i, 5).Formula = "=B" & i & "+" & "C" & i & "-" & "D" & i
            End If
        End If
    Next i
    
    'creates totals formulas
    For Each cell In Range("b7:b" & intLastRow)
        If cell.Font.Bold = True Then
            If InStr(cell.Formula, "-") = False Then
                Cells(cell.Row, 5).Formula = "=sum(" & Sum_QBO_Range(cell.Formula) & ")"
            Else
                Cells(cell.Row, 2).Copy
                Cells(cell.Row, 5).PasteSpecial Paste:=xlPasteFormulas
            End If
        End If
    Next
    
    'formats adjusted column same as current year column
    Columns("B:B").Copy
    Columns("E:E").PasteSpecial Paste:=xlPasteFormats
    
    'totals ajes for income
    Range("C" & intLastRow + 1).Formula = "=sum(c1:c" & intLastRow - 1 & ")"
    Range("D" & intLastRow + 1).Formula = "=sum(D1:D" & intLastRow - 1 & ")"
    
End Function
Function Balance_Sheet_QBO(intIncomeStatement As Integer)
    Worksheets("Balance Sheet").Activate
    
    Dim intAssRow As Integer
    Dim intLiaRow As Integer
    Dim intLastRow As Integer
    
    intLastRow = Cells(7, 1).End(xlDown).Row
    
    Columns("C:E").Insert
    Columns("C:E").ClearFormats
    
    'finds intLiaRow number
    For i = 1 To intLastRow
        If LCase(Cells(i, 1).Value) = "liabilities and equity" Then
            intLiaRow = i
            Exit For
        End If
    Next i
    If intLiaRow = 0 Then intLiaRow = intLastRow
    
    'finds intAssRow number
    For i = 1 To 10
        If LCase(Cells(i, 1).Value) = "assets" Then
            intAssRow = i
            Exit For
        End If
    Next i
    If intAssRow = 0 Then intAssRow = intLiaRow
    
    'formats
    With Range("B" & intAssRow, "F" & intLastRow + 10)
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""_);_(@_)"
        .ColumnWidth = 13.57
    End With
    
    'creates aje formulas and adjusted column formulas
    For i = intAssRow To intLastRow
        If Cells(i, 2).Value <> "" And Cells(i, 2).Font.Bold = False Then
            Cells(i, 3).Formula = "=sumif(dname,a" & i & ",dval)"
            Cells(i, 4).Formula = "=sumif(cname,a" & i & ",cval)"
            If i < intLiaRow Then
                Cells(i, 5).Formula = "=B" & i & "+" & "C" & i & "-" & "D" & i
            Else
                Cells(i, 5).Formula = "=B" & i & "-" & "C" & i & "+" & "D" & i
            End If
        End If
    Next i
    
    For Each cell In Range("b7:b" & intLastRow)
        If cell.Font.Bold = True Then
            If InStr(cell.Formula, "-") = False Then
                Cells(cell.Row, 5).Formula = "=sum(" & Sum_QBO_Range(cell.Formula) & ")"
            Else
                Cells(cell.Row, 2).Copy
                Cells(cell.Row, 5).PasteSpecial Paste:=xlPasteFormulas
            End If
        End If
    Next
    
    'formats adjusted column same as current year column
    Columns("B:B").Copy
    Columns("E:E").PasteSpecial Paste:=xlPasteFormats
    
    'totals ajes for income
    For i = 1 To intLastRow
        If LCase(Cells(i, 1).Value) = "net income" Then
            Cells(i, 3).Formula = "='Income Statement'!C" & intIncomeStatement
            Cells(i, 4).Formula = "='Income Statement'!D" & intIncomeStatement
        End If
    Next i
    
End Function
Function ClearStatus_QBO()
    Application.StatusBar = False
End Function

Function Rename_ISandBS_QBO()
    Worksheets(1).Activate
    For i = 1 To Range("A1").SpecialCells(xlCellTypeLastCell).Row
        If LCase(Cells(i, 1).Value) = "assets" Or LCase(Cells(i, 1).Value) = "liabilities and equity" Then
            ActiveSheet.Name = "Balance Sheet"
            Worksheets(2).Name = "Income Statement"
            Exit For
        End If
        If i = Range("A1").SpecialCells(xlCellTypeLastCell).Row Then
            ActiveSheet.Name = "Income Statement"
            Worksheets(2).Name = "Balance Sheet"
        End If
    Next i
    Worksheets("Balance Sheet").Move Before:=ActiveWorkbook.Sheets(1)
    intLastRowIS = Worksheets("income statement").Range("A1").SpecialCells(xlCellTypeLastCell).Row
    intLastRowBS = Worksheets("balance sheet").Range("A1").SpecialCells(xlCellTypeLastCell).Row
    Worksheets("income statement").Rows(intLastRowIS).Delete
    Worksheets("Balance Sheet").Rows(intLastRowBS).Delete
    
End Function
Function Flat_ISandBS_QBO()
    Dim intLastRow As Integer
    Dim intTrimLen As Integer, intValLen As Integer
    
    intLastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    With ActiveSheet.Cells
        .UnMerge
        .WrapText = False
    End With
    
    Range("B6", "C6").WrapText = True
    
    For i = 1 To intLastRow - 4
        intTrimLen = Len(Trim(Cells(i + 4, 1).Value))
        intValLen = Len(Cells(i + 4, 1).Value)
        If intValLen > intTrimLen Then
            Cells(i + 4, 1).IndentLevel = 2 * ((intValLen - intTrimLen) / 3)
            Cells(i + 4, 1).Value = Trim(Cells(i + 4, 1).Value)
        End If
    Next i
    
    Cells.EntireColumn.AutoFit
End Function
Function AJE_Sheet_QBO()
    Sheets.Add after:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "AJE's"
    
    Range("A1").Value = Worksheets("income statement").Range("A1").Value
    Range("A1:E1").HorizontalAlignment = xlCenterAcrossSelection
    Range("A2").Value = "AJE's"
    Range("A2:E2").HorizontalAlignment = xlCenterAcrossSelection
    Range("A4").Value = 1

    ActiveWorkbook.Names.Add "dName", Worksheets("AJE's").Range("B:B")
    ActiveWorkbook.Names.Add "cName", Worksheets("AJE's").Range("C:C")
    ActiveWorkbook.Names.Add "dVal", Worksheets("AJE's").Range("D:D")
    ActiveWorkbook.Names.Add "cVal", Worksheets("AJE's").Range("E:E")
    
    Range("A:B").ColumnWidth = 2.86
    Range("C:C").ColumnWidth = 45
    With Range("D:E")
        .ColumnWidth = 13.57
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""_);_(@_)"
    End With
    
End Function
Function Sum_QBO_Range(strFormula As String) As String
    Dim arrFormula As Variant
    Dim arrSkips As Variant

        strFormula = Replace(Replace(Replace(Replace(strFormula, "(", ""), ")", ""), "+", ""), "=", "")
        arrFormula = Split(strFormula, "B")
        x = Int(arrFormula(1))
        For i = 1 To UBound(arrFormula)
            If Int(arrFormula(i)) <> x + i - 1 Then
                y = y & i & ","
                x = x + arrFormula(i) - arrFormula(i - 1) - 1
            End If
        Next i
        If y <> Empty Then
            y = Left(y, Len(y) - 1)
            arrSkips = Split(y, ",")
                    Sum_QBO_Range = "E" & arrFormula(1)
            For Each element In arrSkips
                Sum_QBO_Range = Sum_QBO_Range & ":E" & arrFormula(Int(element) - 1) & ",E" & arrFormula(Int(element))
            Next element
            Sum_QBO_Range = Sum_QBO_Range & ":E" & arrFormula(UBound(arrFormula))
        Else
            Sum_QBO_Range = "E" & arrFormula(1) & ":E" & arrFormula(UBound(arrFormula))
        End If
        

End Function


