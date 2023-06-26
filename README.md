# VBA-Challange
Sub multiple_year_stock()
    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim Rng As Range
        Dim Lastrow As Long
        WorksheetName = ws.Name
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Set Rng = ws.Range("J2:J" & Lastrow)
        ws.Range("A2:A" & Lastrow).AdvancedFilter Action:=xlFilterCopy, copyToRange:=Rng, Unique:=True
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Start value"
        ws.Cells(1, 12).Value = "End value"
        ws.Cells(1, 13).Value = "yearly change"
        ws.Cells(1, 14).Value = "percent change"
        ws.Cells(1, 15).Value = "Totalstockvolume"
        ws.Cells(5, 17).Value = "Greatest%increase"
        ws.Cells(6, 17).Value = "Greatest%decrease"
        ws.Cells(7, 17).Value = "Greatestvolumetotal"
        ws.Cells(4, 18).Value = "Ticker"
        ws.Cells(4, 19).Value = "value"
        Dim EndRow As Long
        EndRow = Rng.End(xlDown).Row
        Dim rg As Range
        Dim rg1 As Range
        For x = 2 To EndRow
                ws.Cells(x, 15) = WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(x, 10), ws.Range("G:G"))
                Set rg = ws.Columns("A:A").Find(What:=ws.Cells(x, 10), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
                Set rg1 = ws.Columns("A:A").Find(What:=ws.Cells(x, 10), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
                ws.Cells(x, 11) = ws.Cells(rg.Row, 3).Value
                ws.Cells(x, 12) = ws.Cells(rg1.Row, 6).Value
                ws.Cells(x, 13) = ws.Cells(rg1.Row, 6).Value - ws.Cells(rg.Row, 3).Value
                If ws.Cells(x, 13) > 0 Then
                   ws.Cells(x, 13).Interior.ColorIndex = 4
                Else
                    ws.Cells(x, 13).Interior.ColorIndex = 3
                End If
                ws.Cells(x, 14) = (ws.Cells(rg1.Row, 6).Value - ws.Cells(rg.Row, 3).Value) / ws.Cells(rg.Row, 3).Value
                If ws.Cells(x, 14) > 0 Then
                    ws.Cells(x, 14).Interior.ColorIndex = 4
                Else
                    ws.Cells(x, 14).Interior.ColorIndex = 3
                End If
                ws.Cells(5, 19) = WorksheetFunction.Max(ws.Range("N:N"))
                ws.Cells(6, 19) = WorksheetFunction.Min(ws.Range("N:N"))
             ws.Cells(7, 19) = WorksheetFunction.Max(ws.Range("o:o"))

        Next x
    Next ws
End Sub
