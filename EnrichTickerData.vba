' TODO: run on every worksheet at once
' TODO: conditional formatting

Sub EnrichTickerData()
    Dim ws as worksheet
    Dim tblRange as Range
    Dim tblName as String
    Dim newRow as ListRow
    Dim lastRow as Long
    Dim row As Long
    Dim column As Integer
    Dim ticker As String
    Dim beginPrice As Double
    Dim endPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalvolume as Double

    'Set ws = ThisWorkbook.Worksheets("2018")
    For Each ws In ThisWorkbook.Worksheets
        Set tblRange = ws.Range("I1:L1")
        tblName = "TickerSummary"

        ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = tblName
        'Set tbl = ThisWorkbook.Worksheets("2018").ListObjects(tblName)
        Set tbl = ws.ListObjects(tblName)
        
        Columns("J").NumberFormat = "0.00"
        Columns("K").NumberFormat = "%0.00"
        
        lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
        For row = 1 To lastRow
            ' Headers - print them and grab initial values
            If row = 1 Then
                Range("I1").Value = "Ticker"
                Range("J1").Value = "Yearly Change"
                Range("K1").Value = "Percent Change"
                Range("L1").Value = "Total Stock Volume"

                ' grab the initial values here
                'ticker = Cells(row + 1, "A").Value
                beginPrice = Cells(row + 1, "C")
                totalvolume = Cells(row + 1, "G")
                
                ' output the ticker symbol (since its new)
                'Cells(row + 1, "I").Value = ticker
            
            ' Last entry of this ticker, do final calculations, doesn't equal next ticker
            ElseIf Cells(row, "A").Value <> Cells(row + 1, "A") Then
                    ticker = Cells(row, "A").Value
                    endPrice = Cells(row, "F").Value

                    ' yearly change
                    yearlyChange = endPrice - beginPrice
                    Cells(row, "J").Value = yearlyChange

                    ' percentage change
                    percentChange = ((endPrice - beginPrice) / beginPrice)
                    Cells(row, "K").Value = percentChange

                    ' total stock volume
                    totalvolume = totalvolume + Cells(row, "G").Value
                    Cells(row, "L").Value = totalvolume

                    ' greatest increase, greatest decrease, and greatest total volume

                    ' write to output table
                    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
                    With newRow
                        .Range(1) = ticker
                        .Range(2) = yearlyChange
                        .Range(3) = percentChange
                        .Range(4) = totalvolume
                    End With

                    ' grab the next stock's initial values
                    beginPrice = Cells(row + 1, "C")
                    totalvolume = Cells(row + 1, "G")

            ' same stock - calculate this row
            Else
                totalvolume = totalvolume + Cells(row, "G").Value
            End If
        Next row

        ' conditional formatting - Yearly Change
        'Dim colRange As Range
        'Set colRange = ws.Range("J:J")
        'With colRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            ' Format for values less than 0 (red)
        '    .Interior.ColorIndex = xlAutomatic
        '    .Color = vbRed
        'End With

        'With colRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            ' Format for values greater than 0 (green)
        '    .Interior.ColorIndex = xlAutomatic
        '    .Color = vbGreen
        'End With

        ' conditional formatting - Percent Change
        'Set colRange = ws.Range("K:K")
        'With colRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            ' Format for values less than 0 (red)
        '    .Interior.ColorIndex = xlAutomatic
        '    .Color = vbRed
        'End With

        'With colRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        '    ' Format for values greater than 0 (green)
        '    .Interior.ColorIndex = xlAutomatic
        '    .Color = vbGreen
        'End With

    Next ws
End Sub
