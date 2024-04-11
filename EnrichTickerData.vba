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
    Dim col As ListColumn
    Dim greatestYearlyChange as Double
    Dim greatestPercentChange as Double
    Dim greatestTotalVolume as Double

    ' loop each sheet
    For Each ws In ThisWorkbook.Worksheets
        ' create output table for ticker details
        Set tblRange = ws.Range("I1:L1")
        tblName = "TickerSummary"

        ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = tblName
        Set tbl = ws.ListObjects(tblName)
        
        ' format numbers of new table
        ws.Columns("J").NumberFormat = "0.00"
        ws.Columns("K").NumberFormat = "%0.00"
        
        lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
        For row = 1 To lastRow
            ' Headers - print them and grab initial values
            If row = 1 Then
                Set col = tbl.ListColumns("Column1")
                col.Name = "Ticker"
                Set col = tbl.ListColumns("Column2")
                col.Name = "Yearly Change"
                Set col = tbl.ListColumns("Column3")
                col.Name = "Percent Change"
                Set col = tbl.ListColumns("Column4")
                col.Name = "Total Stock Volume"

                ' grab the initial values here
                beginPrice = Cells(row + 1, "C")
                totalvolume = Cells(row + 1, "G")
            
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

        ' greatest increase, greatest decrease, and greatest total volume

        ' create output table for ticker details
        Set tblRange = ws.Range("O1:Q1")
        greatTblName = "GreatestSummary"

        ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = greatTblName
        Set greatTbl = ws.ListObjects(greatTblName)

        ' set column headers
        Set col = greatTbl.ListColumns("Column1")
        col.Name = " "

        Set col = greatTbl.ListColumns("Column2")
        col.Name = "Ticker"

        Set col = greatTbl.ListColumns("Column3")
        col.Name = "Value"

        ' check every row in ticker details table to bring out the greatest
        Set tbl = ws.ListObjects("TickerSummary")
        lastRow = tbl.DataBodyRange.Rows.Count + tbl.HeaderRowRange.Row - 1

        ' set initial greatest values
        greatestPercentIncrease = 0
        greatestPercentDecrease = 0
        greatestTotalVolume = 0
        Dim greatestPercentIncrease_Ticker as String
        Dim greatestPercentDecrease_Ticker as String
        Dim greatestTotalVolume_Ticker as String

        For row = 2 To lastRow
            ' check greatest increase
            If Cells(row, "K").Value > greatestPercentIncrease Then
                greatestPercentIncrease = Cells(row, "K").Value
                greatestPercentIncrease_Ticker = Cells(row, "I").Value
            End If

            ' check greatest decrease
            If Cells(row, "K").Value < greatestPercentDecrease Then
                greatestPercentDecrease = Cells(row, "K").Value
                greatestPercentDecrease_Ticker = Cells(row, "I").Value
            End If

            ' check total stock volume
            If Cells(row, "L").Value > greatestTotalVolume Then
                greatestTotalVolume = Cells(row, "L").Value
                greatestTotalVolume_Ticker = Cells(row, "I").Value
            End If
        Next row

        ' write out greatest values to greatest table
        Set newRow = greatTbl.ListRows.Add(AlwaysInsert:=True)
        With newRow
            .Range(1) = "Greatest % Increase"
            .Range(2) = greatestPercentIncrease_Ticker
            .Range(3) = greatestPercentIncrease
        End With

        Set newRow = greatTbl.ListRows.Add(AlwaysInsert:=True)
        With newRow
            .Range(1) = "Greatest % Decrease"
            .Range(2) = greatestPercentDecrease_Ticker
            .Range(3) = greatestPercentDecrease
        End With

        Set newRow = greatTbl.ListRows.Add(AlwaysInsert:=True)
        With newRow
            .Range(1) = "Greatest Total Volume"
            .Range(2) = greatestTotalVolume_Ticker
            .Range(3) = greatestTotalVolume
        End With

    Next ws
End Sub
