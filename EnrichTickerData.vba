' TODO: run on every worksheet at once
' TODO: conditional formatting

Sub EnrichTickerData()
    Dim row As Long
    Dim column As Integer
    Dim ticker As String
    Dim beginPrice As Double
    Dim endPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalvolume as Double
    
    Columns("J").NumberFormat = "0.00"
    
    For row = 1 To 22771
        ' Headers - print them and grab initial values
        If row = 1 Then
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"

        ' First seeing this new stock ticker, doesn't equal last ticker
        ElseIf Cells(row, "A").Value <> Cells(row - 1, "A").Value Then
            ' grab the initial values here
            ticker = Cells(row + 1, "A").Value
            beginPrice = Cells(row + 1, "C")
            totalvolume = Cells(row + 1, "G")
            
            ' output the ticker symbol (since its new)
            Cells(row + 1, "I").Value = ticker
        
        ' Last entry of this ticker, do final calculations, doesn't equal next ticker
        ElseIf Cells(row, "A").Value <> Cells(row + 1, "A") Then
                endPrice = Cells(row, "F").Value

                ' yearly change
                yearlyChange = (endPrice - beginPrice) * 100
                Cells(row, "J").Value = yearlyChange

                ' percentage change
                percentChangeChange = ((endPrice - beginPrice) / beginPrice) * 100
                Cells(row, "K").Value = percentChange

                ' total stock volume
                totalvolume = totalvolume + Cells(row, "G").Value
                Cells(row, "L").Value = totalvolume

                ' greatest increase, greatest decrease, and greatest total volume

        ' same stock - calculate this row
        Else
            totalvolume = totalvolume + Cells(row, "G").Value
        End If
    Next row
End Sub
