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
    
    Columns("J").NumberFormat = "0.00"
    
    For row = 1 To 250
        If row = 1 Then
            ' print the headers
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
            
            ' grab the first ticker values from here
            ticker = Cells(row + 1, "A").Value
            
            ' output the ticker symbol (if its new)
            Cells(row + 1, "I").Value = ticker
        Else
            If Cells(row + 1, "A").Value <> Cells(row, "A") Then
                ' last run, final calculate, then print
            Else
                ' calculate next row
                
            End If
            ' output the ticker symbol - TODO: if its new
            'ticker = Cells(row, "A").Value
            'Cells(row, "I").Value = ticker
            
            ' yearly change
            'beginPrice = Cells(row, "C").Value
            'endPrice = Cells(row, "F").Value
            'yearlyChange = (endPrice - beginPrice) * 100
            'Cells(row, "J").Value = yearlyChange
            
            ' percentage change
            'percentChange = ((endPrice - beginPrice) / beginPrice) * 100
            'Cells(row, "K").Value = percentChange
            
            ' total stock volume
            
            ' greatest increase, greatest decrease, and greatest total volume
        End If
    Next row
End Sub
