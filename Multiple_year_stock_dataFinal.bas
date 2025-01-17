Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data():

Dim TickerName As String
Dim TickerTotal As LongLong
Dim YearChg As Double
Dim PercentChg As Double
Dim TickerOpenSum As Double
Dim TickerCloseSum As Double


For x = 1 To Worksheets.Count
Worksheets(x).Select

TickerTotal = 0
TickerOpenSum = 0
TickerCloseSum = 0

SummaryTableRow = 2

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            TickerName = Cells(i, 1).Value
            TickerOpenSum = TickerOpenSum + Cells(i, 3).Value
            TickerCloseSum = TickerCloseSum + Cells(i, 6).Value
            TickerTotal = TickerTotal + Cells(i, 7).Value
            YearChg = TickerOpenSum - TickerCloseSum
            If TickerOpenSum = 0 Then
                PercentChg = 0
            Else
                PercentChg = (YearChg / TickerOpenSum) * 100
            End If
            Range("I" & SummaryTableRow).Value = TickerName
            Range("J" & SummaryTableRow).Value = YearChg
                Range("J" & SummaryTableRow).NumberFormat = "0.00"
                If YearChg > 0 Then
                    Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                ElseIf YearChg < 0 Then
                    Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            Range("K" & SummaryTableRow).Value = PercentChg
                Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            Range("L" & SummaryTableRow).Value = TickerTotal
            SummaryTableRow = SummaryTableRow + 1
            TickerTotal = 0
            TickerOpenSum = 0
            TickerCloseSum = 0
        
        Else
            TickerOpenSum = TickerOpenSum + Cells(i, 3).Value
            TickerCloseSum = TickerCloseSum + Cells(i, 6).Value
            TickerTotal = TickerTotal + Cells(i, 7).Value
        End If
        
    Next i

'Bonus Section

    Dim max As Double
    Dim min As Double
    Dim maxvol As LongLong

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    For i = 2 To lastrow

        If Cells(i, 11) > max Then
            max = Cells(i, 11)
            tickermax = Cells(i, 9)
        End If
    
        If Cells(i, 11) < min Then
            min = Cells(i, 11)
            tickermin = Cells(i, 9)
        End If
    
        If Cells(i, 12) > maxvol Then
            maxvol = Cells(i, 12)
            tickervol = Cells(i, 9)
        End If
    
    Next i

    Range("Q2").Value = max
        Range("Q2").NumberFormat = "0.00%"
    Range("P2").Value = tickermax
    Range("Q3").Value = min
        Range("Q3").NumberFormat = "0.00%"
    Range("P3").Value = tickermin
    Range("Q4").Value = maxvol
    Range("P4").Value = tickervol
    
    max = 0
    min = 0
    maxvol = 0

    Columns("O:Q").Select
    Selection.EntireColumn.AutoFit
    
Next x

End Sub
