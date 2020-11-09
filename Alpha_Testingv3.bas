Attribute VB_Name = "Module1"
Sub alphbetical_testing():

Dim TickerName As String
Dim TickerTotal As Long
Dim YearChg As Double
Dim PercentChg As Double
Dim TickerOpenSum As Double
Dim TickerCloseSum As Double

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
        TickerCloseSum = TickerCloseSum + Cells(i, 5).Value
        TickerTotal = TickerTotal + Cells(i, 7).Value
        YearChg = TickerOpenSum - TickerCloseSum
        PercentChg = YearChg / TickerOpenSum
        Range("I" & SummaryTableRow).Value = TickerName
        Range("J" & SummaryTableRow).Value = YearChg
        Range("K" & SummaryTableRow).Value = PercentChg
            If PercentChg >= 0 Then
                Range("K" & SummaryTableRow).Interior.ColorIndex = 4
            Else
                Range("K" & SummaryTableRow).Interior.ColorIndex = 3
            End If
        Range("L" & SummaryTableRow).Value = TickerTotal
        TickerTotal = 0
        TickerOpenSum = 0
        TickerCloseSum = 0
        
    Else
        TickerOpenSum = TickerOpenSum + Cells(i, 3).Value
        TickerCloseSum = TickerCloseSum + Cells(i, 5).Value
        TickerTotal = TickerTotal + Cells(i, 7).Value
    End If
    
    
    
Next i


End Sub
