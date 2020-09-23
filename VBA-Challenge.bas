Attribute VB_Name = "Module1"

Sub VBA_Challenge()

'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.The total stock volume of the stock.
    'You should also have conditional formatting that will highlight positive change in green and negative change in red.
    'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
    'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
    


    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest Percent Increase"
        ws.Range("N3").Value = "Greatest Percent Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
        Dim LastRowValue As Long
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

        TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                TickerName = ws.Cells(i, 1).Value
          
                ws.Range("I" & SummaryTableRow).Value = TickerName
                ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                TotalTickerVolume = 0

                Open_Price = ws.Range("C" & PreviousAmount)
                Close_Price = ws.Range("F" & i)
                YearlyChange = Close_Price - Open_Price
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                If Open_Price = 0 Then
                    PercentChange = 0
                Else
                    Open_Price = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / Open_Price
                End If
               
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
                
            Next i

            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
       
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
    Next ws

End Sub
