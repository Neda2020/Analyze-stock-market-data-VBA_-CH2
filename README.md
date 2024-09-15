# Analyze-stock-market-data-VBA_-CH2
VBA scripting to analyze generated stock market data-Challenge 2

"-Quarterly Change=Closing Price(the last date of closing price of AAF)−Opening Price(the first date of opening price of AAF)
-Percentage Change=(Opening PriceQuarterly Change​)×100
-Total Volume=∑(Daily Volume)for all days in the quarter for each ticker


Sub StockAnalysis()

    ' Loop through each worksheet
    Dim ws As Worksheet
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String

    ' Initialize variables for greatest calculations
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0

    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        
        ' Define variables
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim TotalVolume As Double
        Dim LastRow As Long
        Dim Row As Long
        Dim StartRow As Long
        
        ' Initialize
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        StartRow = 2
        TotalVolume = 0
        
        ' Create output columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        Dim OutputRow As Integer
        OutputRow = 2
        
        ' Loop through the rows of the data
        For Row = 2 To LastRow
            TotalVolume = TotalVolume + ws.Cells(Row, 7).Value  ' Sum up the volume
            
            ' Check if the ticker symbol changes (end of the stock data for this ticker)
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                ' Set the ticker
                Ticker = ws.Cells(Row, 1).Value
                
                ' Get the open price
                OpenPrice = ws.Cells(StartRow, 3).Value
                
                ' Get the closing price
                ClosePrice = ws.Cells(Row, 6).Value
                
                ' Calculate quarterly change and percentage change
                Dim QuarterlyChange As Double
                Dim PercentChange As Double
                QuarterlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ' Output the results
                ws.Cells(OutputRow, 9).Value = Ticker
                ws.Cells(OutputRow, 10).Value = QuarterlyChange
                ws.Cells(OutputRow, 11).Value = PercentChange
                ws.Cells(OutputRow, 12).Value = TotalVolume
                
                ' Conditional formatting for Quarterly Change and Percent Change
                If QuarterlyChange > 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = vbGreen
                ElseIf QuarterlyChange < 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = vbRed
                Else
                    ws.Cells(OutputRow, 10).Interior.Color = vbWhite
                End If
                
                          
                ' Move to the next output row
                OutputRow = OutputRow + 1
                
                ' Reset for the next ticker
                TotalVolume = 0
                StartRow = Row + 1
            End If
        Next Row
    Next ws
    
    

End Sub

Sub AddMaxMinFormulas()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")

    ' Place the headers for the results
    ws.Cells(7, 17).Value = "Greatest % Increase"
    ws.Cells(8, 17).Value = "Greatest % Decrease"
    ws.Cells(9, 17).Value = "Greatest Total Volume"
    ws.Cells(6, 18).Value = "Ticker"
    ws.Cells(6, 19).Value = "Value"
    
    ' Apply the MAX formula for Greatest % Increase
    ws.Cells(7, 19).Formula = "=MAX(K:K)" ' Adjust if needed based on your data
    
    ' Apply the MIN formula for Greatest % Decrease
    ws.Cells(8, 19).Formula = "=MIN(K:K)" ' Adjust if needed based on your data
    
    ' Apply the MAX formula for Greatest Total Volume
    ws.Cells(9, 19).Formula = "=MAX(L:L)" ' Adjust if needed based on your data

    ' Optionally, set the corresponding tickers (assuming ticker symbols are in column J)
    ws.Cells(7, 18).Formula = "=INDEX(I:I,MATCH(MAX(K:K),K:K,0))"
    ws.Cells(8, 18).Formula = "=INDEX(I:I,MATCH(MIN(K:K),K:K,0))"
    ws.Cells(9, 18).Formula = "=INDEX(I:I,MATCH(MAX(L:L),L:L,0))"
    
    
End Sub
