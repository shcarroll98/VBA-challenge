Attribute VB_Name = "Module1"


Sub TickerAnalysis()
    Dim Ticker As String
    Dim Symbol As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Dollar_Change As Double
    Dim Percent_Change As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim Ticker_Total As Double
    Dim ws As Worksheet
    Dim Summary_Table_Row As Long
    Dim lastrow As Long
    Dim lastrowsummarytable As Long
    Dim j As Long
    Dim k As Long

'Loop through worksheets
    For Each ws In Worksheets
    
    'Initialize Summary Table and Ticker
        Summary_Table_Row = 2
        Ticker_Total = 0

        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
'For loop through ticker symbols
        For I = 2 To lastrow
        'When the next ticker symbol is not the same as the previous ticker symbol
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                Ticker = ws.Cells(I, 1).Value
                'Add stock volume
                Ticker_Total = Ticker_Total + ws.Cells(I, "G").Value

'Initialize Open and Close Price to get Dollar and Percentage Change
                OpenPrice = ws.Cells(I + 1, "C")
                ClosePrice = ws.Cells(I, "F")
'Calculate dollar and percentage change
                Dollar_Change = ClosePrice - OpenPrice
                Percent_Change = Dollar_Change / Cells(2, "C").Value * 100
'Print values for each symbol from the for loop
                ws.Cells(Summary_Table_Row, "I").Value = Ticker
                ws.Cells(Summary_Table_Row, "J").Value = Dollar_Change
                ws.Cells(Summary_Table_Row, "K").Value = Percent_Change
                ws.Cells(Summary_Table_Row, "L").Value = Ticker_Total
 'format cells
               Cells(Summary_Table_Row, "K").NumberFormat = "0.00%"
'print next item from for loop in Next row
                Summary_Table_Row = Summary_Table_Row + 1
            Else
                Ticker_Total = Ticker_Total + ws.Cells(I, "G").Value
            End If
            
'End first for loop
        Next I

'Define rows for color change/ greatest amounts
        lastrowsummarytable = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
'For loop for colors
        For j = 2 To lastrowsummarytable
        'make positive amounts green
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
        'make negative amounts red
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

        GreatestDecrease = ws.Cells(2, "M").Value
        GreatestIncrease = ws.Cells(2, "M").Value
        GreatestVolume = ws.Cells(2, "N").Value
'for loop for greatest % decrease/increase and volume
        For k = 2 To lastrowsummarytable
            If ws.Cells(k, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(k, 11).Value
                Cells(2, "Q").NumberFormat = "0.00%"
                'get the matching symbol
                Cells(2, "P").Value = Cells(k, "I").Value
                
            End If

            If ws.Cells(k, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(k, 11).Value
                Cells(3, "Q").NumberFormat = "0.00%"
                Cells(3, "P").Value = Cells(k, "I").Value
            End If

            If ws.Cells(k, 12).Value > GreatestVolume Then
                GreatestVolume = Cells(k, 12).Value
                Cells(4, "P").Value = Cells(k, "I").Value
            End If
        Next k
'row/column headers
        ws.Cells(2, "Q").Value = GreatestIncrease
        ws.Cells(3, "Q").Value = GreatestDecrease
        ws.Cells(4, "Q").Value = GreatestVolume

        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
    Next ws
End Sub














