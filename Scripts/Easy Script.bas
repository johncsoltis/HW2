Attribute VB_Name = "Module1"
'Loop through one year of stock data for each run and return the total volume each stock had
'Display the ticker symbol to coincide with the total stock volume.

Sub StocksEasy()

'Define variables------
Dim totalvolume As Double
Dim resultrow As Long
Dim i As Long
Dim lastrow As Long
Dim ws As Worksheet

'Loop through worksheets-------
For Each ws In Worksheets
    'Activate workseet
    ws.Activate
    
    'Add labels
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Volume"

    'Find last row on sheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'Define output row
    resultrow = 2

        'Loop through all tickers
        For i = 2 To lastrow
        
        'Add volume to the total ticker volume
        totalvolume = totalvolume + Cells(i, 7).Value
            
            'If next ticker different, print total volume and ticker symbol to results chart
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(resultrow, 9).Value = Cells(i, 1).Value
            Cells(resultrow, 10).Value = totalvolume
            
            'Reset variables for next ticker
            resultrow = resultrow + 1
            totalvolume = 0
    
            End If
    
    Next i
    
Next ws

End Sub
