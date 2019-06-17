Attribute VB_Name = "Module2"
'Create a script that will loop through all the stocks for one year for each run and take the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
'Conditional formatting
    'Highlight positive change in green and negative change in red

Sub Stocks2()

'Define variables------
Dim totalvolume As Double
Dim resultrow As Long
Dim i As Long
Dim lastrow As Long
Dim ws As Worksheet
Dim openvalue As Double
Dim closevalue As Double
Dim change As Double


'Loop through worksheets-------
'For Each ws In Worksheets
    'Activate workseet
    'ws.Activate

    'Find last row on sheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'Define output row
    resultrow = 2

        'Loop through all tickers to find total volume
        For i = 2 To lastrow
            
            'If next ticker is the same, add volume to total
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            totalvolume = totalvolume + Cells(i, 7).Value
    
            'If ticker different
            Else
            'Print ticker and total volume
            Cells(resultrow, 9).Value = Cells(i, 1).Value
            Cells(resultrow, 12).Value = totalvolume
            
            'Find opening and closing values
            closevalue = Cells(i, 6).Value
            openvalue = Cells(i - 261, 3).Value
            
            'Calculate change from open to close
            change = closevalue - openvalue
            
            'Print change and percent change
            Cells(resultrow, 10).Value = change
            Cells(resultrow, 11).Value = change / openvalue
            
            'Reset variables for next ticker
            resultrow = resultrow + 1
            totalvolume = 0
            openvalue = 0
            closevalue = 0
    
            End If
    
        
    Next i
        
    
'Next ws

End Sub


