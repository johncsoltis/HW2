Attribute VB_Name = "Module21"
'Create a script that will loop through all the stocks for one year for each run and take the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
'Conditional formatting
    'Highlight positive change in green and negative change in red

Sub StocksModerate()

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
For Each ws In Worksheets
    'Activate workseet
    ws.Activate
    
     'Create labels
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Find last row on sheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'Define output row
    resultrow = 2
    
    'Set initial openening value
    openvalue = Cells(2, 3).Value

        'Loop through all tickers to find total volume
        For i = 2 To lastrow
        
            'Skip cells that have no data
            If openvalue = O Then
                openvalue = Cells(i + 1, 3)
                Else
            
                    totalvolume = totalvolume + Cells(i, 7).Value
                    
                    'If next ticker different, print volume and ticker to results chart
                    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                    Cells(resultrow, 9).Value = Cells(i, 1).Value
                    Cells(resultrow, 12).Value = totalvolume
            
                    'Find closing value
                    closevalue = Cells(i, 6).Value
                
            
                    'Calculate change from open to close
                    change = closevalue - openvalue
            
                    'Print change and percent change
                    Cells(resultrow, 10).Value = change
                    Cells(resultrow, 11).Value = (change / openvalue)
                    
                    'Format cells as percent and dollar
                    Cells(resultrow, 11).Value = FormatPercent(Cells(resultrow, 11).Value)
                    Cells(resultrow, 10).Value = FormatCurrency(Cells(resultrow, 10).Value)
                    Cells(resultrow, 10).Font.Color = 1
            
                    'Reset variables for next ticker
                    resultrow = resultrow + 1
                    totalvolume = 0
                    openvalue = 0
                    closevalue = 0
                
                    'Set next opening value
                    openvalue = Cells(i + 1, 3).Value
    
                    End If
                    
                End If
        
    Next i
    
    
    'Conditional formatting for yearly change
    'Find last row of results table
    lastresultsrow = Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastresultsrow
        'if change is positive, color green
        If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        
        'if change is negative color red
        Else: Cells(i, 10).Interior.ColorIndex = 3
            
        End If
    
    Next i
        
    
Next ws

'Autofit all columns in workbook
For Each ws In ThisWorkbook.Worksheets
    ws.Cells.EntireColumn.AutoFit
    
Next ws

End Sub


