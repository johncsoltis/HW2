Attribute VB_Name = "Module11"
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
Sub StocksHard()

'Define variables------
Dim totalvolume As Double
Dim resultrow As Long
Dim i As Long
Dim lastrow As Long
Dim ws As Worksheet
Dim openvalue As Double
Dim closevalue As Double
Dim change As Double
Dim max As Double
Dim min As Double
Dim minname As String
Dim maxname As String
Dim volname As String
Dim lastresultrow As Long




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

    'Set values for the results chart
    max = 0
    min = 0
    totalvolume = 0

    'Find last row in results
    lastresultsrow = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Loop through greater/less than and volume summary to find max and min changes and max total volume
    For i = 2 To lastresultsrow

        'Find max and save to variable
        If Cells(i, 11).Value > max Then
        max = Cells(i, 11).Value
        maxname = Cells(i, 9).Value
    
        End If
    
        'Find min and save to variable
        If Cells(i, 11).Value < min Then
        min = Cells(i, 11).Value
        minname = Cells(i, 9).Value
    
        End If
        
        'Find max volume and save to variable
        If Cells(i, 12).Value > totalvolume Then
        totalvolume = Cells(i, 12).Value
        volname = Cells(i, 9).Value
    
        End If
    
    Next i


    'Print summary to chart area
    Cells(2, 15).Value = maxname
    Cells(2, 16).Value = max
    Cells(3, 15).Value = minname
    Cells(3, 16).Value = min
    Cells(4, 15).Value = volname
    Cells(4, 16).Value = totalvolume
    
    'Format max and min as percent
    Cells(2, 16).Value = FormatPercent(Cells(2, 16).Value)
    Cells(3, 16).Value = FormatPercent(Cells(3, 16).Value)
    
    'Add chart labels
    Cells(2, 14).Value = "Biggest % Gain"
    Cells(3, 14).Value = "Biggest % Loss"
    Cells(4, 14).Value = "Most Annual Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    'Conditional formatting for yearly change

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



