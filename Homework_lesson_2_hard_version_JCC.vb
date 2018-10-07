Sub Vba_homework()

'This script loops through each year of stock data
'grabbing the total amount of volume by stock per year,
'as well as the yearly change for each stock, the percent
'change by year, and a summary of the stock with the greatest
'percent increase by year, the greatest percent decrease by year,
'and the largest total stock volume by year.

'Turn off screen updating until script completes
'to increase speed (not sure if it helps for this script)
Application.ScreenUpdating = False

'Set variable for worksheet
Dim WS As Worksheet

'Iterate inner code through each worksheet
For Each WS In Worksheets
    WS.Activate

'Set initial variable for holding stock ticker
Dim stockticker As String

'Set initial variable for holding total volume per ticker
Dim total_volume As Double
total_volume = 0

'Set variable for location of each stock in the summary column
Dim currentStockRow As Double
currentStockRow = 2

'Set variable for yearly change opening price
Dim yearly_opening_price As Double
yearly_opening_price = 0

'Set variable for yearly change closing price
Dim yearly_closing_price As Double
yearly_closing_price = 0

'Set variable for yearly change
Dim yearly_change As Double
yearly_change = 0

'Set variable for yearly percent change
Dim yearly_percent_change As Double

'Set variable for determining and holding the last row #
Dim lastRow As Double
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Label summary cells in worksheet
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

    'Loop through the sheet for stock name and volume
    For i = 2 To lastRow
        
        'Check if this is the first row with non-zero values for a particular stock
        If (Cells(i, 2).Value Like "*0101*" Or (Cells(i - 1, 1).Value <> Cells(i, 1).Value)) And Cells(i, 3) <> 0 Then

            'Store the value of the opening price
            yearly_opening_price = Cells(i, 3).Value
        
        'Avoid division by 0 errors
        ElseIf Cells(i, 3).Value = 0 Then
        i = i + 1
        
            yearly_opening_price = Cells(i, 3).Value
            
        'Check if the next row has the same name as current ticker
        'if it does...
        ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            
            'Add to the ticker volume for that stock
            total_volume = total_volume + Cells(i, 7).Value
            
        'If it's not the same, we've hit the last row for this stock
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Set the stock name
            stockticker = Cells(i, 1).Value

            'Add to the volume total for that stock
            total_volume = total_volume + Cells(i, 7).Value
            
            'Store the value of the closing price
            yearly_closing_price = Cells(i, 6).Value

            'Calculate the yearly change
            yearly_change = yearly_closing_price - yearly_opening_price
                             
            'Calculate the yearly percent change
            yearly_percent_change = (yearly_change / yearly_opening_price)

            'Copy this card's info in the summary table
            Cells(currentStockRow, 9).Value = stockticker
            Cells(currentStockRow, 10).Value = yearly_change
            Cells(currentStockRow, 11).Value = FormatPercent(yearly_percent_change, 2)
            Cells(currentStockRow, 12).Value = total_volume

            'Add one to the summary table row
            currentStockRow = currentStockRow + 1

            'Reset the stock total
            total_volume = 0

        End If

    Next i


    'Set variables for formatted range
    Dim yearly_range As Range
    Set yearly_range = Range("J2", Range("J2").End(xlDown))

    'Set variables for format conditions
    Dim cond1 As FormatCondition
    Dim cond2 As FormatCondition

    'clear any existing conditional formatting
    yearly_range.FormatConditions.Delete

    'define rule for each conditional format
    Set cond1 = yearly_range.FormatConditions.Add(xlCellValue, xlGreater, 0)
    Set cond2 = yearly_range.FormatConditions.Add(xlCellValue, xlLess, 0)
      
    With cond1
    .Interior.Color = vbGreen
    End With

    With cond2
    .Interior.Color = vbRed
    End With

    'Create variable for Percent column
    Dim percent_range As Range
    Set percent_range = Range("K2", Range("K2").End(xlDown))

    'Create variable for volume column
    Dim volume_range As Range
    Set volume_range = Range("L2", Range("L2").End(xlDown))
    
    'Create variables for smallest and largest percent increase and greatest total volume
    Dim greatest_per_inc As Double
    Dim greatest_per_dec As Double
    Dim greatest_tot_vol As Double
    
    'Calculate greatest percent increase for this sheet
    greatest_per_inc = Application.WorksheetFunction.Max(percent_range)
    
    'Calculate greatest percent decrease for this sheet
    greatest_per_dec = Application.WorksheetFunction.Min(percent_range)
    
    'Calculate greatest total volume for this sheet
    greatest_tot_vol = Application.WorksheetFunction.Max(volume_range)

    'Copy greatest percent increase, decrease, and volume increase to summary area
    Cells(2, 17).Value = FormatPercent(greatest_per_inc, 2)
    Cells(3, 17).Value = FormatPercent(greatest_per_dec, 2)
    Cells(4, 17).Value = greatest_tot_vol
   
    'Create variable for intermediate summary range
    Dim summary_range As Range
    Set summary_range = Range("I2:L2").End(xlDown)
    
    'Apply formatting to Yearly Change column
    Dim per_inc_range As Double
    
        For j = 2 To lastRow
        
            If Cells(j, 11).Value = greatest_per_inc Then
            Cells(2, 16) = Cells(j, 11).Offset(0, -2)
        
            ElseIf Cells(j, 11).Value = greatest_per_dec Then
            Cells(3, 16).Value = Cells(j, 11).Offset(0, -2)
        
            ElseIf Cells(j, 12).Value = greatest_tot_vol Then
            Cells(4, 16).Value = Cells(j, 12).Offset(0, -3)
            
        End If
        
    Next j

Next WS

End Sub
