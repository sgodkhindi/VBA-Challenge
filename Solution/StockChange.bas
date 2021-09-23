Attribute VB_Name = "Module1"
Sub StockChange()

'Define Variables
'Variables for Part 1 - Main Assignment
Dim TotalChange As Double
Dim PercentChange As Double
Dim OpenAmount As Double
Dim DispRow As Integer
Dim TotalVolume As Double
Dim LastRow As Double
Dim WSCount As Integer

'Variables for Part 2 - Bonus Assignment
Dim MaxPercent As Double
Dim MinPercent As Double
Dim MaxVolume As Double
Dim MaxTicker As String
Dim MinTicker As String
Dim VolTicker As String

'Part 1 Main Assignment - Finding the Yearly Change, Percent Change & Total Volume of Each Stock

'Find the number of Worksheets in the Excel File
WSCount = ActiveWorkbook.Worksheets.Count

'Outer Loop which goes through each Worksheet of the Workbook
For j = 1 To WSCount
        'Intializing the variables
        TotalChange = 0
        PercentChange = 0
        TotalVolume = 0
        DispRow = 2
        LastRow = 0
        OpenAmount = 0
        
               
        'Assign Headers to the Display Area
        ActiveWorkbook.Worksheets(j).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(j).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(j).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(j).Cells(1, 12).Value = "Total Volume"
        
        'Find The Last Non Blank Row in the Main Data Area Sheet
        LastRow = ActiveWorkbook.Worksheets(j).Cells(Rows.Count, 1).End(xlUp).Row
        
        'Assign the 1st Stock's Open amount
        OpenAmount = ActiveWorkbook.Worksheets(j).Cells(2, 3).Value
                
        'Inner Loop - Loop goes through rows in the sheet from the 1st to the last row
        For i = 2 To LastRow + 1
            TotalVolume = TotalVolume + ActiveWorkbook.Worksheets(j).Cells(i, 7).Value
            
            ' Searches for when the value of the next cell is different than that of the current cell
            If ActiveWorkbook.Worksheets(j).Cells(i + 1, 1).Value <> ActiveWorkbook.Worksheets(j).Cells(i, 1).Value Then
                 TotalChange = ActiveWorkbook.Worksheets(j).Cells(i, 6).Value - OpenAmount
                 
                 'Check of the Opening Amount = 0 to avoid dividing by Zero
                 If OpenAmount <> 0 Then
                    PercentChange = TotalChange / OpenAmount
                 Else
                    PercentChange = 0
                 End If
                 
                 'Assign the computed values to the Display Row
                 ActiveWorkbook.Worksheets(j).Cells(DispRow, 9).Value = ActiveWorkbook.Worksheets(j).Cells(i, 1).Value
                 ActiveWorkbook.Worksheets(j).Cells(DispRow, 10).Value = TotalChange
                 ActiveWorkbook.Worksheets(j).Cells(DispRow, 11).Value = PercentChange
                 ActiveWorkbook.Worksheets(j).Cells(DispRow, 12).Value = TotalVolume
                 
                 'Assign the colors to the Cells depending if they are +ve or -ve
                 If TotalChange < 0 Then
                    ActiveWorkbook.Worksheets(j).Cells(DispRow, 10).Interior.ColorIndex = 3
                    Else
                    ActiveWorkbook.Worksheets(j).Cells(DispRow, 10).Interior.ColorIndex = 4
                 End If
                 
                 
                'Resetting Totals to 0
                 TotalVolume = 0
                 TotalChange = 0
                 PercentChange = 0
                 
                'Advancing to the next display row
                 DispRow = DispRow + 1
                
                'Initializing the Open Amount for the Next Security
                 OpenAmount = ActiveWorkbook.Worksheets(j).Cells(i + 1, 3).Value
            End If
            
         'Move to the next row
         Next i

'Move to the next sheet
Next j

'------------------------------------------------------------------------------------------------------------

'Part 2 Bonus Assignment - Finding the Maximum & Minimum Percent Change & Maximum Total Volume Across All the stocks

'Outer Loop which goes through each Worksheet of the Workbook

For j = 1 To WSCount

'Since WSCount has already been assigned the value in Part 1 Main Assignment

'Initializing the variables
LastRow = 0
MaxPercent = 0
MinPercent = 0
MaxVolume = 0

'Assigning Headers and Labels to the display area of the Worksheet where these will be displayed
 ActiveWorkbook.Worksheets(j).Cells(1, 15).Value = "Ticker"
 ActiveWorkbook.Worksheets(j).Cells(1, 16).Value = "Value"
 ActiveWorkbook.Worksheets(j).Cells(2, 14).Value = "Greatest % Increase"
 ActiveWorkbook.Worksheets(j).Cells(3, 14).Value = "Greatest % Decrease"
 ActiveWorkbook.Worksheets(j).Cells(4, 14).Value = "Greatest Total Volume"

'Find The Last Non Blank Row in the Display Area of the Sheet Prepared in the Part 1 of the Program
LastRow = ActiveWorkbook.Worksheets(j).Cells(Rows.Count, 9).End(xlUp).Row

'Initializing the Variables with the values from the 2nd Row of the Display Area - Percentage Values and Volume
MaxPercent = ActiveWorkbook.Worksheets(j).Cells(2, 11).Value
MinPercent = ActiveWorkbook.Worksheets(j).Cells(2, 11).Value
MaxVolume = ActiveWorkbook.Worksheets(j).Cells(2, 12).Value

'Initializing the Variables with the values from the 2nd Row of the Display Area - Ticker Values
MaxTicker = ActiveWorkbook.Worksheets(j).Cells(2, 9).Value
MinTicker = ActiveWorkbook.Worksheets(j).Cells(2, 9).Value
VolTicker = ActiveWorkbook.Worksheets(j).Cells(2, 9).Value

    'Inner Loop - Loop goes through rows in the sheet from the 1st to the last row
    For i = 2 To LastRow
            
            'Check of the Percentage Change of the Current Cell is HIGHER than the Previous Cell's Value
            'If so assign it to the MaxPercent and the corresponding Ticker values to the MaxTicker
            If ActiveWorkbook.Worksheets(j).Cells(i, 11).Value > MaxPercent Then
                MaxPercent = ActiveWorkbook.Worksheets(j).Cells(i, 11).Value
                MaxTicker = ActiveWorkbook.Worksheets(j).Cells(i, 9).Value
            End If
            
            'Check of the Percentage Change of the Current Cell is LOWER than the Previous Cell's Value
            'If so assign it to the MinPercent and the corresponding Ticker values to the MinTicker
            If ActiveWorkbook.Worksheets(j).Cells(i, 11).Value < MinPercent Then
                MinPercent = ActiveWorkbook.Worksheets(j).Cells(i, 11).Value
                MinTicker = ActiveWorkbook.Worksheets(j).Cells(i, 9).Value
            End If
            
            'Check of the Volume Change of the Current Cell is HIGHER than the Previous Cell's Value
            'If so assign it to the MinVolume and the corresponding Ticker values to the VolTicker
            If ActiveWorkbook.Worksheets(j).Cells(i, 12).Value > MaxVolume Then
                MaxVolume = ActiveWorkbook.Worksheets(j).Cells(i, 12).Value
                VolTicker = ActiveWorkbook.Worksheets(j).Cells(i, 9).Value
            End If
     
     'Move to the next row
     Next i

'Assign the Ticker Values to the Respective Display Areas
ActiveWorkbook.Worksheets(j).Cells(2, 15).Value = MaxTicker
ActiveWorkbook.Worksheets(j).Cells(3, 15).Value = MinTicker
ActiveWorkbook.Worksheets(j).Cells(4, 15).Value = VolTicker

'Assign the Maximum and Minimum Percent Values and the Maximumn Volume numbers to the Respective Display Areas
ActiveWorkbook.Worksheets(j).Cells(2, 16).Value = MaxPercent
ActiveWorkbook.Worksheets(j).Cells(3, 16).Value = MinPercent
ActiveWorkbook.Worksheets(j).Cells(4, 16).Value = MaxVolume

'Move to the next sheet
Next j

End Sub
