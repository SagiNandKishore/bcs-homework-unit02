VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub getTotals()
        
    Dim ws As Worksheet                 'Tracks each individual worksheet in the workbook
    Dim summaryline As Double           'Tracks the rows where the Summary details would be populated.
    Dim total As Double                 'Tracks the total volume of each ticker in the month
    Dim Counter As Double               'Tracks the individual rows in each sheet till the last row with data
    Dim endingRow As Double             'Identifies the last row with data in it
    Dim startingPrice As Double         'Identifies the starting price of the ticker
    Dim closingPrice As Double          'Identifies the closing price of the ticker
    Dim percentageChange As Double      'Identifies the percentage change of the ticker price in the month
    Dim yearlyChange As Double          'Identifies the change of the ticker price in the period
    Dim MaxPctChgRow As Double          'Identifies the row containing maximum percentage change of ticker price in the period
    Dim MinPctChgRow As Double          'Identifies the row containing minimum percentage change of the ticker price in the period
    Dim MaxVolumeRow As Double          'Identifies the row containing maximum volume traded in the period
    Dim MaxPctChg As Double             'Identifies the maximum percentage change of ticker price in period
    Dim MinPctChg As Double             'Identifies the minimum percentage change of ticker price in period
    Dim MaxVolume As Double             'Identifies the maximum volume traded in the period
        
    'The scope of the analysis is the current workbook only.
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        'Initialize the summary row and section
        summaryline = 1
        
        'Setup the headers of the summary section
        Cells(summaryline, 9).Value = "Ticker"
        Cells(summaryline, 10).Value = "Yearly Change"
        Cells(summaryline, 11).Value = "Percent Change"
        Cells(summaryline, 12).Value = "Total Stock Volume"
        Cells(summaryline, 16).Value = "Ticker"
        Cells(summaryline, 17).Value = "Value"
        
        'Keep the summary line moving on to the next row as each row gets filled up
        summaryline = summaryline + 1
        
        'Setup the value change trackers
        MaxPctChg = 0
        MinPctChg = 0
        MaxVolume = 0
        total = 0
        
        'Identify the last row with the data
        endingRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Identify the starting price of the first ticker in the sheet.
        'For other tickers, this value can be captured at the boundary line
        startingPrice = Cells(2, 3).Value
        
        'Loop through all rows in the sheet containing the ticker data
        For Counter = 2 To endingRow
        
            'Keep running sum of ticker volume
            total = total + Cells(Counter, 7)
            
            'When you are at the boundary row between two tickers data, capture
            '       (1.) Closing price
            '       (2.) Starting price of next ticker
            '       (3.) All the details needed for summary pane
            If Cells(Counter, 1).Value <> Cells(Counter + 1, 1).Value Then
                 closingPrice = Cells(Counter, 6).Value
                 yearlyChange = closingPrice - startingPrice
                 
                 'The data for ticker PLNT was missing and the percentage change has been tagged as 0%.
                 If startingPrice = 0 Then
                     percentageChange = 0
                 Else
                     percentageChange = (yearlyChange / startingPrice) * 100
                 End If
                 
                 'Keep track of the maximum percentage change and the row from the summary pane this occurs at
                 'Once we know the row the rest of the values can be auto determined
                 If percentageChange >= MaxPctChg Then
                    MaxPctChgRow = summaryline
                    MaxPctChg = percentageChange
                 End If
                 
                 'Keep track of the lowest percentage change and the row from the summary pane this occurs at
                 'Once we know the row the rest of the values can be auto determined
                 If percentageChange <= MinPctChg Then
                    MinPctChgRow = summaryline
                    MinPctChg = percentageChange
                 End If
                 
                 'Keep track of the greatest ticker volume traded and the row from the summary pane this occurs at
                 'Once we know the row the rest of the values can be auto determined
                 If total >= MaxVolume Then
                    MaxVolumeRow = summaryline
                    MaxVolume = total
                 End If
                 
                 'Populate the summary table details
                 Cells(summaryline, 9).Value = Cells(Counter, 1).Value
                 Cells(summaryline, 10).Value = yearlyChange
                 Cells(summaryline, 11).Value = Str(percentageChange) + " %"
                 Cells(summaryline, 12).Value = total
                 
                 'Add colors based on the direction of price change
                 If yearlyChange > 0 Then
                    Cells(summaryline, 10).Interior.ColorIndex = 4
                 ElseIf yearlyChange < 0 Then
                    Cells(summaryline, 10).Interior.ColorIndex = 3
                 End If
                 
                 'Reset the counters and start processing the next ticker
                 total = 0
                 summaryline = summaryline + 1
                 startingPrice = Cells(Counter + 1, 3).Value
            End If
        Next Counter
        
        'After all rows are processed add the summary section for greatst and lowest changes across the sheet
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P2").Value = Cells(MaxPctChgRow, 9).Value
        Range("P3").Value = Cells(MinPctChgRow, 9).Value
        Range("P4").Value = Cells(MaxVolumeRow, 9).Value
        
        Range("Q2").Value = Str(MaxPctChg) + " %"
        Range("Q3").Value = Str(MinPctChg) + " %"
        Range("Q4").Value = MaxVolume
                
        'Autofit the columns based on data and header
        Columns("I:Q").EntireColumn.AutoFit
    Next

End Sub
Sub clearUpdates()
    'This macro resets the summary tables by deleting the rows
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Columns("I:Q").EntireColumn.Delete
    Next
End Sub
