VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockData()
    
    
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim PreviousTicker As String
        Dim PreviousOpenPrice As Double
        Dim ClosePrice As Double
        Dim LastSheetRow As Double
        Dim SummaryRow As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim Volume As Double
        Dim GreatestPercentIncreaseTicker As String
        Dim GreatestPercentDecreaseTIcker As String
        Dim GreatestVolumeTicker As String
        Dim GreatestPercentIncreaseNumber As Double
        Dim GreatestPercentDecreaseNumber As Double
        Dim GreatestVolumeNumber As Double
                               
        'SummaryRow set the first row to write out the first stock summary
        SummaryRow = 2
        
        Volume = 0
        GreatestPercentIncreaseNumber = 0
        GreatestPercentDecreaseNumber = 0
        GreatestVolumeNumber = 0
        
        WorksheetName = ws.Name
        'MsgBox WorksheetName
        
        'Figure out how many rows there are in total
        
        'FOR TESTING ONLY******************
        'LastSheetRow = 1000
        '**********************************
        
        LastSheetRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastSheetRow)
    
        'Populate the headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Set values for the first time through
        PreviousTicker = ws.Cells(2, 1).Value
        PreviousOpenPrice = ws.Cells(2, 3).Value
        ClosePrice = ws.Cells(2, 6).Value
        
        'Loop through the entore worksheet
        For MainRowNum = 2 To LastSheetRow
            
            'Check if the current stock has changed
            If ws.Cells(MainRowNum, 1).Value = PreviousTicker And MainRowNum < LastSheetRow Then
                'STILL THE SAME STOCK
                StockEndingRow = MainRowNum
                ClosePrice = ws.Cells(MainRowNum, 6).Value
                Volume = Volume + ws.Cells(MainRowNum, 7).Value
            Else
                'DIFFERENT STOCK
                'Write out the summary rows for the previous stock
                ws.Cells(SummaryRow, 9).Value = PreviousTicker
                YearlyChange = ClosePrice - PreviousOpenPrice
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                PercentChange = ((ClosePrice - PreviousOpenPrice) / PreviousOpenPrice)
                ws.Cells(SummaryRow, 11).Value = PercentChange
                If MainRowNum = LastSheetRow Then
                    Volume = Volume + ws.Cells(MainRowNum, 7).Value
                End If
                'Conditional formatting
                If PercentChange > 0 Then
                    'Green if positive
                    ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4
                ElseIf PercentChange < 0 Then
                    'Red if negative
                    ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3
                Else
                    'White if neither
                    ws.Cells(SummaryRow, 11).Interior.ColorIndex = 0
                End If
                'Check and set greatest section
                If PercentChange > GreatestPercentIncreaseNumber Then
                    GreatestPercentIncreaseTicker = PreviousTicker
                    GreatestPercentIncreaseNumber = PercentChange
                End If
                If PercentChange < GreatestPercentDecreaseNumber Then
                    GreatestPercentDecreaseTIcker = PreviousTicker
                    GreatestPercentDecreaseNumber = PercentChange
                End If
                If Volume > GreatestVolumeNumber Then
                    GreatestVolumeTicker = PreviousTicker
                    GreatestVolumeNumber = Volume
                End If
                'Write summary volume
                ws.Cells(SummaryRow, 12).Value = Volume
                'Increment the SummaryRow so it's ready to write out the next summary
                SummaryRow = SummaryRow + 1
                'Reset the Previous values for the new stock and restart the vol indicator
                PreviousTicker = ws.Cells(MainRowNum, 1).Value
                PreviousOpenPrice = ws.Cells(MainRowNum, 3).Value
                Volume = ws.Cells(MainRowNum, 7).Value
                'Put new close price (first close of the new stock)
                ClosePrice = ws.Cells(MainRowNum, 6).Value
            End If
            
        Next MainRowNum
        
        'Finished looping throu the worksheet, so write out greatest section
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = GreatestPercentIncreaseTicker
        ws.Cells(2, 17).Value = GreatestPercentIncreaseNumber
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = GreatestPercentDecreaseTIcker
        ws.Cells(3, 17).Value = GreatestPercentDecreaseNumber
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        ws.Cells(4, 17).Value = GreatestVolumeNumber
        
    Next ws
    
End Sub
