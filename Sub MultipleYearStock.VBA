Sub MultipleYearStock()

    'Declare variables
    Dim WorksheetName As String
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim OutputRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim Opening As Double
    Dim Closing As Double
    Dim Increase As String
    Dim Decrease As String
    Dim Volume As String
    
    
    'Loop through each worksheet
      For Each ws In Worksheets
        WorksheetName = ws.Name
        OutputRow = 2
        TotalVolume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        Opening = 0
        Closing = 0
        Increase = ""
        Decrease = ""
        Volume = ""
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Columns("K:L").AutoFit
        
    'Loop through rows
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
            If i > 2 Then
            'Write data to the column
            ws.Cells(OutputRow, 9).Value = Ticker
            ws.Cells(OutputRow, 10).Value = YearlyChange
            ws.Cells(OutputRow, 11).Value = PercentChange
            ws.Cells(OutputRow, 12).Value = TotalVolume
            OutputRow = OutputRow + 1
            
         End If
         
        'Reset variables for the new Ticker
        Ticker = ws.Cells(i, 1).Value
        YearlyChange = 0
        TotalVolume = 0
        Opening = ws.Cells(i, 3).Value
        
        End If
        
        'Calculate YearlyChange and update TotalVolume
        Closing = ws.Cells(i, 6).Value
        YearlyChange = Closing - Opening
        
        'Color Cells based on YearlyChange
        If YearlyChange < 0 Then
                ws.Cells(OutputRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(OutputRow, 10).Interior.ColorIndex = 4
            End If
            
        'Calculate PercentChange
        If ws.Cells(i, 3).Value <> 0 Then
            PercentChange = (Closing - Opening) / Opening
        Else
            PercentChange = 0
            
        End If
        
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
            'Write data to output columns
            ws.Cells(OutputRow, 9).Value = Ticker
            ws.Cells(OutputRow, 10).Value = YearlyChange
            ws.Cells(OutputRow, 11).Value = PercentChange
            ws.Cells(OutputRow, 11).NumberFormat = "0.00%"
            ws.Cells(OutputRow, 12).Value = TotalVolume
            
       'Track greatest increase, decrease, and volume
       If PercentChange > GreatestIncrease Then
            GreatestIncrease = PercentChange
            Increase = Ticker
            
            End If
       
       If PercentChange < GreatestDecrease Then
            GreatestDecrease = PercentChange
            Decrease = Ticker
       
       End If
       
       
       If TotalVolume > GreatestVolume Then
            GreatestVolume = TotalVolume
            Volume = Ticker
       
       End If
       
       
   Next i
   
            'Write the greatest increase, decrease and volume to the worksheet
            ws.Cells(2, 17).Value = GreatestIncrease
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(2, 16).Value = Increase
            ws.Cells(3, 17).Value = GreatestDecrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(3, 16).Value = Decrease
            ws.Cells(4, 17).Value = GreatestVolume
            ws.Cells(4, 16).Value = Volume
            
   Next ws
   
    
End Sub
