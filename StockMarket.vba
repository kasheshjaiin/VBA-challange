Sub Stock_Market()
    ' Declare variables
    Dim lastrow As Double
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalStockVolume As Double
    Dim TableRow As Integer

    Dim MaxPercentIncrease As Double
    Dim MinPercentDecrease As Double
    Dim MaxTotalVolume As Double
    Dim MaxPercentIncreaseTicker As String
    Dim MinPercentDecreaseTicker As String
    Dim MaxTotalVolumeTicker As String
    
        
   
   ' Loop through all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        
        ' Assigning column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Finding the last row
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables for the data
        TableRow = 2
        Ticker = ws.Cells(2, 1).Value
        OpenPrice = ws.Cells(2, 3).Value
        TotalStockVolume = 0
        MaxPercentIncrease = 0
        MinPercentDecrease = 0
        MaxTotalVolume = 0
        
        ' Loop through the data
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set Closing Price and calculate Yearly Change and Percentage Change
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                PercentageChange = (YearlyChange / OpenPrice)
              
                
         ' Putting Values for Result
                ws.Cells(TableRow, 9).Value = Ticker
                ws.Cells(TableRow, 10).Value = YearlyChange
                ws.Cells(TableRow, 11).Value = PercentageChange
                ws.Cells(TableRow, 11).NumberFormat = "0.00%"
                ws.Cells(TableRow, 12).Value = TotalStockVolume
                
                 ' Check for the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume"
                If PercentageChange > MaxPercentIncrease Then
                    MaxPercentIncrease = PercentageChange
                    MaxPercentIncreaseTicker = Ticker
                

                ElseIf PercentageChange < MinPercentDecrease Then
                    MinPercentDecrease = PercentageChange
                    MinPercentDecreaseTicker = Ticker
             

                ElseIf TotalStockVolume > MaxTotalVolume Then
                    MaxTotalVolume = TotalStockVolume
                    MaxTotalVolumeTicker = Ticker
                End If
                
                 ' Conditional formatting for cell color
                If YearlyChange > 0 Then
                    ws.Cells(TableRow, 10).Interior.ColorIndex = 4 ' Green
                ElseIf YearlyChange < 0 Then
                    ws.Cells(TableRow, 10).Interior.ColorIndex = 3 ' Red
                End If
                
                
                TableRow = TableRow + 1
                
                ' Reset variables for the new ticker
                Ticker = ws.Cells(i + 1, 1).Value
                OpenPrice = ws.Cells(i + 1, 3).Value
                TotalStockVolume = 0
                
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Output the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" information
        
        
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ws.Range("P2").Value = MaxPercentIncreaseTicker
    ws.Range("P3").Value = MinPercentDecreaseTicker
    ws.Range("P4").Value = MaxTotalVolumeTicker

    ws.Range("Q2").Value = MaxPercentIncrease
    ws.Range("Q3").Value = MinPercentDecrease
    ws.Range("Q4").Value = MaxTotalVolume
    
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = MaxTotalVolume
    
   Next ws
    
End Sub

