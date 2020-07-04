Sub VBA_challenge()

    'set initial variable for ticker name
    Dim TickerName As String
    'set initial variable for total stock volume
    Dim StockVolume As Double
    StockVolume = 0
    'Keep track of the location for each ticker in the summary table
    Dim TableRow As Integer
    TableRow = 2
    'set first and last variables for the yearly change calculation
    Dim OpenValue As Double
    Dim CloseValue As Double
    'set yearly chance variable
    Dim YearlyChange As Double
    'set percentage change variable
    Dim PercentChange As Double
    
    'loop through all ticker values
    For i = 2 To 73000

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'set ticker name
            TickerName = Cells(i, 1).Value
            
            'set stock volue
            StockVolume = StockVolume + Cells(i, 7).Value
            
            'print ticker name in table
            Range("I" & TableRow).Value = TickerName
            
            'print stock volume
            Range("N" & TableRow).Value = StockVolume
            
            'register close value
            CloseValue = Cells(i, 6).Value
            
            'TEST print close value TEST
            Range("K" & TableRow).Value = CloseValue
            
            'set value for yearly change
            YearlyChange = (CloseValue - OpenValue)
            
            'print yearly change
            Range("L" & TableRow).Value = YearlyChange
                    
            'set value for percent change
            PercentChange = (CloseValue / OpenValue) - 1
            
            'print percent change
            Range("M" & TableRow).Value = PercentChange
            
            'number formate percent
            Range("M" & TableRow).NumberFormat = "0.00%"
            
            'add TableRow
            TableRow = TableRow + 1
            
            'set stock volume to 0
            StockVolume = 0
            
            'if the previous cell is the same then...
        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            'add to Stock Volume
            StockVolume = StockVolume + Cells(i, 7).Value
        
            'register open value
            OpenValue = Cells(i, 3).Value
            
            'TEST print open value TEST
            Range("J" & TableRow).Value = OpenValue
            
            
        End If
        
            'conditional formating yearly change
        If Cells(i, 12).Value < 0 Then
            Cells(i, 12).Interior.ColorIndex = 3
            Cells(i, 13).Interior.ColorIndex = 3
        ElseIf Cells(i, 12).Value > 0 Then
            Cells(i, 12).Interior.ColorIndex = 4
            Cells(i, 13).Interior.ColorIndex = 4
        End If
        
    Next i
       

End Sub


Sub fast_reset()

    Range("I2:N500").Value = ""
    Range("I2:N500").Interior.ColorIndex = 0

End Sub