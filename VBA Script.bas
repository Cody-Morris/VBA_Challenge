Attribute VB_Name = "Module1"
Sub studywork()

    'we have to iterate through each row
    'want to generate a unique ID
    '   yearly change
    '       first open price - last close price
    '   yearly percent change
    
    'Name the Cells
    Cells(1, 9).Value = "Ticker"
    
    Cells(1, 10).Value = "Yearly Change"
    
    Cells(1, 11).Value = "Percent Change"
    
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim LastRow As Double
    Dim ticker As String
    Dim UniqueCounter As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim OpenPrice As Double
    Dim PercentChange As Double
    Dim StockVolumne As Double
    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ""
    OpenPrice = Range("C2")
    ClosePrice = 0
    UniqueCounter = 1
    PercentChange = 0
    
    For i = 2 To LastRow
    
        ticker = Range("A" & i)
        StockVolumne = StockVolumne + Range("G" & i)
        
        
        If Range("A" & i) <> Range("A" & i + 1) Then
            
            UniqueCounter = UniqueCounter + 1
            Range("I" & UniqueCounter).Value = Range("A" & i).Value
            
            'name, open date, volumne
            'i + 1 is going to close price
            ' yearly change = i + 1 close price - I open price
            
            ClosePrice = Range("F" & i).Value
            'openpriuce is defined
            YearlyChange = ClosePrice - OpenPrice
            Range("J" & UniqueCounter) = YearlyChange
            
        If OpenPrice <> 0 Then
            
            ' percent chang = yearlychange / openprice
            PercentChange = YearlyChange / OpenPrice
        'Else: OpenPrice = 0
        End If
        
            Range("K" & UniqueCounter) = PercentChange
            
            ' printer volumne here
            Range("L" & UniqueCounter) = StockVolumne
                       
            'Update OpenPrice for next Unique ID
            OpenPrice = Range("F" & i + 1).Value
            
        End If
        
    Next i
    

    
    
End Sub

