Sub Stocks_Data()

    ' Declare all of your variable types here
    'Dim ws As Worksheet
    Dim Ticker_symbol As String
    Dim yearly_Change As Double
    Dim open_Price As Double
    Dim Close_Price As Double
    Dim Percent_Change As Double
    Dim Vol As Double
    Dim LastRow As Long
    Dim j As Long
    
    ' Loop through all sheets
    For Each ws In Worksheets
    
        Vol = 0
        j = 2
        
        open_Price = ws.Cells(2, 3).Value 'initialization with the first ticker's open price
        ' Determine the last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
      
        'Set variable names for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through all the rows in the worksheet
        For i = 2 To LastRow
        
        
            'i will loop through ticker A if Cell AA then change ticker
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_symbol = ws.Cells(i, 1).Value
                Close_Price = ws.Cells(i, 6).Value
                yearly_Change = (Close_Price - open_Price)
                
                If (open_Price = 0) Then
                   Percent_Change = 0
                Else
                   Percent_Change = (yearly_Change / open_Price) * 100
                End If
                
                Vol = Vol + ws.Cells(i, 7).Value
                open_Price = ws.Cells(i + 1, 3).Value  'After calculating yearly Change, keep the next ticker's open price
                
                ws.Cells(j, 9).Value = Ticker_symbol
                ws.Cells(j, 10).Value = yearly_Change
                ws.Cells(j, 11).Value = Percent_Change
                ws.Cells(j, 12).Value = Vol
                
                If (yearly_Change < 0) Then
                   ws.Cells(j, 10).Interior.ColorIndex = 3
                Else
                   ws.Cells(j, 10).Interior.ColorIndex = 4
                End If
                
                If (Percent_Change < 0) Then
                   ws.Cells(j, 11).Interior.ColorIndex = 3
                Else
                   ws.Cells(j, 11).Interior.ColorIndex = 4
                End If
                   
                
                j = j + 1
                Vol = 0
            End If ' You need to have an End If here
            
            ' Do other stuff here if you like
        
           Vol = Vol + ws.Cells(i, 7).Value

        ' Go to the next row
        Next i
        
        
        ' Do other stuff here if you like

    
    ' Go to next worksheet
    Next ws
    
End Sub


