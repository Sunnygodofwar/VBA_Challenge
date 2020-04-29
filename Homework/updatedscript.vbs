Sub WallStreetTicker():

    'Loop through all worksheets for output
    
    For Each WS In Worksheets

        ' Name Headers
        
        Cells(1, 9).Value = "Ticker"
        
        Cells(1, 10).Value = "Yearly Change"
        
        Cells(1, 11).Value = "Percent Change"
        
        Cells(1, 12).Value = "Total_Stock_Volume"
        
        Cells(1, 13).Value = "Ticker1"
        
        Cells(1, 14).Value = "Value"

        
        
        ' Declare Variables
        
        
        Dim beginning_open_price As Double
        
        Dim ending_close_price As Double
        
        Dim yearly_change As Double
        
        Dim Ticker_name As String
        
        Dim Percent_change As Double
        
        Dim Volume As Double
        
        
        'Define volume as a value
        
        Volume = 0
        
        
        Dim Row As Double
        
        Row = 2
        
        
        Dim Column As Integer
        
        Column = 1
        
        
        Dim i As Long
        
        
        'Declare an initial opening price
        
        beginning_open_price = Cells(2, Column + 2).Value
        
        For i = 2 To LastRow
        
        
         'Check for changes in Ticker
         
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                Ticker = Cells(i, Column).Value
                
                Ticker = Cells(Row, Column + 8).Value
                
                ending_close_price = Cells(i, Column + 5).Value
                
                yearly_change = ending_close_price - beginning_open_price
                
                Cells(Row, Column + 9).Value = yearly_change
                
                
                'Check for change in percentage %
                
                If (beginning_open_price = 0 And ending_close_price = 0) Then
                
                    Percent_change = 0
                    
                ElseIf (beginning_open_price = 0 And ending_close_price <> 0) Then
                
                    Percent_change = 1
                    
                Else
                
                    Percent_change = yearly_change / beginning_open_price
                    
                    Cells(Row, Column + 10).Value = Percent_change
                    
                    Cells(Row, Column + 10).NumberFormat = "%"
                    
                End If
                
        
                Volume = Volume + Cells(i, Column + 6).Value
                
                Cells(Row, Column + 11).Value = Volume
                
                Row = Row + 1
                
                beginning_open_price = Cells(i + 1, Column + 2)
                
                Volume = 0
            Else
            
                Volume = Volume + Cells(i, Column + 6).Value
                
            End If
            
        Next i
        
    
        
        'Set color changes for positive = green; negative = red
        
        For j = 2 To LastRow
        
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
            
                
                Cells(j, Column + 9).Interior.ColorIndex = 4
                
                
            ElseIf Cells(j, Column + 9).Value < 0 Then
            
            
                Cells(j, Column + 9).Interior.ColorIndex = 3
                
                
            End If
        
        
            Next j
            
            
            Next WS
            
            
        End Sub
            
        
        
                
