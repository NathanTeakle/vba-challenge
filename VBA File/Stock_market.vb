Sub Stock_market()
    'Declare and set worksheet
    Dim ws As Worksheet
    Dim main_spreadsheet As Boolean
    
    main_spreadsheet = True
    
    'Loop through all stocks for one year
    For Each ws In Worksheets
    
        'Create the column headings
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Set initial variables for ticker symbol
        Dim Ticker As String
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
        
        'Set initial variables for yearly change
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
        'Set initial variables for opening and closing prices
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        
        'Set initial variables for row count and last row
        Dim Row_Count As Long
        Row_Count = 2
        Dim Last_Row As Long
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set variables for bonus solution
        Dim max_ticker As String
        max_ticker = " "
        Dim min_ticker As String
        min_ticker = " "
        Dim max_percent As Double
        max_percent = 0
        Dim min_percent As Double
        min_percent = 0
        Dim max_volume_ticker As String
        max_volume_ticker = " "
        Dim max_volume As Double
        max_volume = 0
        
        'Loop through all rows in worksheet
        For i = 2 To Last_Row
        
            'Check if we are still within the same ticker symbol, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set ticker symbol
                Ticker = ws.Cells(i, 1).Value
                
                'Add to ticker volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                'Get closing price
                Closing_Price = ws.Cells(i, 6).Value
                
                'Calculate yearly change and percent change
                Yearly_Change = Closing_Price - Opening_Price
                
                If Opening_Price <> 0 Then
                
                    Percent_Change = (Closing_Price - Opening_Price) / Opening_Price
                    
                Else
                
                    Percent_Change = 0
                    
                End If
                
                'Output ticker symbol, yearly change, percent change and total stock volume to worksheet
                ws.Range("I" & Row_Count).Value = Ticker
                ws.Range("J" & Row_Count).Value = Yearly_Change
                ws.Range("K" & Row_Count).Value = Percent_Change
                ws.Range("L" & Row_Count).Value = Ticker_Volume
                
                'Add one to row count
                Row_Count = Row_Count + 1
                
                'Reset ticker volume and opening price variables for new ticker symbol
                Ticker_Volume = 0
                Opening_Price = 0
                
            Else
            
                'Add to ticker volume and get opening price if it's the first row for a new ticker symbol
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                If Opening_Price = 0 Then
                
                    Opening_Price = ws.Cells(i, 3).Value
                    
                End If
                
            End If
            
            'Bonus Solution: Everything below is related to the Bonus Solution
                If (Percent_Change > max_percent) Then
                    max_percent = Percent_Change
                    max_ticker = Ticker
                
                ElseIf (Percent_Change < min_percent) Then
                    min_percent = Percent_Change
                    min_ticker = Ticker
                End If
                
                If (Ticker_Volume > max_volume) Then
                    max_volume = Ticker_Volume
                    max_volume_ticker = Ticker
                End If
                
        Next i
        
            'Check if we're on the first spreadsheet
            If main_spreadsheet Then
            
                'Counts all new values to a summary on the right of the current spreadsheet
                ws.Range("Q2").Value = (CStr(max_percent))
                ws.Range("Q3").Value = (CStr(min_percent))
                ws.Range("P2").Value = max_ticker
                ws.Range("P3").Value = min_ticker
                ws.Range("Q4").Value = max_volume
                ws.Range("P4").Value = max_volume_ticker
            
            Else
                main_spreadsheet = False
            End If

    Next ws
    
End Sub