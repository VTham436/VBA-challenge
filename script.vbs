Sub stock_data()

  Dim WS As Worksheet
  For Each WS In Worksheets

    
    'creating variables
    Dim ticker As String
    Dim volume As Double
    volume = 0
    Dim ticker_summary As Integer
    ticker_summary = 2
    Dim open_price As Double
    open_price = WS.Cells(2, 3).Value
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'printing column names
    WS.Cells(1, 9).Value = "Ticker Symbol"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Volume"
    
    'for loop that iterates change in cells
    For i = 2 To lastrow
    
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        ticker = WS.Cells(i, 1).Value
        close_price = WS.Cells(i, 6).Value
        yearly_change = (close_price - open_price)
        volume = volume + WS.Cells(i, 7).Value
        WS.Range("I" & ticker_summary).Value = ticker
        WS.Range("L" & ticker_summary).Value = volume
        WS.Range("J" & ticker_summary).Value = yearly_change
        
      
     'percent change
         If (open_price = 0) Then
         percent_change = 0
            Else
            percent_change = (yearly_change / open_price)
        End If
        
     'yearly change
        WS.Range("K" & ticker_summary).Value = percent_change
        WS.Range("K" & ticker_summary).NumberFormat = "0.00%"
        
        volume = 0
        ticker_summary = ticker_summary + 1
    
        open_price = WS.Cells(i + 1, 3).Value
        
        Else
        
        volume = volume + WS.Cells(i, 7).Value
        
        End If
        
    Next i
        
        lastrow_summary_table = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow_summary_table
            If WS.Cells(i, 11).Value > 0 Then
            WS.Cells(i, 10).Interior.ColorIndex = 4
            Else: WS.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
            Next i
      'Greatest values
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        
    For i = 2 To lastrow_summary_table
            If WS.Cells(i, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & lastrow_summary_table)) Then
            WS.Cells(2, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(2, 17).Value = WS.Cells(i, 11).Value
            WS.Cells(2, 17).NumberFormat = "0.00%"
            End If
            
            If WS.Cells(i, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & lastrow_summary_table)) Then
            WS.Cells(3, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(3, 17).Value = WS.Cells(i, 11).Value
            WS.Cells(3, 17).NumberFormat = "0.00%"
            End If
            
            If WS.Cells(i, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & lastrow_summary_table)) Then
            WS.Cells(4, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(4, 17).Value = WS.Cells(i, 12).Value
            
            End If
        Next i
    Next WS
    End Sub