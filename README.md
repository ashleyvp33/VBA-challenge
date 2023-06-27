# VBA-challenge
Sub VBA_challenge()
 Dim ticker_name As String
    Dim summary_row As Integer
    Dim yearly_change As Double
            yearly_change = 0
    Dim pct_change As Double
    Dim total_volume As Double
    Dim opening_price As Double
    Dim closing_price As Double
    Dim greatest_total_vol As Integer
    Dim greatest_total_vol_ticker As String
    Dim greatest_percent_dec As Long
    Dim greatest_percent_dec_ticker As String
    Dim greatest_percent_inc As Long
    Dim greatest_percent_inc_ticker As String
    
    
    For Each ws In Worksheets
        total_volume = 0
        summary_row = 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        opening_price = ws.Cells(2, 3).Value
        
        greatest_total_vol = 0
        greatest_total_vol_ticker = 999999999
        
        greatest_percent_dec = 999999999
        greatest_percent_dec_ticker = 999999999
  
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Value"
    
    For i = 2 To LastRow
        ticker_name = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            closing_price = ws.Cells(i, 6).Value
            yearly_change = closing_price - opening_price
                If opening_price > 0 Then
                    
                    pct_change = (yearly_change / opening_price) * 100
                Else
                pct_change = 0
                End If
                
            opening_price = ws.Cells(i + 1, 3).Value
            
            ws.Cells(summary_row, 9) = ticker_name
            ws.Cells(summary_row, 10) = yearly_change
            ws.Cells(summary_row, 11) = pct_change
            ws.Cells(summary_row, 12) = total_volume
            
            If total_volume > greatest_total_vol Then
                greatest_percent_vol = total_vol
                greatest_total_vol_ticker = ws.Cells(i, 1).Value
            End If
            
            If pct_change > greatest_percent_inc Then
                greatest_percent_inc = pct_change
                greatest_percent_inc_ticker = ws.Cells(i, 1).Value
            End If
            
            
            If pct_change < greatest_percent_dec Then
                greatest_percent_dec = pct_change
                greatest_percent_dec_ticker = ws.Cells(i, 1).Value
            End If
            
            summary_row = summary_row + 1
            total_volume = 0
            
        End If
    Next i
Cells(summary_row, 16).Value = greatest_total_vol

Next ws

End Sub
