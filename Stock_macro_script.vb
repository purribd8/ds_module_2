Sub stock_calculations_no_reset():

    Dim i As Long ' row number
    Dim cell_vol As Double ' contents of column G (cell volume)
    Dim vol_total As Double ' what is going to go in column L
    Dim ticker As String ' what is going to go in column I

    Dim k As Long ' leaderboard row
    
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim lastRow As Long ' declare once
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        ' xpert gave code
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        vol_total = 0
        k = 2
        
        ' Leaderboard Columns (Write)
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Volume"
        
        
        ' open for first ticker
        ticker_open = Cells(2, 3).Value
    
        For i = 2 To lastRow:
            cell_vol = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
    
            ' LOOP dynamically
            ' check if next row ticker is DIFFERENT
            ' if the same, then we only need to add to the vol_total
            ' if DIFFERENT, then we need add last row, write out to the leaderboard
            ' reset the vol_total to 0
    
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
                ' calculate different ticker
                vol_total = vol_total + cell_vol
                
                ' get the closing price of the ticker
                ticker_close = ws.Cells(i, 6).Value
                price_change = ticker_close - ticker_open
                
                
                ' If ticker open is 0
                If (ticker_open > 0) Then
                    percent_change = price_change / ticker_open
                Else
                    percent_change = 0
                End If
    
                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = price_change
                ws.Cells(k, 11).Value = percent_change
                ws.Cells(k, 12).Value = vol_total
                
                ' conditional formatting
                If (price_change > 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4 ' Green
                    ws.Cells(k, 11).Interior.ColorIndex = 4
                ElseIf (price_change < 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3 ' Red
                    ws.Cells(k, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 2 ' White
                    ws.Cells(k, 11).Interior.ColorIndex = 2
                End If
    
                ' reset
                vol_total = 0
                k = k + 1
                ticker_open = ws.Cells(i + 1, 3).Value ' look ahead to get next ticker
            Else
                ' add to the total
                vol_total = vol_total + cell_vol
            End If
        Next i
        
        ' Style my leaderboard
        ws.Columns("K:K").NumberFormat = "0.0%"
        ws.Columns("I:L").AutoFit
    
    Next ws
    
End Sub

Sub stock_calculations_reset():

    Dim i As Long ' row number
    Dim cell_vol As Double ' contents of column G (cell volume)
    Dim vol_total As Double ' what is going to go in column L
    Dim ticker As String ' what is going to go in column I

    Dim k As Long ' leaderboard row
    
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim lastRow As Long ' declare once
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        ' xpert gave code
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        vol_total = 0
        k = 2
        
        ' Leaderboard Columns (Write)
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Volume"
        
        
        ' open for first ticker
        ticker_open = Cells(2, 3).Value
    
        For i = 2 To lastRow:
            cell_vol = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
    
            ' LOOP dynamically
            ' check if next row ticker is DIFFERENT
            ' if the same, then we only need to add to the vol_total
            ' if DIFFERENT, then we need add last row, write out to the leaderboard
            ' reset the vol_total to 0
    
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
                ' calculate different ticker
                vol_total = vol_total + cell_vol
                
                ' get the closing price of the ticker
                ticker_close = ws.Cells(i, 6).Value
                price_change = ticker_close - ticker_open
                
                
                ' If ticker open is 0
                If (ticker_open > 0) Then
                    percent_change = price_change / ticker_open
                Else
                    percent_change = 0
                End If
    
                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = price_change
                ws.Cells(k, 11).Value = percent_change
                ws.Cells(k, 12).Value = vol_total
                
                ' conditional formatting
                If (price_change > 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4 ' Green
                    ws.Cells(k, 11).Interior.ColorIndex = 4
                ElseIf (price_change < 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3 ' Red
                    ws.Cells(k, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 2 ' White
                    ws.Cells(k, 11).Interior.ColorIndex = 2
                End If
    
                ' reset
                vol_total = 0
                k = k + 1
                ticker_open = ws.Cells(i + 1, 3).Value ' look ahead to get next ticker
            Else
                ' add to the total
                vol_total = vol_total + cell_vol
            End If
        Next i
        
        ' Style my leaderboard
        ws.Columns("K:K").NumberFormat = "0.0%"
        ws.Columns("I:L").AutoFit
    
    Next ws
    
End Sub

