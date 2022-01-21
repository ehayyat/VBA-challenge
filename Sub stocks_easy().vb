Sub stocks_easy()
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        'Define variables and set to 0
        Dim ticker As String
        Dim total_vol As Double
        total_vol = 0
        Dim summary_table As Integer
        summary_table = 2
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim price_change As Double
        price_change = 0
        Dim change_percentage As Double
        change_percentage = 0
        Dim Lastrow as Long
        Dim i as Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Create Table Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        '--------------------
        open_price = Cells(2, 3)
        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
                price_change = close_price - open_price
                If open_price <> 0
                    change_percentage = (price_change / open_price) * 100
                End If    
                total_vol = total_vol + ws.Cells(i, 7).Value
                ws.Range("I" & summary_table).Value = ticker
                ws.Range("J" & summary_table).Value = price_change
                'Colors
                If (price_change > 0) Then
                    ws.Range("J" & summary_table).Interior.ColorIndex = 4
                ElseIf (price_change <= 0) Then
                    ws.Range("J" & summary_table).Interior.ColorIndex = 3
                End If
                ws.Range("K" & summary_table).Value = change_percentage
                ws.Range("L" & summary_table).Value = total_vol
                'Reset values
                summary_table = summary_table + 1
                total_vol = 0
                price_change = 0
                close_price = 0
                open_price = ws.Cells(i + 1, 3).Value
            Else
                total_vol = total_vol + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub
