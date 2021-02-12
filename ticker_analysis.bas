Attribute VB_Name = "Module1"
Sub ticker_analysis()
'Dim & set variables for loop
Dim ws As Worksheet
Dim i As Long
Dim lastrow As Long
Dim volume As Double
Dim offset As Integer
Dim open_price As Double
Dim close_price As Double
Dim first_iteration As Integer
Dim percent_change As Double
Dim price_change As Double
offset = 2
first_iteration = 0

' **LOOP**
'*Set Table*
'Set table from I to L
For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Price Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Volume"
'Set table for hard solution
    ws.Cells(2, 14).Value = "Greatest % increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
'*Set Table* End
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Loop
    For i = 2 To lastrow
        'Set
        If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            first_iteration = first_iteration + 1
            volume = volume + ws.Cells(i, 7)
            If first_iteration = 1 Then
                open_price = ws.Cells(i, 3)
            Else
            End If
        Else
            volume = volume + ws.Cells(i, 7)
            ws.Cells(offset, 9) = ws.Cells(i, 1)
            ws.Cells(offset, 12) = volume
            close_price = ws.Cells(i, 6)
            If open_price <> 0 Then
                percent_change = ((close_price - open_price) / open_price)
                price_change = close_price - open_price
            Else
                percent_change = 0
                price_change = 0
            End If
            ws.Cells(offset, 11) = percent_change
            ws.Cells(offset, 11).NumberFormat = "0.00%"
            ws.Cells(offset, 10) = price_change
                'conditional formatting
                If ws.Cells(offset, 10).Value > 0 Then
                    ws.Cells(offset, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(offset, 10).Interior.ColorIndex = 3
                End If
            volume = 0
            offset = offset + 1
            first_iteration = 0
        End If
    Next i
    offset = 2
'**END LOOP**


'**Hard Solution**
    Dim lastrow_new As Long
    lastrow_new = ws.Cells(Rows.Count, 9).End(xlUp).Row

    Dim best_stock As String
    Dim best_percent As Double

    best_stock = ws.Cells(2, 9).Value
    best_percent = ws.Cells(2, 10).Value

    Dim worst_stock As String
    Dim worst_percent As Double

    worst_stock = ws.Cells(2, 9).Value
    worst_percent = ws.Cells(2, 10).Value

    Dim max_vol_stock As String
    Dim max_vol As Double

    max_vol_stock = ws.Cells(2, 9).Value
    max_vol = ws.Cells(2, 12).Value

    Dim x As Integer

    For x = 2 To lastrow_new
        'find best stock %
        If ws.Cells(x, 11).Value > best_percent Then
            best_stock = ws.Cells(x, 9).Value
            best_percent = ws.Cells(x, 11).Value
        End If
        'find worst stock %
        If ws.Cells(x, 11).Value < worst_percent Then
            worst_stock = ws.Cells(x, 9).Value
            worst_percent = ws.Cells(x, 11).Value
        End If
        'find max volume
        If ws.Cells(x, 12).Value > max_vol Then
            max_vol_stock = ws.Cells(x, 9).Value
            max_vol = ws.Cells(x, 12).Value
        End If
    Next x
    'display best %
    ws.Cells(2, 15).Value = best_stock
    ws.Cells(2, 16).Value = best_percent
    ws.Cells(2, 16).NumberFormat = "0.00%"
    'display worst %
    ws.Cells(3, 15).Value = worst_stock
    ws.Cells(3, 16).Value = worst_percent
    ws.Cells(3, 16).NumberFormat = "0.00%"
    'display max volume
    ws.Cells(4, 15).Value = max_vol_stock
    ws.Cells(4, 16).Value = max_vol
    'format column widths
    ws.Columns("I:P").EntireColumn.AutoFit
        
Next ws
End Sub
