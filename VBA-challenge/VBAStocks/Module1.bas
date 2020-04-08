Attribute VB_Name = "Module1"
Sub StockTicker_calculator():
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percente Change"
    Cells(1, 12) = "Total Stock Volume"
        
    len_ticker_column = 70926
    amount = 0
    open_date = 0
    open_value = 0
    close_date = 0
    close_value = 0
    current_ticker_index = 2
    
    
    For i = 2 To len_ticker_column
        ticker_name = Cells(i, 1)
        If ticker_name = Cells(i + 1, 1) Then
            If amount = 0 Then
                open_date = Cells(i, 2)
                open_value = Cells(i, 3)
                close_date = Cells(i, 2)
                close_value = Cells(i, 6)
                amount = Cells(i, 7)
            Else
                If open_date > Cells(i, 2) Then
                    open_date = Cells(i, 2)
                    open_value = Cells(i, 3)
                End If
                If close_date < Cells(i, 2) Then
                    close_date = Cells(i, 2)
                    close_value = Cells(i, 6)
                End If
                amount = Cells(i, 7) + amount
            End If
        Else
            If amount = 0 Then
                open_date = Cells(i, 2)
                open_value = Cells(i, 3)
                close_date = Cells(i, 2)
                close_value = Cells(i, 6)
                amount = Cells(i, 7)
            Else
                If open_date > Cells(i, 2) Then
                    open_date = Cells(i, 2)
                    open_value = Cells(i, 3)
                End If
                If close_date < Cells(i, 2) Then
                    close_date = Cells(i, 2)
                    close_value = Cells(i, 6)
                End If
                amount = Cells(i, 7) + amount
            End If

            Cells(current_ticker_index, 9) = ticker_name
            Cells(current_ticker_index, 10) = close_value - open_value
            If Cells(current_ticker_index, 10) >= 0 Then
                Cells(current_ticker_index, 10).Interior.ColorIndex = 4
            Else
                Cells(current_ticker_index, 10).Interior.ColorIndex = 3
            End If

            If open_value > 0 Then
                Cells(current_ticker_index, 11) = Round(Cells(current_ticker_index, 10) / open_value, 4)
                Cells(current_ticker_index, 11).NumberFormat = "0.00%"
            End If
            
            Cells(current_ticker_index, 12) = amount
            current_ticker_index = current_ticker_index + 1
            amount = 0
            open_date = 0
            open_value = 0
            close_date = 0
            close_value = 0
        End If
    Next i
        
End Sub

