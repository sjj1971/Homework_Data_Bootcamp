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
        ticker_name = Cells(i,1)
        if ticker_name = Cells(i+1,1) Then
            if open_date = 0 Then
                open_date = Cells(i,2)
                open_value = Cells(i,3)
                close_date = Cells(i,2)
                close_value = Cells(i,6)
                amount = Cells(i,7)
            Else
                if open_date > Cells(i,2) Then
                    open_date = Cells(i,2)
                    open_value = Cells(i,3)
                end If
                if close_date < Cells(i,2) Then
                    close_date = Cells(i,2)
                    close_value = Cells(i,6)
                end If
                amount = cells(i,7) + amount
            end If
        Else
            if open_date = 0 Then
                open_date = Cells(i,2)
                open_value = Cells(i,3)
                close_date = Cells(i,2)
                close_value = Cells(i,6)
                amount = Cells(i,7)
            Else
                if open_date > Cells(i,2) Then
                    open_date = Cells(i,2)
                    open_value = Cells(i,3)
                end If
                if close_date < Cells(i,2) Then
                    close_date = Cells(i,2)
                    close_value = Cells(i,6)
                end If
                amount = cells(i,7) + amount
            end If

            Cells(current_ticker_index,9) = ticker_name
            Cells(current_ticker_index,10) = close_value - open_value
            if Cells(current_ticker_index,10)>=0 Then
                Cells(current_ticker_index,10).interior.colorindex = 4
            Else
                Cells(current_ticker_index,10).interior.colorindex = 3
            end If

            if open_value > 0 Then
                Cells(current_ticker_index,11) = Round(Cells(current_ticker_index,10)/open_value,4)
                Cells(current_ticker_index,11).NumberFormat = "0.00%"                
            end If
            
            Cells(current_ticker_index,12) = amount
            current_ticker_index = current_ticker_index + 1
            amount = 0
            open_date = 0
            open_value = 0
            close_date = 0
            close_value = 0
        end If
    next i
        
End Sub
