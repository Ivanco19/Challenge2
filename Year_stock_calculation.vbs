Sub year_stock()

    Dim i, j, n, k As Integer
    Dim ticker, ticker_max, ticker_min, ticker_max_volume As String
    Dim open_ticker, close_ticker, stock_volume, yearly_change, max_value, min_value, max_volume As Double
    
    'PART I. We initialize values for ticker, its open value and the stock volume
    ticker = Cells(2, 1).Value
    open_ticker = Cells(2, 3).Value
    stock_volume = 0
    'This variable helps us insert summarize data in a table (Column I:L)
    j = 2
    
    'This loop provides Tickers information
    n = ActiveSheet.UsedRange.Rows.Count + 1
    For i = 2 To n
    'If our ticker is the same as the cell we are looking at, we sum stock volume
        If Cells(i, 1).Value = ticker Then
            stock_volume = stock_volume + Cells(i, 7).Value
        Else
    'If not, it means we found another ticker, so we must print the values we found until the row before
            close_ticker = Cells(i - 1, 6)
            Cells(j, 9).Value = ticker
            Cells(j, 10).Value = close_ticker - open_ticker
            Cells(j, 11).Value = ((close_ticker - open_ticker) / open_ticker)
            Cells(j, 12).Value = stock_volume
            'We color Green if yearly change is positive and Red if negative
            If Cells(j, 10).Value >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            
     'Finally we have to update/reset our values because we found another ticker
            ticker = Cells(i, 1).Value
            open_ticker = Cells(i, 3).Value
            stock_volume = Cells(i, 7).Value
            j = j + 1
        End If
    Next i
    
    'PART II. This loop provides Greates values
    k = Application.WorksheetFunction.CountA(ActiveSheet.Range("I:I"))
    
    'We will compare to get max and min values. We first need to initialize some variables
    ticker_max = Cells(2, 9)
    ticker_min = Cells(2, 9)
    ticker_max_volume = Cells(2, 9)
    
    max_value = Cells(2, 11)
    min_value = Cells(2, 11)
    max_volume = Cells(2, 12)
    
    'We compare our current values to the current cell
    For i = 3 To k
        If Cells(i, 11) > max_value Then
            max_value = Cells(i, 11)
            ticker_max = Cells(i, 9)
        End If
        If Cells(i, 11) < min_value Then
            min_value = Cells(i, 11)
            ticker_min = Cells(i, 9)
        End If
        If Cells(i, 12) > max_volume Then
            max_volume = Cells(i, 12)
            ticker_max_volume = Cells(i, 9)
        End If
    Next i
    
    'Once done, we print the values
    Cells(2, 15).Value = ticker_max
    Cells(3, 15).Value = ticker_min
    Cells(4, 15).Value = ticker_max_volume
    
    Cells(2, 16).Value = max_value
    Cells(3, 16).Value = min_value
    Cells(4, 16).Value = max_volume
    
End Sub