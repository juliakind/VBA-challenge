Attribute VB_Name = "Module1"
        Sub TickerYK()
        
        Dim ticker As String
        Dim begin_y As Double
        begin_y = 0
        Dim end_y As Double
        end_y = 0
        Dim dif As Double
        Dim total_v As Double
        Dim result As Double
        Dim price_chg As Double
        Dim index_price As Double
        Dim prev_value As Double
        Dim max_value As Double
        Dim min_value As Double
        Dim last As Double
        
        ' Outer for loop that makes sure my code runs through all tabs of the worksheet
             For Each ws In Worksheets
                  last = ws.Cells(Rows.Count, "A").End(xlUp).Row 'formula that sets the last row with the value
                  result = 2
                  ' Setting colum names for the analytic data
                 ws.Range("I1") = "Ticker"
                 ws.Range("J1") = "Price Change"
                 ws.Range("K1") = "% Change"
                 ws.Range("L1") = "Total Volume"
                 
                  total_v = 0
                  prev_value = 0
                  index_price = 1
                  max_value = 0
        
                 ' Inner loop for all rows of information in each worksheet tab
                          For i = 2 To last
                            ' Contitional statement which pulls out stock tickers and summerizes their total volume
                             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                             ws.Cells(result, 9) = ws.Cells(i, 1).Value
                             ws.Cells(result, 12) = total_v
        
                             result = result + 1
                             total_v = 0
        
                         Else
                             total_v = total_v + ws.Cells(i + 1, 7).Value
        
                         End If
                         
                    'Contitional statement which calculate annual prise difference in $ and %
                     If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                             If i = 2 Then
                                 begin_y = ws.Cells(2, 3).Value
                             Else
                                begin_y = ws.Cells(i, 3).Value
                             End If
                             index_price = index_price + 1
                       Else
                            end_y = ws.Cells(i, 6).Value
                            dif = end_y - begin_y
                          ws.Cells(index_price, 10) = dif
                             If begin_y <> 0 Then
                                 price_chg = dif / begin_y
                             
                         
        
                                'Contitional statements which set format of the Column K to percentage and visually separates negative and positive values in Column J by two different color indexes
                                If dif > 0 Then
                                ws.Cells(index_price, 10).Interior.ColorIndex = 4
                                Else
                                ws.Cells(index_price, 10).Interior.ColorIndex = 3
                                End If
        
                              ws.Cells(index_price, 11) = Format(price_chg, "Percent")
                             Else
                                 ws.Cells(index_price, 10) = ""
                                 ws.Cells(index_price, 10).Interior.ColorIndex = 5
                             End If
                             
                     End If
                             
                 Next i
                 ' Block of code which pulls and displays Min and Max percentage change in stocks price and Max stock total volume with the related ticker index
                 ws.Cells(1, 14) = "Ticker"
                 ws.Cells(1, 15) = "Value"
                 ws.Cells(2, 13) = "Max % Increase"
                 ws.Cells(3, 13) = "Min % Decrease"
                 ws.Cells(4, 13) = "Max Total Volume"
                 ws.Cells(2, 15) = Format(WorksheetFunction.Max(ws.Range("K:K")), "Percent")
                 ws.Cells(2, 14) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0), 9).Value
                 ws.Cells(3, 15) = Format(WorksheetFunction.Min(ws.Range("K:K")), "Percent")
                 ws.Cells(3, 14) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0), 9).Value
                 ws.Cells(4, 15) = WorksheetFunction.Max(ws.Range("L:L"))
                 ws.Cells(4, 14) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0), 9).Value
        
            Next ws
        
        End Sub
        
        
        
