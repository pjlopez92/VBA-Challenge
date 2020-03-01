Attribute VB_Name = "Module1"
Sub Stock_market_analyst()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Stock Volume"
        
        Dim ticker_symbol As String
    
        Dim Yearly_Open As Double
        Yearly_Open = 0
    
        Dim Yearly_Close As Double
        Yearly_Close = 0
        
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
        Dim Percent_Change As Double
        Percent_Change = 0
    
        Dim Stock_Volume As Double
        Stock_Volume = 0
    
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
    
        Dim Lastrow As Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
               For i = 2 To Lastrow

                 If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                     Yearly_Open = ws.Cells(i, 3).Value
                                        
                 End If
                 
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7)
                    
                      If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                      
                        ticker_symbol = ws.Cells(i, 1).Value
                        ws.Cells(Summary_Table_Row, 9).Value = ticker_symbol
                                                                                           
                        Yearly_Close = ws.Cells(i, 6).Value
                      
                        Yearly_Change = Yearly_Close - Yearly_Open
                        ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                       
                        ws.Cells(Summary_Table_Row, 12).Value = Stock_Volume
              
                            If Yearly_Change >= 0 Then
                 
                                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    
                            Else
                
                                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    
                            End If
                
                            If Yearly_Open = 0 And Yearly_Close = 0 Then
                    
                                Percent_Change = 0
                    
                                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                    
                                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                    
                             ElseIf Yearly_Open = 0 Then
                   
                                Dim Percent_Change_NS As String
                    
                                Percent_Change_NS = "New Stock"
                    
                                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change_NS
                    
                             Else
                
                                Percent_Change = Yearly_Change / Yearly_Open
                    
                                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                    
                                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                    
                             End If
               
                         Summary_Table_Row = Summary_Table_Row + 1
      
                         Yearly_Open = 0
                     
                         Yearly_Close = 0
                     
                         Yearly_Change = 0
                 
                        Percent_Change = 0
                 
                        Stock_Volume = 0
                     
                 End If
               
            Next i
             
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        Dim Greatest_Stock As String
        
        Dim Greatest_Value As Double
        Greatest_Value = ws.Cells(2, 11).Value

        Dim Lowest_Stock As String
        Dim Lowest_Value As Double
        Lowest_Value = ws.Cells(2, 11).Value

        Dim most_vol_stock As String
        Dim most_vol_value As Double
        most_vol_value = ws.Cells(2, 12).Value
        
        Lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To Lastrow

            If ws.Cells(j, 11).Value > Greatest_Value Then
            
                Greatest_Value = ws.Cells(j, 11).Value
                
                Greatest_Stock = ws.Cells(j, 9).Value
                
            End If

            If ws.Cells(j, 11).Value < Lowest_Value Then
            
                Lowest_Value = ws.Cells(j, 11).Value
                
                Lowest_Stock = ws.Cells(j, 9).Value
                
            End If

            If ws.Cells(j, 12).Value > most_vol_value Then
            
                most_vol_value = ws.Cells(j, 12).Value
                
                most_vol_stock = ws.Cells(j, 9).Value
                
            End If

        Next j

        ws.Cells(2, 16).Value = Greatest_Stock
        ws.Cells(2, 17).Value = Greatest_Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = Lowest_Stock
        ws.Cells(3, 17).Value = Lowest_Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value

        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit

    Next ws

End Sub


