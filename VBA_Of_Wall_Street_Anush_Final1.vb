Sub VBA_Of_Wall_Street()

    For Each ws In Worksheets
       
        'Create new column headers

            ws.Cells(1, 9).Value = "Ticker"  
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Set Ticker as a variable
            Dim Ticker As String

        'Set vol_Total as a variable
            Dim vol_Total As Double
            vol_Total = 0
    
        'Keep track of each ticker location in summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
    
        'Find last row
            Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Set open_Price as a variable
            Dim open_Price As Double
            open_Price = ws.Cells(2,3).Value

        'Set other variables
            Dim close_Price As Double
            Dim Yearly_Change As Double
            Dim Percent_Change As Double

        'Loop Through all Tickers
            For i = 2 To Last_Row
        
                ' Check for change in Ticker
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                        'Set the Ticker
                            Ticker = ws.Cells(i, 1).Value
                        
                        'Set the close_Price
                            close_Price = ws.Cells(i, 6).Value

                        'Calc. vol_Total
                            vol_Total = vol_Total + ws.Cells(i, 7).Value

                        'Calc. Yearly_Change
                            Yearly_Change = close_Price - open_Price

                        'Calc. Percent_Change
                            Percent_Change = Yearly_Change / open_Price

                        ' Print the Ticker name to the Summary Table
                            ws.Range("I" & Summary_Table_Row).Value = Ticker
                    
                        ' Print the Stock Volume to the Summary Table
                            ws.Range("L" & Summary_Table_Row).Value = vol_Total

                        ' Print the Yearly_Change to the Summary Table
                            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                        ' Print the Percent_Change to the Summary Table
                            ws.Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change)

                            ' Highlight positive/negative change in Yearly_Change
                                If Yearly_Change > 0 Then
                                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)

                                ElseIf Yearly_Change < 0 Then
                                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)

                                End If
                    
                        ' Prep open_Price for next ticker
                            open_Price = ws.Cells(i+1,3).Value
                        
                        ' Prep summary table row for next ticker
                            Summary_Table_Row = Summary_Table_Row + 1

                        ' Reset the vol_Total
                            vol_Total = 0

                ' If no change in Ticker

                    Else

                        ' Add to the vol_Total
                            vol_Total = vol_Total + ws.Cells(i, 7).Value

                    End If

            Next i

        'Create new headers
            ws.Cells(1, 16).value = "Ticker"
            ws.Cells(1, 17).value = "Value"
            ws.Cells(2, 15).value = "Greatest % Increase"
            ws.Cells(3, 15).value = "Greatest % Decrease"
            ws.Cells(4, 15).value = "Greatest Total Volume"

        ' Set Greatest % Increase as a variable
            Dim Great_Inc As Double
           Great_Inc = 0

        ' Set Greatest % Decrease as a variable
            Dim Great_Dec As Double
            Great_Dec = 0

        ' Set Greatest Total Volume as a variable
            Dim Great_Vol As Double
            Great_Vol = 0

        'Find last row of Ticker Column
            Last_Row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Loop Through Ticker Column
             For i = 2 To Last_Row
                
                ' Calc and Print Greatest % Increase
                    If ws.Cells(i, 11).Value > Great_Inc Then
                        Great_Inc = ws.Cells(i, 11).Value
                        ws.Range("P2").Value = ws.Cells(i, 9).Value
                        ws.Range("Q2").Value = FormatPercent(Great_Inc)
                        
                    End If

                ' Calc and Print Greatest % Decrease    
                    If ws.Cells(i, 11).Value < Great_Dec Then
                        Great_Dec = ws.Cells(i, 11).Value
                        ws.Range("P3").Value = ws.Cells(i, 9).Value
                        ws.Range("Q3").Value = FormatPercent(Great_Dec)
                        
                    End If

                ' Calc and Print Greatest Total Volume
                    If ws.Cells(i, 12).Value > Great_Vol Then
                        Great_Vol = ws.Cells(i, 12).Value
                        ws.Range("P4").Value = ws.Cells(i, 9).Value
                        ws.Range("Q4").Value = Great_Vol
                        
                    End If
            Next i
        
        'Autofit new columns
        ws.Range("I1:Q1").EntireColumn.AutoFit
   
    Next ws

End Sub
