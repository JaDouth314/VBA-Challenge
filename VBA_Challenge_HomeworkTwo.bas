Attribute VB_Name = "Module1"
Sub Alphabet_Test()

    For Each ws In Worksheets
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        Dim Ticker_Name As String
        Dim Ticker_Open As Double
        Ticker_Open = 0
        Dim Ticker_Close As Double
        Ticker_Close = 0
        Dim Yearly_Change As Double
        Yearly_Change = 0
        Dim Percent_Change As Double
        Percent_Change = 0
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        Ticker_Open = ws.Cells(Summary_Table_Row, 3).Value
              
        For i = 2 To LastRow
        
                    'See if the ticker name is the same, otherwise, run the loop
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    'Grab the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                    
                    'Calculate change in price
                Ticker_Close = ws.Cells(Summary_Table_Row, 6).Value
                Yearly_Change = (ws.Cells(i, 6) - Ticker_Open)
                Percent_Change = (Yearly_Change / Ticker_Open)
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                    'Print the value of the column to the table
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
                ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("L" & Summary_Table_Row).Value = Percent_Change
                ws.Range("L" & Summary_Table_Row).Style = "Percent"
                ws.Range("M" & Summary_Table_Row).Value = Ticker_Volume
                
                    'Add a row to the summary table to grab the next name
                Summary_Table_Row = Summary_Table_Row + 1
                    
                    'Reset the total for the next loop
                Ticker_Open = 0
                Ticker_Open = ws.Cells(i + 1, 3).Value
                Ticker_Close = 0
                Ticker_Volume = 0
                                        
            Else
                
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
            End If
                
            'Loop through the rows
        Next i
        
        Dim YearChange As Long
        YearChange = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        For i = 2 To YearChange
        
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
            
        Next i
    
    Next ws
    
End Sub
