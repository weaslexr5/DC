Attribute VB_Name = "Module1"
Sub STOCK_DATA()
    For Each ws In Worksheets
        'Dim WorksheetName As String
        Dim Ticker As String
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        Dim Brand_Summary_Table As Integer
        Brand_Summary_Table = 2
        'Dim Yearly_Change As Double
        'Dim Percent_Change As Double
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'WorksheetName = ws.Name
       
       For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            'Yearly_Change =
            'Percent_Change =
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            ws.Range("I" & Brand_Summary_Table).Value = Ticker
            'ws.Range("J" & Brand_Summary_Table).Value = Yearly_Change
            'ws.Range("K" & Brand_Summary_Table).Value = Percent_Change
            ws.Range("L" & Brand_Summary_Table).Value = Total_Stock_Volume
            Brand_Summary_Table = Brand_Summary_Table + 1
            Total_Stock_Volume = 0
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        End If
        
        Next i
      
      ws.Cells(1, 9).Value = "Ticker"
      'ws.Cells(1, 10).Value = "Yearly Change"
      'ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 12).Value = "Total Stock Volume"
      ws.Columns("I:L").AutoFit
      
    Next ws
     

End Sub

