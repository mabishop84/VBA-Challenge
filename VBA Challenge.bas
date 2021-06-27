Attribute VB_Name = "Module1"
Sub HW2()




Dim LastRow As LongLong
Dim Ticker_name As String
Dim Initial_Price As LongLong
Dim Closing_Price As LongLong
Dim Percent_Change As LongLong
Dim Price_Change As LongLong
Dim Summary_Table_Row As LongLong
Dim Volume_Total As LongLong



For Each ws In Worksheets



LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
Summary_Table_Row = 2

ws.Range("H1").Value = "Ticker Name"
ws.Range("I1").Value = "Change in Price"
ws.Range("J1").Value = "Pct Change in Price"
ws.Range("k1").Value = "Total Volume"

For i = 2 To LastRow
    
    
   'Determine if first row of new stock
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker_name = ws.Cells(i, 1).Value
    
        Initial_Price = ws.Cells(i, 3).Value
    
        Volume_Total = 0
        
    End If

    
     
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
   
    'Determine if last row of stock
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
        Closing_Price = ws.Cells(i, 6).Value
        Price_Change = Closing_Price - Initial_Price
                If Initial_Price = 0 Then
                Percent_Change = 100
                Else
                Percent_Change = (CLng(Price_Change) / Initial_Price) * 100
                End If
        ws.Range("H" & Summary_Table_Row).Value = Ticker_name
        ws.Range("I" & Summary_Table_Row).Value = Price_Change
                If Price_Change > 0 Then
                ws.Range("I" & Summary_Table_Row).Interior.ColorIndex = 4
                
                
                Else
                ws.Range("I" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                
               
        ws.Range("J" & Summary_Table_Row).Value = CLng(Percent_Change)
        
        ws.Range("K" & Summary_Table_Row).Value = Volume_Total
        Summary_Table_Row = Summary_Table_Row + 1
        
    
    End If
    

 
    
    Next i

Next ws




End Sub



