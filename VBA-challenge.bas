Attribute VB_Name = "Module1"
Sub Challenge_2():
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
        Dim ticker As String
    
        Dim begin As Long
        begin = 2
        Dim j As Long
        j = 2
        Dim Lastrow As Long
        Dim Lastrows As Long
        
        Dim table As Integer
        table = 2
        
        Dim Yearly_Change As Double
        Dim total_stock As Double
        total_stock = 0
        Dim open_price As Double
        Dim close_price As Double
        Dim Percent_Change As Double
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double
        Dim Value As Double
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Value"
        
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & table).Value = ticker
                
                open_price = ws.Cells(begin, 3).Value
                begin = i + 1
                close_price = ws.Cells(i, 6).Value
                Yearly_Change = close_price - open_price
                ws.Range("J" & table).Value = Yearly_Change
                
                If ws.Range("J" & table).Value > 0 Then
                    ws.Range("J" & table).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & table).Interior.ColorIndex = 3
                End If
                
                Percent_Change = ((close_price - open_price) / open_price)
                ws.Range("K" & table).Value = Percent_Change
                
                ws.Columns("K").NumberFormat = "0.00%"
                
        
                total_stock = total_stock + Cells(i, 7).Value
                ws.Range("L" & table).Value = total_stock
                
                total_stock = 0
                table = table + 1
                begin = i + 1
                
            Else
                total_stock = total_stock + Cells(i, 7).Value
                ws.Range("L" & table).Value = total_stock
      
            End If
        Next i
         
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Lastrows = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        Greatest_Increase = ws.Cells(2, 11).Value
        Greatest_Decrease = ws.Cells(2, 11).Value
        Greatest_Volume = ws.Cells(2, 12).Value
        
        For i = 2 To Lastrows
            If ws.Cells(i, 11).Value > Greatest_Increase Then
            Greatest_Increase = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            Else
            Greatest_Increase = Greatest_Increase
            End If

            If ws.Cells(i, 11).Value < Greatest_Decrease Then
            Greatest_Decrease = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
            Else
            Greatest_Decrease = Greatest_Decrease
            End If
            
            If ws.Cells(i, 12).Value > Greatest_Volume Then
            Greatest_Volume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            Else
            Greatest_Volume = Greatest_Volume
            End If
        Next i
        
        ws.Cells(2, 17).Value = Greatest_Increase
        ws.Cells(3, 17).Value = Greatest_Decrease
        ws.Cells(4, 17).Value = Greatest_Volume
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Columns("A:R").AutoFit
    Next ws
    
    
End Sub

