Attribute VB_Name = "Module1"
Sub small_stock()
    '   Declaring variables
    Dim ws As Worksheet
     Dim open_value As Double
     Dim close_value As Double
     Dim count As Integer
     Dim total_volume As Double
     Dim great As Double
     Dim low As Double
     Dim max As Double
     Dim min As Double
     Dim tmax As String
     Dim tmin As String
     Dim vol_max As Double
     Dim vmax As String
     Dim first_val(10000) As Double
     max = (-1000#)
     min = (1000#)
    vol_max = 0
    Dim ticker_name As String
    Dim table_row As Long
    Dim yearly_change As Double
    Dim percentage_change As Double
   
For Each ws In Worksheets
table_row = 2
total_volume = 0
    
    
    
      'Set ws = ActiveSheet

'MsgBox (ws.Name)
        max = -1000
        min = 10000
        vol_max = 0
'headers and formatting
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
   
lastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
' Change_in_Price = Close_Stock_Price - Open_Stock_Price 'Calculates change in price?


    'ws.Range("k:k").NumberFormat = "0.00%"
open_value = ws.Range("C2")
    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Range("I" & table_row).Value = ws.Cells(i, 1).Value
            
            ws.Range("L" & table_row).Value = total_volume
            close_value = ws.Cells(i, 6).Value
            yearly_change = close_value - open_value
            ws.Cells(table_row, 10).Value = yearly_change
            If open_value = 0 Then
              percent_Change = 0
            Else
                percent_Change = yearly_change / open_value
            End If
            ws.Range("K" & table_row).Value = percent_Change
            open_value = ws.Cells(i + 1, 3).Value
            total_volume = 0
            'If open_value <> 0 Then
            'percentage_change = yearly_change / open_value
            If ws.Cells(table_row, 10) < 0 Then
                ws.Cells(table_row, 10).Interior.ColorIndex = 3 'Format color to Red if negative
                ws.Cells(table_row, 10).Font.ColorIndex = 1
            Else
                ws.Cells(table_row, 10).Interior.ColorIndex = 4 ' Format color to Green if positive
                ws.Cells(table_row, 10).Font.ColorIndex = 1
            End If
            table_row = table_row + 1
        Else
            total_volume = total_volume + ws.Cells(i, 7).Value
            ' Openvalue = ws.Cells(i, 3)
        End If
    'Format Percent Change DONE
    
    

        
            ' table_row = table_row + 1
            
            
            'ws.Cells(i, 11) = "%" And percent_change
        'End If
        
        Next i
    
    Columns("A:Q").AutoFit
    
    Next ws
    
        End Sub
        
        
