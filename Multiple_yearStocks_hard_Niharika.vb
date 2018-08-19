Sub Multiple_yearStocks()

Dim ws As Worksheet
Dim WorksheetName As String
Dim Summary_Table_Row As Integer
Dim opening_amt As Double
Dim Yearly_change As Double
Dim Percent_Change As Double
Dim i As Double
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim volume_total As Double
Dim Ticker_Nm As String
Dim highpercent As Integer
Dim lowpercent As Integer
Dim highstockvol As Double
Dim TickerName As String
Dim TickerNm As String
Dim vol_Ticker_Name As String


For Each ws In Worksheets

   WorksheetName = ws.Name
   MsgBox WorksheetName

    Summary_Table_Row = 2
    volume_total = 0

'Set stock_sheet = Worksheets(ws)'
'Set stock_sheet = Worksheets("stock_sheet")'

    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    opening_amt = ws.Cells(2, 3).Value
'MsgBox opening_amt'

    For i = 2 To LastRow

        'when Current row Ticker is <> to Ticker in next row'
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
         'Print the Ticker Name on column I'
 
        Ticker_Nm = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Nm
     
        'Print Total Stock Volume in column L'
        volume_total = volume_total + ws.Cells(i, 7).Value
        ws.Range("L" & Summary_Table_Row).Value = volume_total

        'Print Yearly change in column J'
         Yearly_change = Round((ws.Cells(i, 6).Value - opening_amt), 8)
         ws.Range("J" & Summary_Table_Row).Value = Yearly_change
         ws.Range("J" & Summary_Table_Row).Font.ColorIndex = 1
         
         'Conditional Formatting of Yearly change
         If Yearly_change > 0 Then
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
         Else
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
         End If
         
         divisor = ws.Cells(i, 6).Value
          
         'Print Percent Change in column K'
         If Yearly_change <> 0 And divisor <> 0 Then
         check = Yearly_change / divisor * 100
         Else
'         check = Yearly_change / ws.Cells(i, 6).Value * 100
           ' check1 = (Yearly_change / divisor)
           check = 0
         End If
          '  check = check1 * 100
         Percent_Change = Round(check, 2)
        'Percent_Change = Round((Yearly_change / ws.Cells(i, 6).Value * 100), 3)
         ws.Range("K" & Summary_Table_Row).Value = Percent_Change
       
         opening_amt = ws.Cells(i + 1, 3).Value
        'ws.Range("R" & Summary_Table_Row).Value = ws.Cells(i, 6).Value
    
         Summary_Table_Row = Summary_Table_Row + 1
         volume_total = 0
        'when Current row Ticker is same as Ticker in next row'
         Else
         volume_total = volume_total + ws.Cells(i, 7).Value
    
        End If
        'MsgBox ("End if IF")
    Next i
    'MsgBox ("End if I")

   'To Calculate Percent decrease and increase
    j = 2
    highpercent = ws.Cells(j, 11).Value
    For j = 3 To Summary_Table_Row
         If ws.Cells(j, 11).Value > highpercent Then
         highpercent = ws.Cells(j, 11).Value
         TickerName = ws.Cells(j, 9).Value
         Else
         End If
    Next j
    ws.Range("Q" & 2).Value = "Greatest % Increase"
    ws.Range("R" & 2).Value = TickerName
    ws.Range("S" & 2).Value = highpercent
    

    k = 2
    lowpercent = ws.Cells(k, 11).Value
    For k = 3 To Summary_Table_Row
         If ws.Cells(k, 11).Value < lowpercent Then
         lowpercent = ws.Cells(k, 11).Value
         TickerNm = ws.Cells(k, 9).Value
         Else
         End If
    Next k
    ws.Range("Q" & 3).Value = "Greatest % Decrease"
    ws.Range("R" & 3).Value = TickerNm
    ws.Range("S" & 3).Value = lowpercent
    
    m = 2
    highstockvol = ws.Cells(m, 12).Value
    For m = 3 To Summary_Table_Row
         If ws.Cells(m, 12).Value > highstockvol Then
         highstockvol = ws.Cells(m, 12).Value
         vol_Ticker_Name = ws.Cells(m, 9).Value
         Else
         End If
    Next m
    ws.Range("Q" & 4).Value = "Greatest Total Volume"
    ws.Range("R" & 4).Value = vol_Ticker_Name
    ws.Range("S" & 4).Value = highstockvol
    

'MsgBox ("complete")
Next ws
'MsgBox ("complete")

End Sub




