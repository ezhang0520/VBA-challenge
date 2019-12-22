Option Explicit
Sub Allsheets()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call SB
    Next
    Application.ScreenUpdating = True
End Sub

Sub SB()
   'Dim wb As Workbook: Set wb = ThisWorkbook
   'Dim ws As Worksheet: Set ws = ThisWorkSheet
   Dim Ticker As String
   Dim vol_sum As Double
   Dim position As Double
   Dim start_v As Double
   Dim end_v As Double
   Dim LastRow As Double
   Dim i As Long
   
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    vol_sum = 0
    position = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    start_v = Cells(position, 3).Value
    For i = 2 To LastRow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Ticker = Cells(i, 1).Value
      end_v = Cells(i, 6).Value
      vol_sum = vol_sum + Cells(i, 7).Value
      Range("I" & position).Value = Ticker
      Range("J" & position).Value = end_v - start_v
      If start_v = 0 Or (end_v - start_v) = 0 Then
        Range("K" & position).Value = 0
      Else
        Range("K" & position).Value = (end_v - start_v) / start_v
      End If
      Range("L" & position).Value = vol_sum
      If Cells(i + 1, 1) <> "" Then
       start_v = Cells(i + 1, 3)
       position = position + 1
       vol_sum = 0
      End If
    Else
      vol_sum = vol_sum + Cells(i, 7).Value
    End If
    Next i

    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    Dim increase As Double
    Dim Ticker_i As String
    Dim decrease As Double
    Dim Ticker_d As String
    Dim max_vol As Double
    Dim Ticker_vol As String
    
    increase = 0
    decrease = 0
    max_vol = 0
    
    Ticker_i = Cells(2, 9).Value
    Ticker_d = Cells(2, 9).Value
    Ticker_vol = Cells(2, 9).Value

    End Sub