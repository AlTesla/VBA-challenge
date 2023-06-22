Const SUMMARY_VOL_COLUMN As Integer = 12
Const SUMMARY_PERCENTAGE_COLUMN As Integer = 11
Const SUMMARY_YEAR_CHANGE_COLUMN As Integer = 10
Const SUMMARY_TICKER_COLUMN As Integer = 9
Const TICKER_COLUMN As Integer = 1
Const VOL_COLUMN As Integer = 7


Sub Accounts()
Dim Year_Change_Array As Variant
Dim Row As Long
Dim ws_index As Integer
Dim Greatests_Array As Variant
ws_index = 1
For Each ws In Worksheets
Call Titles(ws)
Call Ticker_Name(ws, ws_index)
 
Row = 2
        'This loop for Each item in SUMMARY_TICKER_COLUMN call Year_Change function
 While IsEmpty(ws.Cells(Row, SUMMARY_TICKER_COLUMN)) = False
   Year_Change_Array = Year_Change(Row, ws, ws_index)

   ws.Cells(Row, SUMMARY_YEAR_CHANGE_COLUMN).Value = Year_Change_Array(1)
   ws.Cells(Row, SUMMARY_PERCENTAGE_COLUMN).Value = FormatPercent(Year_Change_Array(2))
   ws.Cells(Row, SUMMARY_VOL_COLUMN).Value = Year_Change_Array(3)
   ws.Cells(Row, SUMMARY_YEAR_CHANGE_COLUMN).Interior.ColorIndex = Color(Row, ws, SUMMARY_YEAR_CHANGE_COLUMN)
   ws.Cells(Row, SUMMARY_PERCENTAGE_COLUMN).Interior.ColorIndex = Color(Row, ws, SUMMARY_PERCENTAGE_COLUMN)
     Row = Row + 1
 Wend


 Greatests_Array = Greatests(ws)
   ws.Cells(2, 16).Value = ws.Cells(Greatests_Array(1), SUMMARY_TICKER_COLUMN).Value
   ws.Cells(3, 16).Value = ws.Cells(Greatests_Array(2), SUMMARY_TICKER_COLUMN).Value
   ws.Cells(4, 16).Value = ws.Cells(Greatests_Array(3), SUMMARY_TICKER_COLUMN).Value

   ws.Cells(2, 17).Value = FormatPercent(ws.Cells(Greatests_Array(1), SUMMARY_PERCENTAGE_COLUMN).Value)
   ws.Cells(3, 17).Value = FormatPercent(ws.Cells(Greatests_Array(2), SUMMARY_PERCENTAGE_COLUMN).Value)
   ws.Cells(4, 17).Value = ws.Cells(Greatests_Array(3), SUMMARY_VOL_COLUMN).Value

 ws_index = ws_index + 1
Next ws
End Sub


Function Ticker_Name(ws, ws_index)
Dim Ticker_Row As Long
Dim Ticker_Index As Long
Ticker_Row = 2
Ticker_Index = 2
While IsEmpty(ws.Cells(Ticker_Index, TICKER_COLUMN)) = False
ws.Cells(Ticker_Row, SUMMARY_TICKER_COLUMN).Value = ws.Cells(Ticker_Index, TICKER_COLUMN)
Ticker_Index = Ticker_Index + 250 + ws_index
Ticker_Row = Ticker_Row + 1
Wend
End Function


Function Year_Change(Row, ws, ws_index) As Variant
Dim Year_Array(1 To 3) As Double
Dim last_Stock_Row As Long
Dim Stock_Index As Long
Dim Year_Volume_Suma As Double
First_Stock_Row = 0
last_Stock_Row = 0

Stock_Index = (Row - 2) * (250 + ws_index) + 2

First_Stock_Row = Stock_Index

While ws.Cells(Stock_Index + 1, TICKER_COLUMN).Value = ws.Cells(Row, SUMMARY_TICKER_COLUMN).Value
Year_Volume_Suma = Year_Volume_Suma + ws.Cells(Stock_Index, VOL_COLUMN).Value
Stock_Index = Stock_Index + 1

Wend

last_Stock_Row = Stock_Index

Year_Array(1) = ws.Cells(last_Stock_Row, 6).Value - ws.Cells(First_Stock_Row, 3).Value
Year_Array(2) = Year_Array(1) / ws.Cells(First_Stock_Row, 3)
Year_Array(3) = Year_Volume_Suma
Year_Change = Year_Array

End Function


Function Color(Row, ws, Column) As Integer

If ws.Cells(Row, Column).Value > 0 Then
   Color = 4
 ElseIf ws.Cells(Row, Column).Value < 0 Then
  Color = 3
 Else
   Color = 6
 End If
 
End Function

Function Greatests(ws) As Variant
Dim Max_Array(1 To 3) As Integer
Dim Max_Decrease As Integer
Dim Max_Increase As Integer
Dim Max_Vol As Integer
Dim Ticker_Index As Integer
Max_Decrease = 2
Max_Increase = 2
Max_Vol = 2
Ticker_Index = 2

While IsEmpty(ws.Cells(Ticker_Index, SUMMARY_TICKER_COLUMN)) = False
  If ws.Cells(Ticker_Index, SUMMARY_PERCENTAGE_COLUMN).Value <= ws.Cells(Max_Decrease, SUMMARY_PERCENTAGE_COLUMN) Then
    Max_Decrease = Ticker_Index
  ElseIf ws.Cells(Ticker_Index, SUMMARY_PERCENTAGE_COLUMN).Value >= ws.Cells(Max_Increase, SUMMARY_PERCENTAGE_COLUMN) Then
    Max_Increase = Ticker_Index
  End If
  If ws.Cells(Ticker_Index, SUMMARY_VOL_COLUMN).Value >= ws.Cells(Max_Vol, SUMMARY_VOL_COLUMN) Then
    Max_Vol = Ticker_Index
  End If
  Ticker_Index = Ticker_Index + 1
Wend
Max_Array(1) = Max_Decrease
Max_Array(2) = Max_Increase
Max_Array(3) = Max_Vol
Greatests = Max_Array

End Function

Function Titles(ws)
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Decrease"
ws.Range("O3").Value = "Greatest % Increase"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
End Function