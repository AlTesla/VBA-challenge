# VBA-challenge
this repository contains the VBA challenge, Data Analytics Bootcamp assignment

First I begin modifying the script for `credit_charges.xlsx` to works with the `alphabetical_testing.xlsx`, I was able to make the script run on all the sheets but it took 5 minutes to run,
then I was debugging the functions and trying to reduce interations and I got the time down to less that a minute so i thought it would work
with the `Multiple_year_stock_daba.xlsx` and I tried it, but it didn't work at all, i kept debugging and thinking algorithms to reduce iterations,
I found mathematical proportions that helped me find values with a menu and uni functions to give me a list of values, so I managed to reduce 
the time in which the code is executed to less than a minute.

### This is the first version Code of the `Year_Change` function 

    Function Year_Change(Row, ws) As Double
    Dim First_Date As Long
    Dim Last_Date As Long
    Dim First_Stock_Row As Long
    Dim last_Stock_Row As Long
    Dim Stock_Index As Long
    First_Date = 0
    Last_Date = 0
    First_Stock_Row = 0
    last_Stock_Row = 0
    Stock_Index = 2
    
    While IsEmpty(ws.Cells(Stock_Index, TICKET_COLUMN)) = False
        If ws.Cells(Stock_Index, TICKET_COLUMN).Value = ws.Cells(Row, SUMMARY_TICKET_COLUMN).Value Then
          If First_Date = 0 Then
            First_Date = ws.Cells(Stock_Index, 2).Value
            First_Stock_Row = Stock_Index
          End If
          If ws.Cells(Stock_Index, 2).Value < First_Date Then
            First_Date = ws.Cells(Stock_Index, 2).Value
            First_Stock_Row = Stock_Index
          End If
          If Last_Date = 0 Then
            Last_Date = ws.Cells(Stock_Index, 2).Value
            last_Stock_Row = Stock_Index
          End If
          If ws.Cells(Stock_Index, 2).Value > Last_Date Then
          Last_Date = ws.Cells(Stock_Index, 2).Value
          last_Stock_Row = Stock_Index
          End If
        End If
        
      Stock_Index = Stock_Index + 1
    Wend
    Year_Change = ws.Cells(last_Stock_Row, 6).Value - ws.Cells(First_Stock_Row, 3).Value
    ws.Cells(Row, 11).Value = Year_Change / ws.Cells(First_Stock_Row, 3)
    End Function
