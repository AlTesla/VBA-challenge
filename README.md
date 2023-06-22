# VBA-challenge
this repository contains the VBA challenge, Data Analytics Bootcamp assignment
## scripting steps
First I began modifying the script for `credit_charges.xlsx` to works with the `alphabetical_testing.xlsx`, I was able to make the script run on all the sheets but it took 5 minutes to run,
then I was debugging the functions and trying to reduce interations and I got the time down to less that a minute so i thought it would work
with the `Multiple_year_stock_daba.xlsx` and I tried it, but it didn't work at all, i kept debugging and thinking algorithms to reduce iterations,
I found mathematical proportions that helped me find values with less iterations and merge functions to give me a list of values, so I was able to reduce 
the time in which the code is executed to less than a minute.

### First version Code of  `Year_Change` function 
The function returns a value and write another value and makes iterations over all rows in the worksheet

```vba
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
```
### Final Version code of `Year_Change` function
This version return an Array with three values `Yearly change`, `Percent Change` and `Total Stock volume`. Also this Function uses a mathematic proportion to find the values that it's looking for and doesn't use any `If`. So with 
these features I was able to reduce Running time  

```vba
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
```
Finally, I was able to run everything without any problems and getting all the data I was looking for, for the assignment.


![First WorkSheet header](https://github.com/AlTesla/VBA-challenge/blob/main/Header.png?raw=true)


Thanks for reading!
