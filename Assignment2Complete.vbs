Attribute VB_Name = "Module121"
Sub ticker():
Dim ticker As String
Dim tickerPrev As String
Dim tickerNext As String
Dim tickerCurrent As String
Dim firstPrice As Double
Dim lastPrice As Double
Dim Change As Double
Dim percentChange As Double
Dim J As Integer
Dim volumeCounterA As Double
Dim volumeTotal As Long
Dim volumeFirst As Long
Dim Sheet1, Sheet2, Sheet3 As Worksheet
Dim tickHead, chgHead, perChgHead, volHead As String

'Define values that will display at column heads
tickHead = "Ticker"
chgHead = "Chg"
perChgHead = "Percent Change"
volHead = "Total Volume"

'Set worksheets
Set Sheet1 = Worksheets("2018")
Set Sheet2 = Worksheets("2019")
Set Sheet3 = Worksheets("2020")


'Set J value for the rows into which data will be recorded
J = 2

'Find last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Start loop
For i = 2 To lastrow

    
    'Set ticker checking values
    tickerPrev = Range("A" & i - 1).Value
    tickerCurrent = Range("A" & i).Value
    tickerNext = Range("A" & i + 1).Value
    
    
        'Reset volume count, set tickerCurrent as ticker, grab ticker value, set first price, set initial volumeCounterA value
        If tickerCurrent <> tickerPrev Then
        
            volumeCounterA = 0
            ticker = tickerCurrent
            Range("I" & J).Value = ticker
            firstPrice = Range("c" & i).Value
            volumeCounterA = Range("G" & i).Value
        
        'If the tickerCurrent is equal to tickerPrev, add the row's volume to volumeCounterA
        Else
        
            volumeCounterA = volumeCounterA + Range("G" & i).Value
        
        End If
        
        
        'find and set last price, and calculate values
        If tickerNext <> tickerCurrent Then
            
            lastPrice = Range("F" & i).Value
            Change = (lastPrice - firstPrice)
            percentChange = (lastPrice - firstPrice) / firstPrice
            Range("J" & J).Value = Change
            Range("K" & J).Value = percentChange
            Range("K" & J).NumberFormat = "0.00%"
            Range("L" & J) = volumeCounterA
            J = J + 1
            volumeCounterA = 0

        End If
        
        
        'Format Change column
        If Change = 0 Then
            
            Range("J" & J).Interior.ColorIndex = 1
        
        End If
                   
        If Not (Range("J" & 1).Value) And (Change > 0) Then
            
            Range("J" & J - 1).Interior.ColorIndex = 4
        
        End If
            
        If Not (Range("J" & 1).Value) And (Change < 0) Then
            Range("J" & J - 1).Interior.ColorIndex = 3
        
        End If
            
        
        'set column headers, select next sheet, reset i and j
        If i = lastrow Then
            
            Range("I" & 1).Value = tickHead
            Range("J" & 1).Value = chgHead
            Range("K" & 1).Value = perChgHead
            Range("L" & 1).Value = volHead
            ActiveSheet.Next.Select
            i = 1
            J = 2
        
        End If
      
Next i

End Sub



