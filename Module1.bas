Attribute VB_Name = "Module1"
Sub Ticker():

' Define the worksheet parameters
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

' Create headers for where values will be pasted
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' Create a counter for the rows
Dim counter As Long
Dim Rows As Range
counter = 0

' Count Used Rows in Active Sheet
With ActiveSheet.UsedRange
    For Each Rows In .Rows
        If Application.CountA(Rows) > 0 Then
            counter = counter + 1
        End If
    Next
End With

' Define a stock volume counter
Dim stock_volume As Double
stock_volume = 0

' Create a variable that increases along with the tickers
Dim j As Integer
j = 2

' Calculate Sum of vol for each Ticker using a for loop

For i = 2 To counter
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
        stock_volume = stock_volume + ws.Cells(i, 7).Value
        
    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        stock_volume = stock_volume + ws.Cells(i, 7).Value
        
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
        
        ws.Cells(j, 12).Value = stock_volume
        
        j = j + 1
        
        stock_volume = 0
    End If
Next i

' Define variables for the first open, last close, the yearly change and the percent change
Dim ticker_start As Double
Dim ticker_end As Double
Dim ticker_change As Double
Dim ticker_percent_change
j = 2

' For loop to calculate all of the above for each ticker
For i = 2 To counter
    ws.Range("K" & i).NumberFormat = "0.00%"
        If Right(ws.Cells(i, 2), 4) = "0102" Then
            ticker_start = ws.Cells(i, 3)
    
            ElseIf Right(ws.Cells(i, 2), 4) = "1231" Then
                ticker_end = ws.Cells(i, 6)
            
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
    
                ticker_change = ticker_end - ticker_start
                ticker_percent_change = ticker_change / ticker_start
                ws.Cells(j, 10).Value = ticker_change
                ws.Cells(j, 11).Value = ticker_percent_change
             
            
                 
                j = j + 1
            
                
              
        End If
Next i

' Count Rows in output fields (amount of individual tickers)

Dim Tickercounter
Tickercounter = 0

For i = 1 To counter
    If ws.Cells(i, 9).Value <> "" Then
        Tickercounter = Tickercounter + 1
    End If
Next i

' Make both of the greatest percentage increase and decrease fields percentages
ws.Range("Q2", "Q3").NumberFormat = "0.00%"

' Make yearly change a colour according to greater or less than 0
For j = 2 To Tickercounter

    If ws.Cells(j, 10).Value > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    
    ElseIf ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
            
    End If
Next

' Define the variables for Max Volume and Max and Min Percent
Dim MaxPercent As Double
Dim MinPercent As Double
Dim MaxVolume As Double


MaxPercent = 0
MinPercent = 0
MaxVolume = 0

' For loop to work out which ticker value is the largest/smallest
For i = 2 To Tickercounter
    
       
        If ws.Cells(i, 11).Value > MaxPercent Then
            MaxPercent = ws.Cells(i, 11).Value
                
        
            
        ElseIf ws.Cells(i, 11).Value < MinPercent Then
            MinPercent = ws.Cells(i, 11).Value
            
        
        ElseIf ws.Cells(i, 12).Value > MaxVolume Then
            MaxVolume = ws.Cells(i, 12).Value
            
        End If
    
    
Next i

' Make cells show those values
ws.Range("Q2").Value = MaxPercent
ws.Range("Q3").Value = MinPercent
ws.Range("Q4").Value = MaxVolume

' For loop to work out which ticker has the corresponding value
For i = 2 To Tickercounter
    If ws.Range("Q2").Value = ws.Cells(i, 11).Value Then
    ws.Range("P2").Value = ws.Cells(i, 9).Value
    
    ElseIf ws.Range("Q3").Value = ws.Cells(i, 11).Value Then
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    
    ElseIf ws.Range("Q4").Value = ws.Cells(i, 12).Value Then
    ws.Range("P4").Value = ws.Cells(i, 9).Value
    End If
Next i

' Autosize columns
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit

' Go to next worksheet
Next ws


End Sub
