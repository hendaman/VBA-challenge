Attribute VB_Name = "Module1"
Sub Ticker():

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

Dim counter As Long
Dim Rows As Range

' Count Used Rows in Active Sheet
With ActiveSheet.UsedRange
    For Each Rows In .Rows
        If Application.CountA(Rows) > 0 Then
            counter = counter + 1
        End If
    Next
End With

Dim stock_volume As Double
stock_volume = 0

Dim j As Integer
j = 2

' Calculate Sum of vol for each Ticker

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

Dim ticker_start As Double
Dim ticker_end As Double
Dim ticker_change As Double
Dim ticker_percent_change
j = 2

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
            
            
        If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
    
        ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
              
    End If
Next i

' Count Rows in Ticker field

Dim Tickercounter

For i = 1 To counter
    If ws.Cells(i, 9).Value <> "" Then
        Tickercounter = Tickercounter + 1
    End If
Next i
ws.Range("Q2", "Q3").NumberFormat = "0.00%"

Dim MaxPercent As Double
Dim MinPercent As Double
Dim MaxVolume As Double


MaxPercent = 0
MinPercent = 0
MaxVolume = 0


For i = 2 To Tickercounter
    
       
        If ws.Cells(i, 11).Value > MaxPercent Then
            MaxPercent = ws.Cells(i, 11).Value
                
        
            
        ElseIf ws.Cells(i, 11).Value < MinPercent Then
            MinPercent = ws.Cells(i, 11).Value
            
        
        ElseIf ws.Cells(i, 12).Value > MaxVolume Then
            MaxVolume = ws.Cells(i, 12).Value
            
        End If
    
    
Next i

ws.Range("Q2").Value = MaxPercent
ws.Range("Q3").Value = MinPercent
ws.Range("Q4").Value = MaxVolume


For i = 2 To Tickercounter
    If ws.Range("Q2").Value = ws.Cells(i, 11).Value Then
    ws.Range("P2").Value = ws.Cells(i, 9).Value
    
    ElseIf ws.Range("Q3").Value = ws.Cells(i, 11).Value Then
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    
    ElseIf ws.Range("Q4").Value = ws.Cells(i, 12).Value Then
    ws.Range("P4").Value = ws.Cells(i, 9).Value
    End If
Next i

ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit

Next ws

' Calculate Percent Change the same way

' Add fields for Greatest % Increase/Decrease and Greatest Total Volume, headers for Ticker and Value

' Search Percent Change for Greatest % Increase, Decrease and total volume

End Sub
