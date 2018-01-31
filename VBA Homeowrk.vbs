Sub Ticker_Analysis()

'Define Variables
Dim volume As Double
Dim ticker As String
Dim outputrow As Double
Dim lastrow As Double
Dim startPrice As Double
Dim endPrice As Double
Dim ws As Worksheet

Application.ScreenUpdating = False

 
'For all worksheets
For Each ws In Worksheets


    'Define start variables
    ws.Activate
    startPrice = Cells(2, 3)
    outputrow = 2
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

    With ActiveSheet
        'Create output headers
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
        .Range("J:J").NumberFormat = "0.0000"
        .Range("K:K").NumberFormat = "0.00%"
    End With

    For i = 2 To lastrow
        'Define volume calc and ticker value
        volume = volume + Range("G" & i).Value
        ticker = Cells(i, 1).Value

        'Find value of last price
        'find date reset
        
        If Cells(i, 2).Value > Cells(i + 1, 2).Value Then
            
            endPrice = Cells(i, 6).Value
            Range("I" & outputrow).Value = ticker
            Range("J" & outputrow).Value = endPrice - startPrice

            
            'Change cell color on values
            If Range("J" & outputrow).Value >= 0 Then

                Range("J" & outputrow).Interior.Color = vbGreen
            Else
                Range("J" & outputrow).Interior.Color = vbRed

            End If

            'For last row error
            If startPrice <> 0 Then
            
                Range("K" & outputrow).Value = Range("J" & outputrow).Value / startPrice
            Else
                Range("K" & outputrow).Value = 0
                
            End If
            
            
            Range("L" & outputrow).Value = volume

            startPrice = Cells(i + 1, 6)
                       
            outputrow = outputrow + 1

            volume = 0
            
        End If
            
    Next i

Next ws

 
End Sub

Sub Hard()

Dim ticker As String
Dim volume As Double
Dim largestincrease As Double
Dim largestdecrease As Double
Dim ws As Worksheet
Dim lastrow As Double
Dim outputrow As Double

Application.ScreenUpdating = False
 
'For all worksheets
For Each ws In Worksheets

    'Define start variables
    ws.Activate
    outputrow = 2
    lastrow = Cells(Rows.Count, "I").End(xlUp).Row
    volume = 0
    largestincrease = 0
    largestdecrease = 0
        
    
    With ActiveSheet
        'Create output headers
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        .Range("Q2:Q3").NumberFormat = "0.00%"
        .Range("Q4").NumberFormat = "0"
    End With
    
    For i = 2 To lastrow
        
        If Cells(i, 12).Value > volume Then
            volume = Cells(i, 12)
            ticker = Cells(i, 9).Value
            Range("Q4").Value = volume
            Range("P4").Value = ticker
        Else
            Range("Q4").Value = volume
            
        End If
        
        If Cells(i, 11).Value > largestincrease Then
            largestincrease = Cells(i, 11).Value
            ticker = Cells(i, 9).Value
            Range("Q2").Value = largestincrease
            Range("P2").Value = ticker
        Else
            Range("Q2").Value = largestincrease
            
        End If
        
        If Cells(i, 11).Value < largestdecrease Then
            largestdecrease = Cells(i, 11).Value
            ticker = Cells(i, 9).Value
            Range("Q3").Value = largestdecrease
            Range("P3").Value = ticker
        Else
            Range("Q3").Value = largestdecrease
            
        End If
        
    Next i

Next ws
    
End Sub
