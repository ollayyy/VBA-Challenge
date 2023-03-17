Attribute VB_Name = "Module1"
Sub stockChecker()
' set variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim currentRow As Long
    
' set the headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
' last row
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    currentRow = 2
    
'for loop through all rows
    For i = 2 To lastRow
        
 'check for new ticker and set open price and add to volume
        If Cells(i, 1) <> Cells(i - 1, 1) Then
            ticker = Cells(i, 1).Value
            openPrice = Cells(i, 3).Value
            totalVolume = Cells(i, 7).Value
        Else
            totalVolume = totalVolume + Cells(i, 7).Value
        End If
        
 ' calculate yearly change and %change
        If Cells(i, 1) <> Cells(i + 1, 1) Or i = lastRow Then
            closePrice = Cells(i, 6).Value
            yearChange = closePrice - openPrice
            percentChange = yearChange / openPrice
            
            Range("I" & currentRow).Value = ticker
            Range("J" & currentRow).Value = yearChange
' conidtional format
                If yearChange >= 0 Then
                    Range("J" & currentRow).Interior.ColorIndex = 4
                Else
                    Range("J" & currentRow).Interior.ColorIndex = 3
                End If
            Range("K" & currentRow).Value = Format(percentChange, "Percent")
            Range("L" & currentRow).Value = totalVolume
            
            
            
' reset variables
            currentRow = currentRow + 1
            ticker = ""
            openPrice = 0
            closePrice = 0
            yearChange = 0
            percentChange = 0
            totalVolume = 0
        End If
    Next i
    
' set cell headers
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            
                    Dim maxValue As Double
                    Dim checkValueMax As Double
                    Dim lowValue As Double
                    Dim checkValueLow As Double
                    Dim maxVolume As Double
                    Dim checkVolumeMax As Double
            
            Range("O2").Value = "Greatest % Increase"
                maxValue = Range("K2").Value
'find max and min values of % changeand highest volume
                    For i = 2 To lastRow
                        checkValueMax = Cells(i, 11).Value
                        If checkValueMax > maxValue Then
                            maxValue = checkValueMax
                        End If
                    Next i
                    Range("Q2").Value = maxValue
                    Range("Q2").NumberFormat = "0.00%"
                    
                   
                    
            Range("O3").Value = "Greatest % Decrease"
                lowValue = Range("K2").Value
                    For i = 2 To lastRow
                        checkValueLow = Cells(i, 11).Value
                        If checkValueLow < lowValue Then
                            lowValue = checkValueLow
                        End If
                    Next i
                    Range("Q3").Value = lowValue
                    Range("Q3").NumberFormat = "0.00%"
                    
                    
            Range("O4").Value = "Greatest Total Volume"
                maxVolume = Range("L2").Value
                    For i = 2 To lastRow
                        checkVolumeMax = Cells(i, 12).Value
                        If checkVolumeMax > maxVolume Then
                            maxVolume = checkVolumeMax
                        End If
                    Next i
                    Range("Q4").Value = maxVolume
                    Range("P4").Value = Cells(i, 1).Value
                    
                    
' return tickers
                    For i = 2 To lastRow
                        If Cells(i, 11).Value = Range("Q2").Value Then
                            Range("P2").Value = Cells(i, 9).Value
                            End If
                    Next i
                    For i = 2 To lastRow
                        If Cells(i, 11).Value = Range("Q3").Value Then
                            Range("P3").Value = Cells(i, 9).Value
                            End If
                    Next i
                    For i = 2 To lastRow
                        If Cells(i, 12).Value = Range("Q4").Value Then
                            Range("P4").Value = Cells(i, 9).Value
                            End If
                    Next i
                    

                    
            
End Sub



