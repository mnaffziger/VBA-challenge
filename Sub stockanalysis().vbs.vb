Sub stockanalysis()

Dim ws As Worksheet
Dim n As Integer
Dim WScount As Integer

Dim ticker, GreatHighTic, GreatLowTic, GreatVolumeTic As String
Dim Tickercount, format_count As Long

Dim OpeningPrice, ClosingPrice, YearChange, percentchange As Double
Dim GreatHigh, GreatLow, neglow As Long
Dim volume, GreatVolume As Variant
Dim i, j As Long
Dim Lastrow As Long



WScount = ActiveWorkbook.Worksheets.Count ' support.mocrosoft.com: macro to loop through all worksheets in a workbook

For n = 1 To WScount

    With Worksheets(n).Activate
'I tried to utilize the For Each ws in Worksheets, but the code endedup running the A worksheet from the alphabet list test 6 times, or it ran the same calculation for every instance not a non-empty worksheet in the workbook.

    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row 'provided from -NC on Slack
    Tickercount = 0
    volume = 0
    
        With Range("I1:Q1")
            .Font.FontStyle = "Bold"
        End With
        
        With Range("O1:O4")
            .Font.FontStyle = "Bold"
        End With
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Yearly Volume Traded"
        Range("o2").Value = "Greatest % Increase"
        Range("o3").Value = "Greatest % Decrease"
        Range("o4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        For i = 2 To Lastrow
        
            ticker = Cells(i, 1).Value
            volume = volume + Cells(i, 7).Value
            
            If Cells(i - 1, 1).Value <> ticker Then
                OpeningPrice = Cells(i, 3).Value
        
                ElseIf Cells(i + 1, 1).Value <> ticker Then
                    ClosingPrice = Cells(i, 6).Value
                                
                    Tickercount = Tickercount + 1
                    
                    YearChange = ClosingPrice - OpeningPrice
                    percentchange = YearChange / OpeningPrice
    
                    Range("I" & Tickercount + 1).Value = ticker
                    Range("J" & Tickercount + 1).Value = YearChange
                    Range("K" & Tickercount + 1).Value = percentchange
                    Range("L" & Tickercount + 1).Value = volume
                    
                    
                    Range("K" & Tickercount + 1).NumberFormat = "0.00%"
                    Range("J" & Tickercount + 1).NumberFormat = "$#,##0.00"
                    Range("L" & Tickercount + 1).NumberFormat = "#,##0"
                    Range("I:I,L:L").EntireColumn.AutoFit
                   
                    volume = 0
                    YearChange = 0
                    percentchange = 0
            End If
            
        Next i
    
   With Range("I1:L1")
        .Font.FontStyle = "Bold"
        .EntireColumn.AutoFit
    End With
    
    format_count = Cells(Rows.Count, 9).End(xlUp).Row
    
    'don't really need the next 4 lines, but was too aftraid to get rid of them since the code worked....before i realized these lines wiere not commented out
    
    GreatHigh = Cells(2, 11).Value
    GreatLow = Cells(2, 11).Value
    GreatVolume = Cells(2, 12).Value
    
    Dim rng_percent As Range
    Dim rng_volume As Range
    
    'this part about setting ranges was from stackoverflow threads concerning min/max with lookup/match functions
    
    Set rng_percent = Range("K1", "K" & format_count)
    Set rng_volume = Range("L1", "L" & format_count)
    
    GreatHigh = WorksheetFunction.Max(rng_percent)
    GreatHighTic = Cells(WorksheetFunction.Match(GreatHigh, rng_percent.Value2, 0), 9).Value
    
    GreatLow = WorksheetFunction.Min(rng_percent)
    GreatLowTic = Cells(WorksheetFunction.Match(GreatLow, rng_percent.Value2, 0), 9).Value
    
    GreatVolume = WorksheetFunction.Max(rng_volume)
    GreatVolumeTic = Cells(WorksheetFunction.Match(GreatVolume, rng_volume.Value2, 0), 9).Value
    
    
    For j = 2 To format_count
    'original code when the second analysis was an if/then based iteration
    'it was during this section that I tried to optimize CPU runtime and be more efficient...i do not feel as seccessful as i should have been.
    
'        If Cells(j, 11).Value > GreatHigh Then
'            GreatHigh = Cells(j, 11).Value
'            GreatHighTic = Cells(j, 9).Value
'        End If
'
'        If Cells(j, 11).Value < GreatLow Then
'            neglow = Cells(j, 11) * (-1)
'            If Abs(GreatLow) > neglow Then
'                GreatLow = Cells(j, 11).Value
'                GreatLowTic = Cells(j, 9).Value
'            End If
'        End If
'
'        If Cells(j, 12).Value > GreatVolume Then
'            GreatVolume = Cells(j, 12).Value
'            GreatVolumeTic = Cells(j, 9).Value
'        End If
        
        If Cells(j, 10).Value > 0 Then
            Cells(j, 10).Interior.Color = 5296274
        Else:
            Cells(j, 10).Interior.Color = 255
        End If

        Range("P2").Value = GreatHighTic
        Range("P3").Value = GreatLowTic
        Range("P4").Value = GreatVolumeTic
        
        Range("Q2").Value = GreatHigh
        Range("Q3").Value = GreatLow
        Range("Q4").Value = GreatVolume
    
    Next j

    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "#,##0"
    Range("O:O,Q:Q").EntireColumn.AutoFit
                      
        
    End With
   
    ActiveCell.Range("A1").Select
   Next n

End Sub




'---------------------------------------------------------------------------------------

'reference code from Alphabet practice



'Sub stockanalysis()
'
'Dim ws As Worksheet
'Dim n As Integer
'Dim WScount As Integer
'
'Dim ticker, GreatHighTic, GreatLowTic, GreatVolumeTic As String
'Dim Tickercount, format_count As Long
'
'Dim OpeningPrice, ClosingPrice, YearChange, percentchange As Double
'Dim GreatHigh, GreatLow, neglow As Long
'Dim volume, GreatVolume As Variant
'Dim i, j As Long
'Dim Lastrow As Long
'
'
'
'WScount = ActiveWorkbook.Worksheets.Count ' support.mocrosoft.com: macro to loop through all worksheets in a workbook
'
'For n = 1 To WScount
'
'    With Worksheets(n).Activate
'
'
'    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row 'provided from -NC on Slack
'    Tickercount = 0
'    volume = 0
'
'        With Range("I1:Q1")
'            .Font.FontStyle = "Bold"
'        End With
'
'        With Range("O1:O4")
'            .Font.FontStyle = "Bold"
'        End With
'
'        Range("I1").Value = "Ticker"
'        Range("J1").Value = "Yearly Change"
'        Range("K1").Value = "Percent Change"
'        Range("L1").Value = "Yearly Volume Traded"
'        Range("o2").Value = "Greatest % Increase"
'        Range("o3").Value = "Greatest % Decrease"
'        Range("o4").Value = "Greatest Total Volume"
'        Range("P1").Value = "Ticker"
'        Range("Q1").Value = "Value"
'
'        For i = 2 To Lastrow
'
'            ticker = Cells(i, 1).Value
'            volume = volume + Cells(i, 7).Value
'
'            If Cells(i - 1, 1).Value <> ticker Then
'                OpeningPrice = Cells(i, 3).Value
'
'                ElseIf Cells(i + 1, 1).Value <> ticker Then
'                    ClosingPrice = Cells(i, 6).Value
'
'                    Tickercount = Tickercount + 1
'
'                    YearChange = ClosingPrice - OpeningPrice
'                    percentchange = YearChange / OpeningPrice
'
'                    Range("I" & Tickercount + 1).Value = ticker
'                    Range("J" & Tickercount + 1).Value = YearChange
'                    Range("K" & Tickercount + 1).Value = percentchange
'                    Range("L" & Tickercount + 1).Value = volume
'
'
'                    Range("K" & Tickercount + 1).NumberFormat = "0.00%"
'                    Range("J" & Tickercount + 1).NumberFormat = "$#,##0.00"
'                    Range("L" & Tickercount + 1).NumberFormat = "#,##0"
'                    Range("I:I,L:L").EntireColumn.AutoFit
'
'                    volume = 0
'                    YearChange = 0
'                    percentchange = 0
'            End If
'
'        Next i
'
'   With Range("I1:L1")
'        .Font.FontStyle = "Bold"
'        .EntireColumn.AutoFit
'    End With
'
'    format_count = Tickercount + 1
'
'    GreatHigh = Cells(2, 11).Value
'    GreatLow = Cells(2, 11).Value
'    GreatVolume = Cells(2, 12).Value
'
'
'    For j = 2 To format_count
'        If Cells(j, 11).Value > GreatHigh Then
'            GreatHigh = Cells(j, 11).Value
'            GreatHighTic = Cells(j, 9).Value
'        End If
'
'        If Cells(j, 11).Value < GreatLow Then
'            neglow = Cells(j, 11) * (-1)
'            If Abs(GreatLow) > neglow Then
'                GreatLow = Cells(j, 11).Value
'                GreatLowTic = Cells(j, 9).Value
'            End If
'        End If
'
'        If Cells(j, 12).Value > GreatVolume Then
'            GreatVolume = Cells(j, 12).Value
'            GreatVolumeTic = Cells(j, 9).Value
'        End If
'
'        If Cells(j, 10).Value > 0 Then
'            Cells(j, 10).Interior.Color = 5296274
'        Else:
'            Cells(j, 10).Interior.Color = 255
'        End If
'
'        Range("P2").Value = GreatHighTic
'        Range("P3").Value = GreatLowTic
'        Range("P4").Value = GreatVolumeTic
'
'        Range("Q2").Value = GreatHigh
'        Range("Q3").Value = GreatLow
'        Range("Q4").Value = GreatVolume
'
'    Next j
'
'    Range("Q2").NumberFormat = "0.00%"
'    Range("Q3").NumberFormat = "0.00%"
'    Range("Q4").NumberFormat = "#,##0"
'    Range("O:O,Q:Q").EntireColumn.AutoFit
'
'
'    End With
'
'    ActiveCell.Range("A1").Select
'   Next n
'
'End Sub
'
'
'
