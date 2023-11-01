Sub Vba_ChallengePt1()
    Dim ws          As Worksheet
    Dim LastRow     As Long
    Dim TickerCount As Integer
    Dim SumVolume   As Double
    Dim Openn       As Double
    Dim Closee      As Double

    For Each ws In Worksheets
        TickerCount = 0
        SumVolume = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Set headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through the rows of data in the current worksheet
        For a = 2 To LastRow
            If ws.Cells(a, 1).Value <> ws.Cells(a - 1, 1).Value Then
                TickerCount = TickerCount + 1
                Openn = ws.Cells(a, 3).Value
                ws.Cells(1 + TickerCount, 9).Value = ws.Cells(a, 1).Value
            End If
            
           ' Accumulate the trading volume for the current ticker
            SumVolume = SumVolume + ws.Cells(a, 7).Value

            If ws.Cells(a, 1).Value <> ws.Cells(a + 1, 1).Value Then
                Closee = ws.Cells(a, 6).Value
                ws.Cells(1 + TickerCount, 12).Value = SumVolume
                SumVolume = 0
           ' Calculate the yearly change and percent change
                Dim Change As Double
                Change = Closee - Openn
                ws.Cells(1 + TickerCount, 10).Value = Change
           
           ' Calculate percent change
                If Openn = 0 And Closee = 0 Then
                    ws.Cells(1 + TickerCount, 11).Value = 0
                ElseIf Openn = 0 And Closee <> 0 Then
                    ws.Cells(1 + TickerCount, 11).Value = 1
                Else
                    ws.Cells(1 + TickerCount, 11).Value = Change / Openn
                End If
            ' Format the percent change as a percentage
                ws.Cells(1 + TickerCount, 11).NumberFormat = "#0.00%"

                If ws.Cells(1 + TickerCount, 11).Value > 0 Then
                    ws.Cells(1 + TickerCount, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(1 + TickerCount, 11).Value < 0 Then
                    ws.Cells(1 + TickerCount, 11).Interior.ColorIndex = 3
                End If
            End If
        Next a
    Next ws

' After processing all worksheets, call pt2
Call Vba_ChallegePt2
End Sub

Sub Vba_ChallengePt2()
    Dim ws              As Worksheet
    Dim LastRow         As Long
    Dim MaxPercent      As Double
    Dim MaxTicker       As String
    Dim MinPercent      As Double
    Dim MinTicker       As String
    Dim MaxVolume       As Double
    Dim MaxVolumeTicker As String
    
    For Each ws In Worksheets
        ' create headers
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ' Find the last row in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        MaxPercent = 0
        MinPercent = 0
        MaxVolume = 0
        
        For r = 2 To LastRow
         
         'Find max percent increase
            If ws.Cells(r, 11).Value > MaxPercent Then
                MaxPercent = ws.Cells(r, 11).Value
                MaxTicker = ws.Cells(r, 9).Value
            End If
            
         'Find min percent decrease
            If ws.Cells(r, 11).Value < MinPercent Then
                MinPercent = ws.Cells(r, 11).Value
                MinTicker = ws.Cells(r, 9).Value
            End If
            
         'Findn max total volume
            If ws.Cells(r, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(r, 12).Value
                MaxVolumeTicker = ws.Cells(r, 9).Value
            End If
        Next r
        
        ' print max/min ticker and perencentages
        ws.Cells(2, 16).Value = MaxPercent
        ws.Cells(2, 15).Value = MaxTicker
        ws.Cells(3, 16).Value = MinPercent
        ws.Cells(3, 15).Value = MinTicker
        ws.Cells(4, 16).Value = MaxVolume
        ws.Cells(4, 15).Value = MaxVolumeTicker
        
        ' format max/min as percentages
        ws.Cells(2, 16).NumberFormat = "#0.00%"
        ws.Cells(3, 16).NumberFormat = "#0.00%"
        
    Next ws
End Sub


