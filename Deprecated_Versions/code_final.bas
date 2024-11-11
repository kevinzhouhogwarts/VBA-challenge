Attribute VB_Name = "Module1"
'Differences from wordy version: _
Assume all target worksheets will contain data with a header row and 7 columns, so the subroutine to identify the last column holding raw data is removed. _
Some comments have been abbreviated and others removed for cleanliness. _
InsertHeader() adds the header titles for both output sections, the per-ticker calculation and the extreme value calculation. _
Instead of using a loop to assign array values to cells, assign the array directly to a hard-coded range. _
Use one subroutine MainSubroutine for all the calculation code. _
Important change: Instead of creating collections for the initial open price and the final close price, directly create a collection that stores the difference, i.e. the quarterly change.

'Run this subroutine to run the component subroutines for all worksheets in the workbook.
Sub RunAllWorksheets()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Call MainSubroutine(ws)
    Next ws
End Sub

Sub MainSubroutine(ws As Worksheet)
'Insert header titles
    Dim HeaderArray(9 To 16) As String
    HeaderArray(9) = "Ticker"
    HeaderArray(10) = "Quarterly Change"
    HeaderArray(11) = "Percent Change"
    HeaderArray(12) = "Total Stock Volume"
    HeaderArray(13) = ""
    HeaderArray(14) = ""
    HeaderArray(15) = "Ticker"
    HeaderArray(16) = "Value"
    ws.Range("I1:P1") = HeaderArray
    
'Begin main loop to calculate the per-ticker outputs
    Dim Ticker As New Collection 'Declare collections to store output data
    Dim QuarterlyChange As New Collection
    Dim PercentChange As New Collection
    Dim TotalVolume As New Collection
    
    Dim LastDataRow As Long 'Declare variable for last row of raw data
    LastDataRow = ws.Range("A1").End(xlDown).row
    
    Dim VolumeTracker As Double 'Declare and initialize helper variables to use within the loop
    VolumeTracker = 0
    Dim TickerTracker As String
    TickerTracker = ws.Range("A2").Value
    Ticker.Add TickerTracker
    Dim InitialOpenPrice As Double
    InitialOpenPrice = ws.Cells(2, 3).Value
    Dim FinalClosePrice As Double
    
    Dim i As Long
    For i = 2 To LastDataRow
        If ws.Cells(i, 1).Value = TickerTracker Then
            VolumeTracker = VolumeTracker + ws.Cells(i, 7).Value
        Else
            FinalClosePrice = ws.Cells(i - 1, 6).Value
            QuarterlyChange.Add (FinalClosePrice - InitialOpenPrice)
            PercentChange.Add ((FinalClosePrice - InitialOpenPrice) / InitialOpenPrice)
            TotalVolume.Add VolumeTracker
            TickerTracker = ws.Cells(i, 1).Value 'Begin preparation for next iteration of the loop
            Ticker.Add TickerTracker
            VolumeTracker = ws.Cells(i, 7).Value
            InitialOpenPrice = ws.Cells(i, 3).Value
        End If
    Next i
    
    FinalClosePrice = ws.Cells(LastDataRow, 6).Value 'Account for the behavior at the end of the loop
    QuarterlyChange.Add (FinalClosePrice - InitialOpenPrice)
    PercentChange.Add ((FinalClosePrice - InitialOpenPrice) / InitialOpenPrice)
    TotalVolume.Add VolumeTracker
    
'Populate worksheet with the output data
    Dim j As Long
    For j = 1 To Ticker.Count
        ws.Cells(j + 1, 9).Value = Ticker(j)
        ws.Cells(j + 1, 10).Value = QuarterlyChange(j)
        ws.Cells(j + 1, 11).Value = PercentChange(j)
        ws.Cells(j + 1, 12).Value = TotalVolume(j)
    Next j
    ws.Columns(11).NumberFormat = "0.00%"
    
'Calculate the extreme outputs.
    Dim GreatestIncreaseTicker As String 'Declare variables for save the extreme value outputs
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestDecreaseTicker As String
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim GreatestVolumeTicker As String
    Dim GreatestVolume As Double
    GreatestVolume = 0
    
    Dim k As Long 'Loop through output data to calculate extremes
    For k = 1 To Ticker.Count
        If ws.Cells(k + 1, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(k + 1, 11).Value
            GreatestIncreaseTicker = ws.Cells(k + 1, 9).Value
        End If
        If ws.Cells(k + 1, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(k + 1, 11).Value
            GreatestDecreaseTicker = ws.Cells(k + 1, 9).Value
        End If
        If ws.Cells(k + 1, 12).Value > GreatestVolume Then
            GreatestVolume = ws.Cells(k + 1, 12).Value
            GreatestVolumeTicker = ws.Cells(k + 1, 9).Value
        End If
    Next k
    
    Dim ExtremesArray(2 To 4, 14 To 16) As Variant 'Assign the extreme values to an array
    ExtremesArray(2, 14) = "Greatest % Increase"
    ExtremesArray(2, 15) = GreatestIncreaseTicker
    ExtremesArray(2, 16) = GreatestIncrease
    ExtremesArray(3, 14) = "Greatest % Decrease"
    ExtremesArray(3, 15) = GreatestDecreaseTicker
    ExtremesArray(3, 16) = GreatestDecrease
    ExtremesArray(4, 14) = "Greatest Volume"
    ExtremesArray(4, 15) = GreatestVolumeTicker
    ExtremesArray(4, 16) = GreatestVolume

'Populate the sheet with the extreme values
    Dim row As Long
    Dim column As Long
    For row = 2 To 4
        For column = 14 To 16
            ws.Cells(row, column).Value = ExtremesArray(row, column)
        Next column
    Next row
    ws.Range("P2:P3").NumberFormat = "0.00%"
End Sub


