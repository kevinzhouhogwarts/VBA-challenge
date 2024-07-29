Attribute VB_Name = "Module1"
Dim LastDataColumn As Long 'Declare module-level variable for the rightmost data column in the worksheet

'Use this RunAllWorksheets subroutine to run all worksheets in the workbook. To run individually, skip RunAllWorksheets and proceed to the three individual subroutines.
Sub RunAllWorksheets()
    Dim ws As Worksheet
    For Each ws In Worksheets 'Loop through all worksheets in the active workbook
        'Instead of calling ws.Activate here, which seems to be discouraged in VBA, instead parameterize the called subroutines
        Call IdentifyLastDataColumn(ws)
        Call InsertHeader(ws)
        Call Calculate(ws)
    Next ws
End Sub

Sub IdentifyLastDataColumn(ws As Worksheet)
'Loop through the cells of the first row in order to identify the last column with a non-blank value. This could also be done using the End function
    Dim Col As Long
    For Col = 1 To Columns.Count 'Loop from first column to last column
        If IsEmpty(ws.Cells(1, Col)) Then 'Test if the first cell in the column is empty
            LastDataColumn = Col - 1 'If true, then initialize the variable
            Exit For 'If true, end the for loop before the iteration finishes
        End If 'Close if statement
    Next Col 'Close for loop
End Sub

Sub InsertHeader(ws As Worksheet)
'Create a header for the output information section and locate it to the second column after the last column
    Dim header(2 To 5) As String 'Declare array, indexed to match the below counter
    header(2) = "Ticker" 'Assign the header values of the output to the array
    header(3) = "Quarterly Change"
    header(4) = "Percent Change"
    header(5) = "Total Stock Volume"
    Dim i As Long
    For i = 2 To 5 'Loop through cells in the first row to populate with header values
        ws.Cells(1, LastDataColumn + i).Value = header(i)
    Next i
End Sub

Sub Calculate(ws As Worksheet)
    'Use collections to store output data, since - unlike arrays - collections resize automatically while preserving existing values. Use New to avoid having to use Set later, since these collections will definitely be populated.
    Dim TickerCollection As New Collection
    Dim OpenPriceCollection As New Collection
    Dim ClosePriceCollection As New Collection
    Dim TotalVolumeCollection As New Collection
    
    Dim CurrentTotalVolume As Double 'Declare variable to accumulate volume while looping
    CurrentTotalVolume = 0
    
    Dim LastDataRow As Long 'Variable to store last row of the raw data
    LastDataRow = ws.Range("A1").End(xlDown).row  'From the top cell in the raw data ticker column, move down to the last data row

    Dim CurrentTicker As String
    CurrentTicker = ws.Range("A2").Value 'Initialize using the first ticker symbol that appears in the raw data
    TickerCollection.Add CurrentTicker 'Append the first ticker symbol to collection
    OpenPriceCollection.Add ws.Cells(2, 3).Value 'Append the first open price to collection
    
    Dim i As Long
    For i = 2 To LastDataRow
        If ws.Cells(i, 1).Value = CurrentTicker Then 'If the ticker matches...
            CurrentTotalVolume = CurrentTotalVolume + ws.Cells(i, 7).Value '...then accumulate the row volume toward the total volume
        Else
            TotalVolumeCollection.Add CurrentTotalVolume 'Append total calculated volume to collection
            ClosePriceCollection.Add ws.Cells(i - 1, 6).Value 'Append closing price of previous row to collection
            'This completes the output information for the previous ticker. From below, start outputting information for the new ticker
            CurrentTicker = ws.Cells(i, 1).Value 'Update the current ticker
            CurrentTotalVolume = ws.Cells(i, 7).Value 'Reset and update the variable tracking the running total volume
            TickerCollection.Add CurrentTicker 'Append the new current ticker to collection
            OpenPriceCollection.Add ws.Cells(i, 3).Value 'Append the current open price to collection
        End If
    Next i
    
    'Account for the behavior at the end of the loop
    TotalVolumeCollection.Add CurrentTotalVolume
    ClosePriceCollection.Add ws.Cells(LastDataRow, 6).Value
    
    'Assign the output data stored in collections to the cells beneath the headers created in the InsertHerder Sub
    Dim j As Long
    For j = 1 To TickerCollection.Count 'Loop through the total count inside the collections
        ws.Cells(j + 1, LastDataColumn + 2).Value = TickerCollection(j)
        ws.Cells(j + 1, LastDataColumn + 3).Value = ClosePriceCollection(j) - OpenPriceCollection(j)
        ws.Cells(j + 1, LastDataColumn + 4).Value = ws.Cells(j + 1, LastDataColumn + 3).Value / OpenPriceCollection(j)
        ws.Cells(j + 1, LastDataColumn + 5).Value = TotalVolumeCollection(j)
    Next j
    ws.Columns(LastDataColumn + 4).NumberFormat = "0.00%" 'Format the Percent Change column as percent
    
    'Calculate the extreme outputs. Apparently you can use Application.WorksheetFunction to access Min and Max for Ranges, but let's use a for loop instead.
    Dim GreatestIncreaseTicker As String 'Declare variables for save the extreme value outputs
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestDecreaseTicker As String
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim GreatestVolumeTicker As String
    Dim GreatestVolume As Double
    GreatestVolume = 0
    
    Dim k As Integer
    For k = 1 To TickerCollection.Count
        If ws.Cells(k + 1, LastDataColumn + 4).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(k + 1, LastDataColumn + 4).Value
            GreatestIncreaseTicker = ws.Cells(k + 1, LastDataColumn + 2).Value
        End If
        If ws.Cells(k + 1, LastDataColumn + 4).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(k + 1, LastDataColumn + 4).Value
            GreatestDecreaseTicker = ws.Cells(k + 1, LastDataColumn + 2).Value
        End If
        If ws.Cells(k + 1, LastDataColumn + 5).Value > GreatestVolume Then
            GreatestVolume = ws.Cells(k + 1, LastDataColumn + 5).Value
            GreatestVolumeTicker = ws.Cells(k + 1, LastDataColumn + 2).Value
        End If
    Next k
    
    'Assign the outputted extreme values to an array
    Dim ExtremesArray(1 To 4, 14 To 16) As Variant
    ExtremesArray(1, 14) = ""
    ExtremesArray(1, 15) = "Ticker"
    ExtremesArray(1, 16) = "Value"
    ExtremesArray(2, 14) = "Greatest % Increase"
    ExtremesArray(2, 15) = GreatestIncreaseTicker
    ExtremesArray(2, 16) = GreatestIncrease
    ExtremesArray(3, 14) = "Greatest % Decrease"
    ExtremesArray(3, 15) = GreatestDecreaseTicker
    ExtremesArray(3, 16) = GreatestDecrease
    ExtremesArray(4, 14) = "Greatest Total Volume"
    ExtremesArray(4, 15) = GreatestVolumeTicker
    ExtremesArray(4, 16) = GreatestVolume
    'Populate the sheet with the array values
    Dim row As Long
    Dim column As Long
    For row = 1 To 4
        For column = 14 To 16
            ws.Cells(row, column).Value = ExtremesArray(row, column)
        Next column
    Next row
    ws.Range("P2:P3").NumberFormat = "0.00%" 'Format the Percent Change values as percent
End Sub

'Lesson learned: It would have been slightly more concise to directly calculate the difference between final closing and initial opening price in the calculation loop, and pass that to a collection.
'Lesson learned: Since many of the cell references are hard-coded, it was not useful to calculate the LastDataColumn


