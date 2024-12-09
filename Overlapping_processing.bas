Attribute VB_Name = "Overlapping_processing"
Option Explicit

Sub ConsolidateOutages()
    ' Define worksheets
    Dim wsDump As Worksheet, wsResult As Worksheet, wsMenu As Worksheet
    Set wsDump = ThisWorkbook.Sheets("Passive")
    Set wsResult = ThisWorkbook.Sheets("Result")
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    
    ' Get column letters from Menu sheet
    Dim siteIDColumn As String, startTimeColumn As String, endTimeColumn As String, lastColumn As String, durationColumn As String
    siteIDColumn = wsMenu.Range("L15").Value
    startTimeColumn = wsMenu.Range("L23").Value
    endTimeColumn = wsMenu.Range("L24").Value
    lastColumn = wsMenu.Range("L16").Value
    durationColumn = wsMenu.Range("L26").Value
    
    ' Convert column letters to numbers
    Dim siteIDColNum As Long, startTimeColNum As Long, endTimeColNum As Long, lastColNum As Long, durationColumnNum As Long
    siteIDColNum = wsDump.Columns(siteIDColumn).Column
    startTimeColNum = wsDump.Columns(startTimeColumn).Column
    endTimeColNum = wsDump.Columns(endTimeColumn).Column
    lastColNum = wsDump.Columns(lastColumn).Column
    durationColumnNum = wsDump.Columns(durationColumn).Column
    
    ' Define variables
    Dim lastRow As Long, resultRow As Long
    Dim siteName As String, currentStart As Date, currentEnd As Date
    Dim i As Long, j As Long
    Dim dataDump As Variant, dataResult As Variant, resultIndex As Long
    Dim siteDict As Object
    Set siteDict = CreateObject("Scripting.Dictionary")
    
    ' Clear Result sheet
    wsResult.Cells.Clear
    
    ' Sort the data by Site Name and Start Time
    lastRow = wsDump.Cells(wsDump.rows.Count, siteIDColNum).End(xlUp).Row
    wsDump.Sort.SortFields.Clear
    wsDump.Sort.SortFields.Add key:=wsDump.Range(wsDump.Cells(2, siteIDColNum), wsDump.Cells(lastRow, siteIDColNum)), Order:=xlAscending
    wsDump.Sort.SortFields.Add key:=wsDump.Range(wsDump.Cells(2, startTimeColNum), wsDump.Cells(lastRow, startTimeColNum)), Order:=xlAscending
    With wsDump.Sort
        .SetRange wsDump.Range(wsDump.Cells(1, 1), wsDump.Cells(lastRow, lastColNum))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    ' Load data into an array to speed up processing
    dataDump = wsDump.Range(wsDump.Cells(1, 1), wsDump.Cells(lastRow, lastColNum)).Value
    ReDim dataResult(1 To UBound(dataDump, 1), 1 To UBound(dataDump, 2))
    
    ' Write headers to Result data array
    For i = LBound(dataDump, 2) To UBound(dataDump, 2)
        dataResult(1, i) = dataDump(1, i)
    Next i
    
    ' Initialize resultIndex for Result data array
    resultIndex = 2
    
    ' Loop through the data to consolidate outages using dictionary
    For i = 2 To UBound(dataDump, 1)
        If Not siteDict.exists(dataDump(i, siteIDColNum)) Then
            siteDict.Add dataDump(i, siteIDColNum), resultIndex
            siteName = dataDump(i, siteIDColNum)
            currentStart = dataDump(i, startTimeColNum)
            currentEnd = dataDump(i, endTimeColNum)
            ' Copy row to result array
            For j = LBound(dataDump, 2) To UBound(dataDump, 2)
                dataResult(resultIndex, j) = dataDump(i, j)
            Next j
            dataResult(resultIndex, startTimeColNum) = currentStart
            dataResult(resultIndex, endTimeColNum) = currentEnd
            ' Calculate duration
            dataResult(resultIndex, durationColumnNum) = currentEnd - currentStart
            resultIndex = resultIndex + 1
        Else
            ' Extend existing site entry if needed
            Dim existingIndex As Long
            existingIndex = siteDict(dataDump(i, siteIDColNum))
            If dataDump(i, startTimeColNum) <= dataResult(existingIndex, endTimeColNum) Then
                If dataDump(i, endTimeColNum) > dataResult(existingIndex, endTimeColNum) Then
                    dataResult(existingIndex, endTimeColNum) = dataDump(i, endTimeColNum)
                    ' Recalculate duration for extended time
                    dataResult(existingIndex, durationColumnNum) = dataResult(existingIndex, endTimeColNum) - dataResult(existingIndex, startTimeColNum)
                End If
            Else
                ' Create a new entry for non-overlapping time
                siteDict(dataDump(i, siteIDColNum)) = resultIndex
                siteName = dataDump(i, siteIDColNum)
                currentStart = dataDump(i, startTimeColNum)
                currentEnd = dataDump(i, endTimeColNum)
                ' Copy row to result array
                For j = LBound(dataDump, 2) To UBound(dataDump, 2)
                    dataResult(resultIndex, j) = dataDump(i, j)
                Next j
                dataResult(resultIndex, startTimeColNum) = currentStart
                dataResult(resultIndex, endTimeColNum) = currentEnd
                ' Calculate duration
                dataResult(resultIndex, durationColumnNum) = currentEnd - currentStart
                resultIndex = resultIndex + 1
            End If
        End If
    Next i
    
    ' Write the consolidated data back to the Result sheet
    wsResult.Range(wsResult.Cells(1, 1), wsResult.Cells(resultIndex - 1, UBound(dataDump, 2))).Value = dataResult
    
    ' Ensure that start and end times are written with full date and time
    wsResult.Columns(startTimeColNum).NumberFormat = "mm/dd/yyyy h:mm"
    wsResult.Columns(endTimeColNum).NumberFormat = "mm/dd/yyyy h:mm"
    wsResult.Columns(durationColumnNum).NumberFormat = "[h]:mm:ss"
    
    ' Wrap and unwrap the Result sheet
    With wsResult.Cells
        .WrapText = False
    End With
    
  '  wsResult.Activate
    
End Sub
