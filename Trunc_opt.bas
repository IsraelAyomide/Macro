Attribute VB_Name = "Trunc_opt"
Option Explicit
Public Sub TruncateOutagesOpt()
    Dim wsSrc As Worksheet, wsDst As Worksheet, wsMenu As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim startTime As Date, endTime As Date
    Dim currentStart As Date, currentEnd As Date
    Dim startCol As String, endCol As String
    Dim startColNum As Long, endColNum As Long
    Dim sourceData() As Variant
    Dim resultArray() As Variant
    Dim totalSourceDuration As Double
    Dim totalNewDuration As Double
    Dim rows As Collection
    Dim tempRow() As Variant

    ' Set references to worksheets
    Set wsSrc = ThisWorkbook.Sheets("Result")
    Set wsDst = ThisWorkbook.Sheets("Truncated")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    
    wsDst.Cells.Clear

    ' Get column references from MENU sheet
    startCol = wsMenu.Range("L23").Value
    endCol = wsMenu.Range("L24").Value

    ' Convert column letters to numbers
    startColNum = wsSrc.Range(startCol & "1").Column
    endColNum = wsSrc.Range(endCol & "1").Column

    ' Performance optimizations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Get last row of source data
    lastRow = wsSrc.Cells(wsSrc.rows.Count, "A").End(xlUp).Row

    ' Load source data into array
    sourceData = wsSrc.Range("A1:AH" & lastRow).Value

    ' Initialize collection
    Set rows = New Collection

    ' Calculate total source duration for validation (in seconds)
    totalSourceDuration = 0
    For i = 2 To UBound(sourceData, 1)
        totalSourceDuration = totalSourceDuration + DateDiff("s", sourceData(i, startColNum), sourceData(i, endColNum))
    Next i

    ' Copy headers to first row
    ReDim tempRow(1 To UBound(sourceData, 2) + 1)
    For j = 1 To UBound(sourceData, 2)
        tempRow(j) = sourceData(1, j)
    Next j
    tempRow(UBound(sourceData, 2) + 1) = " Final Duration"
    rows.Add tempRow

    ' Process each row
    For i = 2 To UBound(sourceData, 1)
        startTime = sourceData(i, startColNum)
        endTime = sourceData(i, endColNum)
        currentStart = startTime

        Do While currentStart < endTime  ' Changed from <= to < to prevent zero-duration entries
            ' Calculate next period end
            currentEnd = DateAdd("h", 1, currentStart)
            currentEnd = DateAdd("n", -Minute(currentEnd), currentEnd)
            currentEnd = DateAdd("s", -Second(currentEnd), currentEnd)

            ' If next period would exceed end time, use end time instead
            If currentEnd > endTime Then
                currentEnd = endTime
            End If

            ' Only create row if there is a duration
            If DateDiff("s", currentStart, currentEnd) > 0 Then  ' Added explicit duration check
                ReDim tempRow(1 To UBound(sourceData, 2) + 1)

                ' Copy row data
                For j = 1 To UBound(sourceData, 2)
                    tempRow(j) = sourceData(i, j)
                Next j

                ' Set period start and end times
                tempRow(startColNum) = currentStart
                tempRow(endColNum) = currentEnd

                ' Calculate duration directly as end - start
                tempRow(UBound(sourceData, 2) + 1) = Format(currentEnd - currentStart, "hh:mm:ss")

                rows.Add tempRow
            End If

            currentStart = currentEnd
        Loop
    Next i

    ' Convert collection to array
    If rows.Count > 0 Then
        ReDim resultArray(1 To rows.Count, 1 To UBound(sourceData, 2) + 1)
        For i = 1 To rows.Count
            For j = 1 To UBound(sourceData, 2) + 1
                resultArray(i, j) = rows(i)(j)
            Next j
        Next i

        ' Clear destination sheet
        wsDst.Cells.Clear

        ' Write results to sheet
        With wsDst.Range("A1").Resize(UBound(resultArray, 1), UBound(resultArray, 2))
            .Value = resultArray
        End With

        ' Set number format for date/time columns
        wsDst.Columns(startCol & ":" & endCol).NumberFormat = "mm/dd/yyyy h:mm:ss"
        wsDst.Columns(UBound(sourceData, 2) + 1).NumberFormat = "hh:mm:ss"

        ' Calculate total new duration
        totalNewDuration = 0
        For i = 2 To UBound(resultArray, 1)
            totalNewDuration = totalNewDuration + DateDiff("s", resultArray(i, startColNum), resultArray(i, endColNum))
        Next i

        ' Validate durations
        If totalSourceDuration <> totalNewDuration Then
            MsgBox "Warning: Total duration before and after truncation do not match. Please review the data.", vbExclamation
        End If
    Else
        MsgBox "No data to process.", vbExclamation
    End If

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call PopulateStartHour
    Call CalculateAvailability

    ' Disable text wrapping
    With wsDst.Cells
        .WrapText = False
    End With
End Sub

Sub PopulateStartHour()
    Dim ws As Worksheet
    Dim wsMenu As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startTimeCol As String
    Dim startTimeColNum As Long
    Dim hourStartCol As Long
    Dim dataArray As Variant
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("Truncated")
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    
    ' Get the start time column from Menu sheet
    startTimeCol = wsMenu.Range("L23").Value
    startTimeColNum = ws.Columns(startTimeCol).Column
    
    ' Get the last row of data
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    ' Set the Hour Start column
    hourStartCol = 36
    
    ' Read all data into an array at once
    dataArray = ws.Range(ws.Cells(1, startTimeColNum), ws.Cells(lastRow, startTimeColNum)).Value
    
    ' Create results array
    Dim resultsArray As Variant
    ReDim resultsArray(1 To lastRow, 1 To 1)
    
    ' Set header
    resultsArray(1, 1) = "Hour Start"
    
    ' Process the data in memory
    For i = 2 To lastRow
        If Not IsEmpty(dataArray(i, 1)) Then
            resultsArray(i, 1) = Format(dataArray(i, 1), "hh:00")
        End If
    Next i
    
    ' Write results back to worksheet in one operation
    ws.Range(ws.Cells(1, hourStartCol), ws.Cells(lastRow, hourStartCol)).Value = resultsArray
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    
End Sub
Sub CalculateAvailability()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim durationArray As Variant
    Dim resultsArray As Variant
    Dim i As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Truncated")
    
    ' Get the last row of data
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    ' Read duration data into array (Column AI)
    durationArray = ws.Range(ws.Cells(1, 35), ws.Cells(lastRow, 35)).Value
    
    ' Create results array
    ReDim resultsArray(1 To lastRow, 1 To 1)
    
    ' Set header
    resultsArray(1, 1) = "Availability"
    
    ' Calculate availability percentage
    For i = 2 To lastRow
        If Not IsEmpty(durationArray(i, 1)) Then
            ' Convert duration directly to decimal hours and calculate percentage
            resultsArray(i, 1) = (1 - (durationArray(i, 1) * 24)) * 100
        End If
    Next i
    
    ' Write results to column AK (37)
    ws.Range(ws.Cells(1, 37), ws.Cells(lastRow, 37)).Value = resultsArray
    
    ' Format the result column to show 2 decimal places
    ws.Range(ws.Cells(2, 37), ws.Cells(lastRow, 37)).NumberFormat = "0.00"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
Public Sub TruncateOutagesOpt2222()
    Dim wsSrc As Worksheet, wsDst As Worksheet, wsMenu As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim startTime As Date, endTime As Date
    Dim currentStart As Date, currentEnd As Date
    Dim startCol As String, endCol As String
    Dim startColNum As Long, endColNum As Long
    Dim sourceData() As Variant
    Dim resultArray() As Variant
    Dim totalSourceDuration As Double
    Dim totalNewDuration As Double
    Dim rows As Collection
    Dim tempRow() As Variant

    ' Set references to worksheets
    Set wsSrc = ThisWorkbook.Sheets("Result")
    Set wsDst = ThisWorkbook.Sheets("Truncated")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    
    wsDst.Cells.Clear

    ' Get column references from MENU sheet
    startCol = wsMenu.Range("L23").Value
    endCol = wsMenu.Range("L24").Value

    ' Convert column letters to numbers
    startColNum = wsSrc.Range(startCol & "1").Column
    endColNum = wsSrc.Range(endCol & "1").Column

    ' Performance optimizations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Get last row of source data
    lastRow = wsSrc.Cells(wsSrc.rows.Count, "A").End(xlUp).Row

    ' Load source data into array
    sourceData = wsSrc.Range("A1:AH" & lastRow).Value

    ' Initialize collection
    Set rows = New Collection

    ' Calculate total source duration for validation (in seconds)
    totalSourceDuration = 0
    For i = 2 To UBound(sourceData, 1)
        totalSourceDuration = totalSourceDuration + DateDiff("s", sourceData(i, startColNum), sourceData(i, endColNum))
    Next i

    ' Copy headers to first row
    ReDim tempRow(1 To UBound(sourceData, 2) + 1)
    For j = 1 To UBound(sourceData, 2)
        tempRow(j) = sourceData(1, j)
    Next j
    tempRow(UBound(sourceData, 2) + 1) = " Final Duration"
    rows.Add tempRow

    ' Process each row
    For i = 2 To UBound(sourceData, 1)
        startTime = sourceData(i, startColNum)
        endTime = sourceData(i, endColNum)
        currentStart = startTime

        Do While currentStart < endTime  ' Changed from <= to < to prevent zero-duration entries
            ' Calculate next period end
            currentEnd = DateAdd("h", 1, currentStart)
            currentEnd = DateAdd("n", -Minute(currentEnd), currentEnd)
            currentEnd = DateAdd("s", -Second(currentEnd), currentEnd)

            ' If next period would exceed end time, use end time instead
            If currentEnd > endTime Then
                currentEnd = endTime
            End If

            ' Only create row if there is a duration
            If DateDiff("s", currentStart, currentEnd) > 0 Then  ' Added explicit duration check
                ReDim tempRow(1 To UBound(sourceData, 2) + 1)

                ' Copy row data
                For j = 1 To UBound(sourceData, 2)
                    tempRow(j) = sourceData(i, j)
                Next j

                ' Set period start and end times
                tempRow(startColNum) = currentStart
                tempRow(endColNum) = currentEnd

                ' Calculate duration directly as end - start
                tempRow(UBound(sourceData, 2) + 1) = Format(currentEnd - currentStart, "hh:mm:ss")

                rows.Add tempRow
            End If

            currentStart = currentEnd
        Loop
    Next i

    ' Convert collection to array
    If rows.Count > 0 Then
        ReDim resultArray(1 To rows.Count, 1 To UBound(sourceData, 2) + 1)
        For i = 1 To rows.Count
            For j = 1 To UBound(sourceData, 2) + 1
                resultArray(i, j) = rows(i)(j)
            Next j
        Next i

        ' Clear destination sheet
        wsDst.Cells.Clear

        ' Write results to sheet
        With wsDst.Range("A1").Resize(UBound(resultArray, 1), UBound(resultArray, 2))
            .Value = resultArray
        End With

        ' Set number format for date/time columns
        wsDst.Columns(startCol & ":" & endCol).NumberFormat = "mm/dd/yyyy h:mm:ss"
        wsDst.Columns(UBound(sourceData, 2) + 1).NumberFormat = "hh:mm:ss"

        ' Calculate total new duration
        totalNewDuration = 0
        For i = 2 To UBound(resultArray, 1)
            totalNewDuration = totalNewDuration + DateDiff("s", resultArray(i, startColNum), resultArray(i, endColNum))
        Next i

        ' Validate durations
        If totalSourceDuration <> totalNewDuration Then
            MsgBox "Warning: Total duration before and after truncation do not match. Please review the data.", vbExclamation
        End If
    Else
        MsgBox "No data to process.", vbExclamation
    End If

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Call PopulateStartHour
    Call CalculateAvailability

    ' Disable text wrapping
    With wsDst.Cells
        .WrapText = False
    End With
End Sub
