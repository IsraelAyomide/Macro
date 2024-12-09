Attribute VB_Name = "PA_Trend"
Option Explicit

Sub PopulateHourlyPATrend()
    Dim wsTruncated As Worksheet, wsPATrend As Worksheet, wsMenu As Worksheet
    Dim lastRowTruncated As Long, lastRowPATrend As Long
    Dim i As Long, col As Long
    Dim siteID As String
    Dim availabilityDict As Object, hourCountDict As Object
    Dim hourStart As Variant
    Dim availability As Double
    Dim siteIDColumn As Long
    Dim hourStartColumn As Long, availabilityColumn As Long
    Dim endTime As Date
    Dim hourLimit As Long
    Dim hourStartDict As Object
    Dim key As String
    Dim hourNum As Variant
    Dim fullDay As Boolean
    Dim dataRange As Range, cell As Range
    Dim resultArray() As Variant
    Dim cellRow As Long, cellCol As Long
    Dim siteStatus As String

    Set wsTruncated = ThisWorkbook.Sheets("Truncated")
    Set wsPATrend = ThisWorkbook.Sheets("PA Trend")
    Set wsMenu = ThisWorkbook.Sheets("MENU")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    siteIDColumn = wsTruncated.Range(wsMenu.Range("L15").Value & "1").Column
    hourStartColumn = 36
    availabilityColumn = 37

    endTime = wsMenu.Range("L14").Value
    fullDay = (Hour(endTime) = 0 And Minute(endTime) = 0 And Second(endTime) = 0)

    If fullDay Then
        hourLimit = 23
    Else
        hourLimit = Hour(endTime) - 1
    End If

    lastRowTruncated = wsTruncated.Cells(wsTruncated.rows.Count, siteIDColumn).End(xlUp).Row
    lastRowPATrend = wsPATrend.Cells(wsPATrend.rows.Count, 3).End(xlUp).Row

    wsPATrend.Range("T2:AQ" & lastRowPATrend).ClearContents

    Set availabilityDict = CreateObject("Scripting.Dictionary")
    Set hourCountDict = CreateObject("Scripting.Dictionary")

    Set dataRange = wsTruncated.Range(wsTruncated.Cells(2, siteIDColumn), wsTruncated.Cells(lastRowTruncated, availabilityColumn))
    For Each cell In dataRange.Columns(1).Cells
        siteID = cell.Value
        hourStart = cell.Offset(0, hourStartColumn - siteIDColumn).Value

        On Error Resume Next
        hourNum = Hour(CDate(hourStart))
        On Error GoTo 0

        If Not IsError(hourNum) Then
            availability = cell.Offset(0, availabilityColumn - siteIDColumn).Value

            If TypeName(availability) = "String" Then
                If InStr(availability, "%") > 0 Then
                    availability = CDbl(Replace(availability, "%", ""))
                Else
                    availability = CDbl(availability)
                End If
            End If

            key = siteID & "_" & hourNum
            
            If availabilityDict.exists(key) Then
                availabilityDict(key) = availabilityDict(key) + availability
                hourCountDict(key) = hourCountDict(key) + 1
            Else
                availabilityDict(key) = availability
                hourCountDict(key) = 1
            End If
        End If
    Next cell

    Set hourStartDict = CreateObject("Scripting.Dictionary")
    For col = 20 To 43
        hourStart = wsPATrend.Cells(1, col).Value

        If hourStart <> "" Then
            On Error Resume Next
            hourNum = Hour(CDate(hourStart))
            On Error GoTo 0

            If Not IsError(hourNum) Then
                If hourNum <= hourLimit Then
                    hourStartDict(hourNum) = col
                End If
            End If
        End If
    Next col

    Set dataRange = wsPATrend.Range(wsPATrend.Cells(2, 3), wsPATrend.Cells(lastRowPATrend, 3))
    ReDim resultArray(1 To dataRange.rows.Count, 1 To 24)

    For Each cell In dataRange.Cells
        siteID = cell.Value
        cellRow = cell.Row - 1
        siteStatus = cell.Offset(0, 12).Value ' Column O (15 - 3)

        If siteID <> "" Then
            For Each hourNum In hourStartDict.Keys
                key = siteID & "_" & hourNum
                col = hourStartDict(hourNum) - 19
                
                If siteStatus = "Off Air" Then
                    resultArray(cellRow, col) = "-"
                ElseIf availabilityDict.exists(key) Then
                    availability = availabilityDict(key) / hourCountDict(key)
                    resultArray(cellRow, col) = Round(availability, 2)
                Else
                    resultArray(cellRow, col) = 100#
                End If
            Next hourNum
        End If
    Next cell

    wsPATrend.Range("T2:AQ" & lastRowPATrend).Value = resultArray

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
   wsPATrend.Activate

   ' MsgBox "Hourly PA Trend populated successfully!"
End Sub
