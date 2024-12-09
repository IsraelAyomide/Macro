Attribute VB_Name = "Average"
Option Explicit
Sub CalculateRegionAvailabilityAverage()
    Dim wsPA As Worksheet, wsTable As Worksheet
    Dim rngPA As Range, rngTable As Range
    Dim regionDict As Object
    Dim arrPA As Variant, arrTable As Variant
    Dim i As Long, j As Long
    Dim region As String
    Dim colStart As Long, colEnd As Long
    Dim valueSum As Double, valueCount As Long
    Dim overallSum As Double, overallCount As Long

    ' Set worksheet references
    Set wsPA = ThisWorkbook.Sheets("PA Trend")
    Set wsTable = ThisWorkbook.Sheets("Table")

    ' Define ranges
    Set rngPA = wsPA.Range("A1:AQ" & wsPA.Cells(wsPA.rows.Count, "A").End(xlUp).Row)
    Set rngTable = wsTable.Range("A30:A36")
    
    ' Load data into arrays
    arrPA = rngPA.Value
    arrTable = rngTable.Value

    ' Define start and end columns for Availability data
    colStart = 20 ' Column T
    colEnd = 43   ' Column AQ

    ' Create a dictionary to store region sums and counts
    Set regionDict = CreateObject("Scripting.Dictionary")

    ' Initialize the dictionary with region names from the Table sheet
    For i = LBound(arrTable, 1) To UBound(arrTable, 1)
        If Not IsEmpty(arrTable(i, 1)) Then
            regionDict(arrTable(i, 1)) = Array(0, 0) ' (Sum, Count)
        End If
    Next i

    ' Initialize overall sum and count
    overallSum = 0
    overallCount = 0

    ' Iterate over the PA Trend data and calculate sums and counts for each region
    For i = 2 To UBound(arrPA, 1) ' Start from row 2 to skip headers
        region = arrPA(i, 7) ' Column G (Region)
        If regionDict.exists(region) Then
            valueSum = 0
            valueCount = 0

            ' Iterate over the Availability columns T to AQ
            For j = colStart To colEnd
                If IsNumeric(arrPA(i, j)) And Not IsEmpty(arrPA(i, j)) Then
                    valueSum = valueSum + arrPA(i, j)
                    valueCount = valueCount + 1
                End If
            Next j

            ' Update the dictionary with the new sum and count
            If valueCount > 0 Then
                Dim currentData As Variant
                currentData = regionDict(region)
                currentData(0) = currentData(0) + valueSum
                currentData(1) = currentData(1) + valueCount
                regionDict(region) = currentData
            End If
        End If

        ' Calculate overall availability sum and count for all data
        For j = colStart To colEnd
            If IsNumeric(arrPA(i, j)) And Not IsEmpty(arrPA(i, j)) Then
                overallSum = overallSum + arrPA(i, j)
                overallCount = overallCount + 1
            End If
        Next j
    Next i

    ' Write the average Availability values back to the Table sheet
    For i = LBound(arrTable, 1) To UBound(arrTable, 1)
        region = arrTable(i, 1)
        If regionDict.exists(region) Then
            Dim sumAndCount As Variant
            sumAndCount = regionDict(region)
            If sumAndCount(1) > 0 Then
                wsTable.Cells(i + 29, 2).Value = sumAndCount(0) / sumAndCount(1) ' Write average to column B (Availability)
            Else
                wsTable.Cells(i + 29, 2).Value = "N/A"
            End If
        End If
    Next i

    ' Write the overall average availability to cell A39 of the Table sheet
    If overallCount > 0 Then
        wsTable.Cells(39, 1).Value = Format(overallSum / overallCount, "0.00")
    Else
        wsTable.Cells(39, 1).Value = "N/A"
    End If

    ' Clean up
    Set regionDict = Nothing

   ' MsgBox "Region and Overall Availability Averages Calculated Successfully!", vbInformation
End Sub

