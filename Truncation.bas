Attribute VB_Name = "Truncation"
Option Explicit
Sub AdjustDateRanges()
    ' Define worksheets
    Dim wsDump As Worksheet, wsMenu As Worksheet
    Set wsDump = ThisWorkbook.Sheets("Dump")
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    
    ' Get column letters from Menu sheet
    Dim startTimeColumn As String, endTimeColumn As String
    startTimeColumn = wsMenu.Range("L23").Value
    endTimeColumn = wsMenu.Range("L24").Value
    
    ' Convert column letters to numbers
    Dim startTimeColNum As Long, endTimeColNum As Long
    startTimeColNum = wsDump.Columns(startTimeColumn).Column
    endTimeColNum = wsDump.Columns(endTimeColumn).Column
    
    ' Get date range from Menu sheet
    Dim startDate As Date, endDate As Date
    startDate = wsMenu.Range("L13").Value
    endDate = wsMenu.Range("L14").Value
    
    ' Define variables
    Dim lastRow As Long, i As Long
    lastRow = wsDump.Cells(wsDump.rows.Count, startTimeColNum).End(xlUp).Row
    
    ' Loop through the data to adjust dates
    For i = 2 To lastRow
        ' Adjust Start Time if earlier than Start Date
        If wsDump.Cells(i, startTimeColNum).Value < startDate Then
            wsDump.Cells(i, startTimeColNum).Value = startDate
        End If
        ' Adjust End Time if later than End Date
        If wsDump.Cells(i, endTimeColNum).Value > endDate Then
            wsDump.Cells(i, endTimeColNum).Value = endDate
        End If
        ' Set End Time to End Date if blank
        If IsEmpty(wsDump.Cells(i, endTimeColNum).Value) Then
            wsDump.Cells(i, endTimeColNum).Value = endDate
        End If
    Next i
    
   ' MsgBox "Date range adjustment completed successfully!", vbInformation
End Sub
