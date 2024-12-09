Attribute VB_Name = "Filter"
Option Explicit
Sub FilterRCAsAndCopyToPassive()
    Dim wsMenu As Worksheet, wsDump As Worksheet, wsPassive As Worksheet
    Dim rcaDict As Object, data As Variant, result() As Variant
    Dim colToFilter As String, filterCol As Long
    Dim lastRowMenu As Long, lastRowDump As Long, lastColDump As Long
    Dim i As Long, j As Long, resultRow As Long, cell As Range
    Dim startTimeColNum As Long, endTimeColNum As Long, finalDurationNum As Long
    
    Application.ScreenUpdating = False
    
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    Set wsDump = ThisWorkbook.Sheets("Dump")
    Set wsPassive = ThisWorkbook.Sheets("Passive")
    
    ' Get column numbers for start and end time
    With wsDump
        startTimeColNum = Application.Match("Final Outage Start", .rows(1), 0)
        endTimeColNum = Application.Match("Final Outage End", .rows(1), 0)
        finalDurationNum = Application.Match("Final Duration", .rows(1), 0)
    End With
    
    wsPassive.Cells.Clear
    colToFilter = wsMenu.Range("L25").Value
    filterCol = wsDump.Range(colToFilter & "1").Column
    
    lastRowMenu = wsMenu.Cells(wsMenu.rows.Count, "A").End(xlUp).Row
    Set rcaDict = CreateObject("Scripting.Dictionary")
    
    For Each cell In wsMenu.Range("A13:A" & lastRowMenu)
        If Len(cell.Value) > 0 Then rcaDict(cell.Value) = 1
    Next cell
    
    With wsDump
        lastRowDump = .Cells(.rows.Count, filterCol).End(xlUp).Row
        lastColDump = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        ' Copy headers
        .rows(1).Copy wsPassive.rows(1)
        
        ' Get data into array
        data = .Range(.Cells(2, 1), .Cells(lastRowDump, lastColDump))
        ReDim result(1 To UBound(data, 1), 1 To UBound(data, 2))
        
        ' Filter data - keep rows that are NOT in rcaDict
        resultRow = 1
        For i = 1 To UBound(data, 1)
            If Not rcaDict.exists(data(i, filterCol)) Then
                For j = 1 To UBound(data, 2)
                    result(resultRow, j) = data(i, j)
                Next j
                resultRow = resultRow + 1
            End If
        Next i
        
        ' Copy filtered results
        If resultRow > 1 Then
            wsPassive.Range(.Cells(2, 1).Address).Resize(resultRow - 1, lastColDump).Value = result
        End If
    End With
    
    ' Format date columns
    If startTimeColNum > 0 Then wsPassive.Columns(startTimeColNum).NumberFormat = "mm/dd/yyyy h:mm"
    If endTimeColNum > 0 Then wsPassive.Columns(endTimeColNum).NumberFormat = "mm/dd/yyyy h:mm"
    If finalDurationNum > 0 Then wsPassive.Columns(finalDurationNum).NumberFormat = "[h]:mm:ss"
    
    With wsPassive.Cells
        .WrapText = False
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub FilterRCAsAndCopyToActive()
    Dim wsMenu As Worksheet, wsDump As Worksheet, wsActive As Worksheet
    Dim rcaDict As Object, data As Variant, result() As Variant
    Dim colToFilter As String, filterCol As Long
    Dim lastRowMenu As Long, lastRowDump As Long, lastColDump As Long
    Dim i As Long, j As Long, resultRow As Long, cell As Range
    Dim startTimeColNum As Long, endTimeColNum As Long, finalDurationNum As Long
    
    Application.ScreenUpdating = False
    
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    Set wsDump = ThisWorkbook.Sheets("Dump")
    Set wsActive = ThisWorkbook.Sheets("Active")
    
    ' Get column numbers for start and end time
    With wsDump
        startTimeColNum = Application.Match("Final Outage Start", .rows(1), 0)
        endTimeColNum = Application.Match("Final Outage End", .rows(1), 0)
        finalDurationNum = Application.Match("Final Duration", .rows(1), 0)
    End With
    
    wsActive.Cells.Clear
    colToFilter = wsMenu.Range("L25").Value
    filterCol = wsDump.Range(colToFilter & "1").Column
    
    lastRowMenu = wsMenu.Cells(wsMenu.rows.Count, "A").End(xlUp).Row
    Set rcaDict = CreateObject("Scripting.Dictionary")
    
    For Each cell In wsMenu.Range("B13:B" & lastRowMenu)
        If Len(cell.Value) > 0 Then rcaDict(cell.Value) = 1
    Next cell
    
    With wsDump
        lastRowDump = .Cells(.rows.Count, filterCol).End(xlUp).Row
        lastColDump = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        ' Copy headers
        .rows(1).Copy wsActive.rows(1)
        
        ' Get data into array
        data = .Range(.Cells(2, 1), .Cells(lastRowDump, lastColDump))
        ReDim result(1 To UBound(data, 1), 1 To UBound(data, 2))
        
        ' Filter data - keep rows that are NOT in rcaDict
        resultRow = 1
        For i = 1 To UBound(data, 1)
            If rcaDict.exists(data(i, filterCol)) Then
                For j = 1 To UBound(data, 2)
                    result(resultRow, j) = data(i, j)
                Next j
                resultRow = resultRow + 1
            End If
        Next i
        
        ' Copy filtered results
        If resultRow > 1 Then
            wsActive.Range(.Cells(2, 1).Address).Resize(resultRow - 1, lastColDump).Value = result
        End If
    End With
    
    ' Format date columns
    If startTimeColNum > 0 Then wsActive.Columns(startTimeColNum).NumberFormat = "mm/dd/yyyy h:mm"
    If endTimeColNum > 0 Then wsActive.Columns(endTimeColNum).NumberFormat = "mm/dd/yyyy h:mm"
    If finalDurationNum > 0 Then wsActive.Columns(finalDurationNum).NumberFormat = "[h]:mm:ss"
    
    With wsActive.Cells
        .WrapText = False
    End With
    
    Application.ScreenUpdating = True
End Sub

