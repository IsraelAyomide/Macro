Attribute VB_Name = "Process_data"
Option Explicit

Sub ProcessInputData()
    ' Define worksheets
    Dim wsDump As Worksheet, wsMenu As Worksheet
    Set wsDump = ThisWorkbook.Sheets("Dump")
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    
    ' Get column letters from Menu sheet (L17 to L26)
    Dim columnDict As Object
    Set columnDict = CreateObject("Scripting.Dictionary")
    
    columnDict("internalOutageStartCol") = wsMenu.Range("L17").Value
    columnDict("internalOutageEndCol") = wsMenu.Range("L18").Value
    columnDict("internalPrimaryCauseCol") = wsMenu.Range("L19").Value
    columnDict("outageStartTimeCol") = wsMenu.Range("L20").Value
    columnDict("outageEndTimeCol") = wsMenu.Range("L21").Value
    columnDict("primaryCauseCol") = wsMenu.Range("L22").Value
    columnDict("finalOutageStartCol") = wsMenu.Range("L23").Value
    columnDict("finalOutageEndCol") = wsMenu.Range("L24").Value
    columnDict("finalPrimaryCauseCol") = wsMenu.Range("L25").Value
    columnDict("finalDurationCol") = wsMenu.Range("L26").Value
    
    ' Convert column letters to positions within the dataDump array
    Dim colKeys As Variant, colPositions As Object
    Set colPositions = CreateObject("Scripting.Dictionary")
    colKeys = columnDict.Keys
    
    Dim key As Variant    ' Add this declaration before the For Each loop
        For Each key In colKeys
            colPositions(key) = wsDump.Columns(columnDict(key)).Column
    Next key
    
    For Each key In colKeys
        colPositions(key) = wsDump.Columns(columnDict(key)).Column
    Next key
    
    ' Add headers for the new columns
    wsDump.Cells(1, colPositions("finalOutageStartCol")).Value = "Final Outage Start"
    wsDump.Cells(1, colPositions("finalOutageEndCol")).Value = "Final Outage End"
    wsDump.Cells(1, colPositions("finalPrimaryCauseCol")).Value = "Final Primary Cause"
    wsDump.Cells(1, colPositions("finalDurationCol")).Value = "Final Duration"
    
    ' Find maximum column needed
    Dim maxColumnNeeded As Long
    maxColumnNeeded = 0
    For Each key In colKeys
        If colPositions(key) > maxColumnNeeded Then
            maxColumnNeeded = colPositions(key)
        End If
    Next key
    
    ' Validate column range
    If maxColumnNeeded > wsDump.Cells(1, wsDump.Columns.Count).End(xlToLeft).Column Then
        MsgBox "Error: Some required columns are outside the data range.", vbCritical
        Exit Sub
    End If
    
    ' Define variables
    Dim lastRow As Long, i As Long, dataDump As Variant, dataResult As Variant
    lastRow = wsDump.Cells(wsDump.rows.Count, colPositions("primaryCauseCol")).End(xlUp).Row
    
    ' Load data into an array to speed up processing
    dataDump = wsDump.Range(wsDump.Cells(1, 1), wsDump.Cells(lastRow, maxColumnNeeded)).Value
    
    ' Initialize result array
    ReDim dataResult(1 To UBound(dataDump, 1), 1 To maxColumnNeeded)
    
    ' Set headers for new columns
    For i = 1 To UBound(dataDump, 2)
        dataResult(1, i) = dataDump(1, i)
    Next i
    
    ' Populate Final Outage Start, End, and Cause with error handling
    For i = 2 To UBound(dataDump, 1)
        On Error Resume Next
        
        ' Final Outage Start
        If colPositions("internalOutageStartCol") <= UBound(dataDump, 2) And _
           colPositions("outageStartTimeCol") <= UBound(dataDump, 2) And _
           colPositions("finalOutageStartCol") <= UBound(dataDump, 2) Then
            
            If Trim(dataDump(i, colPositions("internalOutageStartCol"))) = "" Then
                dataResult(i, colPositions("finalOutageStartCol")) = dataDump(i, colPositions("outageStartTimeCol"))
            Else
                dataResult(i, colPositions("finalOutageStartCol")) = dataDump(i, colPositions("internalOutageStartCol"))
            End If
        End If
        
        ' Final Outage End
        If colPositions("internalOutageEndCol") <= UBound(dataDump, 2) And _
           colPositions("outageEndTimeCol") <= UBound(dataDump, 2) And _
           colPositions("finalOutageEndCol") <= UBound(dataDump, 2) Then
            
            If Trim(dataDump(i, colPositions("internalOutageEndCol"))) = "" Then
                dataResult(i, colPositions("finalOutageEndCol")) = dataDump(i, colPositions("outageEndTimeCol"))
            Else
                dataResult(i, colPositions("finalOutageEndCol")) = dataDump(i, colPositions("internalOutageEndCol"))
            End If
        End If
        
        ' Final Primary Cause
        If colPositions("internalPrimaryCauseCol") <= UBound(dataDump, 2) And _
           colPositions("primaryCauseCol") <= UBound(dataDump, 2) And _
           colPositions("finalPrimaryCauseCol") <= UBound(dataDump, 2) Then
            
            If Trim(dataDump(i, colPositions("internalPrimaryCauseCol"))) = "" Then
                dataResult(i, colPositions("finalPrimaryCauseCol")) = dataDump(i, colPositions("primaryCauseCol"))
            Else
                dataResult(i, colPositions("finalPrimaryCauseCol")) = dataDump(i, colPositions("internalPrimaryCauseCol"))
            End If
        End If
        
        If Err.Number <> 0 Then
            Debug.Print "Error at row " & i & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    ' Write the results back to the sheet before adjusting dates
    On Error Resume Next
    wsDump.Range(wsDump.Cells(1, colPositions("finalOutageStartCol")), _
                 wsDump.Cells(lastRow, colPositions("finalPrimaryCauseCol"))).Value = _
    Application.Index(dataResult, Evaluate("ROW(1:" & lastRow & ")"), _
                     Array(colPositions("finalOutageStartCol"), _
                           colPositions("finalOutageEndCol"), _
                           colPositions("finalPrimaryCauseCol")))
    
    If Err.Number <> 0 Then
        Debug.Print "Error writing results: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' Run AdjustDateRanges on Final Outage Start and End
    Call AdjustDateRanges
    
    ' Check and delete rows where Final Outage Start is greater than Final Outage End
    Dim j As Long
    For j = lastRow To 2 Step -1
        If wsDump.Cells(j, colPositions("finalOutageStartCol")).Value > _
           wsDump.Cells(j, colPositions("finalOutageEndCol")).Value Then
            wsDump.rows(j).Delete
        End If
    Next j
    
    ' Update lastRow after deletions
    lastRow = wsDump.Cells(wsDump.rows.Count, colPositions("primaryCauseCol")).End(xlUp).Row
    
    ' Calculate Final Duration
    On Error Resume Next
    For i = 2 To lastRow
        wsDump.Cells(i, colPositions("finalDurationCol")).Value = _
            wsDump.Cells(i, colPositions("finalOutageEndCol")).Value - _
            wsDump.Cells(i, colPositions("finalOutageStartCol")).Value
        wsDump.Cells(i, colPositions("finalDurationCol")).NumberFormat = "[h]:mm:ss"
        
        If Err.Number <> 0 Then
            Debug.Print "Error calculating duration at row " & i & ": " & Err.Description
            Err.Clear
        End If
    Next i
    On Error GoTo 0
    
  '  MsgBox "Input data processed successfully!", vbInformation
End Sub
