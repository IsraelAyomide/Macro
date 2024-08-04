Attribute VB_Name = "Main"
' Module: ModuleSharedFunctions

Option Explicit

Function CalculateAverageIfs(rng As Range, criteria As Variant, avgRange As Range, regionRange As Range, region As String) As Double
    Dim cell As Range
    Dim sum As Double
    Dim count As Long
    Dim i As Long
    
    sum = 0
    count = 0
    
    If IsArray(criteria) Then
        For i = LBound(criteria) To UBound(criteria)
            For Each cell In rng
                If LCase(cell.Value) = LCase(criteria(i)) And LCase(regionRange.Cells(cell.Row, 1).Value) = LCase(region) Then
                    sum = sum + avgRange.Cells(cell.Row, 1).Value
                    count = count + 1
                End If
            Next cell
        Next i
    Else
        For Each cell In rng
            If LCase(cell.Value) = LCase(criteria) And LCase(regionRange.Cells(cell.Row, 1).Value) = LCase(region) Then
                sum = sum + avgRange.Cells(cell.Row, 1).Value
                count = count + 1
            End If
        Next cell
    End If
    
    If count > 0 Then
        CalculateAverageIfs = sum / count
    Else
        CalculateAverageIfs = 0
    End If
End Function

Sub UpdateSheet(sheetName As String, criteria As String)
    Dim ws As Worksheet
    Dim wsMenu As Worksheet
    Dim wsBreakdown As Worksheet
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim searchWords As Variant
    Dim wordCount As Long
    Dim total As Long
    Dim avgValue As Double
    Dim totalAvg As Double
    Dim avgCount As Long
    Dim sumValue As Double
    Dim totalSum As Double
    Dim dataRange As Variant
    Dim avgDict As Object
    Dim countDict As Object
    Dim sumDict As Object
    Dim key As String
    Dim lr As Long
    
    ' Set worksheet references
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    Set wsBreakdown = ThisWorkbook.Sheets("breakdown")
    
    ' Find the last column in the sheet
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
    
    ' Add value from MENU!H6 to the first cell in the last column
    ws.Cells(1, lastCol).Value = wsMenu.Range("H6").Value
    
    ' Define the words to search for
    searchWords = Array("ACCESS", "Environmental", "DC Issue", "DCGG Gen", "DIESEL", "GAS", "GRID", "Infrastructure", "OTHERS", "PLANNED ACTIVITY", "RECTIFIER", "Theft or Force Majuere")
    
    ' Initialize dictionaries
    Set avgDict = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    Set sumDict = CreateObject("Scripting.Dictionary")
    
    ' Initialize total
    total = 0
    totalAvg = 0
    avgCount = 0
    totalSum = 0
    
    ' Load the entire breakdown data into an array
    lr = wsBreakdown.Cells(wsBreakdown.Rows.count, "A").End(xlUp).Row
    dataRange = wsBreakdown.Range("A1:K" & lr).Value
    
    ' Loop through the data array
    For i = 2 To UBound(dataRange, 1) ' Start from row 2 to skip header
        If LCase(dataRange(i, 5)) = LCase(criteria) Then
            key = LCase(dataRange(i, 11))
            If key = "acdg gen" Or key = "dcdg gen" Then
                key = "dc issue"
            ElseIf key = "theft" Or key = "force majuere" Then
                key = "theft or force majuere"
            End If
            
            If Not avgDict.exists(key) Then
                avgDict(key) = 0
                countDict(key) = 0
                sumDict(key) = 0
            End If
            
            avgDict(key) = avgDict(key) + dataRange(i, 9)
            countDict(key) = countDict(key) + 1
            sumDict(key) = sumDict(key) + dataRange(i, 9)
        End If
    Next i
    
    ' Populate the counts in the worksheet
    For j = LBound(searchWords) To UBound(searchWords)
        key = LCase(searchWords(j))
        If countDict.exists(key) Then
            ws.Cells(j + 2, lastCol).Value = countDict(key)
            total = total + countDict(key)
        Else
            ws.Cells(j + 2, lastCol).Value = 0
        End If
    Next j
    
    ' Add the total count to the last row
    ws.Cells(UBound(searchWords) + 3, lastCol).Value = total
    
    ' Calculate averages and print starting from row 17
    For j = LBound(searchWords) To UBound(searchWords)
        key = LCase(searchWords(j))
        If avgDict.exists(key) And countDict(key) > 0 Then
            avgValue = avgDict(key) / countDict(key)
            ws.Cells(j + 17, lastCol).Value = avgValue
            ws.Cells(j + 17, lastCol).NumberFormat = "[h]:mm:ss"
            totalAvg = totalAvg + avgValue
            avgCount = avgCount + 1
        Else
            ws.Cells(j + 17, lastCol).Value = "00:00:00"
        End If
    Next j
    
    ' Add the average of the averages to the last row
    If avgCount > 0 Then
        ws.Cells(UBound(searchWords) + 18, lastCol).Value = totalAvg / avgCount
        ws.Cells(UBound(searchWords) + 18, lastCol).NumberFormat = "[h]:mm:ss"
    Else
        ws.Cells(UBound(searchWords) + 18, lastCol).Value = "00:00:00"
    End If
    
    ' Calculate sum of durations and print starting from row 31
    For j = LBound(searchWords) To UBound(searchWords)
        key = LCase(searchWords(j))
        If sumDict.exists(key) Then
            sumValue = sumDict(key)
            ws.Cells(j + 32, lastCol).Value = sumValue
            ws.Cells(j + 32, lastCol).NumberFormat = "[h]:mm:ss"
            totalSum = totalSum + sumValue
        Else
            ws.Cells(j + 32, lastCol).Value = "00:00:00"
        End If
    Next j
    
    ' Add the total sum of the durations to the last row
    ws.Cells(UBound(searchWords) + 33, lastCol).Value = totalSum
    ws.Cells(UBound(searchWords) + 33, lastCol).NumberFormat = "[h]:mm:ss"
    
    ' Remove the filter
    wsBreakdown.AutoFilterMode = False
End Sub
' Module: Abuja
Sub UpdateAbujaSheet()
    UpdateSheet "Abuja", "ABJ"
End Sub

' Module: Asaba
Sub UpdateAsabaSheet()
    UpdateSheet "Asaba", "ASB"
End Sub

' Module: Enugu
Sub UpdateEnuguSheet()
    UpdateSheet "Enugu", "ENG"
End Sub

' Module: Ibadan
Sub UpdateIbadanSheet()
    UpdateSheet "Ibadan", "IBD"
End Sub

' Module: Kano
Sub UpdateKanoSheet()
    UpdateSheet "Kano", "KNO"
End Sub

' Module: Lagos
Sub UpdateLagosSheet()
    UpdateSheet "Lagos", "LGS"
End Sub

' Module: PHC
Sub UpdatePHCSheet()
    UpdateSheet "PHC", "PHC"
End Sub
'---------------------------------------------------------------------------------------------
Attribute VB_Name = "Overall_Counter"
Sub UpdateOverallSheet()
    Dim wsOverall As Worksheet
    Dim wsMenu As Worksheet
    Dim wsBreakdown As Worksheet
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim searchWords As Variant
    Dim wordCount As Long
    Dim total As Long
    Dim avgValue As Double
    Dim totalAvg As Double
    Dim avgCount As Long
    Dim sumValue As Double
    Dim totalSum As Double
    Dim dataRange As Variant
    Dim avgDict As Object
    Dim countDict As Object
    Dim sumDict As Object
    Dim key As String
    
    ' Set worksheet references
    Set wsOverall = ThisWorkbook.Sheets("Overall")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    Set wsBreakdown = ThisWorkbook.Sheets("breakdown")
    
    ' Update breakdown sheet to replace unlisted words with "OTHERS"
    UpdateBreakdownSheet
    
    ' Find the last column in the "Overall" sheet
    lastCol = wsOverall.Cells(1, wsOverall.Columns.count).End(xlToLeft).Column + 1
    
    ' Add value from MENU!H6 to the first cell in the last column
    wsOverall.Cells(1, lastCol).Value = wsMenu.Range("H6").Value
    
    ' Define the words to search for
    searchWords = Array("ACCESS", "Environmental", "DC Issue", "DCGG Gen", "DIESEL", "GAS", "GRID", "Infrastructure", "OTHERS", "PLANNED ACTIVITY", "RECTIFIER", "Theft or Force Majuere")
    
    ' Initialize dictionaries
    Set avgDict = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    Set sumDict = CreateObject("Scripting.Dictionary")
    '-------------------------------------------------------------------
    total = 0
    totalAvg = 0
    avgCount = 0
    totalSum = 0
    '-------------------------------------------------------------------
    dataRange = wsBreakdown.UsedRange.Value
    '-----------------------------------------------------------------------
    For i = 2 To UBound(dataRange, 1)
        key = LCase(dataRange(i, 11))
        If key = "acdg gen" Or key = "dcdg gen" Then
            key = "dc issue"
        ElseIf key = "theft" Or key = "force majuere" Then
            key = "theft or force majuere"
        End If
        
        If Not avgDict.exists(key) Then
            avgDict(key) = 0
            countDict(key) = 0
            sumDict(key) = 0
        End If
        
        avgDict(key) = avgDict(key) + dataRange(i, 9)
        countDict(key) = countDict(key) + 1
        sumDict(key) = sumDict(key) + dataRange(i, 9)
    Next i
    
    ' Populate the counts in the worksheet
    For j = LBound(searchWords) To UBound(searchWords)
        key = LCase(searchWords(j))
        If key = "dc issue" Then
            key = "dc issue"
        ElseIf key = "theft or force majuere" Then
            key = "theft or force majuere"
        End If
        
        If countDict.exists(key) Then
            wsOverall.Cells(j + 2, lastCol).Value = countDict(key)
            total = total + countDict(key)
        Else
            wsOverall.Cells(j + 2, lastCol).Value = 0
        End If
    Next j
    
    ' Add the total count to the last row
    wsOverall.Cells(UBound(searchWords) + 3, lastCol).Value = total
    
    ' Calculate averages and print starting from row 17
    For j = LBound(searchWords) To UBound(searchWords)
        key = LCase(searchWords(j))
        If key = "dc issue" Then
            key = "dc issue"
        ElseIf key = "theft or force majuere" Then
            key = "theft or force majuere"
        End If
        
        If avgDict.exists(key) And countDict(key) > 0 Then
            avgValue = avgDict(key) / countDict(key)
            wsOverall.Cells(j + 17, lastCol).Value = avgValue
            wsOverall.Cells(j + 17, lastCol).NumberFormat = "[h]:mm:ss"
            totalAvg = totalAvg + avgValue
            avgCount = avgCount + 1
        Else
            wsOverall.Cells(j + 17, lastCol).Value = "00:00:00"
        End If
    Next j
    
    ' Add the average of the averages to the last row
    If avgCount > 0 Then
        wsOverall.Cells(UBound(searchWords) + 18, lastCol).Value = totalAvg / avgCount
        wsOverall.Cells(UBound(searchWords) + 18, lastCol).NumberFormat = "[h]:mm:ss"
    Else
        wsOverall.Cells(UBound(searchWords) + 18, lastCol).Value = "00:00:00"
    End If
    
    ' Calculate sum of durations and print starting from row 31
    For j = LBound(searchWords) To UBound(searchWords)
        key = LCase(searchWords(j))
        If key = "dc issue" Then
            key = "dc issue"
        ElseIf key = "theft or force majuere" Then
            key = "theft or force majuere"
        End If
        
        If sumDict.exists(key) Then
            sumValue = sumDict(key)
            wsOverall.Cells(j + 32, lastCol).Value = sumValue
            wsOverall.Cells(j + 32, lastCol).NumberFormat = "[h]:mm:ss"
            totalSum = totalSum + sumValue
        Else
            wsOverall.Cells(j + 32, lastCol).Value = "00:00:00"
        End If
    Next j
    
    ' Add the total sum of the durations to the last row
    wsOverall.Cells(UBound(searchWords) + 33, lastCol).Value = totalSum
    wsOverall.Cells(UBound(searchWords) + 33, lastCol).NumberFormat = "[h]:mm:ss"
End Sub

Sub UpdateBreakdownSheet()
    Dim wsBreakdown As Worksheet
    Dim cell As Range
    Dim validWords As Variant
    Dim word As String
    Dim found As Boolean
    Dim i As Long
    
    ' Set worksheet reference
    Set wsBreakdown = ThisWorkbook.Sheets("breakdown")
    
    ' Define the list of valid words
    validWords = Array("Environmental", "ACCESS", "Diesel", "Infrastructure", "Force Majuere", _
                       "ACDG Gen", "DCDG Gen", "DCGG Gen", "Grid", "Theft", _
                       "Planned Activity", "Gas", "Rectifier")
    
    ' Loop through each cell in column K starting from K2
    For Each cell In wsBreakdown.Range("K2:K" & wsBreakdown.Cells(wsBreakdown.Rows.count, "K").End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            found = False
            word = LCase(cell.Value)  ' Convert cell value to lower case
            
            ' Check if the word is in the list of valid words
            For i = LBound(validWords) To UBound(validWords)
                If word = LCase(validWords(i)) Then  ' Convert valid word to lower case
                    found = True
                    Exit For
                End If
            Next i
            
            ' If the word is not found, change it to "OTHERS"
            If Not found Then
                cell.Value = "OTHERS"
            End If
        End If
    Next cell
End Sub

Function AverageIf(rng As Range, criteria As Variant, avgRange As Range) As Double
    Dim cell As Range
    Dim sum As Double
    Dim count As Long
    Dim i As Long
    
    sum = 0
    count = 0
    
    If IsArray(criteria) Then
        For i = LBound(criteria) To UBound(criteria)
            For Each cell In rng
                If LCase(cell.Value) = LCase(criteria(i)) Then
                    sum = sum + avgRange.Cells(cell.Row, 1).Value
                    count = count + 1
                End If
            Next cell
        Next i
    Else
        For Each cell In rng
            If LCase(cell.Value) = LCase(criteria) Then
                sum = sum + avgRange.Cells(cell.Row, 1).Value
                count = count + 1
            End If
        Next cell
    End If
    
    If count > 0 Then
        AverageIf = sum / count
    Else
        AverageIf = 0
    End If
End Function
'------------------------------------------------------------------------------------
Attribute VB_Name = "Macro_trigger"
Sub Trigger()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    UpdateOverallSheet
    UpdateAbujaSheet
    UpdateEnuguSheet
    UpdateIbadanSheet
    UpdateKanoSheet
    UpdateLagosSheet
    UpdatePHCSheet
    UpdateAsabaSheet
    CopyFormattingAcrossSheets
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Complete in " & Format(Timer - startTime, "0.00 seconds"), vbInformation
End Sub

Sub CopyFormattingAcrossSheets()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim i As Long

    ' Define the names of the sheets to process
    sheetNames = Array("Overall", "Abuja", "Enugu", "Ibadan", "Kano", "Lagos", "PHC", "Asaba")

    ' Loop through each sheet name
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        
        ' Find the last column in the worksheet
        lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        
        ' Copy formatting from the second-to-last column to the last column
        ws.Columns(lastCol - 1).Copy
        ws.Columns(lastCol).PasteSpecial Paste:=xlPasteFormats
        
        ' Clear the clipboard to remove the copied selection
        Application.CutCopyMode = False
    Next i
    
   ' MsgBox "Formatting copied across all specified sheets.", vbInformation
End Sub

