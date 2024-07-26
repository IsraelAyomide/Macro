Attribute VB_Name = "Site_detail_update"
Sub UpdateSiteStatus()
    Dim wsTCH As Worksheet
    Dim wsList As Variant
    Dim ws As Worksheet
    Dim mapDict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim siteID As String
    Dim status As String
    
    ' Set the TCH sheet
    Set wsTCH = ThisWorkbook.Sheets("TCH")
    
    ' Create a dictionary for the TCH site status data
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    ' Load TCH data into the dictionary
    lastRow = wsTCH.Cells(wsTCH.Rows.Count, "F").End(xlUp).Row
    For i = 2 To lastRow
        siteID = Trim(wsTCH.Cells(i, "F").Value)
        status = wsTCH.Cells(i, "G").Value
        mapDict(siteID) = status
    Next i
    
    ' List of worksheets to update
    wsList = Array("2G", "3G", "4G", "5G", "2G & 3G & 4G & 5G")
    
    ' Disable screen updates and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each worksheet and update status
    For Each wsName In wsList
        Set ws = ThisWorkbook.Sheets(wsName)
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        For i = 2 To lastRow
            siteID = Trim(ws.Cells(i, "D").Value)
            If mapDict.exists(siteID) Then
                ws.Cells(i, "N").Value = mapDict(siteID)
            Else
                ws.Cells(i, "N").Value = "Off Air"
            End If
        Next i
    Next wsName
    
    ' Re-enable screen updates and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Site status update complete", vbInformation
End Sub
Sub UpdateSiteDetails()
    Dim wsTCH As Worksheet
    Dim wsList As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim siteID As String
    Dim status As String
    Dim tchData As Variant
    Dim siteDict As Object
    Dim colPairs As Variant
    Dim tchCol As Long
    Dim shtCol As Long
    Dim siteIndex As Long
    
    ' Set the TCH sheet
    Set wsTCH = ThisWorkbook.Sheets("TCH")
    
    ' Load TCH data into an array
    lastRow = wsTCH.Cells(wsTCH.Rows.Count, "F").End(xlUp).Row
    tchData = wsTCH.Range("A1:AI" & lastRow).Value
    
    ' Create a dictionary to map site IDs to row indices
    Set siteDict = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(tchData, 1)
        siteID = Trim(tchData(i, 6)) ' Column F is the 6th column
        siteDict(siteID) = i
    Next i
    
    ' List of worksheets to update
    wsList = Array("2G", "3G", "4G", "5G", "2G & 3G & 4G & 5G")
    
    ' Define column pairs for updating
    ' Format: Array(Array("SheetColA", "TCHColA"), Array("SheetColB", "TCHColB"), ...)
    colPairs = Array(Array(1, 1), Array(2, 2), Array(3, 3), Array(5, 22), Array(6, 26), Array(7, 26), Array(8, 23), Array(9, 35), Array(10, 14), Array(11, 18), Array(12, 13), Array(17, 10), Array(18, 11), Array(19, 21))
    
    ' Disable screen updates and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each worksheet and update details
    For Each wsName In wsList
        Set ws = ThisWorkbook.Sheets(wsName)
        
        ' Filter "On Air" sites
        ws.AutoFilterMode = False
        lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
        ws.Range("A1:N" & lastRow).AutoFilter Field:=14, Criteria1:="On Air"
        
        ' Loop through filtered rows and update columns
        For i = 2 To lastRow
            If Not ws.Rows(i).Hidden Then
                siteID = Trim(ws.Cells(i, "D").Value)
                If siteDict.exists(siteID) Then
                    siteIndex = siteDict(siteID)
                    For Each pair In colPairs
                        shtCol = pair(0)
                        tchCol = pair(1)
                        ws.Cells(i, shtCol).Value = tchData(siteIndex, tchCol)
                    Next pair
                End If
            End If
        Next i
        
        ' Turn off the filter
        ws.AutoFilterMode = False
    Next wsName
    
    ' Re-enable screen updates and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Site details update complete", vbInformation
End Sub

Sub CopyLastColumnToDestinationWorkbook()
    Dim sourceWb As Workbook
    Dim destWb As Workbook
    Dim menuSheet As Worksheet
    Dim sheetNames As Variant
    Dim sourceSheet As Worksheet
    Dim destSheet As Worksheet
    Dim colToCopy As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim destPath As String
    Dim destFileName As String
    Dim fullPath As String
    
    ' Set the source workbook to this workbook
    Set sourceWb = ThisWorkbook
    
    ' Set the menu sheet
    Set menuSheet = sourceWb.Sheets("Menu")
    
    ' Read the path and file name from the menu sheet
    destPath = menuSheet.Range("S10").Value
    destFileName = menuSheet.Range("S11").Value
    fullPath = destPath & "\" & destFileName
    
    ' Disable screen updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Open the destination workbook invisibly
    Set destWb = Workbooks.Open(fullPath, ReadOnly:=False, AddToMru:=False, Notify:=False)
    destWb.Windows(1).Visible = False
    
    ' Array of sheet names
    sheetNames = Array("2G", "3G", "4G", "5G", "2G & 3G & 4G & 5G")
    
    ' Loop through each sheet name
    For Each sheetName In sheetNames
        ' Set the source and destination sheets
        Set sourceSheet = sourceWb.Sheets(sheetName)
        Set destSheet = destWb.Sheets(sheetName)
        
        ' Find the last column with data in the source sheet
        colToCopy = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
        
        ' Find the last row with data in the column to copy
        lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, colToCopy).End(xlUp).Row
        
        ' Find the last column with data in the destination sheet
        lastCol = destSheet.Cells(1, destSheet.Columns.Count).End(xlToLeft).Column + 1
        
        ' Copy the last column to the destination sheet
        sourceSheet.Range(sourceSheet.Cells(1, colToCopy), sourceSheet.Cells(lastRow, colToCopy)).Copy
        
        ' Paste the copied data as values and number formats to preserve date formatting
        With destSheet.Cells(1, lastCol)
            .PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            .PasteSpecial Paste:=xlPasteFormats
        End With
        
        Application.CutCopyMode = False
    Next sheetName
    
    ' Save and close the destination workbook properly
    destWb.Save
    destWb.Close SaveChanges:=True
    
    ' Re-enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Confirm completion
    MsgBox "Data copied successfully", vbInformation
End Sub

Sub UpdateSiteCommercialStatus()
    Dim wsTDI As Worksheet
    Dim wsList As Variant
    Dim ws As Worksheet
    Dim mapDict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim siteID As String
    Dim status As String
    
    ' Set the TCH sheet
    Set wsTDI = ThisWorkbook.Sheets("TDI")
    
    ' Create a dictionary for the TCH site status data
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    ' Load TCH data into the dictionary
    lastRow = wsTDI.Cells(wsTDI.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        siteID = Trim(wsTDI.Cells(i, "A").Value)
        status = wsTDI.Cells(i, "B").Value
        mapDict(siteID) = status
    Next i
    
    ' List of worksheets to update
    wsList = Array("2G", "3G", "4G", "5G", "2G & 3G & 4G & 5G")
    
    ' Disable screen updates and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each worksheet and update status
    For Each wsName In wsList
        Set ws = ThisWorkbook.Sheets(wsName)
        lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        For i = 2 To lastRow
            siteID = Trim(ws.Cells(i, "C").Value)
            If mapDict.exists(siteID) Then
                ws.Cells(i, "O").Value = mapDict(siteID)
            Else
                ws.Cells(i, "O").Value = "Off Air"
            End If
        Next i
    Next wsName
    
    ' Re-enable screen updates and calculations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Site status update complete", vbInformation
End Sub
