Attribute VB_Name = "Macro_trigger"
Sub TCH_Compilation()
    Dim wsCombo As Worksheet
    Dim newCol As Long
    Dim dateValue As String
    
    ' Update individual sheets
    Update2GSheet
    Update3GSheet
    Update4GSheet
    Update5GSheet
    
    ' Update combined sheet and get required values
    UpdateCombinedSheet wsCombo, newCol, dateValue
    
    ' Update "-" values from other sheets
    UpdateFromOtherSheets wsCombo, newCol, dateValue
    
    CopyFormattingAcrossSheets
    
    MsgBox "Complete", vbInformation
End Sub
Sub ReplaceLTE800()
    Dim ws As Worksheet
    Dim cell As Range
    Dim searchRange As Range
    Dim lastRow As Long
    Dim shtName As String
    
    ' Define the sheet name and set the worksheet
    shtName = "TCH" ' Change this to your actual sheet name
    Set ws = ThisWorkbook.Sheets(shtName)
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    ' Define the range to search
    Set searchRange = ws.Range("J1:J" & lastRow) ' Change the range if your data is in a different column
    
    ' Loop through each cell in the range and replace "LTE 800" with "4G"
    For Each cell In searchRange
        cell.Value = Replace(cell.Value, "LTE 800", "4G")
    Next cell
    
    MsgBox "Replacement complete", vbInformation
End Sub

Sub CopyFormattingAcrossSheets()
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim i As Long

    ' Define the names of the sheets to process
    sheetNames = Array("2G & 3G & 4G & 5G", "2G", "3G", "4G", "5G")

    ' Loop through each sheet name
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        
        ' Find the last column in the worksheet
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Copy formatting from the second-to-last column to the last column
        ws.Columns(lastCol - 1).Copy
        ws.Columns(lastCol).PasteSpecial Paste:=xlPasteFormats
        
        ' Clear the clipboard to remove the copied selection
        Application.CutCopyMode = False
    Next i
    
   ' MsgBox "Formatting copied across all specified sheets.", vbInformation
End Sub
