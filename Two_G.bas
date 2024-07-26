Attribute VB_Name = "Two_G"
Sub Update2GSheet()
    Dim ws2G As Worksheet
    Dim wsMAP As Worksheet
    Dim wsMenu As Worksheet
    Dim lastCol As Long
    Dim newCol As Long
    Dim siteID As String
    Dim i As Long
    Dim cell As Range
    Dim dateValue As String
    Dim mapDict As Object
    Dim primaryID As String
    Dim secondaryID As String
    Dim cellValue As Variant
    
    ' Set worksheets
    Set ws2G = ThisWorkbook.Sheets("2G")
    Set wsMAP = ThisWorkbook.Sheets("MAP")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    
    ' Get the date from Menu sheet cell L14
    dateValue = wsMenu.Range("L14").Value
    
    ' Find the last column in the 2G sheet
    lastCol = ws2G.Cells(1, ws2G.Columns.Count).End(xlToLeft).Column
    newCol = lastCol + 1
    
    ' Insert the date in the first row of the new column
    ws2G.Cells(1, newCol).Value = dateValue
    
    ' Create a dictionary for the MAP sheet data
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    ' Load MAP data into the dictionary
    For i = 2 To wsMAP.Cells(wsMAP.Rows.Count, "B").End(xlUp).Row
        If wsMAP.Cells(i, "C").Value = "" Then
            mapDict(wsMAP.Cells(i, "B").Value) = "-"
        Else
            mapDict(wsMAP.Cells(i, "B").Value) = wsMAP.Cells(i, "C").Value
        End If
    Next i
    
    ' Loop through each Site ID in column D of the 2G sheet
    For i = 2 To ws2G.Cells(ws2G.Rows.Count, "D").End(xlUp).Row
        siteID = ws2G.Cells(i, "D").Value
        
        ' Check if the Site ID contains a slash
        If InStr(siteID, "/") > 0 Then
            primaryID = Split(siteID, "/")(0)
            secondaryID = Split(siteID, "/")(1)
            
            ' Try to find the primary ID first
            If mapDict.exists(primaryID) Then
                ws2G.Cells(i, newCol).Value = mapDict(primaryID)
            ' If not found, try the secondary ID
            ElseIf mapDict.exists(secondaryID) Then
                ws2G.Cells(i, newCol).Value = mapDict(secondaryID)
            Else
                ws2G.Cells(i, newCol).Value = "-"
            End If
        Else
            ' Check if the Site ID exists in the dictionary
            If mapDict.exists(siteID) Then
                ' If found, get the corresponding value from the dictionary and print it in the new column in 2G sheet
                ws2G.Cells(i, newCol).Value = mapDict(siteID)
            Else
                ' If not found, print "-"
                ws2G.Cells(i, newCol).Value = "-"
            End If
        End If
    Next i
    
    ' Check status in column N and update the new column to "-" if status is "Off Air"
    For i = 2 To ws2G.Cells(ws2G.Rows.Count, "N").End(xlUp).Row
        If ws2G.Cells(i, "N").Value = "Off Air" Then
            ws2G.Cells(i, newCol).Value = "-"
        End If
    Next i
    
    ' Check values in the new column for less than 0 and greater than 100
    For i = 2 To ws2G.Cells(ws2G.Rows.Count, newCol).End(xlUp).Row
        cellValue = ws2G.Cells(i, newCol).Value
        If IsNumeric(cellValue) Then
            If cellValue < 0 Then
                ws2G.Cells(i, newCol).Value = 0
            ElseIf cellValue > 100 Then
                ws2G.Cells(i, newCol).Value = 100
            End If
        End If
    Next i
End Sub

Sub Update3GSheet()
    Dim ws3G As Worksheet
    Dim wsMAP As Worksheet
    Dim wsMenu As Worksheet
    Dim lastCol As Long
    Dim newCol As Long
    Dim siteID As String
    Dim i As Long
    Dim cell As Range
    Dim dateValue As String
    Dim mapDict As Object
    Dim primaryID As String
    Dim secondaryID As String
    Dim cellValue As Variant
    
    ' Set worksheets
    Set ws3G = ThisWorkbook.Sheets("3G")
    Set wsMAP = ThisWorkbook.Sheets("MAP")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    
    ' Get the date from Menu sheet cell L14
    dateValue = wsMenu.Range("L14").Value
    
    ' Find the last column in the 3G sheet
    lastCol = ws3G.Cells(1, ws3G.Columns.Count).End(xlToLeft).Column
    newCol = lastCol + 1
    
    ' Insert the date in the first row of the new column
    ws3G.Cells(1, newCol).Value = dateValue
    
    ' Create a dictionary for the MAP sheet data
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    ' Load MAP data into the dictionary
    For i = 2 To wsMAP.Cells(wsMAP.Rows.Count, "F").End(xlUp).Row
        If wsMAP.Cells(i, "G").Value = "" Then
            mapDict(wsMAP.Cells(i, "F").Value) = "-"
        Else
            mapDict(wsMAP.Cells(i, "F").Value) = wsMAP.Cells(i, "G").Value
        End If
    Next i
    
    ' Loop through each Site ID in column D of the 3G sheet
    For i = 2 To ws3G.Cells(ws3G.Rows.Count, "D").End(xlUp).Row
        siteID = ws3G.Cells(i, "D").Value
        
        ' Check if the Site ID contains a slash
        If InStr(siteID, "/") > 0 Then
            primaryID = Split(siteID, "/")(0)
            secondaryID = Split(siteID, "/")(1)
            
            ' Try to find the primary ID first
            If mapDict.exists(primaryID) Then
                ws3G.Cells(i, newCol).Value = mapDict(primaryID)
            ' If not found, try the secondary ID
            ElseIf mapDict.exists(secondaryID) Then
                ws3G.Cells(i, newCol).Value = mapDict(secondaryID)
            Else
                ws3G.Cells(i, newCol).Value = "-"
            End If
        Else
            ' Check if the Site ID exists in the dictionary
            If mapDict.exists(siteID) Then
                ' If found, get the corresponding value from the dictionary and print it in the new column in 3G sheet
                ws3G.Cells(i, newCol).Value = mapDict(siteID)
            Else
                ' If not found, print "-"
                ws3G.Cells(i, newCol).Value = "-"
            End If
        End If
    Next i
    
    ' Check status in column N and update the new column to "-" if status is "Off Air"
    For i = 2 To ws3G.Cells(ws3G.Rows.Count, "N").End(xlUp).Row
        If ws3G.Cells(i, "N").Value = "Off Air" Then
            ws3G.Cells(i, newCol).Value = "-"
        End If
    Next i
    
    ' Check values in the new column for less than 0 and greater than 100
    For i = 2 To ws3G.Cells(ws3G.Rows.Count, newCol).End(xlUp).Row
        cellValue = ws3G.Cells(i, newCol).Value
        If IsNumeric(cellValue) Then
            If cellValue < 0 Then
                ws3G.Cells(i, newCol).Value = 0
            ElseIf cellValue > 100 Then
                ws3G.Cells(i, newCol).Value = 100
            End If
        End If
    Next i
End Sub
Sub Update4GSheet()
    Dim ws4G As Worksheet
    Dim wsMAP As Worksheet
    Dim wsMenu As Worksheet
    Dim lastCol As Long
    Dim newCol As Long
    Dim siteID As String
    Dim i As Long
    Dim cell As Range
    Dim dateValue As String
    Dim mapDict As Object
    Dim primaryID As String
    Dim secondaryID As String
    Dim cellValue As Variant
    
    ' Set worksheets
    Set ws4G = ThisWorkbook.Sheets("4G")
    Set wsMAP = ThisWorkbook.Sheets("MAP")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    
    ' Get the date from Menu sheet cell L14
    dateValue = wsMenu.Range("L14").Value
    
    ' Find the last column in the 4G sheet
    lastCol = ws4G.Cells(1, ws4G.Columns.Count).End(xlToLeft).Column
    newCol = lastCol + 1
    
    ' Insert the date in the first row of the new column
    ws4G.Cells(1, newCol).Value = dateValue
    
    ' Create a dictionary for the MAP sheet data
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    ' Load MAP data into the dictionary
    For i = 2 To wsMAP.Cells(wsMAP.Rows.Count, "J").End(xlUp).Row
        If wsMAP.Cells(i, "K").Value = "" Then
            mapDict(wsMAP.Cells(i, "J").Value) = "-"
        Else
            mapDict(wsMAP.Cells(i, "J").Value) = wsMAP.Cells(i, "K").Value
        End If
    Next i
    
    ' Loop through each Site ID in column D of the 4G sheet
    For i = 2 To ws4G.Cells(ws4G.Rows.Count, "D").End(xlUp).Row
        siteID = ws4G.Cells(i, "D").Value
        
        ' Check if the Site ID contains a slash
        If InStr(siteID, "/") > 0 Then
            primaryID = Split(siteID, "/")(0)
            secondaryID = Split(siteID, "/")(1)
            
            ' Try to find the primary ID first
            If mapDict.exists(primaryID) Then
                ws4G.Cells(i, newCol).Value = mapDict(primaryID)
            ' If not found, try the secondary ID
            ElseIf mapDict.exists(secondaryID) Then
                ws4G.Cells(i, newCol).Value = mapDict(secondaryID)
            Else
                ws4G.Cells(i, newCol).Value = "-"
            End If
        Else
            ' Check if the Site ID exists in the dictionary
            If mapDict.exists(siteID) Then
                ' If found, get the corresponding value from the dictionary and print it in the new column in 4G sheet
                ws4G.Cells(i, newCol).Value = mapDict(siteID)
            Else
                ' If not found, print "-"
                ws4G.Cells(i, newCol).Value = "-"
            End If
        End If
    Next i
    
    ' Check status in column N and update the new column to "-" if status is "Off Air"
    For i = 2 To ws4G.Cells(ws4G.Rows.Count, "N").End(xlUp).Row
        If ws4G.Cells(i, "N").Value = "Off Air" Then
            ws4G.Cells(i, newCol).Value = "-"
        End If
    Next i
    
    ' Check values in the new column for less than 0 and greater than 100
    For i = 2 To ws4G.Cells(ws4G.Rows.Count, newCol).End(xlUp).Row
        cellValue = ws4G.Cells(i, newCol).Value
        If IsNumeric(cellValue) Then
            If cellValue < 0 Then
                ws4G.Cells(i, newCol).Value = 0
            ElseIf cellValue > 100 Then
                ws4G.Cells(i, newCol).Value = 100
            End If
        End If
    Next i
End Sub
Sub Update5GSheet()
    Dim ws5G As Worksheet
    Dim wsMAP As Worksheet
    Dim wsMenu As Worksheet
    Dim lastCol As Long
    Dim newCol As Long
    Dim siteID As String
    Dim i As Long
    Dim cell As Range
    Dim dateValue As String
    Dim mapDict As Object
    Dim primaryID As String
    Dim secondaryID As String
    Dim cellValue As Variant
    
    ' Set worksheets
    Set ws5G = ThisWorkbook.Sheets("5G")
    Set wsMAP = ThisWorkbook.Sheets("MAP")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    
    ' Get the date from Menu sheet cell L14
    dateValue = wsMenu.Range("L14").Value
    
    ' Find the last column in the 5G sheet
    lastCol = ws5G.Cells(1, ws5G.Columns.Count).End(xlToLeft).Column
    newCol = lastCol + 1
    
    ' Insert the date in the first row of the new column
    ws5G.Cells(1, newCol).Value = dateValue
    
    ' Create a dictionary for the MAP sheet data
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    ' Load MAP data into the dictionary
    For i = 2 To wsMAP.Cells(wsMAP.Rows.Count, "P").End(xlUp).Row
        If wsMAP.Cells(i, "O").Value = "" Then
            mapDict(wsMAP.Cells(i, "P").Value) = "-"
        Else
            mapDict(wsMAP.Cells(i, "P").Value) = wsMAP.Cells(i, "O").Value
        End If
    Next i
    
    ' Loop through each Site ID in column D of the 5G sheet
    For i = 2 To ws5G.Cells(ws5G.Rows.Count, "D").End(xlUp).Row
        siteID = ws5G.Cells(i, "D").Value
        
        ' Check if the Site ID contains a slash
        If InStr(siteID, "/") > 0 Then
            primaryID = Split(siteID, "/")(0)
            secondaryID = Split(siteID, "/")(1)
            
            ' Try to find the primary ID first
            If mapDict.exists(primaryID) Then
                ws5G.Cells(i, newCol).Value = mapDict(primaryID)
            ' If not found, try the secondary ID
            ElseIf mapDict.exists(secondaryID) Then
                ws5G.Cells(i, newCol).Value = mapDict(secondaryID)
            Else
                ws5G.Cells(i, newCol).Value = "-"
            End If
        Else
            ' Check if the Site ID exists in the dictionary
            If mapDict.exists(siteID) Then
                ' If found, get the corresponding value from the dictionary and print it in the new column in 5G sheet
                ws5G.Cells(i, newCol).Value = mapDict(siteID)
            Else
                ' If not found, print "-"
                ws5G.Cells(i, newCol).Value = "-"
            End If
        End If
    Next i
    
    ' Check status in column N and update the new column to "-" if status is "Off Air"
    For i = 2 To ws5G.Cells(ws5G.Rows.Count, "N").End(xlUp).Row
        If ws5G.Cells(i, "N").Value = "Off Air" Then
            ws5G.Cells(i, newCol).Value = "-"
        End If
    Next i
    
    ' Check values in the new column for less than 0 and greater than 100
    For i = 2 To ws5G.Cells(ws5G.Rows.Count, newCol).End(xlUp).Row
        cellValue = ws5G.Cells(i, newCol).Value
        If IsNumeric(cellValue) Then
            If cellValue < 0 Then
                ws5G.Cells(i, newCol).Value = 0
            ElseIf cellValue > 100 Then
                ws5G.Cells(i, newCol).Value = 100
            End If
        End If
    Next i
End Sub


Sub UpdateCombinedSheet(ByRef wsCombo As Worksheet, ByRef newCol As Long, ByRef dateValue As String)
    'Dim wsCombo As Worksheet
    Dim wsMAP As Worksheet
    Dim wsMenu As Worksheet
    Dim lastCol As Long
    'Dim newCol As Long
    Dim siteID As String
    Dim i As Long
    Dim cell As Range
   ' Dim dateValue As String
    Dim mapDict As Object
    
   ' Set worksheets
    Set wsCombo = ThisWorkbook.Sheets("2G & 3G & 4G & 5G")
    Set wsMAP = ThisWorkbook.Sheets("MAP")
    Set wsMenu = ThisWorkbook.Sheets("MENU")
    
    ' Get the date from Menu sheet cell L14
    dateValue = wsMenu.Range("L14").Value
    
    ' Find the last column in the 2G sheet
    lastCol = wsCombo.Cells(1, wsCombo.Columns.Count).End(xlToLeft).Column
    newCol = lastCol + 1
    
    ' Insert the date in the first row of the new column
    wsCombo.Cells(1, newCol).Value = dateValue
    
    ' Create a dictionary for the MAP sheet data
    Set mapDict = CreateObject("Scripting.Dictionary")
    
    ' Load MAP data into the dictionary
    For i = 2 To wsMAP.Cells(wsMAP.Rows.Count, "B").End(xlUp).Row
        If wsMAP.Cells(i, "C").Value = "" Then
            mapDict(wsMAP.Cells(i, "B").Value) = "-"
        Else
            mapDict(wsMAP.Cells(i, "B").Value) = wsMAP.Cells(i, "C").Value
        End If
    Next i
    
    ' Loop through each Site ID in column D of the 2G sheet
    For i = 2 To wsCombo.Cells(wsCombo.Rows.Count, "D").End(xlUp).Row
        siteID = wsCombo.Cells(i, "D").Value
        
        ' Check if the Site ID contains a slash
        If InStr(siteID, "/") > 0 Then
            primaryID = Split(siteID, "/")(0)
            secondaryID = Split(siteID, "/")(1)
            
            ' Try to find the primary ID first
            If mapDict.exists(primaryID) Then
                wsCombo.Cells(i, newCol).Value = mapDict(primaryID)
            ' If not found, try the secondary ID
            ElseIf mapDict.exists(secondaryID) Then
                wsCombo.Cells(i, newCol).Value = mapDict(secondaryID)
            Else
                wsCombo.Cells(i, newCol).Value = "-"
            End If
        Else
            ' Check if the Site ID exists in the dictionary
            If mapDict.exists(siteID) Then
                ' If found, get the corresponding value from the dictionary and print it in the new column in 2G sheet
                wsCombo.Cells(i, newCol).Value = mapDict(siteID)
            Else
                ' If not found, print "-"
                wsCombo.Cells(i, newCol).Value = "-"
            End If
        End If
    Next i
    
    ' Check status in column N and update the new column to "-" if status is "Off Air"
    For i = 2 To wsCombo.Cells(wsCombo.Rows.Count, "N").End(xlUp).Row
        If wsCombo.Cells(i, "N").Value = "Off Air" Then
            wsCombo.Cells(i, newCol).Value = "-"
        End If
    Next i
    
    ' Check values in the new column for less than 0 and greater than 100
    For i = 2 To wsCombo.Cells(wsCombo.Rows.Count, newCol).End(xlUp).Row
        cellValue = wsCombo.Cells(i, newCol).Value
        If IsNumeric(cellValue) Then
            If cellValue < 0 Then
                wsCombo.Cells(i, newCol).Value = 0
            ElseIf cellValue > 100 Then
                wsCombo.Cells(i, newCol).Value = 100
            End If
        End If
    Next i
    UpdateFromOtherSheets wsCombo, newCol, dateValue
End Sub

Sub UpdateFromOtherSheets(wsCombo As Worksheet, newCol As Long, dateValue As String)
    Dim ws3G As Worksheet
    Dim ws4G As Worksheet
    Dim ws5G As Worksheet
    Dim siteID As String
    Dim i As Long
    Dim cell As Range
    Dim col3G As Long
    Dim col4G As Long
    Dim col5G As Long
    
    ' Set worksheets
    Set ws3G = ThisWorkbook.Sheets("3G")
    Set ws4G = ThisWorkbook.Sheets("4G")
    Set ws5G = ThisWorkbook.Sheets("5G")
    
    ' Find the column with the date in the header for each sheet
    col3G = FindDateColumn(ws3G, dateValue)
    col4G = FindDateColumn(ws4G, dateValue)
    col5G = FindDateColumn(ws5G, dateValue)
    
    ' Filter "-" values in the new column
    wsCombo.AutoFilterMode = False
    wsCombo.Range(wsCombo.Cells(1, 1), wsCombo.Cells(wsCombo.Cells(wsCombo.Rows.Count, "D").End(xlUp).Row, newCol)).AutoFilter Field:=newCol, Criteria1:="-"
    
    ' Loop through filtered rows and update values from other sheets
    For i = 2 To wsCombo.Cells(wsCombo.Rows.Count, "D").End(xlUp).Row
        If wsCombo.Cells(i, newCol).Value = "-" Then
            siteID = wsCombo.Cells(i, "D").Value
            
            ' Check in 3G sheet
            If col3G > 0 Then
                Set cell = ws3G.Columns("D").Find(siteID, LookIn:=xlValues, LookAt:=xlWhole)
                If Not cell Is Nothing Then
                    wsCombo.Cells(i, newCol).Value = ws3G.Cells(cell.Row, col3G).Value
                    GoTo NextIteration
                End If
            End If
            
            ' Check in 4G sheet
            If col4G > 0 Then
                Set cell = ws4G.Columns("D").Find(siteID, LookIn:=xlValues, LookAt:=xlWhole)
                If Not cell Is Nothing Then
                    wsCombo.Cells(i, newCol).Value = ws4G.Cells(cell.Row, col4G).Value
                    GoTo NextIteration
                End If
            End If
            
            ' Check in 5G sheet
            If col5G > 0 Then
                Set cell = ws5G.Columns("D").Find(siteID, LookIn:=xlValues, LookAt:=xlWhole)
                If Not cell Is Nothing Then
                    wsCombo.Cells(i, newCol).Value = ws5G.Cells(cell.Row, col5G).Value
                End If
            End If
        End If
NextIteration:
    Next i
    
    ' Turn off the filter
    wsCombo.AutoFilterMode = False
End Sub

Function FindDateColumn(ws As Worksheet, dateValue As String) As Long
    Dim col As Long
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, col).Value = dateValue Then
            FindDateColumn = col
            Exit Function
        End If
    Next col
    FindDateColumn = 0 ' Return 0 if the date is not found
End Function
