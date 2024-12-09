Attribute VB_Name = "Send_mail"
Option Explicit

Sub SendEmailWithInlineChartsAndTables()
    Dim outApp As Object
    Dim outMail As Object
    Dim chart1Path As String, chart2Path As String, chart3Path As String, chart4Path As String
    Dim excelFilePath As String
    Dim ws As Worksheet
    Dim tempWorkbook As Workbook
    Dim tableSheet As Worksheet
    Dim rngTable1 As Range, rngTable2 As Range, rngTable10 As Range
    Dim table1HTML As String, table2HTML As String, table10HTML As String
    Dim overallAvailability As Double
    Dim availabilityColor As String
    Dim availabilityMessage As String

    ' Paths to the chart images
    chart1Path = "C:\Temp\Overall.png"
    chart2Path = "C:\Temp\Zone.png"
    chart3Path = "C:\Temp\Region.png"
    chart4Path = "C:\Temp\Priority.png"

    ' Path to save the copied sheet
    excelFilePath = "C:\Temp\PA Trend.xlsx"

    ' Copy the 'PA Trend' sheet and save it as a new workbook
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PA Trend")
    If Not ws Is Nothing Then
        ' Delete the existing file if it exists
        If Dir(excelFilePath) <> "" Then
            Kill excelFilePath
        End If

        ' Copy the sheet to a new workbook
        ws.Copy
        Set tempWorkbook = ActiveWorkbook

        ' Save the new workbook to the specified path
        tempWorkbook.SaveAs Filename:=excelFilePath, FileFormat:=xlOpenXMLWorkbook
        tempWorkbook.Close SaveChanges:=False
    Else
        MsgBox "Sheet 'PA Trend' not found in the workbook.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Get Table1, Table2, and Table10 from the "Table" sheet
    Set tableSheet = ThisWorkbook.Sheets("Table")
    Set rngTable1 = tableSheet.ListObjects("Table1").Range
    Set rngTable2 = tableSheet.ListObjects("Table2").Range
    Set rngTable10 = tableSheet.ListObjects("Table10").Range

    ' Get HTML representation of the tables
    table1HTML = GetHTMLFromRange(rngTable1)
    table2HTML = GetHTMLFromRange(rngTable2)
    table10HTML = GetHTMLFromRange(rngTable10)

    ' Retrieve the overall availability value from cell A39 of the Table sheet
    overallAvailability = tableSheet.Cells(39, 1).Value

    ' Determine the color based on the overall availability value
    If overallAvailability >= 99.7 Then
        availabilityColor = "#C6EFCE" ' Light green
    Else
        availabilityColor = "#FFC7CE" ' Light red
    End If

    ' Create the availability message with the appropriate formatting
    availabilityMessage = "<p>Kindly note that the current overall network availability is " & _
                          "<span style='font-size:16pt; font-weight:bold; color:black; background-color:" & availabilityColor & ";'>" & _
                          Format(overallAvailability, "0.00") & "</span></p>"

    ' Create Outlook application
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)

    ' Compose the email
    On Error Resume Next
    With outMail
        ' Set recipients
        .To = "na_performance@ihstowers.com"
        .CC = "israel.fehintola@ihstowers.com; monday.attah@ihstowers.com; ugbede.achobe@ihstowers.com; solomon.ogho@ihstowers.com; olayemi.awofisoye@ihstowers.com; nehme@ihstowers.com; ihs_nocmanagers@ihstowers.com; oritsematosan.onosode@ihstowers.com; oluwafemi.okodugha@ihstowers.com"
    '    .BCC = "igwebuike.udeme@ihstowers.com"
    '    .BCC = "igwebuike.udeme@ihstowers.com"
        ' Set email subject and body
        .Subject = "Hourly PA Trend " & Format(Now, "mm-dd-yyyy hh:00")
        .HTMLBody = "<p>Dear Team,</p>" & _
                    "<p>Please find the hourly PA trend as at " & Format(Now, "h:00 AM/PM") & ":</p>" & _
                    availabilityMessage & _
                    "<p>Additionally, the attached Excel file contains the 'PA Trend' data.</p>" & _
                    "<p><img src=""cid:Overall""></p>" & _
                    "<p><img src=""cid:Zone""></p>" & _
                    "<p>REGIONAL PA ON THE HOUR:</p>" & _
                    table10HTML & _
                    "<p>PA TREND BY REGION:</p>" & _
                    table1HTML & _
                    "<p>PA TREND BY MTN PRIORITY:</p>" & _
                    table2HTML & _
                    "<p>Regards,<br>Israel</p>"

        ' Set email importance (1 = High, 2 = Normal, 0 = Low)
        .Importance = 1 ' Set to high importance

        ' Set "Send on Behalf Of" (replace with your desired sender)
        .SentOnBehalfOfName = "na_performance@ihstowers.com"

        ' Add the first chart as an inline attachment
        .Attachments.Add chart1Path
        .Attachments.Item(.Attachments.Count).PropertyAccessor.SetProperty _
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Overall"

        ' Add the second chart as an inline attachment
        .Attachments.Add chart2Path
        .Attachments.Item(.Attachments.Count).PropertyAccessor.SetProperty _
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Zone"

        ' Attach the copied 'PA Trend' sheet as an Excel file
        .Attachments.Add excelFilePath

        ' Send the email
        .Send
    End With
    On Error GoTo 0

    ' Clean up
    Set outMail = Nothing
    Set outApp = Nothing
End Sub

Function GetHTMLFromRange(rng As Range) As String
    Dim htmlString As String
    Dim cell As Range
    Dim rowRange As Range
    Dim cellValue As Double

    ' Start the HTML table with styling
    htmlString = "<table border='1' cellspacing='0' cellpadding='5' style='border-collapse:collapse; font-family: Arial, sans-serif; font-size: 10pt;'>" & _
                 "<thead style='background-color: #BDD7EE; font-weight: bold;'>"

    ' Create header row
    htmlString = htmlString & "<tr>"
    For Each cell In rng.rows(1).Cells
        htmlString = htmlString & "<th style='border: 1px solid #000; padding: 5px;'>" & cell.Text & "</th>"
    Next cell
    htmlString = htmlString & "</tr></thead><tbody>"

    ' Loop through each row in the range (starting from the second row to skip the header row)
    For Each rowRange In rng.rows
        If rowRange.Row <> rng.rows(1).Row Then
            htmlString = htmlString & "<tr>"
            ' Loop through each cell in the row
            For Each cell In rowRange.Cells
                ' Check if the cell contains a numeric value and format accordingly
                If IsNumeric(cell.Value) Then
                    cellValue = CDbl(cell.Value)
                    If cellValue >= 99.7 Then
                        htmlString = htmlString & "<td style='border: 1px solid #000; padding: 5px; text-align: center; background-color: #C6EFCE;'>" & cell.Text & "</td>"
                    Else
                        htmlString = htmlString & "<td style='border: 1px solid #000; padding: 5px; text-align: center; background-color: #FFC7CE;'>" & cell.Text & "</td>"
                    End If
                Else
                    ' If the cell does not contain a numeric value, use default formatting
                    htmlString = htmlString & "<td style='border: 1px solid #000; padding: 5px; text-align: center;'>" & cell.Text & "</td>"
                End If
            Next cell
            htmlString = htmlString & "</tr>"
        End If
    Next rowRange

    ' Close the HTML table
    htmlString = htmlString & "</tbody></table>"

    ' Return the HTML string
    GetHTMLFromRange = htmlString
End Function


