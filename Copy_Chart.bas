Attribute VB_Name = "Copy_Chart"
Option Explicit

Sub CopyChartsToClipboard()
    Dim ws As Worksheet
    Dim chart1Path As String, chart2Path As String, chart3Path As String, chart4Path As String
    Dim folderPath As String

    ' Define folder path
    folderPath = "C:\Temp\"

    ' Check if the folder exists; if not, create it
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DASHBOARD")

    ' Define file paths for the images
    chart1Path = folderPath & "Overall.png"
    chart2Path = folderPath & "Zone.png"
    chart3Path = folderPath & "Region.png"
    chart4Path = folderPath & "Priority.png"

    ' Export Chart 1
    ws.ChartObjects("Overall").Chart.Export Filename:=chart1Path, FilterName:="PNG"

    ' Export Chart 2
    ws.ChartObjects("Zone").Chart.Export Filename:=chart2Path, FilterName:="PNG"
    ' Export Chart 1
'    ws.ChartObjects("Region").Chart.Export Filename:=chart3Path, FilterName:="PNG"
'
'    ' Export Chart 2
'    ws.ChartObjects("Priority").Chart.Export Filename:=chart4Path, FilterName:="PNG"

   ' MsgBox "Charts exported successfully!", vbInformation
End Sub
