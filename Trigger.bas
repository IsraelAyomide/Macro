Attribute VB_Name = "Trigger"
Option Explicit

Sub Trigger()
Call ProcessInputData
Call FilterRCAsAndCopyToPassive
Call FilterRCAsAndCopyToActive
Call ConsolidateOutages
Call TruncateOutagesOpt
Call PopulateHourlyPATrend
Call CopyChartsToClipboard
Call CalculateRegionAvailabilityAverage
Call SendEmailWithInlineChartsAndTables

MsgBox "Completed", vbInformation
End Sub

