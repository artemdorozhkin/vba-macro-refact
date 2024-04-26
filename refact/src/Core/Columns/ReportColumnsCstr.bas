Attribute VB_Name = "ReportColumnsCstr"
'@Folder "GoodsCollectorProject.src.Core.Columns"
Option Explicit

Public Function NewReportColumns(ByRef Table As Range) As ReportColumns
    Set NewReportColumns = New ReportColumns
    NewReportColumns.RegisterColumns Table
End Function
