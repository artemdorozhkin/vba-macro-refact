Attribute VB_Name = "LT20CollectorCstr"
'@Folder "GoodsCollectorProject.src.Core.GoodsCollector.LT20Collector"
Option Explicit

Public Function NewLT20Collector(ByRef Data As Variant, ByRef Columns As ReportColumns) As LT20Collector
    Set NewLT20Collector = New LT20Collector
    NewLT20Collector.Data = Data
    Set NewLT20Collector.Columns = Columns
End Function
