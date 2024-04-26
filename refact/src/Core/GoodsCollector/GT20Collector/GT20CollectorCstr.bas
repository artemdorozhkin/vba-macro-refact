Attribute VB_Name = "GT20CollectorCstr"
'@Folder "GoodsCollectorProject.src.Core.GoodsCollector.GT20Collector"
Option Explicit

Public Function NewGT20Collector(ByRef Data As Variant, ByRef Columns As ReportColumns) As GT20Collector
    Set NewGT20Collector = New GT20Collector
    NewGT20Collector.Data = Data
    Set NewGT20Collector.Columns = Columns
End Function
