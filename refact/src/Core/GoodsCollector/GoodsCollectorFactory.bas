Attribute VB_Name = "GoodsCollectorFactory"
'@Folder("GoodsCollectorProject.src.Core.GoodsCollector")
Option Explicit

Public Function GetCollector(ByVal Kind As CollectorKind, ByRef Data As Variant, ByRef Columns As ReportColumns) As IGoodsCollector
    Select Case Kind
    Case CollectorKind.GT20
        Set GetCollector = NewGT20Collector(Data, Columns)
    Case CollectorKind.LT20
        Set GetCollector = NewLT20Collector(Data, Columns)
    Case Else
        Call Err.Raise( _
            Number:=9, _
            Source:="GetCollector", _
            Description:="Не удалось определить тип сборщика: " & Kind _
        )
    End Select
End Function

