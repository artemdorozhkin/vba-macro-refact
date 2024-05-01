Attribute VB_Name = "CollectorTypes"
'@Folder("GoodsCollectorProject.src.Core.GoodsCollector")
Option Explicit

Public Enum CollectorKind
    GT20
    LT20
End Enum

Public Function GetCollectorKind(ByRef Config As Config) As CollectorKind
    If Config.GetValue(FormTags.GT20Collector, CastType:=vbBoolean) Then
        GetCollectorKind = GT20
    ElseIf Config.GetValue(FormTags.LT20Collector, CastType:=vbBoolean) Then
        GetCollectorKind = LT20
    Else
        Err.Raise 9, "GetCollectorKind", "Выбран неизвестный тип сборщика."
    End If
End Function
