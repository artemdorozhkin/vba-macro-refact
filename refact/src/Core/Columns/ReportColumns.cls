VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "GoodsCollectorProject.src.Core.Columns"
Option Explicit

Private Type TReportColumns
    Item As TColumn
    Name As TColumn
    Quantity As TColumn
End Type

Private this As TReportColumns

Public Property Get Item() As TColumn
    Item = this.Item
End Property

Public Property Get Name() As TColumn
    Name = this.Name
End Property

Public Property Get Quantity() As TColumn
    Quantity = this.Quantity
End Property

Public Sub RegisterColumns(ByRef Table As Range)
    this.Item = ColumnTypes.FindColumn("Артикул", Table)
    this.Name = ColumnTypes.FindColumn("Наименование", Table)
    this.Quantity = ColumnTypes.FindColumn(Array( _
        "Количество", _
        "Кол-во", _
        "Кол." _
    ), Table)
End Sub

