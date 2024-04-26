VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GT20Collector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "GoodsCollectorProject.src.Core.GoodsCollector.GT20Collector"
Option Explicit

Implements IGoodsCollector

Private Type TGT20Collector
    Data As Variant
    Columns As ReportColumns
End Type

Private this As TGT20Collector

Public Property Get Data() As Variant
    Data = this.Data
End Property
Public Property Let Data(ByVal RHS As Variant)
    this.Data = RHS
End Property

Public Property Get Columns() As ReportColumns
    Set Columns = this.Columns
End Property
Public Property Set Columns(ByVal RHS As ReportColumns)
    Set this.Columns = RHS
End Property

Public Function Collect() As Object
    Dim Goods As Object: Set Goods = NewDictionary()

    Dim Row As Long
    For Row = LBound(this.Data, 1) To UBound(this.Data, 1)
        Dim Quantity As Variant: Quantity = Data(Row, this.Columns.Quantity.Index)
        If Not IsNumeric(Quantity) Then GoTo Continue
        If Not Quantity > 20 Then GoTo Continue

        Dim Key As String: Key = GenerateKey(Row)
        Goods(Key) = Goods(Key) + Quantity
Continue:
    Next

    Set Collect = Goods
End Function

Public Function GenerateKey(ByVal Row As Long) As String
    Dim KeyData As Variant: KeyData = Array( _
        this.Data(Row, this.Columns.Item.Index), _
        this.Data(Row, this.Columns.Name.Index) _
    )

    GenerateKey = KeySerializer.Stringify(KeyData)
End Function

Private Function IGoodsCollector_Collect() As Object
    Set IGoodsCollector_Collect = Collect()
End Function
