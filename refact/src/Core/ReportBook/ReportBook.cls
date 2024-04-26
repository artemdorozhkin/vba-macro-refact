VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "GoodsCollectorProject.src.Core.ReportBook"
Option Explicit

Private Type TReportBook
    Book As Workbook
    Path As String
    Columns As ReportColumns
End Type

Private this As TReportBook

Public Property Get Columns() As ReportColumns
    Set Columns = this.Columns
End Property

Public Property Get Path() As String
    Path = this.Path
End Property
Public Property Let Path(ByVal RHS As String)
    this.Path = RHS
End Property

Public Function Validate() As TCheckResult
    Set this.Book = Workbooks.Open(this.Path)
    On Error GoTo Catch
    Set this.Columns = NewReportColumns(this.Book.ActiveSheet.UsedRange)
Exit Function

Catch:
    Validate.HasError = True
    Validate.Message = Err.Description
    Resume Next
End Function

Public Function GetData() As Variant
    GetData = this.Book.ActiveSheet.UsedRange.Value
End Function

Public Sub SaveData(ByRef Goods As Object)
    Dim DataSheet As Worksheet: Set DataSheet = this.Book.Worksheets.Add()

    Dim Key As Variant
    For Each Key In Goods
        Dim Name As String: Name = KeySerializer.Parse(Key)(1)
        Dim Quantity As Long: Quantity = Goods(Key)

        Dim Row As Long: Row = Row + 1
        DataSheet.Cells(Row, 1).Value = Name
        DataSheet.Cells(Row, 2).Value = Quantity
    Next

    Dim FileName As String: FileName = GenerateFileName()
    this.Book.SaveAs FileName
End Sub

Public Sub CloseReport()
    this.Book.Close SaveChanges:=False
End Sub

Private Function GenerateFileName() As String
    Dim FSO As Object: Set FSO = NewFileSystemObject()
    Dim FileName As String: FileName = Strings.Format(DateTime.Now, "ddmmyyhhnn") & ".xlsx"

    GenerateFileName = FSO.BuildPath(this.Book.Path, FileName)
End Function
