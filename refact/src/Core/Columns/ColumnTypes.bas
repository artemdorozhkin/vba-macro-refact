Attribute VB_Name = "ColumnTypes"
'@Folder "GoodsCollectorProject.src.Core.Columns"
Option Explicit

Public Type TColumn
    Name As String
    Index As Integer
End Type

Public Function FindColumn(ByRef NameOrNames As Variant, ByRef Table As Range) As TColumn
    Dim Names As Variant
    If IsArray(NameOrNames) Then Names = NameOrNames Else Names = Array(NameOrNames)

    Dim Name As Variant
    For Each Name In Names
        Dim FoundCell As Range: Set FoundCell = Table.Find( _
            What:=Name, _
            LookIn:=XlFindLookIn.xlValues, _
            Lookat:=XlLookAt.xlWhole, _
            MatchCase:=False _
        )
        If Not FoundCell Is Nothing Then
            FindColumn.Name = FoundCell.Value
            FindColumn.Index = FoundCell.Column
            Exit Function
        End If
    Next

    Dim ErrMsg As String
    ErrMsg = GenerateErrMsg(Names, Table.Parent.Parent)

    Call Err.Raise( _
         Number:=9, _
         Source:="FindColumn", _
         Description:=ErrMsg _
    )
End Function

Public Function GenerateErrMsg(ByRef Names As Variant, ByRef Book As Workbook)
    GenerateErrMsg = "В книге '" & Book.Name & "' не удалось найти имя столбца по ключевым словам:" & _
        vbNewLine & Strings.Join(Names, vbNewLine)
End Function

