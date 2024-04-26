Attribute VB_Name = "PathChecker"
'@Folder "GoodsCollectorProject.src.Common"
Option Explicit

Public Function Validate(ByVal Path As String) As TCheckResult
    If Len(Path) = 0 Then
        Validate.HasError = True
        Validate.Message = "Путь не указан."
        Exit Function
    End If

    Dim FSO As Object: Set FSO = Utils.NewFileSystemObject()
    If Not FSO.FileExists(Path) Then
        Validate.HasError = True
        Validate.Message = "Файл не существует или передан не корректный путь."
        Exit Function
    End If
End Function
