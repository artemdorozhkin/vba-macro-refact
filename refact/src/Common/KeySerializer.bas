Attribute VB_Name = "KeySerializer"
'@Folder("GoodsCollectorProject.src.Common")
Option Explicit

Const SEPARATOR As String = ";"

Public Function Stringify(ByRef KeyData As Variant) As String
    Stringify = Strings.Join(KeyData, SEPARATOR)
End Function

Public Function Parse(ByVal Key As String) As Variant
    Parse = Strings.Split(Key, SEPARATOR)
End Function
