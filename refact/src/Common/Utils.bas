Attribute VB_Name = "Utils"
'@Folder "GoodsCollectorProject.src.Common"
Option Explicit

Public Sub DisableSettings()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub EnableSetting()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Function NewFileSystemObject() As Object
    Set NewFileSystemObject = CreateObject("Scripting.FileSystemObject")
End Function

Public Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function
