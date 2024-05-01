Attribute VB_Name = "ConfigCstr"
'@Folder("GoodsCollectorProject.src.Common")
Option Explicit

Public Function NewConfig(ByVal AppName As String) As Config
    Set NewConfig = New Config
    NewConfig.SetAppName AppName
End Function
