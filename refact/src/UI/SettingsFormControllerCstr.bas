Attribute VB_Name = "SettingsFormControllerCstr"
'@Folder("GoodsCollectorProject.src.UI")
Option Explicit

Public Function NewSettingsFormController(ByRef Form As MSForms.UserForm, ByRef Config As Config) As SettingsFormController
    Set NewSettingsFormController = New SettingsFormController
    Set NewSettingsFormController.Form = Form
    Set NewSettingsFormController.Config = Config
End Function
