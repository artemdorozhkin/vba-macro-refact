VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsFormController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("GoodsCollectorProject.src.UI")
Option Explicit

Private Type TSettingsFormController
    Config As Config
    Form As MSForms.UserForm
End Type

Private this As TSettingsFormController

Public Property Set Config(ByVal RHS As Config)
    Set this.Config = RHS
End Property

Public Property Set Form(ByVal RHS As MSForms.UserForm)
    Set this.Form = RHS
End Property

Public Function IsPrepairedToRun(ParamArray Controls() As Variant) As TCheckResult
    Dim OptionSelected As Boolean

    Dim Control As Variant
    For Each Control In Controls
        If IsTextBox(Control) Then
            If Strings.Len(Control.Value) = 0 Then
                IsPrepairedToRun.HasError = True
                IsPrepairedToRun.Message = "Не указан путь к файлу."
                Exit Function
            ElseIf Not NewFileSystemObject().FileExists(Control.Value) Then
                IsPrepairedToRun.HasError = True
                IsPrepairedToRun.Message = "Указанный файл не найден."
                Exit Function
            End If
        ElseIf IsOptionButton(Control) Then
            If Not OptionSelected Then OptionSelected = Control.Value
        Else
            GoTo Continue
        End If
Continue:
    Next

    If Not OptionSelected Then
        IsPrepairedToRun.HasError = True
        IsPrepairedToRun.Message = "Не выбран ни один тип сборки."
    End If
End Function

Public Sub LoadValues()
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If Strings.Len(Control.Tag) = 0 Then GoTo Continue

        Dim CastType As VbVarType
        If IsTextBox(Control) Then
            CastType = vbString
        ElseIf IsOptionButton(Control) Then
            CastType = vbBoolean
        Else
            GoTo Continue
        End If

        Dim Value As Variant
        Value = this.Config.GetValue(Control.Tag, CastType:=CastType)
        If IsEmpty(Value) Then GoTo Continue
        Control.Value = Value
Continue:
    Next
End Sub

Public Sub SaveValues()
    Dim Control As MSForms.Control
    For Each Control In this.Form.Controls
        If Not IsTextBox(Control) _
        And Not IsOptionButton(Control) Then GoTo Continue

        Me.SaveValue Control
Continue:
    Next
End Sub

Public Sub SaveValue(ByRef Control As Control)
    this.Config.SetValue Control.Tag, Control.Value
End Sub
