VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Настройки"
   ClientHeight    =   5184
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7836
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "GoodsCollectorProject.src.UI"
Option Explicit

Private Type TSettingsForm
    Controller As SettingsFormController
End Type

Private this As TSettingsForm

Private Sub UserForm_Initialize()
    SetTags

    Set this.Controller = NewSettingsFormController( _
        Form:=Me, _
        Config:=NewConfig(Constants.APP_NAME_ENG).SetSection("Settings") _
    )

    this.Controller.LoadValues
End Sub

Private Sub SetTags()
    Me.FilePathTextBox.Tag = FormTags.FilePath
    Me.GT20OptionButton.Tag = FormTags.GT20Collector
    Me.LT20OptionButton.Tag = FormTags.LT20Collector
End Sub

Private Sub SelectFilePathCommandButton_Click()
    Dim Path As Variant
    Path = Application.GetOpenFilename( _
        Title:="Укажите путь к файлу отчета", _
        FileFilter:="Excel Файлы (*.xls*), *.xls*" _
    )

    If IsBoolType(Path) Then Exit Sub
    Me.FilePathTextBox.Value = Path
End Sub

Private Sub RunCommandButton_Click()
    Dim Check As TCheckResult
    Check = this.Controller.IsPrepairedToRun( _
        Me.FilePathTextBox, _
        Me.GT20OptionButton, _
        Me.LT20OptionButton _
    )

    If Check.HasError Then
        MsgBox Check.Message, vbExclamation, "Ошибка запуска"
        Exit Sub
    End If

    this.Controller.SaveValues

    App.Main
    Unload Me
End Sub

Private Sub CloseCommandButton_Click()
    Unload Me
End Sub

