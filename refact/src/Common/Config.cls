VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("GoodsCollectorProject.src.Common")
Option Explicit

Private Type TConfig
    AppName As String
    Section As String
End Type

Private this As TConfig

Public Function SetAppName(ByVal Name As String) As Config
    this.AppName = Name
    Set SetAppName = Me
End Function

Public Function SetSection(ByVal Name As String) As Config
    this.Section = Name
    Set SetSection = Me
End Function

Public Sub SetValue(ByVal Key As String, ByVal Value As String)
    CheckSettings

    SaveSetting this.AppName, this.Section, Key, Value
End Sub

Public Function GetValue( _
    ByVal Key As String, _
    Optional ByVal DefaultValue As Variant = Empty, _
    Optional ByVal CastType As VbVarType = VbVarType.vbString _
) As Variant
    CheckSettings

    GetValue = Cast(GetSetting(this.AppName, this.Section, Key, DefaultValue), CastType)
End Function

Private Sub CheckSettings()
    If Len(this.AppName) = 0 Then Err.Raise 91, TypeName(Me), "Не задана переменная AppName."
    If Len(this.Section) = 0 Then Err.Raise 91, TypeName(Me), "Не задана переменная Section."
End Sub

Private Function Cast(ByVal Value As Variant, ByVal CastType As VbVarType) As Variant
    Select Case CastType
    Case vbString: Cast = CStr(Value)
    Case vbBoolean
        On Error Resume Next
        Cast = CBool(Value)
        Cast = IIf(Err.Number = 0, Cast, False)
    End Select
End Function
