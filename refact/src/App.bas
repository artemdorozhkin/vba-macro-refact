Attribute VB_Name = "App"
'@Folder "GoodsCollectorProject.src"
Option Explicit

'@EntryPoint
Public Sub Main()
    On Error GoTo Catch

    #If DEV Then
        MsgBox "Макрос запущен в режиме разработки.", vbInformation, "DEV MODE"
    #End If

    Utils.DisableSettings

    Dim Config As Config: Set Config = NewConfig(Constants.APP_NAME_ENG).SetSection("Settings")

    #If DEV Then
        Const FilePath As String = "C:\dev\projects\vba\refact\data.xlsx"
        Dim Kind As CollectorKind: Kind = CollectorKind.GT20
    #Else
        Dim FilePath As String
        FilePath = Config.GetValue(FormTags.FilePath, vbString)
        Dim Kind As CollectorKind
        Kind = CollectorTypes.GetCollectorKind(Config)
    #End If

    Dim Result As TCheckResult: Result = PathChecker.Validate(FilePath)
    If Result.HasError Then
        MsgBox Result.Message, vbExclamation, "Ошибка пути к файлу"
        GoTo ExitSub
    End If

    Dim Book As ReportBook: Set Book = NewReportBook(FilePath)
    Result = Book.Validate()
    If Result.HasError Then
        MsgBox Result.Message, vbExclamation, "Ошибка файла"
        GoTo ExitSub
    End If

    Dim Collector As IGoodsCollector
    Set Collector = GoodsCollectorFactory.GetCollector( _
        Kind, Book.GetData(), Book.Columns _
    )

    Dim Goods As Object: Set Goods = Collector.Collect()
    Book.SaveData Goods
    Book.CloseReport

    MsgBox "Данные успешно сформированы.", vbInformation, "Выполнено"
ExitSub:
    ' Чистый выход
    Utils.EnableSetting
Exit Sub

Catch:
    ' Ловим ошибку
    Call MsgBox( _
        "Критическая ошибка." & vbNewLine & vbNewLine & Err.Description, _
        vbCritical, _
        "Ошибка выполнения" _
    )
    Debug.Print "#" & Err.Number, Err.Description
    Resume ExitSub
End Sub

