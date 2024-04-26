Attribute VB_Name = "Module1"
Option Explicit

Sub Обработка()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim f As String, bm As Boolean
    Dim twb, workb As Workbook
    Dim sh As Worksheet
    Dim lr As Long, lc As Long, i As Long, ii As Long, iii As Long
    Dim arr
    Dim tovary_cotoryx_bolshe_than_20 As Object: Set tovary_cotoryx_bolshe_than_20 = CreateObject("Scripting.Dictionary")
    Dim tovary_cotoryx_menshe_chem_20 As Object: Set tovary_cotoryx_menshe_chem_20 = CreateObject("Scripting.Dictionary")
    Dim q, v, g
    Dim sh2 As Worksheet

    f = Sheets(1).Range("B1").Value
    bm = Sheets(1).Range("B2").Value = "0"
    
    If Len(f) = 0 Then
        MsgBox "Путь не указан!", vbCritical, "Ошибка"
        Exit Sub
    End If

    Set twb = ThisWorkbook
    Workbooks.Open f
    Set workb = ActiveWorkbook
    Set sh = workb.ActiveSheet

    If sh.Range("B1").Value <> "Артикул" And sh.Range("B1").Value <> "АРТИКУЛ" Then
        MsgBox "Не верный формат файла!", vbCritical, "Ошибка"
        Exit Sub
    End If

    If LCase(sh.Range("C1").Value) <> "наименование" Then
        MsgBox "Не верный формат файла!", vbCritical, "Ошибка"
        Exit Sub
    End If

    If sh.Range("E1").Value Like "*оличество" = False Then
        MsgBox "Не верный формат файла!", vbCritical, "Ошибка"
        Exit Sub
    End If

    lr = workb.ActiveSheet.Rows(Rows.Count).End(xlUp).Row
    lc = workb.ActiveSheet.Columns(Columns.Count).End(xlToLeft).Column
    arr = workb.ActiveSheet.Range(Cells(1, 1), Cells(lr, lc)).Value

    For i = 1 To lr
        If IsNumeric(arr(i, 5)) Then
            If arr(i, 5) > 20 Then
                If tovary_cotoryx_bolshe_than_20.Exists(arr(i, 2) & ";" & arr(i, 3)) = True Then
                    q = tovary_cotoryx_bolshe_than_20(arr(i, 2) & ";" & arr(i, 3)) = arr(i, 5)
                    tovary_cotoryx_bolshe_than_20(arr(i, 2) & ";" & arr(i, 3)) = q
                Else
                    tovary_cotoryx_bolshe_than_20.Add arr(i, 2) & ";" & arr(i, 3), arr(i, 5)
                End If
            
            ElseIf arr(i, 5) = 0 Then
            
            Else
                
                If tovary_cotoryx_menshe_chem_20.Exists(arr(i, 2) & ";" & arr(i, 3)) = True Then
                    q = tovary_cotoryx_menshe_chem_20(arr(i, 2) & ";" & arr(i, 3)) = arr(i, 5)
                    tovary_cotoryx_menshe_chem_20(arr(i, 2) & ";" & arr(i, 3)) = q
                Else
                    tovary_cotoryx_menshe_chem_20.Add arr(i, 2) & ";" & arr(i, 3), arr(i, 5)
                End If
    
            End If
        End If
    Next

    workb.Sheets.Add
    Set sh2 = workb.Sheets(1)

    If bm Then
        v = tovary_cotoryx_bolshe_than_20.Items()
        g = tovary_cotoryx_bolshe_than_20.Keys()
    Else
        v = tovary_cotoryx_menshe_chem_20.Items()
        g = tovary_cotoryx_menshe_chem_20.Keys()
    End If
    For ii = LBound(g) To UBound(g)
        sh2.Cells(ii + 1, 1).Value = Split(g(ii), ";")(0)
        sh2.Cells(ii + 1, 2).Value = Split(g(ii), ";")(1)
    Next
    For iii = LBound(v) To UBound(v)
        sh2.Cells(iii + 1, 3).Value = v(iii)
    Next

    workb.SaveAs twb.Path & "\" & Format(Now, "ddmmyyhhnn") & ".xlsx"

    workb.Close False
    
    MsgBox "Готово!", vbInformation, "Готово"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
