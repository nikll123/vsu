Attribute VB_Name = "checkAll"
Function getWordText(fileName) As String
    On Error GoTo err
    Set appWord = CreateObject("Word.Application")
    Set WrdDoc = appWord.Documents.Open(fileName)
    getWordText = WrdDoc.Content.text
    WrdDoc.Close
    GoTo fine
err:
    MsgBox "Error on the file opening: " & fileName & vbCrLf & err.Description
    getWordText = "err"
    
fine:
    If IsObject(appWord) Then
        appWord.Quit
        Set appWord = Nothing
    End If
End Function


Sub test()
    Debug.Print getPbdFileName(Now())
    'qq = getWordText("R:\БР\бр\Бойове розпорядження з РЕБ рРЕБ Бетон нова 158 омбр 102025.docx11")
End Sub


Function getPbdFileName(d1 As Date) As String
    date1 = Format(d1, "dd.mm.yyyy")
    path1 = "R:\БР\ПідсумковеБД\"
    Const dateStart = #8/2/2026#  ' 2 aug 2026
    pbdNumStart = 430
    pbdNum = pbdNumStart + Int(d1) - dateStart
    fn = Trim(Str(pbdNum)) + ".ПБД_Ф_5.18_рРЕБ_" & date1 & ".docx"
    getPbdFileName = path1 & fn
End Function

Sub checkGBD()
    Debug.Assert ActiveSheet.Name = "Табель"
    Dim date1 As Date
    Dim txt1 As String
    Dim txt0 As String
    Dim fio As String
    If Not testDate2(Format(ActiveCell.Value)) Then
        MsgBox "Неверная дата: " & Format(ActiveCell.Value)
        Exit Sub
    End If
    date1 = ActiveCell.Value
    
    c = ActiveCell.Column
    cFio = 8
    fn = getPbdFileName(date1)
    Debug.Print fn
    txt1 = getWordText(fn)
    If txt1 = "err" Then Exit Sub
    resMsg = ""
    'R:\БР\ПідсумковеБД\488.ПБД_Ф_5.18_рРЕБ_29.09.2026.docx
    nnn = Mid(fn, 20, 3)
    ddd = Mid(fn, 40, 10)
    txt0 = "№ " & nnn & " дск"
    If InStr2(txt1, txt0) = 0 Then
        MsgBox "Нет правильного номера дск: " & txt0
        Exit Sub
    End If
    txt0 = "21:00 " & ddd & ". Карта 1:25000, видання 2025 року"
    If InStr2(txt1, txt0) = 0 Then
        MsgBox "Нет правильной фразы с датой: " & txt0
        Exit Sub
    End If
    
    For r = 9 To 63
        fio = ActiveSheet.Cells(r, cFio).Value
        cnt = InStr2(txt1, fio)
        money = ActiveSheet.Cells(r, c).Value
        If 100 = money And cnt <> 1 Or money <> 100 And cnt > 0 Then
            resMsg = resMsg & "Проблемы с " & fio & ": " & money & " " & cnt & vbCrLf
        Else
            Debug.Print fio & " - " & cnt & " - "; money
        End If
    Next
    If resMsg = "" Then
        resMsg = date1 & vbCrLf & " Все ок"
    End If
    MsgBox resMsg
End Sub

Function InStr2(strWhere As String, strWhat As String) As Integer
    InStr2 = 0
    If strWhat = "" Then Exit Function
    in1 = InStr(1, strWhere, strWhat, vbTextCompare)
    If in1 > 0 Then
        InStr2 = 1
        in2 = InStr(in1 + 1, strWhere, strWhat, vbTextCompare)
        If in2 > 0 Then
            InStr2 = 2
            Exit Function
        End If
    End If
End Function
