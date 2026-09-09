Attribute VB_Name = "Tabel_PD"
'InStr(44,"zsxg xh dfТяпко Олексій  fxcТяпко Олексій Тяпко Олексій h ","Тяпко Олексій ",vbTextCompare)

''''''''Function GetOpenFilen() As String
''''''''    Dim filePath As Variant
''''''''
''''''''    ' Вызываем диалоговое окно
''''''''    filePath = Application.GetOpenFilename(FileFilter:="Документы Word", "*.doc; *.docx",Title:="Выберите файл для работы",MultiSelect:=False)
''''''''
''''''''    ' Проверяем, выбрал ли пользователь файл (а не нажал Отмена)
''''''''    If filePath <> False Then
''''''''        MsgBox "Вы выбрали файл: " & filePath, vbInformation
''''''''        ' Здесь можно написать код открытия файла, например:
''''''''        ' Workbooks.Open filePath
''''''''    Else
''''''''        MsgBox "Выбор отменен.", vbExclamation
''''''''    End If
''''''''End Function

Function textDocWord() As String
'    Dim filePath As String
'    filePath = "C:\C:\YourFolder\Document.docx" ' Замените на реальный путь
    Dim fd As FileDialog
    Dim wordPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    
    ' Создаем объект диалогового окна выбора файла
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Выберите файл Word"
        .Filters.Clear
        ' Ограничиваем выбор только файлами Word
        .Filters.Add "Документы Word", "*.doc; *.docx"
        .AllowMultiSelect = False ' Только один файл
        
        ' Если пользователь выбрал файл и нажал OK
        If .Show = -1 Then
            wordPath = .SelectedItems(1)
            Set wdApp = CreateObject("Word.Application")
            wdApp.Visible = True ' Скрываем процесс Word
            Set wdDoc = wdApp.Documents.Open(wordPath)
            textDocWord = wdDoc.Content.text
            wdDoc.Close False
            wdApp.Quit
            Set wdDoc = Nothing
            Set wdApp = Nothing
        Else
            MsgBox "Выбор отменен пользователем.", vbExclamation
        End If
    End With
End Function
Function kvoVchod(TextDe As String, TextScho As String) As Double
    Dim i As Integer
    Dim r As Integer
    i = 0
    Do
        i = i + 1
        i = InStr(i, TextDe, TextScho, vbTextCompare)
        If i > 0 Then
           kvoVchod = kvoVchod + 1
        End If
    Loop While i > 0
End Function

Sub Ceels()

    On Error GoTo filesNotOpened
    Dim i As Integer
    Dim nomColDana As Integer
    Dim docWord As String
    Dim kvoE As Double
    Dim fso As Object
    Dim txtFile As Object
    Dim filePath As String
    Set wsTabel = Workbooks("табель рРЕБ.xlsm").Worksheets("Табель")
'    Set wsTabel = Workbooks("табель рРЕБ Серпень.xlsm").Worksheets("Табель")
    docWord = textDocWord()
    i = InStr(1, docWord, "Карта 1:25000, видання 2025 року", vbTextCompare)
    textData = Mid(docWord, i - 12, 10)
    For Each colD In wsTabel.Range("I7:AM7")
        If colD.Value = textData Then
            nomColDana = colD.Column
        End If
    Next
    On Error GoTo 0
    filePath = "C:\tmp\test.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(filePath, True)
    'filePath = " Прювіт, это первая строка!" & vbLf
    txtFile.WriteLine textData & " Прювіт, это первая строка!"
    For Each c1 In wsTabel.Range("H9:H63")
        'Debug.Print c1.Value
        If c1.Value <> "" Then
            c1.Select
            kvoE = kvoVchod(docWord, c1.Value)
            If kvoE = 0 And wsTabel.Cells(c1.Row, nomColDana).Value = 100 Then
                'filePath = filePath & "В підсумковом повинен бути: " & c1.Value & vbLf
                'txtFile.Write "В підсумковом немає: " & c1.Value
                txtFile.WriteLine "В підсумковом немає: " & c1.Value
            End If
            If kvoE = 1 And wsTabel.Cells(c1.Row, nomColDana).Value <> 100 Then
                'filePath = filePath & "В підсумковом не повинно бути: " & c1.Value & vbLf
                'txtFile.Write "В підсумковом не повинно бути: " & c1.Value
                txtFile.WriteLine "В підсумковом не повинно бути: " & c1.Value
            End If
            If kvoE > 1 Then
                'filePath = filePath & "В підсумковом дублюється: " & c1.Value & vbLf
                'txtFile.Write "В підсумковом дублюється: " & c1.Value
                txtFile.WriteLine "В підсумковом дублюється: " & c1.Value
            End If
            'Debug.Print kvoE
     '       If kvoE=0 and wsTabel.Cells(c1.Row, c1.Column).Value
        '        Set CelSt = wsStuff.Cells(5, c1.Column)
        '        If Left(Cel.Value, 1) <> "=" Then
        '            If ws526.Cells(c1.Row, c1.Column).Value = c1.Value Then
        '               Cel.Interior.Color = CelSt.Interior.Color 'шо було
        '            Else
        '               ws526.Cells(c1.Row, c1.Column).Value = c1.Value
        '               Cel.Interior.Color = 65535 'жовтий
        '            End If
        '        End If
        '        If Abs(Rng.Value) < 0.01 Then Rng.Value = 0
        End If
    Next
    'filePath = filePath & "Повний бардак"
    txtFile.Write textData & " Закінчили"
    Workbooks.Open Filename:=filePath
    Exit Sub
    
filesNotOpened:
    MsgBox "Файли: табель рРЕБ.xlsx та Підсумкове_.docx  повинні бути вікриті!"

End Sub

