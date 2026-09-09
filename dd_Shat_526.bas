Attribute VB_Name = "dd_Shat_526"
Dim wbStuff As Object
Dim wsStuff As Object
Dim wb526 As Object
Dim ws526 As Object

Sub Perenos(Str1 As Integer, Kol1 As Integer, Str2 As Integer, Kol2 As Integer)
    Cel = ws526.Cells(Str1, Kol1)
    If ws526.Cells(Str1, Kol1).Value = wsStuff.Cells(Str2, Kol2).Value Then
        Cel.Pattern = xlNone 'без залівки
    Else
        ws526.Cells(Str1, Kol1).Value = wsStuff.Cells(Str2, Kol2).Value
       Cel.Color = 65535 'жовтий
    End If
End Sub

Sub Ceels()
On Error GoTo filesNotOpened
Dim wsRpu As Object
Dim wsTabel As Object
Set wsStuff = Workbooks("штатка.xlsx").Worksheets("526-518")
Set ws526 = Workbooks("рРЕБ Ф5.26 - .xlsx").Worksheets("БЧС")
On Error GoTo 0

'    Filename1 = "C:\Users\Admin\OneDrive\Документи\штатка_левая.xlsx"
'    Set wbStuff = Workbooks.Open(Filename1)
'    Set wsStuff = wbStuff.Sheets("526-518")
'    Filename2 = "C:\Users\Admin\OneDrive\Документи\рРЕБ Ф5.26 - Бяка.xlsx"
'    Set wb526 = Workbooks.Open(Filename2)
'    Set ws526 = wb526.Sheets("БЧС")
'        For j = 3 To 150 ' поиск данных по котлу в штатке
'            If wsHist.Cells(j, 19) = "котел" Then
'                kotel1 = wsHist.Cells(j, 25)
'                Exit For
'            End If
'        Next j
    For Each c1 In wsStuff.Range("C6:AE13")
        c1.Select
        Set Cel = ws526.Cells(c1.Row, c1.Column)
        Set CelSt = wsStuff.Cells(5, c1.Column)
        If Left(Cel.Value, 1) <> "=" Then
            If ws526.Cells(c1.Row, c1.Column).Value = c1.Value Then
               Cel.Interior.Color = CelSt.Interior.Color 'шо було
            Else
               ws526.Cells(c1.Row, c1.Column).Value = c1.Value
               Cel.Interior.Color = 65535 'жовтий
            End If
        End If
'        If Abs(Rng.Value) < 0.01 Then Rng.Value = 0
    Next
    
Exit Sub

filesNotOpened:
    MsgBox "Файли: рРЕБ Ф5.26 - .xlsx та штатка.xlsx  повинні бути вікриті!"

End Sub
''''''
''''''Sub tabelFillSub(action1 As action)
''''''
''''''On Error GoTo filesNotOpened
''''''Dim wsRpu As Object
''''''Dim wsTabel As Object
''''''Set wsRpu = Workbooks("1 РПУ.xlsx").Worksheets(1)
''''''Set wsTabel = Workbooks(constTabelRebFileName).Worksheets(1)
''''''On Error GoTo 0
''''''Dim dateFill As String
''''''wsRpu.Activate
''''''dateFill = ActiveCell.Value
''''''If Not testDate2(dateFill) Then
''''''    MsgBox "Неверная дата " & dateFill
''''''    Exit Sub
''''''End If
''''''If action1 = action.Fill Then
''''''    If MsgBox("Перенести данные за " & dateFill & " из РПУ1 в Табель?", vbOKCancel, "Табель") = vbCancel Then
''''''        Exit Sub
''''''    End If
''''''End If
''''''
''''''Dim r2 As Long
''''''Dim r1 As Long
''''''r0 = ActiveCell.Row
''''''c1 = ActiveCell.Column
''''''rmax = 20
''''''r1 = r0
''''''While Cells(r1, c1).Value = dateFill And r0 - rmax < r1
''''''    r1 = r1 - 1
''''''Wend
''''''r1 = r1 + 1
''''''Debug.Assert r0 - rmax < r1
''''''
''''''r2 = r0
''''''While Cells(r2, c1).Value = dateFill And r0 + rmax > r2
''''''    r2 = r2 + 1
''''''Wend
''''''r2 = r2 - 1
''''''Debug.Assert r0 + rmax > r1
''''''
''''''
''''''Debug.Print dateFill
'''''''Debug.Print r1
'''''''Debug.Print r2
''''''Dim colDate As Long
''''''colDate = findDateCol(wsTabel, dateFill)
''''''If colDate = -1 Then
''''''    Exit Sub
''''''End If
''''''Dim colIx As Long
''''''count100 = 0
''''''
''''''' columns B-L
''''''For colIx = 2 To 12
''''''    cnt = tableFill2(wsRpu, r1, r2, colIx, wsTabel, colDate, action1)
''''''    count100 = count100 + cnt
''''''Next
''''''cntTabel = tableCalcDate100Count(dateFill)
''''''If count100 <> cntTabel Then
''''''    MsgBox "Different counts РПУ:" & count100 & " vs. Табель: " & cntTabel & " for the " & date1
''''''End If
''''''Exit Sub
''''''
''''''filesNotOpened:
''''''MsgBox "Files: 1 РПУ.xlsx and " & constTabelRebFileName & "  must be opened!"
''''''
''''''End Sub
