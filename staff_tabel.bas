Attribute VB_Name = "staff_tabel"
Enum statusFill
    seekingR1      ' поиск первой строки  с датой r1
    seekingR2      ' поиск последней строки  с датой r2
    Done
End Enum

Enum action
    Fill
    check
End Enum
'
'Sub qq()
'Debug.Print StringToMD5Hex("1")
''Debug.Print MD5Hex("1")
'End Sub



Function StringToMD5Hex(ByVal s As String) As String
Dim enc As Object
Dim bytes() As Byte
Dim bytes1() As Byte
Dim pos As Long
Dim outstr As String

'Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

bytes = StrConv(s, vbFromUnicode)
'bytes = enc.ComputeHash_2(bytes)
bytes1 = GetHashOfByteArray(bytes, CALG_MD5)


For pos = LBound(bytes1) To UBound(bytes1) Step 2
   outstr = outstr & Left(LCase(Chr(bytes1(pos))), 1)
Next pos

StringToMD5Hex = outstr
Set enc = Nothing
End Function

Function concat1(r As Range) As String
Dim i As Integer
i = 1
For Each c In r
    concat1 = concat1 & i & c.text
    i = i + 1
Next
End Function

Function RangeToMD5(r As Range) As String
RangeToMD5 = StringToMD5Hex(concat1(r))
End Function

Function RangeToMD5Short(r As Range) As String
RangeToMD5Short = RangeToMD5(r)
RangeToMD5Short = Right(RangeToMD5Short, 5)
End Function

Sub SelectionMd5Show()
res = RangeToMD5Short(Selection)
MsgBox res
End Sub

Function countif1(r As Range, txt As String) As Long
    Dim row1 As Range
    For Each row1 In r.Rows
        If Len(row1.Cells(1, 2).text) <> 0 And row1.Cells(1, 1).text = txt Then
            countif1 = countif1 + 1
        End If
    Next

End Function

Sub ActiveSheetCopyPasteValues()
    int_ActiveSheetCopyPasteValues (False)
End Sub

Sub int_ActiveSheetCopyPasteValues(showMsg As Boolean)
    If showMsg Then
        If MsgBox("Убрать формулы?", vbOKCancel, "Инструменты") = vbCancel Then Exit Sub
    End If
    clearFilter ActiveSheet
    ActiveSheet.Cells.Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    ActiveSheet.Cells(1, 1).Select
    Application.CutCopyMode = False
End Sub

Sub ActiveSheetClearCondFormats()
    ActiveSheet.Cells.FormatConditions.Delete
End Sub

Sub tabelFill()
    tabelFillSub action.Fill
End Sub

Sub tabelCheck()
    tabelFillSub action.check
End Sub


Sub tabelFillSub(action1 As action)

On Error GoTo filesNotOpened
Dim wsRpu As Object
Dim wsTabel As Object
Set wsRpu = Workbooks("1 РПУ.xlsx").Worksheets(1)
Set wsTabel = Workbooks(constTabelRebFileName).Worksheets(1)
On Error GoTo 0
Dim dateFill As String
wsRpu.Activate
dateFill = ActiveCell.Value
If Not testDate2(dateFill) Then
    MsgBox "Неверная дата " & dateFill
    Exit Sub
End If
If action1 = action.Fill Then
    If MsgBox("Перенести данные за " & dateFill & " из РПУ1 в Табель?", vbOKCancel, "Табель") = vbCancel Then
        Exit Sub
    End If
End If

Dim r2 As Long
Dim r1 As Long
r0 = ActiveCell.Row
c1 = ActiveCell.Column
rmax = 20
r1 = r0
While Cells(r1, c1).Value = dateFill And r0 - rmax < r1
    r1 = r1 - 1
Wend
r1 = r1 + 1
Debug.Assert r0 - rmax < r1

r2 = r0
While Cells(r2, c1).Value = dateFill And r0 + rmax > r2
    r2 = r2 + 1
Wend
r2 = r2 - 1
Debug.Assert r0 + rmax > r1


Debug.Print dateFill
'Debug.Print r1
'Debug.Print r2
Dim colDate As Long
colDate = findDateCol(wsTabel, dateFill)
If colDate = -1 Then
    Exit Sub
End If
Dim colIx As Long
count100 = 0

' columns B-L
For colIx = 2 To 14
    cnt = tableFill2(wsRpu, r1, r2, colIx, wsTabel, colDate, action1)
    count100 = count100 + cnt
Next
cntTabel = tableCalcDate100Count(dateFill)
If count100 <> cntTabel Then
    MsgBox "Different counts РПУ:" & count100 & " vs. Табель: " & cntTabel & " for the " & date1
End If
Exit Sub

filesNotOpened:
MsgBox "Files: 1 РПУ.xlsx and " & constTabelRebFileName & "  must be opened!"

End Sub


Function tableFill2(wsRpu As Object, r1 As Long, r2 As Long, colIx As Long, wsTabel As Object, colDate As Long, action1 As action) As Integer
    tableFill2 = 0
    Dim s As String
    Dim fio As String
    Dim errMsg As String
    For rowno = r1 To r2
        s = wsRpu.Cells(rowno, colIx).Value
        eolPos = InStr(s, vbLf)
        fio = Mid(s, eolPos + 1)
        If fio <> "" Then
            tableFill2 = tableFill2 + 1
            Debug.Print fio
            rowFio = findFioRow(wsTabel, fio)
            If rowFio = -1 Then
                tableFill2 = -1
                Exit Function
            End If
            If action1 = action.Fill Then
                wsTabel.Cells(rowFio, colDate).Value = 100
            Else
                If wsTabel.Cells(rowFio, colDate).Value <> 100 Then
                    errMsg = errMsg & fio & " -  error" & vbCr
                End If
            End If
        End If
    Next rowno
    If action1 = action.check And errMsg <> "" Then
        MsgBox errMsg
    End If
    
End Function


Public Function findFioRow(wsTabel As Object, fio As String, Optional hideMsg As Boolean) As Long
colFio = 8
row1 = 9
row2 = 95
For findFioRow = row1 To row2
    If wsTabel.Cells(findFioRow, colFio).Value = fio Then
        Exit Function
    End If
Next
If hideMsg = False Then
    MsgBox "Fio is not found: " & fio
End If
findFioRow = -1
End Function

Function findDateCol(wsTabel As Object, date1 As String) As Long
rowDate = 7
col1 = 9
col2 = 39
For findDateCol = col1 To col2
    If wsTabel.Cells(rowDate, findDateCol) = date1 Then
        Exit Function
    End If
Next
MsgBox "date is not found: " & date1
findDateCol = -1
End Function

Function tableCalcDate100Count(date1 As String)
    On Error GoTo filesNotOpened
    Dim wsTabel As Object
    Set wsTabel = Workbooks(constTabelRebFileName).Worksheets(1)
    On Error GoTo 0
    tableCalcDate100Count = 0
    colIx = findDateCol(wsTabel, date1)
    For r = 9 To 92
        If wsTabel.Cells(r, colIx).Value = 100 Then
            tableCalcDate100Count = tableCalcDate100Count + 1
        End If
    Next

Exit Function
filesNotOpened:
    MsgBox "Files: " & constTabelRebFileName & " must be opened!"
End Function

'выборка данных для сводной таблицы
Sub createPivotTableData()
    Set dataSheet1 = ActiveSheet
    Set dataSheet2 = Worksheets.Add
    Dim i As Integer
    i = 0
    For r = 9 To 100
        If dataSheet1.Cells(r, 1).Value <> Empty Then
            fio = dataSheet1.Cells(r, 8).text
            For c = 9 To 39
                date1 = dataSheet1.Cells(7, c).Value
                If dataSheet1.Cells(r, c).Value <> Empty Then
                    i = i + 1
                    dataSheet2.Cells(i, 1).Value = fio
                    dataSheet2.Cells(i, 2).Value = date1
                    dataSheet2.Cells(i, 3).Value = dataSheet1.Cells(r, c).Value
                End If
            Next
        End If
    Next
End Sub

Function getStaffDate(d1 As Date)
    year1 = Year(d1)
    month1 = Month(d1)
    day1 = Day(d1)
    
    day2 = Right("0" & day1, 2)
    month2 = Right("0" & month1, 2)
    year2 = Right(year1, 2)
    getStaffDate = year2 & month2 & day2
End Function

' not in usage
Sub zzz_checkRpu1VsTabel()

On Error GoTo filesNotOpened
Dim wsRpu As Object
Dim wsTabel As Object
Set wsRpu = Workbooks("1 РПУ.xlsx").Worksheets(1)
Set wsTabel = Workbooks(constTabelRebFileName).Worksheets(1)
On Error GoTo 0
Dim date1 As String
date1 = wsTabel.Cells(2, 7).Value
If Not testDate2(date1) Then
    MsgBox "Неверная дата " & dateFill
    Exit Sub
End If
year1 = Year(date1)
month1 = Month(date1)
day2 = Day(date1)
day1 = 1
Dim r0 As Long
Dim date2 As String
r0 = 1
For d = day1 To day2
    day2 = Right("0" & day1, 2)
    month2 = Right("0" & month1, 2)
    year2 = year1
    date2 = day2 & "." & month2 & "." & year2
    r1 = getDateRow(wsRpu, r0, date2)
    While wsRpu.Cells(r1, 1) = date2
        foreach c1 in (1,2,3)
    Wend
    
    
Next d


Exit Sub

filesNotOpened:
MsgBox "Files: 1 РПУ.xlsx and " & constTabelRebFileName & "  must be opened!"

End Sub
' not in usage
Function getDateRow(ws As Object, r0 As Long, date1 As String) As Long
    For r = r0 To 3000
        If ws.Cells(r, 1).Value = date1 Then
            getDateRow = r
            Exit Function
        End If
    Next r
End Function

Sub compare2Sheet()
Dim wsTabel As Object
Debug.Assert ActiveWindow.SelectedSheets.Count = 2
If ActiveWindow.SelectedSheets.Count <> 2 Then
    MsgBox "Exactly two sheets must be selected"
    Exit Sub
End If

intShift = 15
row1 = 2
Dim ws1 As Object
Dim ws2 As Object
Set ws1 = ActiveWindow.SelectedSheets(1)
Set ws2 = ActiveWindow.SelectedSheets(2)

Set nwb = Workbooks.Add
Set nws = nwb.ActiveSheet

ws1.Activate
ws1.Select
clearFilter ws1
ws1.Range("R1:S82").Copy

nws.Activate
nws.Cells(1, 1).Value = ws1.Name
nws.Cells(row1, 1).Select
Selection.PasteSpecial xlPasteValues

ws1.Activate
ws1.Range("AK1:AN82").Copy

nws.Activate
nws.Cells(row1, 3).Select
Selection.PasteSpecial xlPasteValues

ws1.Activate
ws1.Range("CA1:CD82").Copy

nws.Activate
nws.Cells(row1, 7).Select
Selection.PasteSpecial xlPasteValues



ws2.Activate
clearFilter ws2
ws2.Range("R1:S82").Copy
nws.Activate
nws.Cells(1, 1 + intShift).Value = ws2.Name
nws.Cells(row1, 1 + intShift).Select
Selection.PasteSpecial xlPasteValues
ws2.Activate
ws2.Range("AK1:AN82").Copy
nws.Activate
nws.Cells(row1, 3 + intShift).Select
Selection.PasteSpecial xlPasteValues

ws2.Activate
ws2.Range("CA1:CD82").Copy

nws.Activate
nws.Cells(row1, 7 + intShift).Select
Selection.PasteSpecial xlPasteValues

For r = 1 To 82
    For c = 1 To intShift
        If nws.Cells(r, c).Value <> nws.Cells(r, c + intShift).Value Then
            nws.Cells(r, c).Interior.Color = 65535
            nws.Cells(r, c + intShift).Interior.Color = 65535
        End If
    Next c
Next r

nws.Cells(1, 1).Select
End Sub
'' проверка 30К табель со штаткой
'Sub getDayMd5Short30()
'    res = int_check_30_100(30)
''    If res <> "" Then
''        MsgBox res
''    End If
'End Sub
'' проверка 100К табель со штаткой
'Sub getDayMd5Short100()
'    res = int_check_30_100(100)
''    If res <> "" Then
''        MsgBox res
''    End If
'End Sub
' проверка 30k 100k табель со штаткой



Sub check_30_100()
    res = int_check_30_100()
End Sub


Function int_check_30_100() As String
int_check_30_100 = ""
Dim date1 As String
Dim wbTabel As Object
Dim wsTabel As Object
Set wbTabel = ActiveWorkbook
If wbTabel.Name <> constTabelRebFileName Then
    MsgBox "Відкривати макрос можно тільки з : " & constTabelRebFileName
    Exit Function
End If
Set wsTabel = ActiveSheet
If ActiveSheet.Name <> "Табель" Then
    MsgBox "Аркуш повинен : " & ActiveSheet.Name
    Exit Function
End If
Dim wbStuff  As Object
Set wbStuff = Workbooks(constStuffWbName)
Debug.Assert Not IsNull(wbStuff)
Dim wsStuff  As Object

date1 = ActiveCell.text
If Not testDate2(date1) Then
    MsgBox "Неверная дата : " & date1
    Exit Function
End If
If date1 = Date Then
    wsName = "штатка"
Else
    Dim date2 As String
    date2 = Right(date1, 2) & Mid(date1, 4, 2) & Left(date1, 2)
    
    wsName = getLastHistSheetName(date2)
End If

Set wsStuff = wbStuff.Worksheets(wsName)
Debug.Assert Not IsNull(wsStuff)

colDate = findDateCol(wsTabel, date1)
colRepShift = 0
chkWsColRes = 0
Dim rngAddr As String
Dim chkWs As Object
Set chkWb = Workbooks.Add()
Set chkWs = ActiveSheet
For Each fee In Range("[PERSONAL.xlsb]Sheet1!A5:J5")
    If fee > 0 Then
        'Данные из табеля
        colFio = 8
        chkRow = 3
        chkWsColRes = 1 + colRepShift
        'wbTabel.Activate
        chkWs.Cells(1, chkWsColRes).Value = constTabelRebFileName & " " & date1
        For r = 9 To 92
            If wsTabel.Cells(r, colDate).Value = fee Then
        '        fios = fios & Cells(r, colFio).Value
        '        Debug.Print Cells(r, colFio).Value
                chkWs.Cells(chkRow, chkWsColRes).Value = wsTabel.Cells(r, colFio).Value
                chkRow = chkRow + 1
            End If
        Next r
        Dim dataRange As String
        Dim sortCol1 As String
        sortColChar = Chr(64 + chkWsColRes)
        sortCol1 = sortColChar & "3"
        dataRange = sortCol1 & ":" & sortColChar & "100"
    
        sortData chkWs, sortCol1, dataRange
        chkWs.Cells(2, chkWsColRes).Value = "=PERSONAL.xlsb!RangeToMD5Short(" & dataRange & ")"
        
        'данные из штатки
        chkWsColRes = 5 + colRepShift
        colFio = 19
        colVynagoroda = 86
        chkRow = 3
        chkWs.Cells(1, chkWsColRes).Value = constStuffWbName & "!" & wsStuff.Name
        For r = 2 To 81
            If wsStuff.Cells(r, colVynagoroda).Value = fee Then
                chkWs.Cells(chkRow, chkWsColRes).Value = wsStuff.Cells(r, colFio).Value
                chkRow = chkRow + 1
            End If
        Next r
        
        sortColChar = Chr(64 + chkWsColRes)
        sortCol1 = sortColChar & "3"
        dataRange = sortCol1 & ":" & sortColChar & "100"
        sortData chkWs, sortCol1, dataRange
        chkWs.Cells(2, chkWsColRes).Value = "=PERSONAL.xlsb!RangeToMD5Short(" & dataRange & ")"
        chkWs.Cells(1, 3 + colRepShift).Value = fee
        chkWs.Cells(2, 3 + colRepShift).Value = "=R2C" & 1 + colRepShift & "=R2C" & 5 + colRepShift
    
        chkWb.Activate
        rngAddr = Range(Cells(3, 1 + colRepShift), Cells(100, 5 + colRepShift)).Address
        setUniqueCondFmt rngAddr
        
        colRepShift = 7 + colRepShift
        
    End If
Next

chkWs.Cells(2, 3).Select

End Function

Sub setUniqueCondFmt(range1 As String)
    Range(range1).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlUnique
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    

    
End Sub

Sub sortData(ws As Object, keyCell As String, range1 As String)
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 key:=Range(keyCell) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range(range1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
       '.SortMethod = xlPin
        .Apply
    
    End With
    
End Sub





Sub test_getTabelPeriods()
Dim txt1 As String
t1 = "02.12.2025;з 14.12.2025 по 16.12.2025;з 18.12.2025 по 20.12.2025;з 22.12.2025 по 24.12.2025;з 26.12.2025 по 27.12.2025;29.12.2025"
't1 = "з 01.12.2025 по 31.12.2025"
txt1 = t1
res = getTabelPeriods(txt1, 1)
Debug.Print res
Debug.Print "--"
txt1 = t1
res = getTabelPeriods(txt1, 2)
Debug.Print res
End Sub


Sub testSort()
With ActiveSheet.Sort
    .SetRange Range("A3:A100")
    .Header = xlNo
    .MatchCase = False
    '.Orientation = xlTopToBottom
   '.SortMethod = xlPin
    .Apply
End With
End Sub

Function getTabelPeriods(txt1 As String, par As Integer) As String
getTabelPeriods = ""
Debug.Assert Right(txt1, 1) <> ";"
txt1 = Replace(txt1, vbLf, "")
txt1 = txt1 & ";"
Dim txt2 As String
Do
    pos1 = InStr(txt1, ";")
    If pos1 > 0 Then
        txt2 = Left(txt1, pos1 - 1)
        r1 = getPeriodsParse1(txt2, par)
    Else
        Exit Do
    End If
    txt1 = Mid(txt1, pos1 + 1)
    getTabelPeriods = getTabelPeriods & r1 & vbCr & vbLf
Loop
getTabelPeriods = Mid(getTabelPeriods, 1, Len(getTabelPeriods) - 2)
End Function

Function getPeriodsParse1(txt2 As String, par1 As Integer) As String
getPeriodsParse1 = ""
Debug.Assert testDate2(txt2) Or testDatePeriod(txt2) Or txt2 = ""
Debug.Assert par1 = 1 Or par1 = 2
If txt2 = "" Then
    getPeriodsParse1 = ""
Else
    If testDate2(txt2) Then
        getPeriodsParse1 = txt2
    Else
        If par1 = 1 Then
            getPeriodsParse1 = Mid(txt2, 3, 10)
        Else
            getPeriodsParse1 = Mid(txt2, 17, 10)
        End If
    End If
End If
End Function

Sub compareBrData()

If Not UCase(Left(ActiveWorkbook.Name, 2)) = "БР" Then
    MsgBox "Active workbook must be БР"
    Exit Sub
End If

If Not testDate2(Selection.Cells(1).text) Then
    MsgBox "wrong  data value " & Selection.Cells(1).text
    Exit Sub
End If
'Debug.Assert testDate2(Selection.Cells(7).text)
date1 = Selection.Cells(-5).Value
date2 = Selection.Cells(1).Value
Debug.Assert date2 - date1 = 1
row1 = Selection.Cells(-5).Row
row2 = Selection.Cells(1).Row
For c = 2 To 21
    For r = 0 To 5
        Debug.Print row1 + r & " x " & c & " ->> " & row2 + r & " x " & c
        Debug.Print Cells(row1 + r, c).Value
        Debug.Print Cells(row2 + r, c).Value
        
        If Cells(row1 + r, c).Value <> Cells(row2 + r, c).Value Then
            Cells(row1 + r, c).Interior.Color = vbGreen
            Cells(row2 + r, c).Interior.Color = vbYellow
        End If

    Next r
Next c
End Sub

Sub copyToClipBoard_intoline()
    res = int_getFioDataInline(Selection)
    CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "text", res
End Sub

Function int_getFioDataInline(rng As Range) As String
For Each c1 In rng
    s1 = Trim(c1.text)
    If s1 <> "" Then
        s1 = Replace(c1.text, vbLf, " ")
        sss = sss & s1 & ", "
    End If
Next
int_getFioDataInline = sss
End Function
