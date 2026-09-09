Attribute VB_Name = "egos"
Public Const dateColShift = 3

Enum repType
    egos          ' поиск первой строки  с датой r1
    kotel         ' поиск последней строки с датой r2
    monthStaffList
End Enum

Sub egosCreate()
    repCreate repType.egos
End Sub

Sub kotelCreate()
    repCreate repType.kotel
End Sub

Sub monthStuffList()
    repCreate monthStaffList
End Sub

Sub repCreate(repType1 As repType)
If ActiveWorkbook.Name <> constStuffWbName Then ' штатка
    MsgBox constStuffWbName & " must be active!"
    Exit Sub
End If
Dim dateStaff As String
Dim monthStuff As String
wsName = ActiveSheet.Name
dateStaff = Left(wsName, 6)
If Not testDate(dateStaff) Then ' дата штатки
    MsgBox "Sheet name " & dateStaff & " is not recognized as date!"
    Exit Sub
End If

If Not Right(wsName, 2) = "_r" Then ' дата штатки
    MsgBox wsName & " - wrong worksheet suffix '_r' is required"
    Exit Sub
End If

monthStuff = Val(Mid(dateStaff, 3, 2))
yearStuff = Val(2000 + Mid(dateStaff, 1, 2))

Dim wbStuff As Object
Dim wsStuff As Object
Dim wbEgos As Object
Dim wsEgos As Object

Set wbStuff = int_getWbStaff()
Set wsStuff = int_getWsStaff(Left(dateStaff, 4) & "01")
Set wbEgos = Workbooks.Add()
Set wsEgos = wbEgos.ActiveSheet
'Set wsFio = wbEgos.Worksheets.Add()
wsEgos.Name = "egos"
'wsFio.Name = "FIOs"

Dim statusesRange As Range
Set statusesRange = wbStuff.Sheets("egos").Range("A2:C16")
Dim arStatuses As Variant
arStatuses = statusesRange.Value

Dim date1 As Date
Dim date2 As Date
Dim fio As String
Dim rowCur As Integer
Dim curHistSheetName As String

date1 = DateSerial(yearStuff, monthStuff, 1)
date2 = DateAdd("m", 1, date1) - 1

If repType1 <> monthStaffList Then
    If repType1 = repType.egos Then
        wsEgos.Cells(1, 1) = "≈∆ќ——"
    ElseIf repType1 = repType.kotel Then
        wsEgos.Cells(1, 1) = " ќ“≈Ћ"
    ElseIf repType1 = repType.monthStaffList Then
        wsEgos.Cells(1, 1) = "персонал за мес€ц"
    Else
        Debug.Assert False
    End If
    
    For i = 0 To (date2 - date1)  ' даты в шапке отчета
        wsEgos.Cells(1, i + dateColShift).Value = date1 + i
    Next i
Else
End If

' fios
rLast = 1
For i = 0 To (date2 - date1)
    curDate = date1 + i
    csName = Right(Year(curDate), 2) & Right((100 + Month(curDate)), 2) & Right(100 + Day(curDate), 2) & "_r"
    If csName <> "" Then
        On Error Resume Next
        Set wsHist = Nothing
        Set wsHist = wbStuff.Worksheets(csName)
        On Error GoTo 0
        If wsHist Is Nothing Then
            'Debug.Print csName & " is not found"
            MsgBox "Error" & csName & " is not found"

            Exit Sub
        End If
        
        For r = 2 To 81
            fio = wsHist.Cells(r, 19).Value
            inn = wsHist.Cells(r, 20).Value
            If fio <> "" Then
                milTitle = wsHist.Cells(r, 18)
                For j = 2 To rLast
                    If wsEgos.Cells(j, 2) = fio Then
                        Exit For
                    End If
                Next j
                If j > rLast Then
                    rLast = rLast + 1
                    wsEgos.Cells(rLast, 1) = milTitle
                    wsEgos.Cells(rLast, 2) = fio
                    If repType1 = monthStaffList Then
                        wsEgos.Cells(rLast, 3) = inn
                    End If
                End If
            End If
        Next r
    End If
Next i

If repType1 = monthStaffList Then
    Exit Sub
End If

rowCur = 1
For i = 2 To 100
    fio = wsEgos.Cells(i, 2).Value
    If fio <> "" Then
'        milTitle = wsStuff.Cells(i, 18).Value
        rowCur = rowCur + 1
'        wsEgos.Cells(rowCur, 1).Value = milTitle
'        wsEgos.Cells(rowCur, 2).Value = fio
        getFioEgos fio, date1, date2, wsEgos, wsStuff, wbStuff, arStatuses, rowCur, repType1
    End If
Next i

' на котле
rowCur = rowCur + 3

For i = 0 To (date2 - date1)
    curDate = date1 + i '- 1 ' сдвиг на 1 день назад потому что котел подаетс€ на один день вперед
    csName = Right(Year(curDate), 2) & Right((100 + Month(curDate)), 2) & Right(100 + Day(curDate), 2) & "_r"
    res1 = "No Data"
    If csName <> "" Then
        On Error Resume Next
        Set wsHist = Nothing
        Set wsHist = wbStuff.Worksheets(csName)
        On Error GoTo 0
        If wsHist Is Nothing Then
            Debug.Print csName & " is not found"
            Exit Sub
        End If
        kotel1 = "не найдено"
        For j = 110 To 150 ' поиск данных по котлу в штатке
            If wsHist.Cells(j, 19) = "котел" Then
                kotel1 = wsHist.Cells(j, 25)
                Exit For
            End If
        Next j
        

        wsEgos.Cells(rowCur, i + dateColShift).Value = kotel1
    End If
    'Debug.Print res1

Next i



End Sub

Sub getFioEgos(fio As String, date1 As Date, date2 As Date, wsEgos As Object, wsStuff As Object, wbStuff As Object, arStatuses As Variant, fioRow1 As Integer, repType1 As repType)
Dim curHistSheetName As String
Dim wsHist As Object
Dim location1 As String
Dim status1 As String
colLocation = 79  ' локаци€
colStatus = 80    ' статус

Dim dateTreshold
dateTreshold = DateValue("09.11.2025")

For i = 0 To (date2 - date1)
    curDate = date1 + i
    curHistSheetName = Right(Year(curDate), 2) & Right((100 + Month(curDate)), 2) & Right(100 + Day(curDate), 2)
    csName = curHistSheetName & "_r" 'int_getHistSheetName(curHistSheetName, "last")
    If curDate = dateTreshold Then
        colStatus = 80
    End If
    res1 = "No Data"
    If csName <> "" Then
        On Error Resume Next
        Set wsHist = Nothing
        Set wsHist = wbStuff.Worksheets(csName)
        On Error GoTo 0
        If wsHist Is Nothing Then
            Debug.Print csName & " is not found"
            Exit Sub
        End If
        
        fioRow = findFioRowStuff(wsHist, fio)
        If fioRow <> -1 Then
            location1 = wsHist.Cells(fioRow, colLocation).Value
            status1 = wsHist.Cells(fioRow, colStatus).Value
            res1 = ""
'            If status1 = "в на€вност≥" Then
'                res1 = location1
'            Else
'                res1 = status1
'            End If

            wsEgos.Cells(fioRow1, i + dateColShift).Value = convertStatus(status1, arStatuses, repType1)
        End If
    End If
    'Debug.Print res1
    
Next i
End Sub

Public Function findFioRowStuff(wsStuff As Object, fio As String) As Long
If Len(fio) <> 0 Then
    colFio = 19
    row1 = 2
    row2 = 81
    For findFioRowStuff = row1 To row2
        If wsStuff.Cells(findFioRowStuff, colFio).Value = fio Then
            Exit Function
        End If
    Next
End If
'MsgBox "Fio is not found (stuff): " & fio
findFioRowStuff = -1
'Debug.Assert False
End Function

Function convertStatus(status1 As String, arStatuses As Variant, repType1 As repType) As String
    convertStatus = ""
    If repType1 = egos Then
        col = 2
    Else
        col = 3
    End If
    For i = LBound(arStatuses) To UBound(arStatuses)
        If arStatuses(i, 1) = status1 Then
            convertStatus = arStatuses(i, col)
            Exit For
        End If
    Next i
End Function


