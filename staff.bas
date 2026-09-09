Attribute VB_Name = "staff"
' [$-uk-UA-x-genlower]Д ММММ ГГГГ "г."


Public Const staffDataRange As String = "A1:BZ73"
Public Const staffStatusDataRange As String = "AK2:AN73"
Public Const hashDataCell As String = "B83"
Public Const staffDateTimeCell As String = "B88" ' date time of
Public Const staffDateCell As String = "B84"  'псевдодата данных на листе штатка
Public Const staffDateCell2 As String = "B87"  'дата данных на листе штатка в формате даты
'Public Const servantHashCell As String = "B78"
Public Const headersRange As String = "A1:BZ1"
Public Const staffFioRange As String = "S2:S81"
Public Const headersHashCell As String = "B8"   ' хеш заголовков на листе спр1
Public Const constStuffWbName As String = "штатка.xlsx"
Public Const constStuffWsName As String = "штатка"
Public Const constPidsWbName As String = "пидс.xlsx"
Public Const constTabelRebFileName As String = "табель рРЕБ.xlsm"

Public Const const526os As String = "CR5"
Public Const constWsStaffNayav As String = "X118"
Public Const constwb526path As String = "C:\RR\БЧС\5.26 БЧС\"
Public Const constwb526fileName As String = "рРЕБ БЧС Форма 5.26 - .xlsx"
Public Const constWsName526 = "БЧС"
'Public Const constStaffPath As String = "R:\ШТАТКА"
 
Public Const constWsName518 = "xx омбр"
Public Const constwb518fileName As String = "рРЕБ Ф5.18.xlsx"

Dim chkRegEx As New RegExp

Function checkHeaders() As Boolean
Set wsPersonalStaff = int_getWsPersonalStaff()
Set wsDict = int_getWsDict1()

'headersHash = wsDict.Range(headersHashCell).Value
'hash1 = RangeToMD5Short(wsPersonalStaff.Range(headersRange))
'checkHeaders = headersHash = hash1

checkHeaders = True
chkRegEx.Pattern = "\d\d.\d\d.202\d"
For i = 1 To 78 ' BZ
    txtDict = wsDict.Cells(9, i).Value
    txtStuff = wsPersonalStaff.Cells(1, i).Value
    
    txtDict1 = chkRegEx.Replace(txtDict, "dd.mm.yyyy")
    txtStuff1 = chkRegEx.Replace(txtStuff, "dd.mm.yyyy")

    If txtDict1 <> txtStuff1 Then
        checkHeaders = False
        MsgBox "Header difference: " & vbCr & txtDict & vbCr & txtStuff1
        Exit Function
    End If
Next i

End Function

Function int_checkServants() As Boolean
Set wsStaff = int_getWsStaff()
'wsStaff.Range()
histSheetName = getLastHistSheetName()

End Function

Public Function testDate(date1 As String) As Boolean
chkRegEx.Pattern = "^2\d{5}$" ' 260101
testDate = chkRegEx.test(date1)
End Function

Public Function testDate2(date1 As String) As Boolean
chkRegEx.Pattern = "^([0-9]{2}).([0-9]{2}).([0-9]{4})$"
testDate2 = chkRegEx.test(date1)  ' 26.10.2026
End Function

Public Function testDatePeriod(date1 As String) As Boolean
chkRegEx.Pattern = "^з ([0-9]{2}).([0-9]{2}).([0-9]{4}) по ([0-9]{2}).([0-9]{2}).([0-9]{4})$"
testDatePeriod = chkRegEx.test(date1) ' з 26.10.2026 по 26.10.2026
End Function

Function testPersonalFileName(date1 As String) As Boolean
chkRegEx.Pattern = "^2\d{5}_ррб.xlsx$" ' 261012_ррб.xlsx
testPersonalFileName = chkRegEx.test(date1)
End Function

Sub copyNewData()
Dim wsPersonalStaff As Object
Dim wsTabel As Object

Set wsPersonalStaff = int_getWsPersonalStaff()

If MsgBox("Скопировать данные из персонала? " & wsPersonalStaff.Name, vbOKCancel, constStuffWbName) = vbCancel Then Exit Sub
If Not checkHeaders() Then
    MsgBox "Personal data headers are changed"
    Exit Sub
End If



If IsNull(wsPersonalStaff) Or wsPersonalStaff Is Nothing Then GoTo NoPersonalFileOpen

Set wsStaff = int_getWsStaff()
If wsStaff Is Nothing Then Exit Sub

If wsStaff.Cells(1, 1).Value = "" Then
    wsPersonalStaff.Activate
    Range(staffDataRange).Select
    Selection.copy
    wsStaff.Activate
    Range("A1").Select
    wsStaff.Range("A1").PasteSpecial xlPasteValues
    dateStuff = Left(wsPersonalStaff.Name, 6)
    wsStaff.Range(staffDateCell).Value = dateStuff
Else
    MsgBox "The sheet " & constStuffWsName & " is not empty!"
End If

Exit Sub

NoPersonalFileOpen:
MsgBox "Personal file is not opened or is not editable!"
Exit Sub

ManyPersonalFilesOpen:
MsgBox "Too many personal files are opened!"
Exit Sub

FileStaffNotOpened:
MsgBox constStuffWbName & " is not opened!"
Exit Sub


filesNotOpened:
MsgBox "Files: 1 РПУ.xlsx and табель рРЕБ.xlsm must be opened!"
End Sub

Function int_getWb(wbName As String) As Object
On Error GoTo Error
Set int_getWb = Workbooks(wbName)
Exit Function

Error:
Set int_getWb = Nothing
MsgBox wbName & " is not opened or editable!"
End
End Function

Public Function int_getWbStaff() As Object
Set int_getWbStaff = int_getWb(constStuffWbName)
End Function


Public Function int_getStaffDate() As String ' 260215
Set wsStaff = int_getWsStaff()
Set a1 = wsStaff.Range(staffDateCell)
int_getStaffDate = a1.Value
End Function


Public Function int_getStaffDate2() As Date ' 15.02.2026
Set wsStaff = int_getWsStaff()
Set a1 = wsStaff.Range(staffDateCell2)
int_getStaffDate2 = a1.Value
End Function

Public Function int_getWsStaff(Optional wsName As String = constStuffWsName) As Object
Set wbStaff = int_getWbStaff()
On Error GoTo Error
Set int_getWsStaff = wbStaff.Worksheets(wsName)
Exit Function

Error:
MsgBox "Error of reading of the sheet " & wsName & "!"
End
End Function

Function int_getWsDict1() As Object
Set wbStaff = int_getWbStaff()
On Error GoTo Error
wsName = "спр1"
Set int_getWsDict1 = wbStaff.Worksheets(wsName)
Exit Function

Error:
MsgBox "Error of reading of the sheet " & wsName & "!"
Debug.Assert False
End
End Function

Function int_getWsPersonalStaff() As Object
personalWbName = ""
wsName = ""
Set int_getWsPersonalStaff = Nothing
For Each wb In Workbooks
    'Debug.Print wb.Name
    If testPersonalFileName(wb.Name) Then
        If Not personalWbNameIsFound Then
            personalWbNameIsFound = True
            wsName = Left(wb.Name, 10)
            Set int_getWsPersonalStaff = wb.Worksheets(wsName)
        Else
            Set int_getWsPersonalStaff = Nothing
            GoTo Error
        End If
    End If
Next wb

If int_getWsPersonalStaff Is Nothing Then GoTo Error

Exit Function

Error:
MsgBox "A single personal file must be opened!"
End
End Function


Function int_clearStaffData() As Integer
int_clearStaffData = -1
Dim wsStaff As Object
Set wsStaff = int_getWsStaff()
clearFilter wsStaff
wsStaff.Activate
wsStaff.Range(staffDataRange).ClearContents
wsStaff.Range(staffDateCell).ClearContents
int_clearStaffData = 0
End Function

Sub clearStaffData()
If MsgBox("Удалить текущие данные?", vbOKCancel, constStuffWbName) = vbCancel Then Exit Sub
Dim res As Integer
res = int_clearStaffData()
If res = 0 Then
    'msg = "Данные удалены"
Else
    MsgBox "Что-то пошло не так"
End If

End Sub

Function getStateFioRow(fio As String) As Long
    Set wsState = int_getWsStaff()

    For Each c In wsState.Range(staffFioRange).Cells
        'Debug.Print c.Value
        If fio = c.Value Then
            getStateFioRow = c.row
            Exit For
        End If
    Next
    
End Function

Function int_getNewHistSheetName(histDate As String) As String
    int_getNewHistSheetName = int_getHistSheetName(histDate, "next")
End Function

Public Function getLastHistSheetName(histDate As String) As String
    getLastHistSheetName = int_getHistSheetName(histDate, "last")
End Function


Public Function int_getHistSheetName(histDate As String, nextLast As String) As String
int_getHistSheetName = ""
Set wbStaff = int_getWbStaff()
Dim wsHist As Object
On Error Resume Next
Set wsHist = wbStaff.Worksheets(histDate)
On Error GoTo 0
int_getHistSheetName = ""
wsNameLast = ""
If wsHist Is Nothing Then
    If nextLast = "next" Then
        int_getHistSheetName = histDate
    End If
Else
    Set wsHist = Nothing
    wsNameTest = histDate
    For i = 1 To 10
        wsNameLast = wsNameTest
        wsNameTest = histDate & "_" & i
        On Error Resume Next
        Set wsHist = Nothing
        Set wsHist = wbStaff.Worksheets(wsNameTest)
        On Error GoTo 0
        If wsHist Is Nothing Then
            If nextLast = "next" Then
                int_getHistSheetName = wsNameTest
            Else
                int_getHistSheetName = wsNameLast
            End If
            Exit For
        End If

    Next i
End If

'If int_getHistSheetName = "" Then
'    MsgBox "закомментировать? Error of creating of the next history sheet name!"
'    'End
'End If

End Function

Sub createHistCopy()
    int_createHistCopy ""
End Sub

Sub test1()
    int_createHistCopy ("r")
End Sub

Function int_createHistCopy(marker As String) As Worksheet
Set wbStaff = int_getWbStaff()
Set wsStaff = int_getWsStaff()

Set a1 = wsStaff.Range(staffDateCell)
Dim staffDate As String
staffDate = a1.Value
If staffDate = "" Then
    MsgBox "дата данных не найдена " & staffDateCell
    Exit Function
End If
If Not testDate(staffDate) Then
    MsgBox "Error date!"
    Exit Function
End If

If marker = "" Then
    If MsgBox("Создать архивную копию (" & staffDate & ")?", vbOKCancel, constStuffWbName) = vbCancel Then
        Exit Function
    End If
End If



    
If marker = "" Then
    lastHistPageName = getLastHistSheetName(staffDate)
    If lastHistPageName <> "" Then
        hashDataLast = wbStaff.Worksheets(lastHistPageName).Range(hashDataCell).Value
        hashData = wsStaff.Range(hashDataCell).Value
        chkRegEx.Pattern = "^[a-fA-F0-9]{5}$"
        If chkRegEx.test(hashDataLast) And chkRegEx.test(hashData) Then
            If hashDataLast = hashData Then
                If MsgBox("Текущие данные идентичны последней архивной копии (" & lastHistPageName & "), продолжить?", vbOKCancel, constStuffWbName) = vbCancel Then
                    Exit Function
                End If
            End If
        Else
            MsgBox "hash codes are incorrect: " & hashDataLast & " vs " & hashData
        End If
    End If
    nextHistPageName = int_getNewHistSheetName(staffDate)
Else
    nextHistPageName = staffDate & "_" & marker
    On Error Resume Next

    Set testWs = wbStaff.Worksheets(nextHistPageName)
    On Error GoTo 0
    If IsObject(testWs) Then
        If MsgBox("Worksheet " & nextHistPageName & " already exists! Recreate it?", vbOKCancel) = vbOK Then
            wbStaff.Worksheets(nextHistPageName).Delete
        Else
            Set int_createHistCopy = testWs
            Exit Function
        End If
    End If
End If
wsStaff.copy After:=wsStaff
Set newWS = wbStaff.Worksheets(wsStaff.Index + 1)
newWS.Name = nextHistPageName
ActiveSheet.Cells(1, 19).FormatConditions.Delete ' удаление условного форматирования с заголовка ФИО
ActiveSheet.Cells(1, 18).FormatConditions.Delete ' удаление условного форматирования с заголовка званием
ActiveSheetCopyPasteValues ' замена формул значениями
newWS.Range(staffDateTimeCell).Value = Now
    
wsStaff.Activate
Range("A1").Select
Set int_createHistCopy = newWS

End Function

Sub copyNewDataToPersonal()

staffDate = int_getStaffDate()
If MsgBox("Скопировать данные (" & staffDate & ") в персонал?", vbOKCancel, constStuffWbName) = vbCancel Then Exit Sub
If int_checkOtpuskCurrent(0) <> 0 Then
    If MsgBox("Продолжить?", vbYesNo) = vbNo Then Exit Sub
End If

Dim wsPersonalStaff As Object
Dim wsTabel As Object
Dim wsStaff As Object

'On Error GoTo FileStaffNotOpened
''Set wsStaff = int_getWsStaff()
'On Error GoTo 0

Set wsStaff = int_createHistCopy("r")
wsName = staffDate & "_ррб"
personalWbName = wsName & ".xlsx"
On Error GoTo NoPersonalFileOpen
Set wb = Workbooks(personalWbName)
On Error GoTo 0
Set wsPersonalStaff = wb.Worksheets(wsName)

clearFilter wsStaff

wsStaff.Range(staffStatusDataRange).copy
wsPersonalStaff.Range(staffStatusDataRange).PasteSpecial xlPasteValues

'int_createHistCopy ("r")

MsgBox "Готово"

Exit Sub

NoPersonalFileOpen:
MsgBox "Personal file " & personalWbName & " is not opened or is not editable!"
Exit Sub

ManyPersonalFilesOpen:
MsgBox "Too many personal files are opened!"
Exit Sub

FileStaffNotOpened:
MsgBox constStuffWbName & " is not opened!"
Exit Sub


filesNotOpened:
MsgBox "Files: 1 РПУ.xlsx and табель рРЕБ.xlsm must be opened!"
Exit Sub

NoComment:
MsgBox "Error on reading of A1 comment"
End Sub

Function clearFilter(ws As Object) As Boolean
If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
    ws.ShowAllData
    clearFilter = True
End If
End Function

Sub bchs526()
Set wsStaff = int_getWsStaff()

End Sub



Function int_getWs526() As Object
Set int_getWs526 = Nothing
'Workbooks.Open constwb526path & constwb526fileName
If Not ActiveWorkbook.Name = constwb526fileName Then
    Exit Function
End If


Set wb526 = int_getWb(constwb526fileName)
Set int_getWs526 = wb526.Worksheets(constWsName526)
On Error GoTo 0
Exit Function
End Function


Sub upd526()
Set ws526 = int_getWs526()
Set wb526 = int_getWb(constwb526fileName)
Set wsStuff = int_getWsStaff()
On Error GoTo err1
ok = False
For i = 1 To 1000
    If ws526.Cells(2, i).Value = "Особовий склад" Then
        date1 = Left(Now(), 10)
        ws526.Cells(5, i).Value = wsStuff.Range(constWsStaffNayav).Value
        title1 = Left(ws526.Cells(1, 5).Value, 46)
        Debug.Assert title1 = "БОЙОВИЙ ТА ЧИСЕЛЬНИЙ СКЛАД" & vbLf & "роти РЕБ станом на "
        ws526.Cells(1, 5).Value = title1 & date1
        
        wb526.Save
        
        On Error GoTo err2
        newFileName = Left(wb526.Name, 19) & " " & date1 & ".xlsx"
        newFileName = constwb526path & newFileName
        wb526.SaveAs newFileName
        'wb526.Close
        MsgBox newFileName & " is created."
        ok = True
        Exit For
    End If
Next i

If ok Then
    Exit Sub
End If

err1:
MsgBox "Error on updating the " & constwb526fileName
End

err2:
MsgBox "Error on updating the " & newFileName
End
' C:\RR\5.26 БЧС\рРЕБ БЧС Форма 5.26 23.02.2026.xlsx
' C:\RR\БЧС\5.26 БЧС
End Sub

Function int_getWs518() As Object
int_getWs518 = Nothing
If Not ActiveWorkbook.Name = constwb526fileName Then
    Exit Function
End If

Set wb518 = int_getWb(constwb518fileName)
Set int_getWs518 = wb518.Worksheets(constWsName518)
On Error GoTo 0
Exit Function
End Function

Sub upd518()
Set ws518 = int_getWs518()
Set wb518 = int_getWb(constwb518fileName)
Set wsStuff = int_getWsStaff()
On Error GoTo err1
ok = False

Debug.Assert ws518.Cells(36, 5).Value = "РПУ ЖУРАВЛИНЕ"
'ws518.Cells(36, 13).Value = wsStuff

For i = 1 To 1000
    If ws526.Cells(2, i).Value = "Особовий склад" Then
        date1 = Left(Now(), 10)
        ws526.Cells(5, i).Value = wsStuff.Range(constWsStaffNayav).Value
        title1 = Left(ws526.Cells(1, 5).Value, 46)
        Debug.Assert title1 = "БОЙОВИЙ ТА ЧИСЕЛЬНИЙ СКЛАД" & vbLf & "роти РЕБ станом на "
        ws526.Cells(1, 5).Value = title1 & date1
        
        wb526.Save
        
        On Error GoTo err2
        newFileName = Left(wb526.Name, 19) & " " & date1 & ".xlsx"
        newFileName = constwb526path & newFileName
        wb526.SaveAs newFileName
        'wb526.Close
        MsgBox newFileName & " is created."
        ok = True
        Exit For
    End If
Next i

If ok Then
    Exit Sub
End If

err1:
MsgBox "Error on updating the " & constwb526fileName
End

err2:
MsgBox "Error on updating the " & newFileName
End

End Sub

Sub staffTodayOpen()
    fn = getStaffDate(Now())
    'ffn = constStaffPath & fn & "_ррб.xlsx"
    Set wsPersonal = Workbooks("PERSONAL.xlsb").Worksheets(1)
    path1 = wsPersonal.Cells(1, 1)
    Debug.Assert Len(path1) > 0
    ffn = wsPersonal.Cells(1, 1) & fn & "_ррб.xlsx"
    On Error Resume Next
    Set wb = Workbooks.Open(fileName:=ffn, Password:=fn)
    If IsEmpty(wb) Then
        MsgBox "что-то пошло не так" & vbCrLf & ffn
    End If
    On Error GoTo 0
End Sub

Public Sub SortArray(ByRef arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            ' Change '>' to '<' for descending order
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub


Public Sub Sort2DArray(ByRef arr() As String, ByVal SortColumn As Long, Optional ByVal Ascending As Boolean = True)
    Dim i As Long, j As Long, k As Long
    Dim lowRow As Long, upRow As Long
    Dim lowCol As Long, upCol As Long
    Dim temp As String
    Dim shouldSwap As Boolean
    
    ' Get array boundaries
    lowRow = LBound(arr, 1)
    upRow = UBound(arr, 1)
    lowCol = LBound(arr, 2)
    upCol = UBound(arr, 2)
    
    ' Bubble Sort routine
    For i = lowRow To upRow - 1
        For j = i + 1 To upRow
            
            ' StrComp returns 1 if arr(i) > arr(j), and -1 if arr(i) < arr(j)
            ' vbTextCompare makes it case-insensitive (use vbBinaryCompare for case-sensitive)
            If Ascending Then
                shouldSwap = (StrComp(arr(i, SortColumn), arr(j, SortColumn), vbTextCompare) = 1)
            Else
                shouldSwap = (StrComp(arr(i, SortColumn), arr(j, SortColumn), vbTextCompare) = -1)
            End If
            
            ' Swap entire rows if criteria met
            If shouldSwap Then
                For k = lowCol To upCol
                    temp = arr(i, k)
                    arr(i, k) = arr(j, k)
                    arr(j, k) = temp
                    Next k
            End If
            
        Next j
    Next i
End Sub


Public Function chekIfWbIsOpened(wbName As String, Optional isActive As Boolean = False)
    On Error GoTo filenotopened
    Set dummy = Workbooks(wbName)
    chekIfWbIsOpened = True
    If isActive And ActiveWorkbook.Name <> wbName Then
         GoTo fileIsNotActive
    End If
        

    GoTo exitFunction
filenotopened:
    MsgBox ("The file " & wbName & " is not opened!")
    chekIfWbIsOpened = False
    
fileIsNotActive:

exitFunction:
    On Error GoTo 0
End Function


Public Function removeSubstring(s1 As String, s2 As String) As String
    res = s1
    pos = InStr(1, s1, s2, vbTextCompare)
    If pos > 0 Then
        s1 = Left(s1, pos - 1) & Mid(s1, pos + Len(s2))
    End If
    removeSubstring = s1
End Function


Sub exportMacro()
repoDir = "C:\Users\SecundUser\OneDrive\Documents\vsu\persxls\"
For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
    m.Export (repoDir & m.Name & ".bas")
Next
'Set wbp = Workbooks("PERSONAL.xlsb").VBProject.VBComponents(1).Export()

End Sub
