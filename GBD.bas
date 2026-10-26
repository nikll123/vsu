Attribute VB_Name = "GBD"

'Dim chkRegEx As New RegExp
Function checkEnv() As Boolean
    wbName = "1 РПУ.xlsx"
    checkEnv = True
    If ActiveWorkbook.Name <> wbName Then
        MsgBox "The activebook must be " & wbName
        checkEnv = False
    End If
    wsName = "Лист2"
    If ActiveSheet.Name <> wsName Then
        MsgBox "The activesheet must be " & wsName
        checkEnv = False
    End If
End Function


Sub gbdGenerate()
    If Not checkEnv() Then
        Exit Sub
    End If
    Dim txt1 As String
    date1 = Cells(29, 1)
    txt1 = date1 & vbCrLf & vbCrLf
    
    txt1 = txt1 & "Карта 1:25000, видання 2025 року" & vbCrLf & vbCrLf
    
    txt1 = txt1 & "1. Без змін" & vbCrLf
    txt1 = txt1 & bchs()
    txt1 = txt1 & "3. Без змін" & vbCrLf  ' 3. СПІВВІДНОШЕНЯ СИЛ ТА ЗАСОБІВ
    txt1 = txt1 & br4()                   ' 4. КОРОТКЕ ВИКЛАДЕННЯ ЗАВДАННЯ ПІДПОРЯДКОВАНИМ ПІДРОЗДІЛАМ
    txt1 = txt1 & part5()                 ' 5. КОРОТКЕ РІШЕННЯ КОМАНДИРА І ПОСТАВЛЕНІ  ЗАВДАННЯ ПІДПОРЯДКОВАНИМ ПІДРОЗДІЛАМ
    txt1 = txt1 & part6()                 ' 6. ХІД БОЙОВИХ ДІЙ
    brs = partBRs()                       ' 7. УТОЧНЕНІ РІШЕННЯ ТА РОЗПОРЯДЖЕННЯ  ЯКІ ВІДДАВАЛИСЯ
    lastBrHashRow = 85
    lastBrHash = Cells(lastBrHashRow, 5)
    newBrHash = StringToMD5Hex(brs)
    If lastBrHash = newBrHash Then
        brs = "7. Без змін" & vbCrLf
    Else
        Cells(lastBrHashRow, 5) = newBrHash
    End If
    txt1 = txt1 & brs
    txt1 = txt1 & part8()                 ' 8. ВТРАТИ СВОЇХ ВІЙСЬК ТА ПРОТИВНИКА
    txt1 = txt1 & part9()                 ' 9. ТРОФЕЇ ТА ЗАХОПЛЕНІ ПОЛОНЕНІ
    txt1 = txt1 & part10()                ' 10. ВИСНОВОК ПРО НОВІ СПОСОБИ БОЙОВИХ ДІЙ
    txt1 = txt1 & sign()

    Dim repDate As Date
    repDate = Cells(29, 1)
    writeTxtPids txt1, repDate
    
End Sub

Function bchs() As String
    chgBchs = Cells(1, 5)
    bchs = "2. "
    If Cells(1, 5) Then ' БЧС без изменений
        bchs = bchs & "Без змін"
    Else
        For r = 14 To 18
            If Cells(r, 2) > 0 Then
                bchs = bchs & Cells(r, 1) & " " & Cells(r, 2) & "; "
            End If
            If Cells(r, 4) > 0 Then
                bchs = bchs & Cells(r, 3) & " " & Cells(r, 4) & "; "
            End If
        Next r
'        bchs = removeSubstring(bchs, " хворі 0; ")
'        bchs = removeSubstring(bchs, " переведення 0; ")
'        bchs = removeSubstring(bchs, "  ; ")
        bchs = bchs & Cells(r, 1) & ": " & Cells(r, 2) & ", " & Cells(r, 3) & ", " & Cells(r, 4)
    End If
    
    bchs = bchs & vbCrLf
End Function


Function br4() As String
    r0 = 46
    br4 = "4. "
    Debug.Assert Cells(r0, 1) = br4
    For r = r0 + 1 To 77
        v = Cells(r, 1)
        If v <> "" Then
              v = v & vbCrLf
            If r <> 47 Then
                v = "   " & v
            End If
        End If

        br4 = br4 & v
    Next r
End Function

Function part5() As String
    part5 = "5. Направити наступних військовослужбовців для техничної підтримки та живучості позиції "
    v = Cells(28, 2)
    v = Replace(v, vbLf, " ")
    part5 = part5 & v & ": "
    'chkRegEx.Pattern = "\d\d.\d\d.202\d"
    cntTotal = 0
    For c = 2 To 2 '  B-B выезд РВЗ
        For r = 29 To 40
            v = Cells(r, c)
            v = Replace(v, vbLf, " ")
            If v <> "" Then
                part5 = part5 & v & ", "
                cntTotal = cntTotal + 1
            End If
        Next r
    Next c
    part5 = Left(part5, Len(part5) - 2) & ". "
    
    part5 = part5 & "Наступним військовослужбовцям продовжити бойове чергування на позиціях (комплексах): "
    
    For c = 4 To 12  ' E - L
        reb = Cells(28, c)
        reb = Replace(reb, vbLf, " ")
        rebTxt = ""
        rebTxt = rebTxt & reb & ": "
        cnt = 0
        For r = 29 To 40
            v = Cells(r, c)
            v = Replace(v, vbLf, " ")
            If v <> "" Then
                cnt = cnt + 1
                rebTxt = rebTxt & v & ", "
                cntTotal = cntTotal + 1
            End If
        Next r
        If cnt > 0 Then
            rebTxt = Left(rebTxt, Len(rebTxt) - 2) & "; "
            part5 = part5 & rebTxt
        End If
    Next c
    
    l1 = Len(part5)
    part5 = Replace(part5, "  ", " ")
    While l1 <> Len(part5)
        part5 = Replace(part5, "  ", " ")
        l1 = Len(part5)
    Wend
    part5 = Left(part5, Len(part5) - 2) & "."
    part5 = Trim(part5) & vbCrLf
    
    cntCheck = Cells(43, 2)
    If cntTotal <> cntCheck Then
        MsgBox "Count check failure " & cntTotal & " vs. " & cntCheck
    End If
End Function


Function part6() As String
    r0 = 79
    part6 = "6. "
    Debug.Assert Cells(r0, 1) = part6
    r = r0 + 1
    part6 = part6 & Cells(r, 1) & " " & _
                    Cells(r, 2) & " " & _
                    Cells(r, 3) & " " & _
                    Cells(r, 4) & vbCrLf
End Function

Function partBRs() As String
    r0 = 85
    partBRs = "7. "
    Debug.Assert Cells(r0, 1) = partBRs
    For r = r0 + 1 To r0 + 36
        If Cells(r, 1) <> "" Then
            Str1 = txtPack(Cells(r, 1) & vbCrLf & Cells(r, 2) & ";")
            partBRs = partBRs & s1 & Str1 & vbCrLf
'            If s1 = "" Then
'                s1 = "   "
'            End If
        End If
    Next
End Function


Function txtPack(txt1 As String) As String
    l1 = Len(txt1)
    txt1 = Replace(txt1, "  ", " ")
    While l1 <> Len(txt1)
        txt1 = Replace(txt1, "  ", " ")
        l1 = Len(txt1)
    Wend
    txtPack = txt1
End Function

Function part8() As String
    r0 = 124
    part8 = "8. "
    Debug.Assert Cells(r0, 1) = part8
    r = r0 + 1
    part8 = part8 & Cells(r, 1) & vbCrLf
End Function

Function part9() As String
    part9 = "9. Не було" & vbCrLf
End Function


Function part10() As String
    part10 = "10. Не виявлено" & vbCrLf
End Function

Function sign() As String
    sign = "Командир роти РЕБ капітан               Микола Хамаганов"
End Function

Function writeTxtPids(txt As String, repDate As Date)
   Set fso = CreateObject("Scripting.FileSystemObject")
   Dim strPath As String
   strPath = fso.GetSpecialFolder(2) & "\" & repDate & "_" & fso.GetTempName()
   writeTxtPids = writeTxtFileAndOpen(txt, strPath)
End Function

Function writeTxtFileAndOpen(txt As String, fileName As String)
    fn = int_writeTxtFile(txt, fileName)
    If fn <> "" Then
        retVal = openTxtFile(fileName)
    End If
End Function

Public Function openTxtFile(fileName As String) As Integer
    openTxtFile = Shell("notepad.exe " & fileName, vbNormalFocus)
End Function

Public Function deleteFile(fileName As String) As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.deleteFile (fileName)
End Function


Function int_writeTxtFile(txt As String, fileName As String) As Integer
    int_writeTxtFilen = fileName
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(fileName)
    oFile.WriteLine txt
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
    Exit Function
    
errHndl:
    Debug.Print err.Description
    int_writeTxtFilen = ""
End Function


Public Function readTxtFile(fileName As String)
    Dim fso As Object
    Const ForReading = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.OpenTextFile(fileName, ForReading)
    readTxtFile = oFile.ReadAll()
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Function



Function getDate(shift As Integer) As Date
    Debug.Assert checkEnv()
    getDate = ActiveSheet.Cells(29, 1) + shift
End Function

Function getDateDataRow(d1 As Date) As Long
    For i = 1 To 1000000
        If ActiveWorkbook.Sheets("рпу").Cells(i, 1) = d1 Then
            getDateDataRow = i
            Exit Function
        End If
    Next i
End Function


Sub copyGBDData_int(d1 As Date)
    If MsgBox("Copy data for " & d1, vbYesNo, "Question") = vbNo Then
        Exit Sub
    End If

    If Not checkEnv() Then
        Exit Sub
    End If
    If Not chekIfWbIsOpened("штатка.xlsx") Then
        Exit Sub
    End If


    ' 1. Store the user's current calculation setting
    Dim originalCalculation As XlCalculation
    originalCalculation = Application.Calculation

    ' 2. Suppress recalculation (set to Manual)
    Application.Calculation = xlCalculationManual
    
    ' Optional: Suppress screen flickering and background events for extra speed
    Application.ScreenUpdating = False
    Application.EnableEvents = False


    r1 = getDateDataRow(d1)
    Set wsSource = ActiveWorkbook.Sheets("рпу")
    Set wsTarget = ActiveSheet
    For r = 0 To 9
        rTarget = 29 + r
        rSource = r1 + r
        For c = 1 To 12
            If wsTarget.Cells(rTarget, c) <> wsSource.Cells(rSource, c).Value Then
                wsTarget.Cells(rTarget, c) = wsSource.Cells(rSource, c).Value
                rColor = 28
            Else
                rColor = 26
            End If
            If c <> 1 Then
                wsTarget.Cells(rTarget, c).Interior.Color = wsTarget.Cells(rColor, 1).Interior.Color
            End If
        Next c
    Next
    
    copyBchs d1, "dataCurDate"
    copyBchs d1 - 1, "dataPrevDate"
    
    
    Application.Calculation = xlCalculationAutomatic   '  originalCalculation
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub copyBchs(d1 As Date, txtPos As String)
    Dim dd As String
    dd = Format(d1, "yymmdd")
    sn = getLastHistSheetName(dd)
    Set wsStaff = Workbooks("штатка.xlsx").Sheets(sn)
    Set wsTarget = ActiveSheet
    hash1 = RangeToMD5(wsStaff.Range("S134:S141")) ' ЖБД диапазон 1
    hash2 = RangeToMD5(wsStaff.Range("U134:U141")) ' ЖБД диапазон 2
    Debug.Assert hash1 = "790455133cb9dd7bfcff9eae1b77936d"
    Debug.Assert hash2 = "63f11f011be33b6d99bf31fb63e41c3b"
'    Dim vArray() As Variant
'    Dim rrr As Range
'    Set rrr = wsStaff.Range("S132:V136")
'    vArray = rrr
    If txtPos = "dataCurDate" Then
        r1 = 13
        r2 = 17
        r3 = 18
        r4 = 19
    Else  '  "dataPrevDate"
        r1 = 3
        r2 = 7
        r3 = 8
        r4 = 9
    End If
    
    Range(Cells(r1, 1), Cells(r2, 4)) = wsStaff.Range("S134:V138").Value
    Range(Cells(r3, 1), Cells(r4, 4)) = wsStaff.Range("S140:V141").Value
    
End Sub

Sub copyGBDData()
    copyGBDData_int (getDate(1))
End Sub

Sub copyGBDDataBack()
    copyGBDData_int (getDate(-1))
End Sub
