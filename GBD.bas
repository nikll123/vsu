Attribute VB_Name = "GBD"

'Dim chkRegEx As New RegExp

Sub gbdGenerate()
    wbName = "1 РПУ.xlsx"
    If ActiveWorkbook.Name <> wbName Then
        MsgBox "The activebook must be " & wbName
        Exit Sub
    End If
    wsName = "Лист2"
    If ActiveSheet.Name <> wsName Then
        MsgBox "The activesheet must be " & wsName
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
    lastBrHashRow = 61
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

    Dim repDate As Date
    repDate = Cells(29, 1)
    writeTxtFileAndOpen txt1, repDate
    
End Sub

Function bchs() As String
    chgBchs = Cells(1, 5)
    bchs = "2. "
    If Cells(1, 5) Then ' БЧС без изменений
        bchs = bchs & "Без змін" & vbCrLf
    Else
        For r = 14 To 19
            bchs = bchs & IIf(r = 14, "", "   ") & Cells(r, 1) & " " & Cells(r, 2) & "; " & Cells(r, 3) & " " & Cells(r, 4) & "; " & vbCrLf
        Next r
    End If
End Function


Function br4() As String
    br4 = "4. "
    For r = 47 To 50
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
    
    For c = 5 To 12  ' E - L
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
    part6 = "6. "
    part6 = part6 & Cells(55, 1) & " " & _
                    Cells(55, 2) & " " & _
                    Cells(55, 3) & " " & _
                    Cells(55, 4) & vbCrLf
End Function

Function partBRs() As String
    partBRs = "7. "
    s1 = ""
    For r = 61 To 79
        If Cells(r, 1) <> "" Then
        Str1 = txtPack(Cells(r, 1) & " " & Cells(r, 2) & ";")
            partBRs = partBRs & s1 & Str1 & vbCrLf
            If s1 = "" Then
                s1 = "   "
            End If
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
    part8 = "8. "
    part8 = part8 & Cells(82, 1) & vbCrLf
    
End Function

Function part9() As String
    part9 = "9. Не було" & vbCrLf
End Function


Function part10() As String
    part10 = "10. Не виявлено" & vbCrLf
End Function


Function writeTxtFileAndOpen(txt As String, repDate As Date)

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
strPath = fso.GetSpecialFolder(2) & "\" & repDate & "_" & fso.GetTempName()
Set oFile = fso.CreateTextFile(strPath)
oFile.WriteLine txt
oFile.Close
Set fso = Nothing
Set oFile = Nothing
retVal = Shell("notepad.exe " & strPath, vbNormalFocus)

End Function
