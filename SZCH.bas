Attribute VB_Name = "SZCH"

'
'Public Enum fioParam
'fioNormal      ' Surname Name1 Name2 (Compress spaces)
'fioFirstUpper  ' SURNAME Name1 Name2
'fioInit        ' SURNAME N.N.
'fioNameSur     ' Name1 SURNAME
'
'End Enum

'Аксьон Олександр Михайлович -->Аксьон Олександр Михайлович
Public Function fioNormal(fioArg As String) As String
fioNormal = strAllTrim(fioArg)
End Function

'Аксьон Олександр Михайлович --> АКСЬОН Олександр Михайлович
Public Function fioSurnameUpper(fioArg As String) As String
Dim fio As String
fio = fioNormal(fioArg)
surname = GetWordByIndex(fio, 1)
name1 = GetWordByIndex(fio, 2)
name2 = GetWordByIndex(fio, 3)
fioSurnameUpper = UCase(surname) + " " + name1 + " " + name2
End Function

'Аксьон Олександр Михайлович --> АКСЬОН О.М.
Public Function fioInitSurnameUpper(fioArg As String) As String
Dim fio As String
fio = fioNormal(fioArg)
surname = GetWordByIndex(fio, 1)
name1 = GetWordByIndex(fio, 2)
name2 = GetWordByIndex(fio, 3)
fioInitSurnameUpper = UCase(surname) + " " + Left(name1, 1) + "." + Left(name2, 1) + "."
End Function

'Аксьон Олександр Михайлович --> АОМ
Public Function fioCode(fioArg As String, len1 As Integer) As String
Dim fio As String
fio = fioNormal(fioArg)
surname = GetWordByIndex(fio, 1)
name1 = GetWordByIndex(fio, 2)
name2 = GetWordByIndex(fio, 3)

fioCode = Left(surname, len1) + Left(name1, len1) + Left(name2, len1)

End Function

'Аксьон Олександр Михайлович --> О. АКСЬОН
Public Function n1FioSurnameUpper(fioArg As String) As String
n1FioSurnameUpper = fio2i1fUpper(fioArg)
End Function

'Аксьон Олександр Михайлович --> О. АКСЬОН
Public Function fio2i1fUpper(fioArg As String) As String
Dim fio As String
fio = fioNormal(fioArg)
surname = GetWordByIndex(fio, 1)
name1 = GetWordByIndex(fio, 2)
name2 = GetWordByIndex(fio, 3)
fio2i1fUpper = Left(name1, 1) + ". " + UCase(surname)
End Function

'Аксьон Олександр Михайлович --> О. АКСЬОН
Public Function fio2ifUpper(fioArg As String) As String
Dim fio As String
fio = fioNormal(fioArg)
surname = GetWordByIndex(fio, 1)
name1 = GetWordByIndex(fio, 2)
fio2ifUpper = name1 + " " + UCase(surname)
End Function

'Аксьон Олександр Михайлович --> Аксьон О.М.
Public Function fioInit(fioArg As String) As String
Dim fio As String
fio = fioNormal(fioArg)
surname = GetWordByIndex(fio, 1)
name1 = GetWordByIndex(fio, 2)
name2 = GetWordByIndex(fio, 3)
fioInit = surname + " " + Left(name1, 1) + "." + Left(name2, 1) + "."
End Function

'Аксьон Олександр Михайлович --> Аксьон О.М.
'Public Function fioNameSurnameUpper(fioArg As String) As String
'Dim fio As String
'fio = fioNormal(fioArg)
'surname = GetWordByIndex(fio, 1)
'name1 = GetWordByIndex(fio, 2)
'name2 = GetWordByIndex(fio, 3)
'fioNameSurnameUpper = surname + " " + Left(name1, 1) + "." + Left(name2, 1) + "."
'End Function


' выдает слово из строки по заданному номеру
Public Function GetWordByIndex(text As String, i As Integer) As String
text = strAllTrim(text)
a = Split(" " + text, " ")
On Error Resume Next
GetWordByIndex = a(i)
On Error GoTo 0
End Function


' удаление всех повторяющихся пробелов
Public Function strAllTrim(text As String) As String
    text = Trim(text)
    a = Split(text, " ")
    For Each w In a
        If Len(w) > 0 Then
            strAllTrim = strAllTrim + " " + w
        End If
    Next w
    strAllTrim = Trim(strAllTrim)
End Function

'оператор 1 комплексу --> Оператор 1 комплексу
Public Function str1Upper(text As String) As String
    text = Trim(text)
    str1Upper = UCase(Left(text, 1)) & Mid(text, 2)
End Function
