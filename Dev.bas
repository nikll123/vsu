Attribute VB_Name = "Dev"
Public Const constRepoDir As String = "R:\Media\macro\NN\"


'Sub importMacro()
'    for
'End Sub

Sub exportMacroAll()
For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
    If m.Type = 1 Then ' vbext_ct_StdModule
        m.Export (getFileName(m.Name))
        'Debug.Print m.Name
    End If
Next
MsgBox "Done: " & constRepoDir
End Sub

Sub removeMacroAll()
If MsgBox("Удалить все макросы?", vbOKCancel, constStuffWbName) = vbCancel Then Exit Sub
For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
    If m.Type = 1 Then ' vbext_ct_StdModule
        'Debug.Print m.Name
        Workbooks("PERSONAL.xlsb").VBProject.VBComponents.Remove (m)
    End If
Next
End Sub

Function getFileName(moduleName As String)
    getFileName = constRepoDir & moduleName & ".bas"
End Function


Sub compareMacroAll()
    res = ""
    For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
        If m.Type = 1 Then ' vbext_ct_StdModule
            'Debug.Print m.Name
            textExcel = m.CodeModule.Lines(1, m.CodeModule.CountOfLines)
            textFile = readTxtFile(getFileName(m.Name))
            textFile = Mid(textFile, InStr(1, textFile, vbCrLf) + 2) ' first line is not used
            textFile = Left(textFile, Len(textFile) - 2) ' the last line is not used
            
            md5Excel = StringToMD5Hex(textExcel)
            md5File = StringToMD5Hex(textFile)
                
            If Not md5Excel = md5File Then
                res = res & vbCr & m.Name
            End If
            
        End If
    Next
    
    If Len(res) > 0 Then
        res = "Changed modules: " & res
    Else
        res = "No changes"
    End If
    
    Debug.Print res

End Sub
