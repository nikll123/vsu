Attribute VB_Name = "Dev"
Public Const constRepoDir As String = "C:\repos\NN\"


'Sub importMacro()
'    for
'End Sub
Sub exportMacroAll()
    int_exportMacro ("")
End Sub
Sub exportMacroChanged()
    Dim mlist As String
    mlist = getChangedModules()
    int_exportMacro "," & mlist
    MsgBox "Done: " & constRepoDir & vbCr & mlist
End Sub

Sub int_exportMacro(macroList As String)
    Dim fn As String
    For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
        If m.Type = 1 Then ' vbext_ct_StdModule
            doExprot = True
            If macroList = "" Or InStr(1, macroList, "," & m.Name & ",") > 0 Then
                fn = getFileName(m.Name)
                deleteFile fn
                m.Export (fn)
            'Debug.Print m.Name
            End If
        End If
    Next
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


Function getChangedModules() As String
    getChangedModules = ""
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
                getChangedModules = getChangedModules & m.Name & ","
            End If
            
        End If
    Next
    
    Debug.Print getChangedModules

End Function
