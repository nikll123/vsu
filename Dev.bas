Attribute VB_Name = "Dev"
Public Const constRepoDir As String = "R:\Media\macro\NN\"


'Sub importMacro()
'    for
'End Sub

Sub exportMacroAll()
For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
    If m.Type = 1 Then ' vbext_ct_StdModule
        m.Export (constRepoDir & m.Name & ".bas")
        Debug.Print m.Name
    End If
Next
MsgBox "Done: " & constRepoDir
End Sub

Sub removeMacroAll()
If MsgBox("Удалить все макросы?", vbOKCancel, constStuffWbName) = vbCancel Then Exit Sub
For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
    If m.Type = 1 Then ' vbext_ct_StdModule
        Debug.Print m.Name
        Workbooks("PERSONAL.xlsb").VBProject.VBComponents.Remove (m)
    End If
Next
End Sub


Sub compareMacroAll()
For Each m In Workbooks("PERSONAL.xlsb").VBProject.VBComponents
    If m.Type = 1 Then ' vbext_ct_StdModule
        Debug.Print m.Name
        m.Export (constRepoDir & m.Name & ".bas")
      
        'Workbooks("PERSONAL.xlsb").VBProject.VBComponents.Remove (m)
    End If
Next
End Sub
