Sub KeyPressHandler()
    If ActiveDocument.Variables("CodeMode").Value = False Then
        ActiveDocument.Variables("CodeMode").Delete
        ActiveDocument.Variables.Add Name:="CodeMode", Value:=True
        CodeModeOn
    Else
        ActiveDocument.Variables("CodeMode").Delete
        ActiveDocument.Variables.Add Name:="CodeMode", Value:=False
        CodeModeOff
    End If

End Sub