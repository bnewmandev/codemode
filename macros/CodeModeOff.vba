Sub CodeModeOff()

    Set rngStart = ActiveDocument.Bookmarks("StartPos").Range
    Set CodeRange = ActiveDocument.Range(rngStart.Start, Selection.Range.End)
    CodeRange.MoveEndWhile Chr(32), wdBackward
    Set newCodeRange = ActiveDocument.Range
    newCodeRange.Collapse wdCollapseEnd
    newCodeRange.FormattedText = CodeRange
    CodeRange.Select
    Selection.NoProofing = True
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = 2829099
    Selection.MoveRight
    Selection.InsertAfter " "
    Selection.Font.Name = "Calibri"
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
    Selection.MoveRight
    
End Sub