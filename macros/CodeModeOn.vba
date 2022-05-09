Sub CodeModeOn()

    ActiveDocument.Bookmarks.Add Name:="StartPos" Range:=Selection.Range
    Selection.Font.Name = "Courier New"

End Sub