Sub CopyRangeA()

    For x = 1 To Sheets.Count
        Worksheets(x).Activate
        Worksheets(x).Range("B6").Select
    Next x
    
End Sub
