// Digunakan untuk select cell tertentu di semua sheet, digunakan untuk import image agar lokasi pas ke cell tertentu Sub SetSelectedCell()

Sub CopyRangeA()
    For x = 1 To Sheets.Count
        Worksheets(x).Activate
        Worksheets(x).Range("B6").Select
    Next x
End Sub

