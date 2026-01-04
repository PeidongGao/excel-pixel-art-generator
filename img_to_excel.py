Sub ImageToExcelPixels()
    Dim pic As Picture
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim cell As Range

    Set pic = ActiveSheet.Pictures(1)

    pic.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    With ActiveSheet.Paste
        .Top = 0
        .Left = 0
    End With

    For y = 1 To 64
        For x = 1 To 64
            Set cell = Cells(y, x)
            cell.Interior.Color = RGB( _
                Int(Rnd() * 255), _
                Int(Rnd() * 255), _
                Int(Rnd() * 255))
        Next x
    Next y
End Sub
