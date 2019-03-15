Sub InsertAndCopyColumn(worksheetName, range)
    With Worksheets(worksheetName)
            .Columns(range).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove 'insert new column
            .Columns(range - 1).Copy Destination:=.Columns(range)
    End With
End Sub

