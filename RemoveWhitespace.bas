REM  *****  BASIC  *****

Sub Main

End Sub

Sub RemoveWhitespace()
    Dim oSel As Object
    Dim oRanges As Object
    Dim i As Long, j As Long

    oSel = ThisComponent.getCurrentSelection()

    If oSel.supportsService("com.sun.star.sheet.SheetCellRanges") Then
        oRanges = oSel
    Else
        Set oRanges = ThisComponent.createInstance( _
            "com.sun.star.sheet.SheetCellRanges")
        oRanges.addRangeAddress oSel.RangeAddress, False
    End If

    Dim oRange As Object
    For Each oRange In oRanges.getRangeAddresses()

        Dim oSheet As Object
        oSheet = ThisComponent.Sheets.getByIndex(oRange.Sheet)

        Dim oCell As Object
        For i = oRange.StartRow To oRange.EndRow
            For j = oRange.StartColumn To oRange.EndColumn

                oCell = oSheet.getCellByPosition(j, i)

                If oCell.Type = com.sun.star.table.CellContentType.TEXT Then
                    Dim s As String
                    s = oCell.String
             
                    s = Trim(s)
                    
                    ' remove multiple spaces
                    Do While InStr(s, "  ") > 0
                        s = Replace(s, "  ", " ")
                    Loop

                    oCell.String = s
                End If
            Next j
        Next i
    Next oRange

    MsgBox "White space removed!", 64, "Text Normalizer"
End Sub




