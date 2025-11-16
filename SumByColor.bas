REM  *****  BASIC  *****

Sub Main

End Sub

Public Function SumByColor(rangeText As String, colorCell As String) As Double
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oRange As Object
    Dim oRefCell As Object
    Dim oCell As Object
    Dim total As Double
    Dim targetColor As Long
    Dim i As Long, j As Long
    Dim numRows As Long, numCols As Long

    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    ' Convert range text to a range object
    oRange = oSheet.getCellRangeByName(rangeText)
    ' Convert reference cell text to a single cell object
    oRefCell = oSheet.getCellRangeByName(colorCell).getCellByPosition(0, 0)

    targetColor = oRefCell.CellBackColor
    total = 0

    numCols = oRange.Columns.Count
    numRows = oRange.Rows.Count

    For i = 0 To numRows - 1
        For j = 0 To numCols - 1
            oCell = oRange.getCellByPosition(j, i)
            If IsNumeric(oCell.Value) Then
                If oCell.CellBackColor = targetColor Then
                    total = total + oCell.Value
                End If
            End If
        Next j
    Next i

    SumByColor = total
End Function




