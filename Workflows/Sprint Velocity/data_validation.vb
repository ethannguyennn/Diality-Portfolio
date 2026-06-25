' Module1 (or any standard module)
Public Sub ApplySprintValidations()
    Dim ws As Worksheet
    Dim sel As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sprint Template")
    Set sel = ThisWorkbook.Worksheets("Selection")

    Dim lastRow As Long
    lastRow = ws.UsedRange.Rows.Count
    If lastRow < 4 Then Exit Sub

    ' Standard dropdown columns
    Dim validations(7, 1) As String
    validations(0, 0) = "A": validations(0, 1) = "=Selection!$F$2:$F$101"
    validations(1, 0) = "B": validations(1, 1) = "=Selection!$B$2:$B$5"
    validations(2, 0) = "C": validations(2, 1) = "=Selection!$A$2:$A$26"
    validations(3, 0) = "D": validations(3, 1) = "=Selection!$A$2:$A$26"
    validations(4, 0) = "E": validations(4, 1) = "=Selection!$D$2:$D$5"
    validations(5, 0) = "F": validations(5, 1) = "=Selection!$H$2:$H$4"
    validations(6, 0) = "G": validations(6, 1) = "=Selection!$K$2:$K$3"
    validations(7, 0) = "J": validations(7, 1) = "=Selection!$J$2:$J$5"

    Dim i As Integer
    For i = 0 To 7
        With ws.Range(validations(i, 0) & "4:" & validations(i, 0) & lastRow).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=validations(i, 1)
            .InCellDropdown = True
        End With
    Next i

    ' Column I: Bug/Dialin vs normal
    Dim bugSource As String:    bugSource = "=Selection!$L$2:$L$13"
    Dim normalSource As String: normalSource = "=Selection!$E$2:$E$7"

    Dim types As Variant
    types = ws.Range("E4:E" & lastRow).Value

    Dim r As Long
    r = 0
    Do While r < UBound(types, 1)
        Dim isBug As Boolean
        isBug = (types(r + 1, 1) = "Bug" Or types(r + 1, 1) = "Dialin ticket")

        Dim j As Long
        j = r
        Do While j + 1 < UBound(types, 1)
            Dim nextIsBug As Boolean
            nextIsBug = (types(j + 2, 1) = "Bug" Or types(j + 2, 1) = "Dialin ticket")
            If nextIsBug <> isBug Then Exit Do
            j = j + 1
        Loop

        With ws.Range("I" & (r + 4) & ":I" & (j + 4)).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=IIf(isBug, bugSource, normalSource)
            .InCellDropdown = True
        End With
        r = j + 1
    Loop
End Sub
