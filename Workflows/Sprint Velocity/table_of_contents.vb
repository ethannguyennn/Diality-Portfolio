Private Const TOC_SHEET As String = "Velocity"

' -- Fires when user clicks a hyperlink on any sheet -----------------
' Unhides the target Sprint Velocity sheet BEFORE Excel navigates to it
' so the navigation succeeds without throwing a "reference not valid" error.
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)

    ' Only act on clicks originating from the TOC tab
    If Sh.Name <> TOC_SHEET Then Exit Sub

    ' Parse sheet name out of SubAddress  (e.g. 'Sprint 5 Velocity'!A1)
    Dim subAddr As String
    subAddr = Replace(Target.SubAddress, "'", "")   ' strip single quotes around names with spaces

    Dim targetSheet As String
    If InStr(subAddr, "!") > 0 Then
        targetSheet = Left(subAddr, InStr(subAddr, "!") - 1)
    Else
        targetSheet = subAddr
    End If

    If targetSheet = "" Then Exit Sub

    ' Unhide the sheet if it exists and is currently hidden
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Me.Worksheets(targetSheet)
    On Error GoTo 0

    If Not ws Is Nothing Then
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    End If

End Sub

' -- Fires the instant the user navigates AWAY from any sheet ---------
' Auto-hides Sprint Velocity sheets the moment focus leaves them.
' Pattern match: "Sprint # Velocity" — all other tabs are untouched,
' including the Velocity TOC itself.
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)

    Dim n As String
    n = Sh.Name

    ' Must start with "Sprint ", end with " Velocity",
    ' and be at least 17 chars long (shortest valid: "Sprint 1 Velocity")
    If Len(n) >= 17 _
        And Left(n, 7) = "Sprint " _
        And Right(n, 9) = " Velocity" Then
        Sh.Visible = xlSheetHidden
    End If

End Sub

' -- Fires automatically when the workbook is opened -----------------
' Hides all Sprint Velocity sheets on load so the tab bar is always
' clean regardless of how the file was last closed.
Private Sub Workbook_Open()

    Dim ws As Worksheet
    Dim n As String

    For Each ws In Me.Worksheets
        n = ws.Name
        If Len(n) >= 17 _
            And Left(n, 7) = "Sprint " _
            And Right(n, 9) = " Velocity" Then
            ws.Visible = xlSheetHidden
        End If
    Next ws

End Sub
