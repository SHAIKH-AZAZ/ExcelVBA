Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' ---- Column G handler ----
    If Not Intersect(Target, Me.Columns("G")) Is Nothing Then
        Handle_L_Column_Change Target
        GoTo SafeExit
    End If

    ' ---- Columns I:N handler ----
    If Not Intersect(Target, Me.Columns("I:N")) Is Nothing Then
        Handle_MS_Columns_Change Target
        GoTo SafeExit
    End If

SafeExit:
    Application.EnableEvents = True
End Sub


Private Sub Handle_L_Column_Change(ByVal Target As Range)

    If Target.Cells.Count > 1 Then Exit Sub
    If Trim(Target.Value) = "" Then Exit Sub

    ' Ensure correct row context for your existing macro
    Target.Select

    ' Place shape
    PlaceShapeFromCode

End Sub
Private Sub Handle_MS_Columns_Change(ByVal Target As Range)

    Dim r As Long

    ' Ignore multi-cell paste if needed
    If Target.Cells.Count > 1 Then Exit Sub

    r = Target.Row

    ' If no shape code in column L, nothing to do
    If Trim(Me.Cells(r, "G").Value) = "" Then Exit Sub

    ' Maintain row context (your macros depend on ActiveCell)
    Me.Cells(r, "G").Select

    ' 1. Delete existing shape
    DeleteShapeFromRow

    ' 2. Recreate shape with updated M:S values
    PlaceShapeFromCode

End Sub