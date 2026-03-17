Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False

    If target.Cells.CountLarge > 1 Then GoTo SafeExit

    ' Column H = shape code
    If Not Intersect(target, Me.Columns("H")) Is Nothing Then
        target.Select
        DeleteShapeFromRow
        
        If Trim$(target.Value) <> "" Then
            PlaceShapeFromCode
        End If
        
        GoTo SafeExit
    End If

    ' M:W = parameters
    If Not Intersect(target, Me.Columns("M:W")) Is Nothing Then
        Dim r As Long
        r = target.Row

        If Trim$(Me.Cells(r, "H").Value) <> "" Then
            Me.Cells(r, "H").Select
            DeleteShapeFromRow
            PlaceShapeFromCode
        End If
    End If

SafeExit:
    Application.EnableEvents = True
End Sub
