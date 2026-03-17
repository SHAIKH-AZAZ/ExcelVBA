Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' Column G = shape code
    If Not Intersect(Target, Me.Columns("G")) Is Nothing Then
        If Target.Cells.Count = 1 Then
            If Trim$(Target.Value) <> "" Then
                Target.Select
                PlaceShapeFromCode
            End If
        End If
        GoTo SafeExit
    End If

    ' L:V = A..K parameters
    If Not Intersect(Target, Me.Columns("L:V")) Is Nothing Then
        If Target.Cells.Count = 1 Then
            Dim r As Long
            r = Target.Row

            If Trim$(Me.Cells(r, "G").Value) <> "" Then
                Me.Cells(r, "G").Select
                DeleteShapeFromRow
                PlaceShapeFromCode
            End If
        End If
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

