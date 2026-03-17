Option Explicit

Public Sub PlaceShapeFromCode()
    Dim ws As Worksheet, lib As Worksheet
    Dim srcShape As Shape, newShp As Shape
    Dim targetCell As Range, originalCell As Range
    Dim shpCode As String, finalName As String
    Dim r As Long
    Dim keys() As Variant, values() As Variant
    Dim i As Long
    Dim shpCountBefore As Long

    If TypeName(Selection) <> "Range" Then Exit Sub

    Set ws = ActiveSheet
    Set originalCell = ActiveCell
    r = originalCell.Row

    On Error Resume Next
    Set lib = ThisWorkbook.Worksheets("ShapeLibrary")
    On Error GoTo 0
    If lib Is Nothing Then
        MsgBox "Critical Error: The sheet 'ShapeLibrary' is missing.", vbCritical
        Exit Sub
    End If

    shpCode = Trim$(ws.Cells(r, "H").Value)
    
    'Delete any old shape on this row first
    DeleteAllShapesFromRow ws, r
    
    If shpCode = "" Then Exit Sub

    On Error Resume Next
    Set srcShape = lib.Shapes(shpCode)
    On Error GoTo 0
    If srcShape Is Nothing Then
        MsgBox "Shape Code '" & shpCode & "' not found in library.", vbExclamation
        Exit Sub
    End If

    finalName = shpCode & "_" & r

    shpCountBefore = ws.Shapes.Count
    srcShape.Copy
    ws.Paste Destination:=ws.Cells(r, "AA")

    If ws.Shapes.Count <= shpCountBefore Then Exit Sub
    Set newShp = ws.Shapes(ws.Shapes.Count)

    newShp.Name = finalName

    Set targetCell = ws.Cells(r, "AA")
    With newShp
        .Placement = xlMoveAndSize
        .Left = targetCell.Left + (targetCell.Width - .Width) / 2
        .Top = targetCell.Top + (targetCell.Height - .Height) / 2
    End With

    ReDim keys(1 To 11)
    ReDim values(1 To 11)

    For i = 1 To 11
        keys(i) = "{" & Chr$(64 + i) & "}"
        values(i) = CStr(ws.Cells(r, 12 + i).Value)   'M:W
    Next i

    ProcessTextRecursively newShp, keys, values

    originalCell.Select
End Sub
Public Sub DeleteShapeFromRow()
    Dim ws As Worksheet
    Dim r As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    r = ActiveCell.Row

    DeleteAllShapesFromRow ws, r
End Sub


Private Sub DeleteAllShapesFromRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim i As Long
    Dim nm As String
    
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Right$(nm, Len("_" & CStr(r))) = "_" & CStr(r) Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub



Private Sub ProcessTextRecursively(ByVal shp As Shape, ByRef keys() As Variant, ByRef values() As Variant)
    Dim subShp As Shape
    Dim txt As String
    Dim i As Long

    If shp.Type = msoGroup Then
        For Each subShp In shp.GroupItems
            ProcessTextRecursively subShp, keys, values
        Next subShp
    Else
        On Error Resume Next
        If shp.TextFrame2.HasText Then
            txt = shp.TextFrame2.TextRange.Text
            For i = LBound(keys) To UBound(keys)
                txt = Replace(txt, CStr(keys(i)), CStr(values(i)))
            Next i
            shp.TextFrame2.TextRange.Text = txt
        End If
        On Error GoTo 0
    End If
End Sub

