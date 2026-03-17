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

    ' Library
    On Error Resume Next
    Set lib = ThisWorkbook.Worksheets("ShapeLibrary")
    On Error GoTo 0
    If lib Is Nothing Then
        MsgBox "Critical Error: The sheet 'ShapeLibrary' is missing.", vbCritical
        Exit Sub
    End If

    ' Shape code from column G
    shpCode = Trim$(ws.Cells(r, "G").Value)
    If shpCode = "" Then Exit Sub

    ' Get source shape from library
    On Error Resume Next
    Set srcShape = lib.Shapes(shpCode)
    On Error GoTo 0
    If srcShape Is Nothing Then
        MsgBox "Shape Code '" & shpCode & "' not found in library.", vbExclamation
        Exit Sub
    End If

    finalName = shpCode & "_" & r

    ' Delete old shape if exists
    On Error Resume Next
    ws.Shapes(finalName).Delete
    On Error GoTo 0

    ' Paste + capture reliably
    shpCountBefore = ws.Shapes.Count
    srcShape.Copy
    ws.Paste Destination:=ws.Cells(r, "Z")

    If ws.Shapes.Count <= shpCountBefore Then Exit Sub ' paste failed
    Set newShp = ws.Shapes(ws.Shapes.Count)

    ' Name
    newShp.Name = finalName

    ' Position in Z
    Set targetCell = ws.Cells(r, "Z")
    With newShp
        .Placement = xlMove
        .Left = targetCell.Left + (targetCell.Width - .Width) / 2
        .Top = targetCell.Top + (targetCell.Height - .Height) / 2
    End With

    ' Text replacement {A}..{K} from L..V
    ReDim keys(1 To 11)
    ReDim values(1 To 11)

    For i = 1 To 11
        keys(i) = "{" & Chr$(64 + i) & "}"          ' {A}..{K}
        values(i) = CStr(ws.Cells(r, 11 + i).Value) ' L..V
    Next i

    ProcessTextRecursively newShp, keys, values

    originalCell.Select
End Sub


Public Sub DeleteShapeFromRow()
    Dim ws As Worksheet
    Dim shpCode As String, targetName As String
    Dim r As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set ws = ActiveSheet
    r = ActiveCell.Row

    shpCode = Trim$(ws.Cells(r, "G").Value)
    If shpCode = "" Then Exit Sub

    targetName = shpCode & "_" & r

    On Error Resume Next
    ws.Shapes(targetName).Delete
    On Error GoTo 0
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

