Option Explicit

Sub Copy_FilterColumns_ValuesOnly_WithShapes()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim srcHeaderRow As Long, dstHeaderRow As Long
    Dim srcStartRow As Long, dstStartRow As Long
    Dim lastRow As Long
    Dim map As Variant
    Dim i As Long
    Dim srcCol As Long, dstCol As Long
    Dim srcCodeCol As Long, dstShapeCol As Long

    Set wsSrc = ThisWorkbook.Worksheets("base_format")
    Set wsDst = ThisWorkbook.Worksheets("Filter_format")

    srcHeaderRow = 2
    dstHeaderRow = 2
    srcStartRow = srcHeaderRow + 1
    dstStartRow = dstHeaderRow + 1

    'SourceHeader -> DestinationHeader
    map = Array( _
        Array("Description", "MEMBER / Discription"), _
        Array("Bar Mark", "BAR MARKS"), _
        Array("Dia", "ø of Bars"), _
        Array("No of Elmts", "No of MEMBER"), _
        Array("No of Bars", "No of Bars"), _
        Array("Total No", "Total No"), _
        Array("Code", "Code"), _
        Array("Cutting Length", "Cutting length"), _
        Array("A", "A"), Array("B", "B"), Array("C", "C"), _
        Array("D", "D"), Array("E", "E"), Array("F", "F"), _
        Array("Weight (Kg)", "Weight (Kg)") _
    )

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Last row based on Code
    srcCol = FindHeaderCol(wsSrc, srcHeaderRow, "Code")
    If srcCol = 0 Then
        MsgBox "Source header 'Code' not found.", vbExclamation
        GoTo CleanExit
    End If
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCol).End(xlUp).Row
    If lastRow < srcStartRow Then GoTo CleanExit

    'Clear destination mapped columns (data area)
    For i = LBound(map) To UBound(map)
        dstCol = FindHeaderCol(wsDst, dstHeaderRow, CStr(map(i)(1)))
        If dstCol > 0 Then
            wsDst.Range(wsDst.Cells(dstStartRow, dstCol), wsDst.Cells(wsDst.Rows.Count, dstCol)).ClearContents
        End If
    Next i

    'Delete destination shapes in Shape column
    dstShapeCol = FindHeaderCol(wsDst, dstHeaderRow, "Shape")
    If dstShapeCol > 0 Then DeleteShapesInColumnFromRow wsDst, dstShapeCol, dstStartRow

    'COPY: Values + Formats (NO formulas)
    For i = LBound(map) To UBound(map)
        Dim sHdr As String, dHdr As String
        sHdr = CStr(map(i)(0))
        dHdr = CStr(map(i)(1))

        srcCol = FindHeaderCol(wsSrc, srcHeaderRow, sHdr)
        dstCol = FindHeaderCol(wsDst, dstHeaderRow, dHdr)

        If srcCol > 0 And dstCol > 0 Then
            'Column width
            wsDst.Columns(dstCol).ColumnWidth = wsSrc.Columns(srcCol).ColumnWidth

            With wsDst.Range(wsDst.Cells(dstStartRow, dstCol), wsDst.Cells(dstStartRow + (lastRow - srcStartRow), dstCol))
                .Value = wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Value
                .NumberFormat = wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).NumberFormat
                .Font.Name = wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Font.Name
                .Font.Size = wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Font.Size
            End With

            'Optional: Copy full cell formatting (borders, fill, etc.)
            wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Copy
            wsDst.Cells(dstStartRow, dstCol).PasteSpecial xlPasteFormats
        End If
    Next i

    Application.CutCopyMode = False

    'Copy shapes by NAME = Code_Row (e.g., 12_5)
    srcCodeCol = FindHeaderCol(wsSrc, srcHeaderRow, "Code")
    dstShapeCol = FindHeaderCol(wsDst, dstHeaderRow, "Shape")
    If srcCodeCol > 0 And dstShapeCol > 0 Then
        CopyShapesByCodeRow wsSrc, wsDst, srcCodeCol, srcStartRow, lastRow, dstStartRow, dstShapeCol
    End If

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


'--- Find column number by header text (exact)
Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long, txt As String
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        txt = Trim$(CStr(ws.Cells(headerRow, c).Value2))
        If StrComp(txt, Trim$(headerText), vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Private Function ShapeExists(ByVal ws As Worksheet, ByVal shpName As String) As Boolean
    On Error Resume Next
    ShapeExists = Not ws.Shapes(shpName) Is Nothing
    On Error GoTo 0
End Function

Private Sub DeleteShapesInColumnFromRow(ByVal ws As Worksheet, ByVal targetCol As Long, ByVal startRow As Long)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        If Not ws.Shapes(i).TopLeftCell Is Nothing Then
            If ws.Shapes(i).TopLeftCell.Column = targetCol And ws.Shapes(i).TopLeftCell.Row >= startRow Then
                ws.Shapes(i).Delete
            End If
        End If
        On Error GoTo 0
    Next i
End Sub

'--- Copy shape whose name is Code_SourceRow (ex: 12_5) and paste into destination Shape cell on same data row
Private Sub CopyShapesByCodeRow( _
    ByVal wsSrc As Worksheet, ByVal wsDst As Worksheet, _
    ByVal srcCodeCol As Long, _
    ByVal srcStartRow As Long, ByVal srcLastRow As Long, _
    ByVal dstStartRow As Long, ByVal dstShapeCol As Long)

    Dim r As Long, dstRow As Long
    Dim codeVal As Variant
    Dim shpName As String
    Dim shp As Shape
    Dim dstCell As Range
    Dim pasted As Shape

    For r = srcStartRow To srcLastRow
        codeVal = wsSrc.Cells(r, srcCodeCol).Value2
        If Len(Trim$(CStr(codeVal))) > 0 Then

            shpName = Trim$(CStr(codeVal)) & "_" & CStr(r)

            If ShapeExists(wsSrc, shpName) Then
                Set shp = wsSrc.Shapes(shpName)
                dstRow = dstStartRow + (r - srcStartRow)
                Set dstCell = wsDst.Cells(dstRow, dstShapeCol)

                ' Set pasted = SafeCopyPasteShape(shp, wsDst)
                Set pasted = SafeCopyPasteShape(shp, wsDst, dstCell)

                If Not pasted Is Nothing Then
                    With pasted
                        .Left = dstCell.Left + 2
                        .Top = dstCell.Top + 2
                        .Width = shp.Width
                        .Height = shp.Height
                        .Placement = xlMoveAndSize
                        .Name = MakeUniqueShapeName(wsDst, shpName) 'avoid name collision
                    End With
                End If
            End If
        End If
    Next r
End Sub

'=========================
' Safe copy/paste with retry + CopyPicture fallback
'=========================
Private Function SafeCopyPasteShape(ByVal shp As Shape, ByVal wsDst As Worksheet, ByVal focusCell As Range) As Shape
    Dim attempt As Long
    Dim ok As Boolean

    wsDst.Activate

    For attempt = 1 To 3
        ok = False
        On Error Resume Next

        '1) Normal copy
        shp.Copy
        If Err.Number = 0 Then
            wsDst.Paste
            If Err.Number = 0 Then ok = True
        End If

        '2) Fallback: copy as picture
        If Not ok Then
            Err.Clear
            shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            If Err.Number = 0 Then
                wsDst.Paste
                If Err.Number = 0 Then ok = True
            End If
        End If

        On Error GoTo 0
        If ok Then
            Set SafeCopyPasteShape = wsDst.Shapes(wsDst.Shapes.Count)

            '? Deselect the pasted shape by selecting a cell
            Application.CutCopyMode = False
            focusCell.Select

            Exit Function
        End If

        DoEvents
    Next attempt

    Set SafeCopyPasteShape = Nothing
End Function


'=========================
' Ensure shape name is unique in destination sheet
'=========================
Private Function MakeUniqueShapeName(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim nm As String, n As Long
    nm = baseName
    n = 1

    Do While ShapeExists(ws, nm)
        n = n + 1
        nm = baseName & "_" & CStr(n)
    Loop

    MakeUniqueShapeName = nm
End Function




