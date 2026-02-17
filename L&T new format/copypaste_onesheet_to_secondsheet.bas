Option Explicit

'=========================
' MAIN MACRO
'=========================
Public Sub Copy_FilterColumns_ValuesOnly_WithShapes_Refined()

    Const SRC_SHEET As String = "base_format"
    Const DST_SHEET As String = "Filter_format"
    Const SRC_HEADER_ROW As Long = 2
    Const DST_HEADER_ROW As Long = 2

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim srcStartRow As Long, dstStartRow As Long
    Dim srcCodeCol As Long, dstShapeCol As Long
    Dim lastRow As Long, dstLastRow As Long

    Dim map As Variant
    Dim i As Long
    Dim srcCol As Long, dstCol As Long

    Dim calcMode As XlCalculation
    Dim scrUpdate As Boolean, evt As Boolean

    On Error GoTo EH

    Set wsSrc = ThisWorkbook.Worksheets(SRC_SHEET)
    Set wsDst = ThisWorkbook.Worksheets(DST_SHEET)

    srcStartRow = SRC_HEADER_ROW + 1
    dstStartRow = DST_HEADER_ROW + 1

    '---- SourceHeader -> DestinationHeader (edit if needed)
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

    '---- Speed settings
    scrUpdate = Application.ScreenUpdating
    evt = Application.EnableEvents
    calcMode = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '---- Find last row using Code column
    srcCodeCol = FindHeaderColSmart(wsSrc, SRC_HEADER_ROW, "Code")
    If srcCodeCol = 0 Then Err.Raise vbObjectError + 101, , "Source header 'Code' not found."

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCodeCol).End(xlUp).Row
    If lastRow < srcStartRow Then GoTo CleanExit

    dstLastRow = dstStartRow + (lastRow - srcStartRow)

    '---- Find destination Shape column and clear old shapes only in that area
    dstShapeCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, "Shape")
    If dstShapeCol > 0 Then
        DeleteShapesInColumnRange wsDst, dstShapeCol, dstStartRow, dstLastRow
    End If

    '---- Clear destination mapped columns ONLY for target rows (not whole sheet)
    For i = LBound(map) To UBound(map)
        dstCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, CStr(map(i)(1)))
        If dstCol > 0 Then
            wsDst.Range(wsDst.Cells(dstStartRow, dstCol), wsDst.Cells(dstLastRow, dstCol)).ClearContents
        End If
    Next i

    '---- Copy columns: VALUES only + formats + width
    For i = LBound(map) To UBound(map)

        srcCol = FindHeaderColSmart(wsSrc, SRC_HEADER_ROW, CStr(map(i)(0)))
        dstCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, CStr(map(i)(1)))

        If srcCol > 0 And dstCol > 0 Then
            'Width
            wsDst.Columns(dstCol).ColumnWidth = wsSrc.Columns(srcCol).ColumnWidth

            'Values (prevents #REF!)
            wsDst.Range(wsDst.Cells(dstStartRow, dstCol), wsDst.Cells(dstLastRow, dstCol)).Value = _
                wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Value

            'Number formats
            wsDst.Range(wsDst.Cells(dstStartRow, dstCol), wsDst.Cells(dstLastRow, dstCol)).NumberFormat = _
                wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).NumberFormat

            'Cell formats (borders/fill/etc.)
            wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Copy
            wsDst.Cells(dstStartRow, dstCol).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        End If

    Next i

    '---- Copy shapes by name: Code_Row (e.g. 12_5)
    dstShapeCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, "Shape")
    If dstShapeCol > 0 Then
        CopyShapesByCodeRow_FitToCell wsSrc, wsDst, srcCodeCol, srcStartRow, lastRow, dstStartRow, dstShapeCol
    End If

    'Finish: select a cell so no shape stays selected
    ' wsDst.Range("A1").Select
    ClearShapeSelectionSafe
    
CleanExit:
    Application.ScreenUpdating = scrUpdate
    Application.EnableEvents = evt
    Application.Calculation = calcMode
    Exit Sub

EH:
    'Restore settings then show error
    Application.ScreenUpdating = scrUpdate
    Application.EnableEvents = evt
    Application.Calculation = calcMode
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbExclamation
End Sub


'=========================
' SHAPES: Copy by name (Code_Row) and fit into Shape cell
'=========================
Private Sub CopyShapesByCodeRow_FitToCell( _
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
        If Len(Trim$(CStr(codeVal))) = 0 Then GoTo NextR

        shpName = Trim$(CStr(codeVal)) & "_" & CStr(r)

        If Not ShapeExists(wsSrc, shpName) Then GoTo NextR

        Set shp = wsSrc.Shapes(shpName)
        dstRow = dstStartRow + (r - srcStartRow)
        Set dstCell = wsDst.Cells(dstRow, dstShapeCol)

        Set pasted = SafeCopyPasteShape(shp, wsDst)

        If Not pasted Is Nothing Then
            'Move & size with cells
            pasted.Placement = xlMoveAndSize

            'Avoid name collision
            pasted.Name = MakeUniqueShapeName(wsDst, shpName)

            'Fit inside the target cell (padding=2, keepAspect=True)
            FitShapeToCell pasted, dstCell, 2, True
        End If

NextR:
    Next r
End Sub


'=========================
' Safe copy/paste shape (retry + CopyPicture fallback)
' returns the newly pasted shape
'=========================
Private Function SafeCopyPasteShape(ByVal shp As Shape, ByVal wsDst As Worksheet) As Shape
    Dim attempt As Long
    Dim ok As Boolean

    For attempt = 1 To 3
        ok = False
        On Error Resume Next

        shp.Copy
        If Err.Number = 0 Then
            wsDst.Paste
            If Err.Number = 0 Then ok = True
        End If

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
            ClearShapeSelectionSafe   '? deselect shape safely
            Exit Function
        End If

        DoEvents
    Next attempt

    Set SafeCopyPasteShape = Nothing
End Function



'=========================
' Fit a shape inside a cell (resize + center)
'=========================
Private Sub FitShapeToCell(ByVal shp As Shape, ByVal tgt As Range, Optional ByVal padding As Double = 2, Optional ByVal keepAspect As Boolean = True)

    Dim maxW As Double
    Dim maxH As Double
    Dim rW As Double
    Dim rH As Double
    Dim sc As Double

    maxW = tgt.Width - (2 * padding)
    maxH = tgt.Height - (2 * padding)
    If maxW < 2 Then maxW = 2
    If maxH < 2 Then maxH = 2

    On Error Resume Next
    shp.LockAspectRatio = IIf(keepAspect, msoTrue, msoFalse)
    On Error GoTo 0

    'Start at cell top-left
    shp.Left = tgt.Left + padding
    shp.Top = tgt.Top + padding

    If keepAspect Then
        rW = maxW / shp.Width
        rH = maxH / shp.Height
        sc = IIf(rW < rH, rW, rH)

        shp.Width = shp.Width * sc
        shp.Height = shp.Height * sc

        'Center in cell
        shp.Left = tgt.Left + (tgt.Width - shp.Width) / 2
        shp.Top = tgt.Top + (tgt.Height - shp.Height) / 2
    Else
        shp.Width = maxW
        shp.Height = maxH
    End If
End Sub


'=========================
' Delete shapes in a specific column between rows
'=========================
Private Sub DeleteShapesInColumnRange(ByVal ws As Worksheet, ByVal targetCol As Long, ByVal startRow As Long, ByVal endRow As Long)
    Dim i As Long
    Dim tl As Range

    For i = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        Set tl = ws.Shapes(i).TopLeftCell
        On Error GoTo 0

        If Not tl Is Nothing Then
            If tl.Column = targetCol Then
                If tl.Row >= startRow And tl.Row <= endRow Then
                    ws.Shapes(i).Delete
                End If
            End If
        End If

        Set tl = Nothing
    Next i
End Sub


'=========================
' Header find with normalization (handles extra spaces + special chars)
'=========================
Private Function FindHeaderColSmart(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    Dim cellTxt As String

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        cellTxt = NormalizeHeader(CStr(ws.Cells(headerRow, c).Value2))
        If cellTxt = NormalizeHeader(headerText) Then
            FindHeaderColSmart = c
            Exit Function
        End If
    Next c

    FindHeaderColSmart = 0
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, ChrW(8470), "No") ' ? -> No
    s = Replace(s, "ø", "o")         ' ø -> o
    s = Replace(s, "Ø", "o")
    s = Replace(s, "º", "")          ' º remove
    s = Replace(s, "  ", " ")
    NormalizeHeader = LCase$(s)
End Function


'=========================
' Shape helpers
'=========================
Private Function ShapeExists(ByVal ws As Worksheet, ByVal shpName As String) As Boolean
    On Error Resume Next
    ShapeExists = Not ws.Shapes(shpName) Is Nothing
    On Error GoTo 0
End Function

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


Private Sub ClearShapeSelectionSafe()
    On Error Resume Next
    'This clears selection without selecting any cell
    Application.CommandBars.ExecuteMso "Escape"
    On Error GoTo 0
End Sub
