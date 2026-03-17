Option Explicit

'=========================================================
' COPY FILTERED COLUMNS + VALUES ONLY + SHAPES
' FIXED:
' 1) old copied shapes are removed reliably
' 2) pasted shapes are resized to fit the target shape area
' 3) target area uses the merged WIDTH, but current ROW height
'    so shapes do not stack on top of each other
' 4) safer workbook detection to avoid Error 9
'=========================================================
Public Sub Copy_FilterColumns_ValuesOnly_WithShapes_Refined()

    Const SRC_SHEET As String = "BBS_for_checking"
    Const DST_SHEET As String = "BBS_for_yard"
    Const SRC_HEADER_ROW As Long = 6
    Const DST_HEADER_ROW As Long = 2
    Const SHAPE_PREFIX As String = "BBSCOPY_"

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim srcStartRow As Long, dstStartRow As Long
    Dim srcCodeCol As Long, dstCodeCol As Long, dstShapeCol As Long
    Dim lastRow As Long, dstLastRow As Long
    Dim oldDstLastRow As Long, clearLastRow As Long

    Dim map As Variant
    Dim i As Long
    Dim srcCol As Long, dstCol As Long

    Dim calcMode As XlCalculation
    Dim scrUpdate As Boolean, evt As Boolean

    On Error GoTo EH

    Set wb = GetWorkbookWithSheets(SRC_SHEET, DST_SHEET)
    If wb Is Nothing Then
        Err.Raise vbObjectError + 200, , _
            "Could not find both sheets '" & SRC_SHEET & "' and '" & DST_SHEET & _
            "' in ThisWorkbook or ActiveWorkbook."
    End If

    Set wsSrc = wb.Worksheets(SRC_SHEET)
    Set wsDst = wb.Worksheets(DST_SHEET)

    srcStartRow = SRC_HEADER_ROW + 1
    dstStartRow = DST_HEADER_ROW + 1

    '---- SourceHeader -> DestinationHeader
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

    '---- Find source Code column
    srcCodeCol = FindHeaderColSmart(wsSrc, SRC_HEADER_ROW, "Code")
    If srcCodeCol = 0 Then
        Err.Raise vbObjectError + 101, , "Source header 'Code' not found."
    End If

    '---- Source last row
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCodeCol).End(xlUp).Row
    If lastRow < srcStartRow Then
        lastRow = srcStartRow - 1
    End If

    '---- Destination last existing row (for cleanup if previous run had more rows)
    dstCodeCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, "Code")
    If dstCodeCol > 0 Then
        oldDstLastRow = wsDst.Cells(wsDst.Rows.Count, dstCodeCol).End(xlUp).Row
    Else
        oldDstLastRow = dstStartRow - 1
    End If
    If oldDstLastRow < dstStartRow Then oldDstLastRow = dstStartRow - 1

    '---- Current destination last row based on source
    If lastRow >= srcStartRow Then
        dstLastRow = dstStartRow + (lastRow - srcStartRow)
    Else
        dstLastRow = dstStartRow - 1
    End If

    If oldDstLastRow > dstLastRow Then
        clearLastRow = oldDstLastRow
    Else
        clearLastRow = dstLastRow
    End If

    '---- Shape column
    dstShapeCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, "Shape")

    '---- Delete old copied shapes in destination shape area
    If dstShapeCol > 0 And clearLastRow >= dstStartRow Then
        DeleteShapesInShapeArea wsDst, dstShapeCol, dstStartRow, clearLastRow, SHAPE_PREFIX
    End If

    '---- Clear destination mapped columns for current + old leftover rows
    If clearLastRow >= dstStartRow Then
        For i = LBound(map) To UBound(map)
            dstCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, CStr(map(i)(1)))
            If dstCol > 0 Then
                wsDst.Range(wsDst.Cells(dstStartRow, dstCol), wsDst.Cells(clearLastRow, dstCol)).ClearContents
            End If
        Next i
    End If

    '---- No source data
    If lastRow < srcStartRow Then GoTo CleanExit

    '---- Copy columns: VALUES only + formats + width
    For i = LBound(map) To UBound(map)

        srcCol = FindHeaderColSmart(wsSrc, SRC_HEADER_ROW, CStr(map(i)(0)))
        dstCol = FindHeaderColSmart(wsDst, DST_HEADER_ROW, CStr(map(i)(1)))

        If srcCol > 0 And dstCol > 0 Then

            'Width
            wsDst.Columns(dstCol).ColumnWidth = wsSrc.Columns(srcCol).ColumnWidth

            'Values only
            wsDst.Range(wsDst.Cells(dstStartRow, dstCol), wsDst.Cells(dstLastRow, dstCol)).Value2 = _
                wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Value2

            'Formats only
            wsSrc.Range(wsSrc.Cells(srcStartRow, srcCol), wsSrc.Cells(lastRow, srcCol)).Copy
            wsDst.Cells(dstStartRow, dstCol).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False

        End If
    Next i

    '---- Copy shapes by name: Code_Row (example: 83ST_31)
    If dstShapeCol > 0 Then
        CopyShapesByCodeRow_FitToCell wsSrc, wsDst, srcCodeCol, srcStartRow, lastRow, dstStartRow, dstShapeCol, SHAPE_PREFIX
    End If

CleanExit:
    ClearShapeSelectionSafe
    Application.ScreenUpdating = scrUpdate
    Application.EnableEvents = evt
    Application.Calculation = calcMode
    Exit Sub

EH:
    ClearShapeSelectionSafe
    Application.ScreenUpdating = scrUpdate
    Application.EnableEvents = evt
    Application.Calculation = calcMode
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbExclamation
End Sub


'=========================================================
' COPY SHAPES BY Code_Row AND FIT THEM INSIDE TARGET AREA
'=========================================================
Private Sub CopyShapesByCodeRow_FitToCell( _
    ByVal wsSrc As Worksheet, ByVal wsDst As Worksheet, _
    ByVal srcCodeCol As Long, _
    ByVal srcStartRow As Long, ByVal srcLastRow As Long, _
    ByVal dstStartRow As Long, ByVal dstShapeCol As Long, _
    ByVal shapePrefix As String)

    Dim r As Long, dstRow As Long
    Dim codeVal As String
    Dim shpKey As String
    Dim shp As Shape
    Dim pasted As Shape
    Dim dstArea As Range

    For r = srcStartRow To srcLastRow

        codeVal = Trim$(CStr(wsSrc.Cells(r, srcCodeCol).Value2))
        If Len(codeVal) = 0 Then GoTo NextR

        shpKey = codeVal & "_" & CStr(r)

        Set shp = FindShapeSmart(wsSrc, shpKey)
        If shp Is Nothing Then GoTo NextR

        dstRow = dstStartRow + (r - srcStartRow)

        'Important:
        'This uses merged WIDTH but current ROW height.
        'So even if the shape area is merged across columns,
        'the shape is still positioned row-by-row and won't stack.
        Set dstArea = GetShapeTargetRange(wsDst, dstRow, dstShapeCol)

        Set pasted = SafeCopyPasteShape(shp, wsDst)
        If Not pasted Is Nothing Then
            pasted.Placement = xlMoveAndSize
            pasted.Name = MakeUniqueShapeName(wsDst, shapePrefix & shpKey)
            FitShapeToCell pasted, dstArea, 2, True
        End If

NextR:
        Set shp = Nothing
        Set pasted = Nothing
        Set dstArea = Nothing
    Next r
End Sub


'=========================================================
' SAFE COPY/PASTE SHAPE
'=========================================================
Private Function SafeCopyPasteShape(ByVal shp As Shape, ByVal wsDst As Worksheet) As Shape
    Dim attempt As Long
    Dim beforeCount As Long, afterCount As Long
    Dim ok As Boolean

    For attempt = 1 To 3

        ok = False
        beforeCount = wsDst.Shapes.Count

        On Error Resume Next
        Err.Clear

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

        afterCount = wsDst.Shapes.Count
        If ok And afterCount > beforeCount Then
            Set SafeCopyPasteShape = wsDst.Shapes(afterCount)
            Application.CutCopyMode = False
            ClearShapeSelectionSafe
            Exit Function
        End If

        Application.CutCopyMode = False
        DoEvents
    Next attempt

    Set SafeCopyPasteShape = Nothing
End Function


'=========================================================
' FIT SHAPE INSIDE TARGET RANGE
'=========================================================
Private Sub FitShapeToCell(ByVal shp As Shape, ByVal tgt As Range, _
                           Optional ByVal padding As Double = 2, _
                           Optional ByVal keepAspect As Boolean = True)

    Dim maxW As Double, maxH As Double
    Dim rW As Double, rH As Double, sc As Double
    Dim w0 As Double, h0 As Double

    maxW = tgt.Width - (2 * padding)
    maxH = tgt.Height - (2 * padding)

    If maxW < 2 Then maxW = 2
    If maxH < 2 Then maxH = 2

    w0 = shp.Width
    h0 = shp.Height

    If w0 <= 0 Then w0 = 1
    If h0 <= 0 Then h0 = 1

    On Error Resume Next
    shp.LockAspectRatio = IIf(keepAspect, msoTrue, msoFalse)
    On Error GoTo 0

    If keepAspect Then
        rW = maxW / w0
        rH = maxH / h0

        If rW < rH Then
            sc = rW
        Else
            sc = rH
        End If

        If sc <= 0 Then sc = 1

        shp.Width = w0 * sc
        shp.Height = h0 * sc
    Else
        shp.Width = maxW
        shp.Height = maxH
    End If

    shp.Left = tgt.Left + (tgt.Width - shp.Width) / 2
    shp.Top = tgt.Top + (tgt.Height - shp.Height) / 2
End Sub


'=========================================================
' DELETE OLD COPIED SHAPES IN DESTINATION SHAPE AREA
' - deletes shapes with our prefix
' - also deletes shapes physically overlapping the target area
'   so old bad copies from previous versions get removed too
'=========================================================
Private Sub DeleteShapesInShapeArea(ByVal ws As Worksheet, ByVal shapeCol As Long, _
                                    ByVal startRow As Long, ByVal endRow As Long, _
                                    ByVal namePrefix As String)

    Dim i As Long
    Dim areaFirst As Range, areaLast As Range
    Dim areaLeft As Double, areaTop As Double, areaRight As Double, areaBottom As Double
    Dim shpLeft As Double, shpTop As Double, shpRight As Double, shpBottom As Double
    Dim deleteIt As Boolean
    Dim nm As String

    If endRow < startRow Then Exit Sub

    Set areaFirst = GetShapeTargetRange(ws, startRow, shapeCol)
    Set areaLast = GetShapeTargetRange(ws, endRow, shapeCol)

    areaLeft = areaFirst.Left
    areaTop = areaFirst.Top
    areaRight = areaFirst.Left + areaFirst.Width
    areaBottom = areaLast.Top + areaLast.Height

    For i = ws.Shapes.Count To 1 Step -1
        deleteIt = False
        nm = ws.Shapes(i).Name

        If Len(namePrefix) > 0 Then
            If LCase$(Left$(nm, Len(namePrefix))) = LCase$(namePrefix) Then
                deleteIt = True
            End If
        End If

        If Not deleteIt Then
            shpLeft = ws.Shapes(i).Left
            shpTop = ws.Shapes(i).Top
            shpRight = shpLeft + ws.Shapes(i).Width
            shpBottom = shpTop + ws.Shapes(i).Height

            If Not (shpRight <= areaLeft Or shpLeft >= areaRight Or shpBottom <= areaTop Or shpTop >= areaBottom) Then
                deleteIt = True
            End If
        End If

        If deleteIt Then ws.Shapes(i).Delete
    Next i
End Sub


'=========================================================
' GET TARGET SHAPE RANGE FOR ONE ROW
' Uses:
' - merged WIDTH from shape area
' - current row HEIGHT only
' This prevents shapes from stacking in one merged block
'=========================================================
Private Function GetShapeTargetRange(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal shapeCol As Long) As Range
    Dim ma As Range
    Dim firstCol As Long, lastCol As Long

    Set ma = ws.Cells(rowNum, shapeCol).MergeArea
    firstCol = ma.Column
    lastCol = ma.Column + ma.Columns.Count - 1

    Set GetShapeTargetRange = ws.Range(ws.Cells(rowNum, firstCol), ws.Cells(rowNum, lastCol))
End Function


'=========================================================
' FIND SHAPE:
' 1) exact name
' 2) name starts with key
'=========================================================
Private Function FindShapeSmart(ByVal ws As Worksheet, ByVal shpKey As String) As Shape
    Dim shp As Shape

    On Error Resume Next
    Set FindShapeSmart = ws.Shapes(shpKey)
    On Error GoTo 0
    If Not FindShapeSmart Is Nothing Then Exit Function

    For Each shp In ws.Shapes
        If LCase$(Left$(shp.Name, Len(shpKey))) = LCase$(shpKey) Then
            Set FindShapeSmart = shp
            Exit Function
        End If
    Next shp

    Set FindShapeSmart = Nothing
End Function


'=========================================================
' HEADER FIND
'=========================================================
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
    s = Replace(s, ChrW(8470), "No")
    s = Replace(s, "ø", "o")
    s = Replace(s, "Ø", "o")
    s = Replace(s, "º", "")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeHeader = LCase$(s)
End Function


'=========================================================
' SHAPE HELPERS
'=========================================================
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
    Application.CommandBars.ExecuteMso "Escape"
    On Error GoTo 0
End Sub


'=========================================================
' WORKBOOK / SHEET HELPERS
'=========================================================
Private Function WorksheetExists(ByVal wb As Workbook, ByVal wsName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Function GetWorkbookWithSheets(ByVal srcSheetName As String, ByVal dstSheetName As String) As Workbook
    Dim wbActive As Workbook

    If WorksheetExists(ThisWorkbook, srcSheetName) And WorksheetExists(ThisWorkbook, dstSheetName) Then
        Set GetWorkbookWithSheets = ThisWorkbook
        Exit Function
    End If

    Set wbActive = ActiveWorkbook
    If Not wbActive Is Nothing Then
        If WorksheetExists(wbActive, srcSheetName) And WorksheetExists(wbActive, dstSheetName) Then
            Set GetWorkbookWithSheets = wbActive
            Exit Function
        End If
    End If

    Set GetWorkbookWithSheets = Nothing
End Function

