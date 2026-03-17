Option Explicit

Sub PlaceShapeFromCode()
    ' -- Object Variables
    Dim ws As Worksheet
    Dim lib As Worksheet
    Dim srcShape As Shape
    Dim newShp As Shape
    Dim targetCell As Range
    Dim originalCell As Range ' <--- 1. Variable is already here
    
    ' -- Data Variables
    Dim shpCode As String
    Dim r As Long
    Dim finalName As String
    
    ' -- Text Replacement Variables
    Dim keys() As Variant
    Dim values() As Variant
    Dim i As Integer
    
    ' ---------------- SILENT MODE ----------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo CleanExit
    ' ---------------------------------------------

    Set ws = ActiveSheet
    
    ' 1. Validate Environment
    ' Check if Library exists
    On Error Resume Next
    Set lib = ThisWorkbook.Sheets("ShapeLibrary")
    On Error GoTo CleanExit
    
    If lib Is Nothing Then
        MsgBox "Critical Error: The sheet 'ShapeLibrary' is missing.", vbCritical
        GoTo CleanExit
    End If
    
    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    
    ' --- NEW CODE START ---
    Set originalCell = ActiveCell ' <--- 2. Capture the starting location here
    ' --- NEW CODE END ---
    
    r = ActiveCell.Row
    
    ' 2. Retrieve Shape Code (Column G)
    shpCode = Trim(ws.Cells(r, "G").Value)
    If shpCode = "" Then GoTo CleanExit
    
    ' 3. Fetch Source Shape
    On Error Resume Next
    Set srcShape = lib.Shapes(shpCode)
    On Error GoTo CleanExit
    
    If srcShape Is Nothing Then
        MsgBox "Shape Code '" & shpCode & "' not found in library.", vbExclamation
        GoTo CleanExit
    End If
    
    ' 4. Clean Pre-existing Shape
    finalName = shpCode & "_" & r
    On Error Resume Next
    ws.Shapes(finalName).Delete
    On Error GoTo CleanExit
    
    ' 5. Copy & Paste
    srcShape.Copy
    ws.Paste Destination:=ws.Cells(r, "S")
    
    ' Capture the new shape
    If TypeName(Selection) = "Drawing" Or TypeName(Selection) = "Group" Or TypeName(Selection) = "Picture" Then
        Set newShp = Selection.ShapeRange.Item(1)
    Else
        Set newShp = ws.Shapes(ws.Shapes.Count)
    End If
    
    ' 6. Assign Unique Metadata (Name)
    newShp.Name = finalName
    
    ' 7. Positioning
    Set targetCell = ws.Cells(r, "S")
    With newShp
      .Placement = xlMove
      .Left = targetCell.Left + (targetCell.Width - .Width) / 2
      .Top = targetCell.Top + (targetCell.Height - .Height) / 2
    End With
    
    ' 8. Text Replacement (Recursive)
    ReDim keys(1 To 7)
    ReDim values(1 To 7)
    
    For i = 1 To 7
        keys(i) = "{" & Chr(64 + i) & "}"
        values(i) = CStr(ws.Cells(r, 8 + i).Value)
    Next i
    
    ProcessTextRecursively newShp, keys, values
    
    ' 9. Reset Selection
    ' --- CHANGED CODE START ---
    originalCell.Select   ' <--- 3. Go back to the original cell
    ' --- CHANGED CODE END ---
    
CleanExit:
    ' Restore Excel State
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub





Sub DeleteShapeFromRow()
    Dim ws As Worksheet
    Dim shpCode As String
    Dim r As Long
    Dim targetName As String
    
    ' ---------------- SILENT MODE ----------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo CleanExit
    ' ---------------------------------------------
    
    Set ws = ActiveSheet
    
    ' Ensure valid selection
    If TypeName(Selection) <> "Range" Then GoTo CleanExit
    
    r = ActiveCell.Row
    
    ' 1. Get Shape Code from Column L
    shpCode = Trim(ws.Cells(r, "L").Value)
    If shpCode = "" Then GoTo CleanExit
    
    ' 2. Construct Unique Identifier
    ' Logic: ShapeCode + "_" + RowNumber
    targetName = shpCode & "_" & r
    
    ' 3. Find and Delete
    ' We use On Error Resume Next because if the shape is already gone,
    ' we don't want the macro to crash. We just want to exit silently.
    On Error Resume Next
    ws.Shapes(targetName).Delete
    On Error GoTo CleanExit
    
CleanExit:
    ' -------- RESTORE EXCEL STATE --------
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub







Private Sub ProcessTextRecursively(shp As Shape, keys() As Variant, values() As Variant)
    Dim subShp As Shape
    Dim txt As String
    Dim i As Integer
    
    ' Case A: The shape is a Group
    If shp.Type = msoGroup Then
        For Each subShp In shp.GroupItems
            ProcessTextRecursively subShp, keys, values ' Recursive Call
        Next subShp
        
    ' Case B: The shape is a single object
    Else
        ' Check for TextFrame
        On Error Resume Next
        If shp.TextFrame2.HasText Then
            txt = shp.TextFrame2.TextRange.Text
            ' Perform Batch Replacement
            For i = LBound(keys) To UBound(keys)
                txt = Replace(txt, keys(i), values(i))
            Next i
            shp.TextFrame2.TextRange.Text = txt
        End If
        On Error GoTo 0
    End If
End Sub
