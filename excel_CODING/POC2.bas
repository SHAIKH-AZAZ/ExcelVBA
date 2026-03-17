Sub PlaceShapeFromCode()

    Dim ws As Worksheet
    Dim lib As Worksheet
    Dim shpCode As String
    Dim srcShape As Shape
    Dim newShp As Shape
    Dim r As Long
    
    Set ws = ActiveSheet              ' where D,E,F,H values exist
    Set lib = Sheets("ShapeLibrary")  ' sheet that contains template shapes
    
    ' Get active row of selection (ex: if you click L5 ? row=5)
    r = ActiveCell.Row
    
    ' 1) Read shape code from column L
    shpCode = ws.Cells(r, "L").Value
    
    If shpCode = "" Then
        MsgBox "Shape code missing in column L" & r, vbExclamation
        Exit Sub
    End If
    
    ' 2) Try to get matching shape from library
    On Error Resume Next
    Set srcShape = lib.Shapes(CStr(shpCode))
    On Error GoTo 0
    
    If srcShape Is Nothing Then
        MsgBox "Shape code " & shpCode & " not found in ShapeLibrary!", vbCritical
        Exit Sub
    End If
    
    ' 3) Copy the shape
    srcShape.Copy
    
    ' 4) Paste on active sheet
    ws.Paste
    Set newShp = ws.Shapes(ws.Shapes.Count)
    
    ' 5) Position at H(row)
    Dim targetCell As Range
    Set targetCell = ws.Cells(r, "W")
    
' ----- Center shape inside target cell W(row) -----
newShp.Left = targetCell.Left + (targetCell.Width - newShp.Width) / 2
newShp.Top = targetCell.Top + (targetCell.Height - newShp.Height) / 2
' --------------------------------------------------

    
    ' 6) Read variable values
    Dim Avalue As String, Bvalue As String, Cvalue As String
    Avalue = ws.Cells(r, "M").Value
    Bvalue = ws.Cells(r, "N").Value
    Cvalue = ws.Cells(r, "O").Value
    Dvalue = ws.Cells(r, "P").Value
    Evalue = ws.Cells(r, "Q").Value
    
    ' 7) Replace text inside copied shape
    Dim inner As Shape
    Dim txt As String
    
    If newShp.Type = msoGroup Then
        
        For Each inner In newShp.GroupItems
            If inner.TextFrame2.HasText Then
                txt = inner.TextFrame2.TextRange.Text
            txt = Replace(txt, "{A}", Avalue)
            txt = Replace(txt, "{B}", Bvalue)
            txt = Replace(txt, "{C}", Cvalue)
            txt = Replace(txt, "{D}", Dvalue)
            txt = Replace(txt, "{E}", Evalue)
                inner.TextFrame2.TextRange.Text = txt
            End If
        Next inner
    
    Else
        ' single shape
        If newShp.TextFrame2.HasText Then
            txt = newShp.TextFrame2.TextRange.Text
            txt = Replace(txt, "{A}", Avalue)
            txt = Replace(txt, "{B}", Bvalue)
            txt = Replace(txt, "{C}", Cvalue)
            txt = Replace(txt, "{D}", Dvalue)
            txt = Replace(txt, "{E}", Evalue)
            newShp.TextFrame2.TextRange.Text = txt
        End If
    End If
    
    ''  MsgBox "Shape placed at W" & r & " successfully!", vbInformation

End Sub




