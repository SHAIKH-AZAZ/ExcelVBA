Sub CopyAndUpdateShape()

    Dim original As Shape
    Dim newShp As Shape
    Dim base As Range
    Dim txt As String
    
    ' Name of the template to copy
    Dim templateName As String
    templateName = "ShapeLibrary"   ' <<< your template name here
    
    ' Get the original shape/group
    Set original = ActiveSheet.Shapes("L shape ")
     
    If original Is Nothing Then
        MsgBox "Template shape/group not found!", vbCritical
        Exit Sub
    End If
    
    ' ---- 1. COPY THE SHAPE/GROUP ----
    original.Copy
    
    ' ---- 2. PASTE THE COPY ----
    ActiveSheet.Paste
    Set newShp = ActiveSheet.Shapes(ActiveSheet.Shapes.Count)  ' the newly created copy
    
    ' ---- 3. MOVE THE NEW COPY (example: offset down 5 rows) ----
    newShp.Top = newShp.Top + 50   ' adjust as needed
    newShp.Left = newShp.Left + 100  ' (optional)
    
    ' ---- 4. UPDATE TEXT INSIDE THE COPY ----
    ' Use its NEW location for relative referencing
    Set base = newShp.TopLeftCell
    
    Dim Avalue As Variant
    Avalue = "sample"    ' or base.Offset(0,1).Value etc.
    
    ' If the object is a GROUP:
    If newShp.Type = msoGroup Then
        
        Dim inner As Shape
        
        For Each inner In newShp.GroupItems
            If inner.TextFrame2.HasText Then
                txt = inner.TextFrame2.TextRange.Text
                txt = Replace(txt, "{A}", Avalue)
                inner.TextFrame2.TextRange.Text = txt
            End If
        Next inner
        
    Else
        ' If it is a single shape
        txt = newShp.TextFrame2.TextRange.Text
        txt = Replace(txt, "{A}", Avalue)
        newShp.TextFrame2.TextRange.Text = txt
    End If


End Sub

