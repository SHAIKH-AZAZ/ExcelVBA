Sub ListAllShapes()

    Dim shp As Shape 
    Dim ws As Worksheet
    Dim i As Long 

    Dim templateName As String 
    templateName = "ShapeLibrary"

    i = 1
    ws.Columns("A").ClearContents
    For Each shp In ws.Shapes
        ws.Cells(i, 1).Value = shp.Name   'Write name in column A
        i = i + 1
    Next shp
    MsgBox "Shape names extracted successfully!", vbInformation


End Sub