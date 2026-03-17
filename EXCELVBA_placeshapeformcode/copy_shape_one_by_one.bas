Option Explicit

Sub Copy_Only_Pictures_To_NewSheet_OneByOne()
    Dim src As Worksheet, dst As Worksheet
    Dim shp As Shape, newShp As Shape
    Dim i As Long, copied As Long, failed As Long
    Dim log As String

    Set src = ActiveSheet
    Set dst = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    dst.Name = MakeUniqueSheetName("Pics_Copy_" & src.Name)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = src.Shapes.Count To 1 Step -1
        Set shp = src.Shapes(i)

        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            Application.StatusBar = "Copying picture " & (copied + 1) & " : " & shp.Name

            On Error Resume Next
            shp.Copy
            dst.Pictures.Paste   '<< paste into Pictures collection (more reliable for images)
            If Err.Number <> 0 Then
                failed = failed + 1
                log = log & vbCrLf & shp.Name & " | Err=" & Err.Number
                Err.Clear
                On Error GoTo 0
                GoTo NextOne
            End If
            On Error GoTo 0

            Set newShp = dst.Shapes(dst.Shapes.Count)
            newShp.Left = shp.Left
            newShp.Top = shp.Top
            newShp.Width = shp.Width
            newShp.Height = shp.Height

            copied = copied + 1
        End If

NextOne:
        Set newShp = Nothing
    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Done!" & vbCrLf & "Pictures copied: " & copied & vbCrLf & _
           "Failed: " & failed & IIf(failed > 0, vbCrLf & "Failures:" & log, ""), vbInformation
End Sub

Private Function MakeUniqueSheetName(ByVal baseName As String) As String
    Dim nm As String, n As Long
    nm = Left$(baseName, 31)
    If Not SheetExists(nm) Then MakeUniqueSheetName = nm: Exit Function
    n = 1
    Do
        nm = Left$(baseName, 28) & "_" & Format$(n, "00")
        n = n + 1
    Loop While SheetExists(nm)
    MakeUniqueSheetName = nm
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

