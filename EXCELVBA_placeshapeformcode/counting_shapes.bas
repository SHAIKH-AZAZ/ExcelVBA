Option Explicit

'1) Simple: total shapes count on ActiveSheet
Sub Count_Shapes_ActiveSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    MsgBox "Total shapes on active sheet (" & ws.Name & "): " & ws.Shapes.Count, vbInformation, "Shape Count"
End Sub

'2) Detailed: total + pictures + non-pictures
Sub Count_Shapes_Detailed_ActiveSheet()
    Dim ws As Worksheet
    Dim i As Long
    Dim total As Long, pics As Long, nonPics As Long
    
    Set ws = ActiveSheet
    total = ws.Shapes.Count
    
    For i = 1 To total
        If ws.Shapes(i).Type = msoPicture Or ws.Shapes(i).Type = msoLinkedPicture Then
            pics = pics + 1
        Else
            nonPics = nonPics + 1
        End If
    Next i
    
    MsgBox "Active sheet: " & ws.Name & vbCrLf & _
           "Total shapes: " & total & vbCrLf & _
           "Pictures: " & pics & vbCrLf & _
           "Other shapes: " & nonPics, _
           vbInformation, "Detailed Shape Count"
End Sub

'3) Count ONLY shapes currently selected (optional)
Sub Count_Selected_Shapes()
    Dim cnt As Long
    
    On Error Resume Next
    cnt = Selection.ShapeRange.Count
    On Error GoTo 0
    
    If cnt = 0 Then
        MsgBox "No shapes selected.", vbExclamation, "Selected Shape Count"
    Else
        MsgBox "Selected shapes: " & cnt, vbInformation, "Selected Shape Count"
    End If
End Sub



