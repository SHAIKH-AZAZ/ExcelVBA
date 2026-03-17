Option Explicit

'1) Delete ALL shapes on ActiveSheet (includes pictures, charts, SmartArt, lines, textboxes, etc.)
Sub Delete_All_Shapes_ActiveSheet()
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ActiveSheet
    Application.ScreenUpdating = False
    
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    Application.ScreenUpdating = True
End Sub


'2) Delete ONLY pictures/images on ActiveSheet (keeps other shapes like lines/boxes/textboxes)
Sub Delete_Only_Pictures_ActiveSheet()
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ActiveSheet
    Application.ScreenUpdating = False
    
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Type = msoPicture Or ws.Shapes(i).Type = msoLinkedPicture Then
            ws.Shapes(i).Delete
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub


'3) Prompt-based: choose what to delete on ActiveSheet
Sub Delete_Shapes_Prompt_ActiveSheet()
    Dim ans As VbMsgBoxResult
    
    ans = MsgBox("YES = Delete ALL shapes (including pictures, charts, textboxes, lines)" & vbCrLf & _
                 "NO  = Delete ONLY pictures/images" & vbCrLf & _
                 "CANCEL = Do nothing", _
                 vbYesNoCancel + vbQuestion, "Delete Shapes?")
    
    If ans = vbYes Then
        Delete_All_Shapes_ActiveSheet
    ElseIf ans = vbNo Then
        Delete_Only_Pictures_ActiveSheet
    End If
End Sub



