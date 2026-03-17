Attribute VB_Name = "Module1"
Private Sub WorksheetAfterPrint()
On Error Resume Next
With ActiveSheet.Range("B1:R5")
Set Code = .Find("Code")
End With
If Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = "HideColumn" Then
Range(Cells(, Code.Column), Cells(, Code.Column)).Columns("A:G").EntireColumn.Hidden = False
ElseIf Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = "UnhideColumn" Then
Range(Cells(, Code.Column + 7), Cells(, Code.Column + 7)).Columns("A:B").EntireColumn.Hidden = True
End If
Dim PauseTime, Start
PauseTime = 1.5
Start = Timer
Do While Timer < Start + PauseTime
DoEvents
Loop
On Error Resume Next
ActiveSheet.OLEObjects("PrintOptionInformer").Visible = False
End Sub
Sub Options()
Attribute Options.VB_Description = "t"
Attribute Options.VB_ProcData.VB_Invoke_Func = "t\n14"
If Application.EnableEvents = False Then
Application.EnableEvents = True
MsgBox "The event actions were disabled due to unusual operations." & Chr(10) & "Now the event actions are restored.", , "BBS Program"
Exit Sub
End If
If ActiveWorkbook.CodeName <> "BBSFormat" Then
ThisWorkbook.Activate
End If
On Error GoTo ThisWorkbookOptionForm
Application.Run "'" & ActiveWorkbook.Name & "'!OptionsFormShow"
Exit Sub
ThisWorkbookOptionForm:
MsgBox "Close the other BBS files and try again to set the options.", vbInformation, "BBS Program"
End Sub
Private Sub OptionsFormShow()
On Error Resume Next
OptionsForm.Show
If Err.Number = 75 Then
Dim VersionError As String
VersionError = IIf(Application.Version = 12, "Please try it in Excel-2003 or Excel-2010.", "Reinstall or repair the Excel")
MsgBox "Unable to open the form in this version of Excel." & Chr(10) & VersionError, , "BBS Program"
Exit Sub
End If
On Error GoTo 0
End Sub
Private Sub SortBBSFormat()
If Application.EnableEvents = False Then
Application.EnableEvents = True
MsgBox "The event actions were disabled due to unusual operations." & Chr(10) & "Now the event actions are restored.", , "BBS Program"
Exit Sub
End If
If ActiveWorkbook.CodeName <> "BBSFormat" Then
ThisWorkbook.Activate
End If
Dim wb As Workbook
Dim wbname As String
AlreadyOpen = False
For Each wb In Workbooks
If wb.CodeName = "BBSMacroFile" Then
AlreadyOpen = True
wbname = wb.Name
Exit For
End If
Next wb
If AlreadyOpen = True Then
Sheet0.OLEObjects("BBSProgram").Object.Caption = wbname
End If
If AlreadyOpen = False Then
MsgBox "Open BBS Program File", , "BBS Program"
Exit Sub
End If
On Error Resume Next
Application.ScreenUpdating = False
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Visible = xlSheetVisible
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Activate
Application.ScreenUpdating = True
If Right(ActiveSheet.Name, 7) = "_Sorted" Then
MsgBox "Sorted shteet already created", , "BBS Program"
Exit Sub
End If
On Error GoTo ThisWorkbookOptionForm
Application.Run "'" & Sheet0.OLEObjects("BBSProgram").Object.Caption & "'!OpenSortForm"
Exit Sub
ThisWorkbookOptionForm:
MsgBox "Close the other BBS files and try again to set the options.", vbInformation, "BBS Program"
End Sub
Private Sub SortingFormShow()
On Error Resume Next
SortingForm.Show
If Err.Number = 75 Then
Dim VersionError As String
VersionError = IIf(Application.Version = 12, "Please try it in Excel-2003 or Excel-2010.", "Reinstall or repair the Excel")
MsgBox "Unable to open the form in this version of Excel." & Chr(10) & VersionError, , "BBS Program"
Exit Sub
End If
On Error GoTo 0
End Sub
