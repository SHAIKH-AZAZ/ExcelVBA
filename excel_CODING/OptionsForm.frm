VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionsForm 
   Caption         =   "BBS Program_ FORMAT OPTIONS_Dimensions in Millimeter"
   ClientHeight    =   7050
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   12710
   OleObjectBlob   =   "OptionsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "OptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckNewVersions_Click()
BBSProgramDownload = "http://www.bendingschedule.com/download.html"
On Error GoTo CannotOpen
ActiveWorkbook.FollowHyperlink Address:=BBSProgramDownload, NewWindow:=True
Unload Me
Exit Sub
CannotOpen:
MsgBox "Internet connection problem" & Chr(10) & "Cannot open " & BBSProgramDownload, vbInformation, "BBS Program"
End Sub
Private Sub Consolidate_Click()
Unload Me
On Error Resume Next
ConsolidateForm.Show
If Err.Number = 75 Then
Dim VersionError As String
VersionError = IIf(Application.Version = 12, "Please try it in Excel-2003 or Excel-2010.", "Reinstall or repair the Excel")
MsgBox "Unable to open the form in this version of Excel." & Chr(10) & VersionError, , "BBS Program"
Exit Sub
End If
On Error GoTo 0
End Sub
Private Sub Link_Registration_Details_Click()
On Error Resume Next
Registration_Details.Show
If Err.Number = 75 Then
Dim VersionError As String
VersionError = IIf(Application.Version = 12, "Please try it in Excel-2003 or Excel-2010.", "Reinstall or repair the Excel")
MsgBox "Unable to open the form in this version of Excel." & Chr(10) & VersionError, , "BBS Program"
Exit Sub
End If
On Error GoTo 0
End Sub
Private Sub NoChange_Click()
If NoChange.Value = True Then
Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = ""
End If
End Sub
Private Sub HideColumn_Click()
If HideColumn.Value = True Then
Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = "HideColumn"
End If
End Sub
Private Sub NoChangeWhileTyping_Click()
If NoChangeWhileTyping.Value = True Then
Sheet0.OLEObjects("Capitalization").Object.Caption = ""
End If
End Sub
Private Sub CapitalizeFirstLetter_Click()
If CapitalizeFirstLetter.Value = True Then
Sheet0.OLEObjects("Capitalization").Object.Caption = "First"
End If
End Sub
Private Sub Open_Sorting_Form_Click()
On Error Resume Next
Application.ScreenUpdating = False
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Visible = xlSheetVisible
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Activate
Application.ScreenUpdating = True
Unload Me
If Right(ActiveSheet.Name, 7) = "_Sorted" Then
MsgBox "Sorted shteet already created", , "BBS Program"
Exit Sub
Else
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "")).Visible = xlSheetVisible
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "")).Activate
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
On Error GoTo ThisWorkbookSortingFormShow
Application.Run "'" & Sheet0.OLEObjects("BBSProgram").Object.Caption & "'!OpenSortForm"
End If
Exit Sub
ThisWorkbookSortingFormShow:
On Error GoTo MacroNotExist
On Error Resume Next
SortingForm.Show
If Err.Number = 75 Then
Dim VersionError As String
VersionError = IIf(Application.Version = 12, "Please try it in Excel-2003 or Excel-2010.", "Reinstall or repair the Excel")
MsgBox "Unable to open the form in this version of Excel." & Chr(10) & VersionError, , "BBS Program"
Exit Sub
End If
On Error GoTo 0
Exit Sub
MacroNotExist:
MsgBox "BBS Program (another file opened in background) is old version." & Chr(10) & "Use the latest version of BBS Program to create Tag.", , "BBS Program"
Application.EnableEvents = True
End Sub
Private Sub OpenOptimizeForm_Click()
On Error Resume Next
Application.ScreenUpdating = False
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Optimized").Visible = xlSheetVisible
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Optimized").Activate
Application.ScreenUpdating = True
Unload Me
If Right(ActiveSheet.Name, 4) = "_Optimized" Then
MsgBox "Optimized sheet already created", , "BBS Program"
Exit Sub
Else
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Visible = xlSheetVisible
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Activate
If Right(ActiveSheet.Name, 7) <> "_Sorted" Then
MsgBox "Sort the sheet before Optimization.", , "BBS Program"
Exit Sub
End If
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
On Error GoTo MacroNotExist
Application.Run "'" & Sheet0.OLEObjects("BBSProgram").Object.Caption & "'!OpenOptimizeForm"
Exit Sub
MacroNotExist:
MsgBox "BBS Program (another file opened in background) is old version." & Chr(10) & "Use the latest version of BBS Program to create Tag.", , "BBS Program"
Application.EnableEvents = True
End Sub
Private Sub OpenTagForm_Click()
On Error Resume Next
Application.ScreenUpdating = False
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Tag").Visible = xlSheetVisible
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Tag").Activate
Application.ScreenUpdating = True
Unload Me
If Right(ActiveSheet.Name, 4) = "_Tag" Then
MsgBox "Tags already created", , "BBS Program"
Exit Sub
Else
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Visible = xlSheetVisible
Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted").Activate
If Right(ActiveSheet.Name, 7) <> "_Sorted" Then
MsgBox "Sort the sheet before Tag.", , "BBS Program"
Exit Sub
End If
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
On Error GoTo MacroNotExist
Application.Run "'" & Sheet0.OLEObjects("BBSProgram").Object.Caption & "'!OpenTagForm"
Exit Sub
MacroNotExist:
MsgBox "BBS Program (another file opened in background) is old version." & Chr(10) & "Use the latest version of BBS Program to create Tag.", , "BBS Program"
Application.EnableEvents = True
End Sub
Private Sub SmartCapitalization_Click()
If SmartCapitalization.Value = True Then
Sheet0.OLEObjects("Capitalization").Object.Caption = "Smart"
End If
End Sub
Private Sub TermsOfUse_Click()
On Error Resume Next
TermsOfUseForm.Show
If Err.Number = 75 Then
Dim VersionError As String
VersionError = IIf(Application.Version = 12, "Please try it in Excel-2003 or Excel-2010.", "Reinstall or repair the Excel")
MsgBox "Unable to open the form in this version of Excel." & Chr(10) & VersionError, , "BBS Program"
Exit Sub
End If
On Error GoTo 0
End Sub
Private Sub UnhideColumn_Click()
If UnhideColumn.Value = True Then
Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = "UnhideColumn"
End If
End Sub
Private Sub OpenProgramOptions_Click()
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
If AlreadyOpen = False Then
MsgBox "BBS Program file not open", , "BBS Program"
Exit Sub
End If
Unload Me
Workbooks(wbname).Windows(1).Visible = True
Workbooks(wbname).Activate
Application.Run "'" & wbname & "'!OpenProgramOptionsForm"
End Sub
Private Sub CloseCommand_Click()
Unload Me
End Sub
Private Sub AutomaticallyOpenProgramFile_Click()
On Error Resume Next
If AutomaticallyOpenProgramFile.Value = True And Sheet0.AutomaticallyOpenProgramFile.Object.Caption = "False" Then
Sheet0.AutomaticallyOpenProgramFile.Object.Caption = "True"
End If
If AutomaticallyOpenProgramFile.Value = False And Sheet0.AutomaticallyOpenProgramFile.Object.Caption = "True" Then
Sheet0.AutomaticallyOpenProgramFile.Object.Caption = "False"
End If
End Sub
Private Sub ChangeProgramFile_Click()
If Application.Version < 12 Then
NewBBSProgramFile = Application.GetOpenFilename("Excel 2000-2003 Files (*.xls),*.xls", , "Select BBS Program File", , False)
Else
NewBBSProgramFile = Application.GetOpenFilename("Excel Macro-Enabled Workbook (*.xlsm),*.xlsm", , "Select BBS Program File", , False)
End If
If NewBBSProgramFile <> False Then
If NewBBSProgramFile = ThisWorkbook.FullName Then
MsgBox "You selected this file." & Chr(10) & "Select BBS Program file.", vbCritical, "BBS Program"
Exit Sub
End If
Dim MethodOfBendingString, FilePathString As String
FilePathString = Replace(Left(NewBBSProgramFile, InStrRev(NewBBSProgramFile, "\")) & "[" & Right(NewBBSProgramFile, Len(NewBBSProgramFile) - Len(Left(NewBBSProgramFile, InStrRev(NewBBSProgramFile, "\")))), "'", "''")
MethodOfBendingString = ExecuteExcel4Macro("'" & FilePathString & "]Sheet1'!R1C26")
If Not IsError(MethodOfBendingString) Then
If MethodOfBendingString = "Manual Bending" Or MethodOfBendingString = "Machine Bending" Then  'error
MethodOfBending = MethodOfBendingString
ProgramFileFullName = NewBBSProgramFile
AutomaticallyOpenProgramFile.Value = True
Sheet0.ProgramFileFullName.Caption = NewBBSProgramFile
Sheet0.AutomaticallyOpenProgramFile.Caption = "True"
Else
MsgBox "The file you selected is not a BBS Program", vbExclamation, "BBS Program"
End If
Else
MsgBox "The file you selected is not a BBS Program", vbExclamation, "BBS Program"
End If
End If
Set MethodOfBendingString = Nothing
Set NewBBSProgramFile = Nothing
End Sub
Private Sub OpenTemplateSheet_Click()
Application.ScreenUpdating = False
Sheet0.Visible = xlSheetVisible
Sheet0.Select
ActiveWindow.ScrollRow = 3
ActiveWindow.ScrollColumn = 20
ActiveSheet.Range("T3").Select
Application.ScreenUpdating = True
Unload Me
End Sub
Private Sub UserForm_Activate()
On Error Resume Next: TemplateAvailable = Sheet0.Range("A1")
If Err Then
MsgBox "************WARNING************" & Chr(10) & "Template sheet has been Deleted." & Chr(10) & "This file cannot work without Template sheet." & Chr(10) & "Create a copy of Template sheet from another BBS Format to this file." & Chr(10) & "Template sheet is a hidden sheet in the BBS Format." & Chr(10) & "You can access Template sheet by clicking the ''View Template'' button in the Format Options (Press Control T).", vbCritical, "BBS Program              WARNING: DO NOT DELETE TEMPLATE SHEET"
Unload Me: Exit Sub: End If: On Error GoTo 0
Dim MethodOfBendingString, FilePathString As String
FilePathString = Replace(Left(Sheet0.ProgramFileFullName, InStrRev(Sheet0.ProgramFileFullName, "\")) & "[" & Right(Sheet0.ProgramFileFullName, Len(Sheet0.ProgramFileFullName) - Len(Left(Sheet0.ProgramFileFullName, InStrRev(Sheet0.ProgramFileFullName, "\")))), "'", "''")
On Error Resume Next
If Dir(Sheet0.ProgramFileFullName) = "" Then
Sheet0.OLEObjects("ProgramFileFullName").Object.Caption = "Need to locate file"
Sheet0.OLEObjects("AutomaticallyOpenProgramFile").Object.Caption = "False"
End If
On Error GoTo 0
ProgramFileFullName.Value = Sheet0.ProgramFileFullName
If Sheet0.AutomaticallyOpenProgramFile.Object.Caption = "True" Then
AutomaticallyOpenProgramFile.Value = True
Else
AutomaticallyOpenProgramFile.Value = False
End If
Set MethodOfBendingString = Nothing
If Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = "" Then
NoChange.Value = True
ElseIf Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = "HideColumn" Then
HideColumn.Value = True
ElseIf Sheet0.OLEObjects("PrintOptionLabel").Object.Caption = "UnhideColumn" Then
UnhideColumn.Value = True
End If
If Sheet0.OLEObjects("Capitalization").Object.Caption = "" Then
NoChangeWhileTyping.Value = True
ElseIf Sheet0.OLEObjects("Capitalization").Object.Caption = "First" Then
CapitalizeFirstLetter.Value = True
ElseIf Sheet0.OLEObjects("Capitalization").Object.Caption = "Smart" Then
SmartCapitalization.Value = True
End If
Dim TagedSheet As Worksheet, SortedSheet As Worksheet, OptimizedSheet As Worksheet
On Error Resume Next
Set SortedSheet = Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Sorted")
Set OptimizedSheet = Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Optimized")
Set TagedSheet = Worksheets(Replace(Replace(Replace(ActiveSheet.Name, "_Optimized", ""), "_Tag", ""), "_Sorted", "") & "_Tag")
On Error GoTo 0
If Not TagedSheet Is Nothing Then
OpenTagForm.Enabled = False
OpenTagForm.Caption = "Tags Completed"
End If
If Not OptimizedSheet Is Nothing Then
OpenOptimizeForm.Enabled = False
OpenOptimizeForm.Caption = "Sort Completed"
End If
If Not SortedSheet Is Nothing Then
Open_Sorting_Form.Enabled = False
Open_Sorting_Form.Caption = "Sort Completed"
Else
OpenTagForm.Enabled = False
OpenOptimizeForm.Enabled = False
End If
If OpenTagForm.Enabled = False And OpenOptimizeForm.Enabled = False And Open_Sorting_Form.Enabled = False Then
CloseCommand.SetFocus
End If
End Sub
