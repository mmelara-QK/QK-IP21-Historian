VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAtHistory 
   Caption         =   "History Pull"
   ClientHeight    =   6790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   OleObjectBlob   =   "frmAtHistory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAtHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblTags_Click()

End Sub

Private Sub lblUnits_Click()

End Sub

Private Sub UserForm_Initialize()
    txtPeriod.Value = "1"
    optMin.Value = True
End Sub

Private Sub btnPickTags_Click()
    Dim v As Variant
    v = Application.InputBox("Select the TAG HEADER range (one row across).", "Pick Tag Range", Type:=8)
    If v = False Then Exit Sub            ' user cancelled
    txtTags.Value = SheetQualifiedAddress(v)
End Sub

Private Sub btnPickStart_Click()
    Dim v As Variant
    v = Application.InputBox("Select the Start DateTime cell.", "Pick Start Cell", Type:=8)
    If v = False Then Exit Sub            ' user cancelled
    txtStart.Value = SheetQualifiedAddress(v)
End Sub

Private Sub btnPickEnd_Click()
    Dim v As Variant
    v = Application.InputBox("Select the End DateTime cell.", "Pick End Date Cell", Type:=8)
    If v = False Then Exit Sub            ' user cancelled
    txtEnd.Value = SheetQualifiedAddress(v)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub btnPickOutput_Click()
    Dim v As Variant
    v = Application.InputBox("Pick the TOP-LEFT output cell (where the table will spill).", "Pick Output Cell", Type:=8)
    If v = False Then Exit Sub            ' user cancelled
    txtOutput.Value = SheetQualifiedAddress(v)
End Sub

Private Sub btnInsertFormula_Click()
    On Error GoTo Fail

    If Len(Trim$(txtTags.Value)) = 0 Then Err.Raise vbObjectError + 1, , "Pick a tag header range."
    If Len(Trim$(txtStart.Value)) = 0 Then Err.Raise vbObjectError + 1, , "Pick a start cell."
    If Len(Trim$(txtEnd.Value)) = 0 Then Err.Raise vbObjectError + 1, , "Pick an end cell."
    If Len(Trim$(txtOutput.Value)) = 0 Then Err.Raise vbObjectError + 1, , "Pick an output cell."

    Dim period As Long
    If Not IsNumeric(txtPeriod.Value) Then Err.Raise vbObjectError + 1, , "Period must be numeric."
    period = CLng(txtPeriod.Value)

    Dim pu As Long
    pu = GetSelectedPU() ' your existing PU option logic

    Dim outCell As Range
    Set outCell = Application.Range(txtOutput.Value)

    ' Build formula with fully-qualified addresses so it keeps working
    Dim tagsAddr As String, startAddr As String, endAddr As String
    tagsAddr = txtTags.Value
    startAddr = txtStart.Value
    endAddr = txtEnd.Value

    Dim f As String
    f = "=AtHistoryData(" & tagsAddr & "," & startAddr & "," & endAddr & "," & period & "," & pu & ")"

    outCell.Formula2 = f

    MsgBox "Formula inserted. Press F9 to refresh; it will only re-query if inputs changed.", vbInformation
    Exit Sub

Fail:
    MsgBox "Could not insert formula: " & Err.Description, vbExclamation
End Sub


Private Sub btnRun_Click()
    On Error GoTo Fail

    If Len(Trim$(txtTags.Value)) = 0 Then Err.Raise vbObjectError + 1, , "Pick a tag header range."
    If Len(Trim$(txtStart.Value)) = 0 Then Err.Raise vbObjectError + 1, , "Pick a start time cell."
    If Len(Trim$(txtEnd.Value)) = 0 Then Err.Raise vbObjectError + 1, , "Pick an end time cell."

    Dim period As Long
    If Not IsNumeric(txtPeriod.Value) Then Err.Raise vbObjectError + 1, , "Period must be a number."
    period = CLng(txtPeriod.Value)
    If period <= 0 Then Err.Raise vbObjectError + 1, , "Period must be > 0."

    Dim pu As Long
    pu = GetSelectedPU()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tagRange As Range, startCell As Range, endCell As Range
    Set tagRange = RangeFromAddress(txtTags.Value)
    Set startCell = RangeFromAddress(txtStart.Value)
    Set endCell = RangeFromAddress(txtEnd.Value)

    ' Basic validation
    If tagRange.Rows.Count <> 1 Then Err.Raise vbObjectError + 1, , "Tag header range must be a single row."
    If Not IsDate(startCell.Value) Then Err.Raise vbObjectError + 1, , "Start cell does not contain a DateTime."
    If Not IsDate(endCell.Value) Then Err.Raise vbObjectError + 1, , "End cell does not contain a DateTime."
    If CDate(endCell.Value) <= CDate(startCell.Value) Then Err.Raise vbObjectError + 1, , "End time must be after start time."

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    RunHistoryForTagRowTimestampAligned ws, tagRange, startCell, endCell, period, pu

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Done. Check AtHistory_Log if any tags show Error.", vbInformation
    Unload Me
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Could not run: " & Err.Description, vbExclamation
End Sub

Private Function GetSelectedPU() As Long
    If optDay.Value Then GetSelectedPU = 0: Exit Function
    If optHour.Value Then GetSelectedPU = 1: Exit Function
    If optMin.Value Then GetSelectedPU = 2: Exit Function
    If optSec.Value Then GetSelectedPU = 3: Exit Function
    GetSelectedPU = 2
End Function

Private Function RangeFromAddress(ByVal addr As String) As Range
    ' addr is a fully qualified A1 with workbook reference if you used Address(..., True)
    ' but Range() can still resolve it in the active workbook context.
    Set RangeFromAddress = Application.Range(addr)
End Function

Private Function SheetQualifiedAddress(ByVal r As Range) As String
    ' No workbook path; only Sheet!$A$1 style
    SheetQualifiedAddress = "'" & r.Parent.Name & "'!" & r.Address(True, True, xlA1)
End Function


