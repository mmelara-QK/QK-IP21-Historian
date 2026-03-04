
Option Explicit

Public Sub ShowHistoryGUI()
    frmAtHistory.Show
End Sub

Public Sub Run_AtHistoryPrompted()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tagRange As Range, startCell As Range, endCell As Range

    On Error Resume Next
    Set tagRange = Application.InputBox("Select the TAG HEADER range (one row, multiple columns).", "Tags", Type:=8)
    If tagRange Is Nothing Then Exit Sub

    Set startCell = Application.InputBox("Select the Start DateTime cell.", "Start Time", Type:=8)
    If startCell Is Nothing Then Exit Sub

    Set endCell = Application.InputBox("Select the End DateTime cell.", "End Time", Type:=8)
    If endCell Is Nothing Then Exit Sub
    On Error GoTo 0

    Dim period As Long
    period = CLng(Application.InputBox("Enter Period (P). Example: 1", "Period", 1, Type:=1))

    Dim pu As Long
    pu = CLng(Application.InputBox("Enter Period Units (PU): 0=Day, 1=Hour, 2=Minute, 3=Second", "Period Units", 2, Type:=1))

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    RunHistoryForTagRowTimestampAligned ws, tagRange, startCell, endCell, period, pu

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Done. Check 'AtHistory_Log' if any tags returned Error.", vbInformation
End Sub
