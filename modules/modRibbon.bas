Option Explicit

' Ribbon load (optional, but good practice)
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    ' You can store ribbon in a module-level variable if you need to refresh it later.
End Sub

' Button 1: Open GUI
Public Sub BtnHistoryGUI_Click(control As IRibbonControl)
    ShowHistoryGUI
End Sub

' Button 2: Clear cache (only if you implemented caching)
Public Sub BtnClearCache_Click(control As IRibbonControl)
    On Error Resume Next
    AtHistoryClearCache ' if you add this function (see below)
    On Error GoTo 0
    MsgBox "Cache cleared.", vbInformation
End Sub

Public Sub BtnAbout_Click(control As IRibbonControl)
    MsgBox "IP21 Historian Add-in" & vbCrLf & _
           "Version: 1.2" & vbCrLf & _
           "Credits: Mauricio Melara" & vbCrLf & _
           "Support: mauricio.melara@quikrete-cement.com", vbInformation, "About"
End Sub

