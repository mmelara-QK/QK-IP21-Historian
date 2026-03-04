Option Explicit

' --- Windows timezone API (32/64-bit safe) ---
#If VBA7 Then
    Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#Else
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#End If

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

' Return the current local->UTC bias in minutes (includes DST if active)
Public Function LocalToUtcBiasMinutes() As Long
    Dim tzi As TIME_ZONE_INFORMATION
    Dim mode As Long
    mode = GetTimeZoneInformation(tzi)

    ' mode: 0=unknown, 1=standard time, 2=daylight time
    Select Case mode
        Case 2
            LocalToUtcBiasMinutes = tzi.Bias + tzi.DaylightBias
        Case Else
            LocalToUtcBiasMinutes = tzi.Bias + tzi.StandardBias
    End Select
End Function

' Convert Excel local DateTime -> Unix epoch milliseconds (UTC)
Public Function ExcelLocalDateToUnixMs(ByVal dtLocal As Date) As Double
    Dim biasMin As Long
    biasMin = LocalToUtcBiasMinutes()

    Dim dtUtc As Date
    ' Windows bias is minutes added to LOCAL to get UTC (e.g., Eastern Standard Bias=300)
    dtUtc = DateAdd("n", biasMin, dtLocal)

    ExcelLocalDateToUnixMs = (CDbl(dtUtc) - CDbl(DateSerial(1970, 1, 1))) * 86400000#
End Function

' Convert Unix ms (UTC) -> Excel local DateTime
Public Function UnixMsToExcelLocalDate(ByVal unixMs As Double) As Date
    Dim dtUtc As Date
    dtUtc = DateSerial(1970, 1, 1) + (unixMs / 86400000#)

    Dim biasMin As Long
    biasMin = LocalToUtcBiasMinutes()

    ' Local = UTC - bias
    UnixMsToExcelLocalDate = DateAdd("n", -biasMin, dtUtc)
End Function
