Option Explicit

' ===== CONFIG =====
Private Const BASE_URL As String = "http://phyvqktwatipa03/Web21/ProcessData/AtProcessDataREST.dll"
Private Const HISTORY_ENDPOINT As String = "/History"
Private Const DEFAULT_DATA_SOURCE As String = "localhost"
Private Const DEFAULT_FIELD As String = "VAL"
Private Const DEFAULT_HF As Long = 0
Private Const DEFAULT_RT As Long = 1
Private Const DEFAULT_STEPPED As Long = 0

' ===== SESSION =====
Private mHttp As Object

Private Function GetHttp() As Object
    If mHttp Is Nothing Then
        Set mHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        On Error Resume Next
        mHttp.SetAutoLogonPolicy 0
        On Error GoTo 0
    End If
    Set GetHttp = mHttp
End Function

Private Function MsKey(ByVal ms As Double) As String
    MsKey = Format$(ms, "0") ' no scientific notation
End Function

' ===== CACHE =====
' key -> 2D Variant array result
Private Function Cache() As Object
    Static d As Object
    If d Is Nothing Then Set d = CreateObject("Scripting.Dictionary")
    Set Cache = d
End Function

Public Function AtHistoryData(ByVal tagHeaderRange As Range, ByVal startDT As Variant, ByVal endDT As Variant, ByVal period As Long, ByVal pu As Long) As Variant
    Dim full As Variant
    full = AtHistoryTable(tagHeaderRange, startDT, endDT, period, pu)

    ' If AtHistoryTable returned an error string, pass it through
    If Not IsArray(full) Then
        AtHistoryData = full
        Exit Function
    End If

    Dim rL As Long, rU As Long, cL As Long, cU As Long
    rL = LBound(full, 1): rU = UBound(full, 1)
    cL = LBound(full, 2): cU = UBound(full, 2)

    ' If only header row exists, return empty string
    If rU <= rL Then
        AtHistoryData = ""
        Exit Function
    End If

    ' Create data-only array: rows 2..end
    Dim outArr() As Variant
    Dim r As Long, c As Long, rr As Long

    ReDim outArr(1 To (rU - rL), 1 To (cU - cL + 1))

    rr = 1
    For r = rL + 1 To rU
        For c = cL To cU
            outArr(rr, c - cL + 1) = full(r, c)
        Next c
        rr = rr + 1
    Next r

    AtHistoryData = outArr
End Function

Private Function MakeCacheKey(ByVal tagRange As Range, ByVal startDT As Date, ByVal endDT As Date, ByVal period As Long, ByVal pu As Long) As String
    Dim parts As String
    Dim c As Range

    parts = "S=" & CStr(CDbl(startDT)) & "|E=" & CStr(CDbl(endDT)) & "|P=" & CStr(period) & "|PU=" & CStr(pu) & "|Tags="

    For Each c In tagRange.Cells
        parts = parts & Trim$(CStr(c.Value)) & ";"
    Next c

    MakeCacheKey = parts
End Function

' ===== HTTP =====
Private Function PostHistory(ByVal payloadXml As String, ByRef outStatus As Long, ByRef outStatusText As String, ByRef outResp As String) As String
    Dim http As Object
    Set http = GetHttp()

    Dim url As String
    url = BASE_URL & HISTORY_ENDPOINT

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-16"
    http.SetRequestHeader "Accept", "application/json"
    http.Send payloadXml

    outStatus = http.Status
    outStatusText = http.StatusText
    outResp = http.ResponseText

    If outStatus <> 200 Then
        Err.Raise vbObjectError + 514, "PostHistory", "HTTP " & outStatus & " - " & outStatusText
    End If

    PostHistory = outResp
End Function

Private Function BuildPayload(ByVal tagName As String, ByVal startMs As Double, ByVal endMs As Double, ByVal period As Long, ByVal pu As Long) As String
    BuildPayload = _
        "<Q f=""d"" allQuotes=""1"">" & _
            "<Tag>" & _
                "<N><![CDATA[" & tagName & "]]></N>" & _
                "<D><![CDATA[" & DEFAULT_DATA_SOURCE & "]]></D>" & _
                "<F><![CDATA[" & DEFAULT_FIELD & "]]></F>" & _
                "<HF>" & DEFAULT_HF & "</HF>" & _
                "<St>" & MsKey(startMs) & "</St>" & _
                "<Et>" & MsKey(endMs) & "</Et>" & _
                "<RT>" & DEFAULT_RT & "</RT>" & _
                "<S>" & DEFAULT_STEPPED & "</S>" & _
                "<P>" & period & "</P>" & _
                "<PU>" & pu & "</PU>" & _
            "</Tag>" & _
        "</Q>"
End Function

' ===== FAST PARSER (RegEx): get "t" and "v" pairs =====
Private Function ParseSamplesToDict(ByVal jsonText As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim re As Object, matches As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """t""\s*:\s*(\d+)\s*,\s*""v""\s*:\s*(null|[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)"
    re.Global = True
    re.IgnoreCase = True
    re.Multiline = True

    If re.Test(jsonText) Then
        Set matches = re.Execute(jsonText)
        For Each m In matches
            Dim tKey As String: tKey = m.SubMatches(0)
            Dim vRaw As String: vRaw = LCase$(m.SubMatches(1))
            If vRaw <> "null" Then d(tKey) = CDbl(m.SubMatches(1))
        Next m
    End If

    Set ParseSamplesToDict = d
End Function

' ===== SORT =====
Private Sub QuickSortDoubles(ByRef arr() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, tmp As Double
    i = lo: j = hi
    pivot = arr((lo + hi) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortDoubles arr, lo, j
    If i < hi Then QuickSortDoubles arr, i, hi
End Sub

Public Sub AtHistoryClearCache()
    On Error Resume Next
    Cache.RemoveAll  ' if Cache() returns the static dictionary
    On Error GoTo 0
End Sub

' ==========================================================
'  PUBLIC UDF: spills a table [Time | Tag1 | Tag2 | ...]
' ==========================================================
Public Function AtHistoryTable(ByVal tagHeaderRange As Range, ByVal startDT As Variant, ByVal endDT As Variant, ByVal period As Long, ByVal pu As Long) As Variant
    On Error GoTo Fail

    ' validate dates
    If Not IsDate(startDT) Or Not IsDate(endDT) Then
        AtHistoryTable = "Start/End must be DateTime"
        Exit Function
    End If

    Dim s As Date, e As Date
    s = CDate(startDT)
    e = CDate(endDT)
    If e <= s Then
        AtHistoryTable = "End must be after Start"
        Exit Function
    End If

    If tagHeaderRange.Rows.Count <> 1 Then
        AtHistoryTable = "Tag range must be 1 row"
        Exit Function
    End If

    ' cache key
    Dim key As String
    key = MakeCacheKey(tagHeaderRange, s, e, period, pu)

    If Cache().Exists(key) Then
        AtHistoryTable = Cache()(key)
        Exit Function
    End If

    ' pull data
    Dim startMs As Double, endMs As Double
    startMs = ExcelLocalDateToUnixMs(s)
    endMs = ExcelLocalDateToUnixMs(e)

    Dim tagCount As Long: tagCount = tagHeaderRange.Columns.Count

    ' per-tag dictionary: timestampKey -> value
    Dim perTag() As Object
    ReDim perTag(1 To tagCount)

    Dim errors() As Boolean
    ReDim errors(1 To tagCount)

    Dim allTimes As Object: Set allTimes = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To tagCount
        Dim tagName As String
        tagName = Trim$(CStr(tagHeaderRange.Cells(1, i).Value))

        If Len(tagName) = 0 Then
            Set perTag(i) = CreateObject("Scripting.Dictionary")
        Else
            On Error GoTo TagFail

            Dim st As Long, stText As String, respText As String
            Dim payload As String, resp As String
            payload = BuildPayload(tagName, startMs, endMs, period, pu)
            resp = PostHistory(payload, st, stText, respText)

            Set perTag(i) = ParseSamplesToDict(resp)

            Dim k As Variant
            For Each k In perTag(i).Keys
                If Not allTimes.Exists(CStr(k)) Then allTimes.Add CStr(k), True
            Next k

            On Error GoTo Fail
        End If

ContinueTag:
    Next i

    If allTimes.Count = 0 Then
        AtHistoryTable = "No data"
        Exit Function
    End If

    ' build sorted time list
    Dim times() As Double
    ReDim times(1 To allTimes.Count)

    Dim idx As Long: idx = 1
    Dim tk As Variant
    For Each tk In allTimes.Keys
        times(idx) = CDbl(tk)
        idx = idx + 1
    Next tk

    QuickSortDoubles times, LBound(times), UBound(times)

    ' output array: rows = 1 header + N data, cols = 1 time + tagCount
    Dim n As Long: n = UBound(times)
    Dim outArr() As Variant
    ReDim outArr(1 To n + 1, 1 To tagCount + 1)

    ' headers
    outArr(1, 1) = "Time"
    For i = 1 To tagCount
        outArr(1, i + 1) = CStr(tagHeaderRange.Cells(1, i).Value)
    Next i

    ' map time key -> row index in output
    Dim rowByTime As Object: Set rowByTime = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To n
        rowByTime(MsKey(times(r))) = r + 1 ' +1 for header row
        outArr(r + 1, 1) = UnixMsToExcelLocalDate(times(r)) ' datetime
    Next r

    ' fill tag columns aligned
    For i = 1 To tagCount
        If errors(i) Then
            outArr(2, i + 1) = "Error"
        Else
            Dim d As Object: Set d = perTag(i)
            Dim kk As Variant
            For Each kk In d.Keys
                If rowByTime.Exists(CStr(kk)) Then
                    outArr(rowByTime(CStr(kk)), i + 1) = d(kk)
                End If
            Next kk
        End If
    Next i

    ' store cache and return
    Cache().Add key, outArr
    AtHistoryTable = outArr
    Exit Function

TagFail:
    errors(i) = True
    Set perTag(i) = CreateObject("Scripting.Dictionary")
    Resume ContinueTag

Fail:
    AtHistoryTable = "Error: " & Err.Description
End Function


