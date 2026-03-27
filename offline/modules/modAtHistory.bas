Option Explicit

' ====== CONFIG ======
Private Const BASE_URL As String = "http://phyvqktwatipa03/Web21/ProcessData/AtProcessDataREST.dll"
Private Const HISTORY_ENDPOINT As String = "/History"

Private Const DEFAULT_DATA_SOURCE As String = "localhost"
Private Const DEFAULT_FIELD As String = "VAL"
Private Const DEFAULT_HF As Long = 0        ' 0 Raw
Private Const DEFAULT_RT As Long = 1        ' retrieval type
Private Const DEFAULT_STEPPED As Long = 0   ' 0/1

' ====== SESSION (reuse WinHTTP to keep cookies/session) ======
Private mHttp As Object

Private Function GetHttp() As Object
    If mHttp Is Nothing Then
        Set mHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        On Error Resume Next
        mHttp.SetAutoLogonPolicy 0 ' Always try current Windows credentials
        On Error GoTo 0
    End If
    Set GetHttp = mHttp
End Function

Private Function MsKey(ByVal ms As Double) As String
    ' Force an integer string with no scientific notation
    MsKey = Format$(ms, "0")
End Function

' ====== LOGGING ======
Private Function EnsureLogSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("AtHistory_Log")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "AtHistory_Log"
        ws.Range("A1:K1").Value = Array( _
            "When", "Sheet", "TagHeaderCell", "TagName", _
            "HTTP Status", "HTTP StatusText", "RespSnippet", _
            "ErrNumber", "ErrDescription", "PayloadSnippet", "URL" _
        )
        ws.Columns("A:K").EntireColumn.AutoFit
    End If

    Set EnsureLogSheet = ws
End Function

Private Sub AppendLog( _
    ByVal logWs As Worksheet, _
    ByVal sheetName As String, _
    ByVal tagCellAddr As String, _
    ByVal tagName As String, _
    ByVal httpStatus As Long, _
    ByVal httpStatusText As String, _
    ByVal respText As String, _
    ByVal errNum As Long, _
    ByVal errDesc As String, _
    ByVal payload As String, _
    ByVal url As String _
)
    Dim r As Long
    r = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1

    logWs.Cells(r, 1).Value = Now
    logWs.Cells(r, 2).Value = sheetName
    logWs.Cells(r, 3).Value = tagCellAddr
    logWs.Cells(r, 4).Value = tagName
    logWs.Cells(r, 5).Value = httpStatus
    logWs.Cells(r, 6).Value = httpStatusText
    logWs.Cells(r, 7).Value = Left$(respText, 300)
    logWs.Cells(r, 8).Value = errNum
    logWs.Cells(r, 9).Value = Left$(errDesc, 300)
    logWs.Cells(r, 10).Value = Left$(payload, 300)
    logWs.Cells(r, 11).Value = url
End Sub

' ====== HTTP POST (captures status + response) ======
Private Function PostHistoryRequest( _
    ByVal fullUrl As String, _
    ByVal payloadXml As String, _
    ByRef outStatus As Long, _
    ByRef outStatusText As String, _
    ByRef outRespText As String, _
    Optional ByVal timeoutMs As Long = 30000 _
) As String

    Dim http As Object
    Set http = GetHttp()

    http.Open "POST", fullUrl, False

    ' Match the sample page more closely:
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-16"
    http.SetRequestHeader "Accept", "application/json"

    http.SetTimeouts timeoutMs, timeoutMs, timeoutMs, timeoutMs
    http.Send payloadXml

    outStatus = http.Status
    outStatusText = http.StatusText
    outRespText = http.ResponseText

    If outStatus <> 200 Then
        Err.Raise vbObjectError + 514, "PostHistoryRequest", "HTTP " & outStatus & " - " & outStatusText
    End If

    PostHistoryRequest = outRespText
End Function

' ====== PAYLOAD ======
Private Function BuildHistoryPayload( _
    ByVal tagName As String, _
    ByVal startMs As Double, _
    ByVal endMs As Double, _
    ByVal period As Long, _
    ByVal periodUnits As Long _
) As String

    BuildHistoryPayload = _
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
                "<PU>" & periodUnits & "</PU>" & _
            "</Tag>" & _
        "</Q>"
End Function

' ====== FAST PARSER: Extract "t" and "v" from JSON using RegEx ======
Private Function ParseSamplesToDict(ByVal jsonText As String) As Object
    ' Returns Dictionary where key = ms timestamp string, value = Double (or Variant if needed)

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim re As Object, matches As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Match: "t":1769662800000,"v":0.2498299628  (allows scientific notation and negatives)
    re.Pattern = """t""\s*:\s*(\d+)\s*,\s*""v""\s*:\s*(null|[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)"
    re.Global = True
    re.IgnoreCase = True
    re.Multiline = True

    If re.Test(jsonText) Then
        Set matches = re.Execute(jsonText)

        For Each m In matches
            Dim tKey As String
            tKey = m.SubMatches(0) ' timestamp digits already non-scientific

            Dim vRaw As String
            vRaw = LCase$(m.SubMatches(1))

            If vRaw <> "null" Then
                d(tKey) = CDbl(m.SubMatches(1))
            Else
                ' If value is null, leave blank (or you could store Empty)
                ' d(tKey) = Empty
            End If
        Next m
    End If

    Set ParseSamplesToDict = d
End Function

' ====== SORT ======
Private Sub QuickSortDoubles(ByRef arr() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, tmp As Double

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

' ====== MAIN: one row of tags, timestamp union, aligned output ======
Public Sub RunHistoryForTagRowTimestampAligned( _
    ByVal ws As Worksheet, _
    ByVal tagHeaderRange As Range, _
    ByVal startCell As Range, _
    ByVal endCell As Range, _
    ByVal period As Long, _
    ByVal periodUnits As Long _
)
    Dim wb As Workbook: Set wb = ws.Parent
    Dim logWs As Worksheet: Set logWs = EnsureLogSheet(wb)

    Dim startMs As Double, endMs As Double
    startMs = ExcelLocalDateToUnixMs(CDate(startCell.Value))
    endMs = ExcelLocalDateToUnixMs(CDate(endCell.Value))

    Dim fullUrl As String
    fullUrl = BASE_URL & HISTORY_ENDPOINT

    ' Store results per tag column: colKey -> Dictionary(timeMs->value)
    Dim tagResults As Object: Set tagResults = CreateObject("Scripting.Dictionary")
    Dim errorTags As Object: Set errorTags = CreateObject("Scripting.Dictionary")
    Dim allTimes As Object: Set allTimes = CreateObject("Scripting.Dictionary")

    Dim cell As Range

    ' Variables used for error logging (must be in procedure scope)
    Dim st As Long, stText As String, respText As String
    Dim payload As String, tagName As String

    For Each cell In tagHeaderRange.Cells
        tagName = Trim$(CStr(cell.Value))
        If Len(tagName) = 0 Then GoTo ContinueNextTag

        ' reset per tag so logs aren't stale
        st = 0: stText = "": respText = "": payload = ""

        On Error GoTo TagErr

        payload = BuildHistoryPayload(tagName, startMs, endMs, period, periodUnits)

        Dim resp As String
        resp = PostHistoryRequest(fullUrl, payload, st, stText, respText)

        Dim d As Object
        Set d = ParseSamplesToDict(resp)

        Set tagResults(CStr(cell.Column)) = d

        Dim k As Variant
        For Each k In d.Keys
            If Not allTimes.Exists(CStr(k)) Then allTimes.Add CStr(k), True
        Next k

        On Error GoTo 0

ContinueNextTag:
        ' next tag
    Next cell

    ' If nothing succeeded, just ensure error marks are present and exit
    If allTimes.Count = 0 Then
        For Each cell In tagHeaderRange.Cells
            tagName = Trim$(CStr(cell.Value))
            If Len(tagName) > 0 And errorTags.Exists(CStr(cell.Column)) Then
                ws.Cells(tagHeaderRange.Row + 1, cell.Column).Value = "Error"
            End If
        Next cell
        Exit Sub
    End If

    ' Build sorted union timestamp array
    Dim times() As Double
    ReDim times(1 To allTimes.Count)

    Dim idx As Long: idx = 1
    Dim tk As Variant
    For Each tk In allTimes.Keys
        times(idx) = CDbl(tk)
        idx = idx + 1
    Next tk

    QuickSortDoubles times, LBound(times), UBound(times)

    ' Write Time column (left of first tag)
    Dim headerRow As Long, firstTagCol As Long, timeCol As Long
    headerRow = tagHeaderRange.Row
    firstTagCol = tagHeaderRange.Cells(1, 1).Column
    timeCol = firstTagCol - 1

    ws.Cells(headerRow, timeCol).Value = "Time"

    Dim n As Long: n = UBound(times)
    Dim timeArr() As Variant
    ReDim timeArr(1 To n, 1 To 1)

    Dim i As Long
    For i = 1 To n
        timeArr(i, 1) = UnixMsToExcelLocalDate(times(i))
    Next i

    With ws.Range(ws.Cells(headerRow + 1, timeCol), ws.Cells(headerRow + n, timeCol))
        .Value = timeArr
        .NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With

    ' Map timeMs string -> row offset
    Dim rowByTime As Object: Set rowByTime = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        rowByTime(MsKey(times(i))) = i
    Next i

    ' Write each tag's column aligned to timestamps
    For Each cell In tagHeaderRange.Cells
        tagName = Trim$(CStr(cell.Value))
        If Len(tagName) = 0 Then GoTo NextWrite

        Dim colKey As String: colKey = CStr(cell.Column)

        If errorTags.Exists(colKey) Then
            ws.Cells(headerRow + 1, cell.Column).Value = "Error"
            GoTo NextWrite
        End If

        If Not tagResults.Exists(colKey) Then
            ws.Cells(headerRow + 1, cell.Column).Value = "Error"
            GoTo NextWrite
        End If

        Dim dataDict As Object
        Set dataDict = tagResults(colKey)

        Dim outArr() As Variant
        ReDim outArr(1 To n, 1 To 1)

        Dim kk As Variant, r As Long
        For Each kk In dataDict.Keys
            If rowByTime.Exists(kk) Then
                r = CLng(rowByTime(kk))
                outArr(r, 1) = dataDict(kk)
            End If
        Next kk

        ws.Range(ws.Cells(headerRow + 1, cell.Column), ws.Cells(headerRow + n, cell.Column)).Value = outArr

NextWrite:
    Next cell

    Exit Sub

TagErr:
    ' Per your requirement: write "Error" under the tag and continue
    errorTags(CStr(cell.Column)) = True
    ws.Cells(tagHeaderRange.Row + 1, cell.Column).Value = "Error"

    ' Log details
    AppendLog logWs, ws.Name, cell.Address(False, False), tagName, st, stText, respText, Err.Number, Err.Description, payload, fullUrl

    Err.Clear
    Resume ContinueNextTag
End Sub

