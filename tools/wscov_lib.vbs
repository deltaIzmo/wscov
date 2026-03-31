' Shared library loaded with ExecuteGlobal.
' Keep this file free of Option Explicit so it can be loaded dynamically.

Const WSCOV_EXIT_OK = 0
Const WSCOV_EXIT_TEST_FAILED = 1
Const WSCOV_EXIT_USAGE = 2
Const WSCOV_EXIT_PARSE = 3
Const WSCOV_EXIT_INSTRUMENT = 4
Const WSCOV_EXIT_RUNTIME = 5

Const WSCOV_ERR_USAGE = 5202
Const WSCOV_ERR_PARSE = 5203
Const WSCOV_ERR_INSTRUMENT = 5204
Const WSCOV_ERR_RUNTIME = 5205
Const WSCOV_ERR_JSON = 5290

Function CreateList()
  Dim list
  Set list = CreateObject("Scripting.Dictionary")
  list.Add "__count", 0
  Set CreateList = list
End Function

Function IsList(value)
  IsList = False
  If IsObject(value) Then
    If TypeName(value) = "Dictionary" Then
      On Error Resume Next
      IsList = value.Exists("__count")
      Err.Clear
      On Error GoTo 0
    End If
  End If
End Function

Sub ListAdd(list, item)
  Dim indexKey
  indexKey = CStr(CLng(list("__count")))
  list.Add indexKey, item
  list("__count") = CLng(list("__count")) + 1
End Sub

Function ListCount(list)
  ListCount = CLng(list("__count"))
End Function

Sub PrintInstrumentUsage()
  WScript.Echo "Usage: cscript //nologo tools\wscov_instrument.vbs <input.wsc> <output.instrumented.wsc> <coverage-map.json>"
End Sub

Sub PrintRunUsage()
  WScript.Echo "Usage: cscript //nologo tools\wscov_run.vbs <instrumented.wsc> [componentId] <testsDir> <coverage-map.json> <outDir>"
End Sub

Function ParseArgsInstrument(args)
  Dim parsed
  Set parsed = CreateObject("Scripting.Dictionary")

  If args.Count <> 3 Then
    PrintInstrumentUsage
    RaiseUsage "Invalid arguments."
  End If

  parsed.Add "inputWscPath", ToAbsoluteLocalPath(CStr(args(0)))
  parsed.Add "outputWscPath", ToAbsoluteLocalPath(CStr(args(1)))
  parsed.Add "coverageMapPath", ToAbsoluteLocalPath(CStr(args(2)))

  Set ParseArgsInstrument = parsed
End Function

Function ParseArgsRun(args)
  Dim parsed, argCount
  Set parsed = CreateObject("Scripting.Dictionary")
  argCount = args.Count

  If argCount <> 4 And argCount <> 5 Then
    PrintRunUsage
    RaiseUsage "Invalid arguments."
  End If

  parsed.Add "instrumentedWscPath", ToAbsoluteLocalPath(CStr(args(0)))

  If argCount = 4 Then
    parsed.Add "componentId", ""
    parsed.Add "testsDir", ToAbsoluteLocalPath(CStr(args(1)))
    parsed.Add "coverageMapPath", ToAbsoluteLocalPath(CStr(args(2)))
    parsed.Add "outDir", ToAbsoluteLocalPath(CStr(args(3)))
  Else
    parsed.Add "componentId", Trim(CStr(args(1)))
    parsed.Add "testsDir", ToAbsoluteLocalPath(CStr(args(2)))
    parsed.Add "coverageMapPath", ToAbsoluteLocalPath(CStr(args(3)))
    parsed.Add "outDir", ToAbsoluteLocalPath(CStr(args(4)))
  End If

  Set ParseArgsRun = parsed
End Function

Function ToAbsoluteLocalPath(pathValue)
  Dim value, fso
  value = Trim(CStr(pathValue))
  If value = "" Then
    RaiseUsage "Path must not be empty."
  End If

  If InStr(1, value, "://", vbTextCompare) > 0 Then
    RaiseUsage "URL paths are not allowed: " & value
  End If

  If Left(value, 2) = "\\" Then
    RaiseUsage "UNC paths are not allowed: " & value
  End If

  Set fso = CreateObject("Scripting.FileSystemObject")
  ToAbsoluteLocalPath = fso.GetAbsolutePathName(value)
End Function

Sub EnsureFileExists(filePath)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(filePath) Then
    RaiseUsage "File does not exist: " & filePath
  End If
End Sub

Sub EnsureFolderExists(folderPath)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(folderPath) Then
    RaiseUsage "Folder does not exist: " & folderPath
  End If
End Sub

Sub EnsureDirExists(folderPath)
  Dim fso, parentPath
  Set fso = CreateObject("Scripting.FileSystemObject")

  If folderPath = "" Then
    Exit Sub
  End If

  If fso.FolderExists(folderPath) Then
    Exit Sub
  End If

  parentPath = fso.GetParentFolderName(folderPath)
  If parentPath <> "" And (Not fso.FolderExists(parentPath)) Then
    EnsureDirExists parentPath
  End If

  On Error Resume Next
  fso.CreateFolder folderPath
  If Err.Number <> 0 Then
    Dim message
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseRuntime "Failed to create folder: " & folderPath & " (" & message & ")"
  End If
  On Error GoTo 0
End Sub

Function ReadTextFile(filePath)
  Dim stream, message
  On Error Resume Next
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2 ' adTypeText
  stream.Charset = "utf-8"
  stream.Open
  stream.LoadFromFile filePath
  ReadTextFile = stream.ReadText(-1)
  stream.Close
  If Err.Number <> 0 Then
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseRuntime "Failed to read file as UTF-8: " & filePath & " (" & message & ")"
  End If
  On Error GoTo 0
End Function

Sub WriteTextFileUtf8(filePath, text)
  Dim fso, parentPath, stream, message
  Set fso = CreateObject("Scripting.FileSystemObject")
  parentPath = fso.GetParentFolderName(filePath)
  If parentPath <> "" Then
    EnsureDirExists parentPath
  End If

  On Error Resume Next
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2 ' adTypeText
  stream.Charset = "utf-8"
  stream.Open
  stream.WriteText CStr(text)
  stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
  stream.Close
  If Err.Number <> 0 Then
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseRuntime "Failed to write UTF-8 file: " & filePath & " (" & message & ")"
  End If
  On Error GoTo 0
End Sub

Sub WriteAllLines(filePath, lines)
  Dim text, i
  text = ""

  If IsObject(lines) Then
    If IsList(lines) Then
      For i = 0 To ListCount(lines) - 1
        If i > 0 Then
          text = text & vbCrLf
        End If
        text = text & CStr(lines(CStr(i)))
      Next
      WriteTextFileUtf8 filePath, text
      Exit Sub
    End If
  End If

  If IsArray(lines) Then
    For i = LBound(lines) To UBound(lines)
      If i > LBound(lines) Then
        text = text & vbCrLf
      End If
      text = text & CStr(lines(i))
    Next
    WriteTextFileUtf8 filePath, text
    Exit Sub
  End If

  RaiseRuntime "WriteAllLines expects list or array."
End Sub

Function JsonEscape(value)
  Dim s, i, ch, codepoint
  s = CStr(value)
  JsonEscape = ""

  For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    Select Case ch
      Case """"
        JsonEscape = JsonEscape & Chr(92) & Chr(34)
      Case "\"
        JsonEscape = JsonEscape & Chr(92) & Chr(92)
      Case "/"
        JsonEscape = JsonEscape & "\/"
      Case vbBack
        JsonEscape = JsonEscape & "\b"
      Case vbFormFeed
        JsonEscape = JsonEscape & "\f"
      Case vbCr
        JsonEscape = JsonEscape & "\r"
      Case vbLf
        JsonEscape = JsonEscape & "\n"
      Case vbTab
        JsonEscape = JsonEscape & "\t"
      Case Else
        codepoint = AscW(ch)
        If codepoint < 32 Then
          JsonEscape = JsonEscape & "\u" & Right("0000" & Hex(codepoint), 4)
        Else
          JsonEscape = JsonEscape & ch
        End If
    End Select
  Next
End Function

Function BuildCoverageMapJson(sourceWscPath, instrumentedWscPath, components)
  Dim sb, i, j, component, points, point
  sb = ""
  sb = sb & "{" & vbCrLf
  sb = sb & "  ""formatVersion"": ""0.1""," & vbCrLf
  sb = sb & "  ""generatedAt"": """ & JsonEscape(FormatIsoTimestamp(Now)) & """," & vbCrLf
  sb = sb & "  ""sourceWscPath"": """ & JsonEscape(sourceWscPath) & """," & vbCrLf
  sb = sb & "  ""instrumentedWscPath"": """ & JsonEscape(instrumentedWscPath) & """," & vbCrLf
  sb = sb & "  ""components"": [" & vbCrLf

  If Not IsList(components) Then
    RaiseRuntime "BuildCoverageMapJson expects components as list."
  End If

  For i = 0 To ListCount(components) - 1
    Set component = components(CStr(i))
    If i > 0 Then
      sb = sb & "," & vbCrLf
    End If

    sb = sb & "    {" & vbCrLf
    sb = sb & "      ""componentId"": """ & JsonEscape(CStr(component("componentId"))) & """," & vbCrLf
    sb = sb & "      ""language"": """ & JsonEscape(CStr(component("language"))) & """," & vbCrLf
    sb = sb & "      ""scriptIndex"": " & CStr(component("scriptIndex")) & "," & vbCrLf
    sb = sb & "      ""points"": [" & vbCrLf

    Set points = component("points")
    For j = 0 To ListCount(points) - 1
      Set point = points(CStr(j))
      If j > 0 Then
        sb = sb & "," & vbCrLf
      End If

      sb = sb & "        {" & vbCrLf
      sb = sb & "          ""id"": """ & JsonEscape(CStr(point("id"))) & """," & vbCrLf
      sb = sb & "          ""kind"": """ & JsonEscape(CStr(point("kind"))) & """," & vbCrLf
      sb = sb & "          ""proc"": """ & JsonEscape(CStr(point("proc"))) & """," & vbCrLf
      sb = sb & "          ""origScriptLine"": " & CStr(point("origScriptLine")) & "," & vbCrLf
      sb = sb & "          ""origScriptLineEnd"": " & CStr(point("origScriptLineEnd")) & "," & vbCrLf
      sb = sb & "          ""snippet"": """ & JsonEscape(CStr(point("snippet"))) & """" & vbCrLf
      sb = sb & "        }"
    Next

    sb = sb & vbCrLf & "      ]" & vbCrLf
    sb = sb & "    }"
  Next

  sb = sb & vbCrLf & "  ]" & vbCrLf
  sb = sb & "}" & vbCrLf

  BuildCoverageMapJson = sb
End Function

Function ParseCoverageMapJson(jsonText)
  Dim position, mapObj, formatVersion, components
  Set mapObj = CreateObject("Scripting.Dictionary")
  Set components = CreateList()
  position = 1
  formatVersion = ""

  JsonSkipWhitespace jsonText, position
  JsonExpectChar jsonText, position, "{"

  Do
    Dim key
    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, "}") Then
      Exit Do
    End If

    key = JsonReadString(jsonText, position)
    JsonSkipWhitespace jsonText, position
    JsonExpectChar jsonText, position, ":"

    Select Case key
      Case "formatVersion"
        formatVersion = JsonReadString(jsonText, position)
      Case "components"
        Set components = JsonParseComponentsArray(jsonText, position)
      Case Else
        JsonSkipValue jsonText, position
    End Select

    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, ",") Then
      ' continue
    ElseIf JsonTryConsumeChar(jsonText, position, "}") Then
      Exit Do
    Else
      JsonRaise "Expected ',' or '}' in root object."
    End If
  Loop

  JsonSkipWhitespace jsonText, position
  If position <= Len(jsonText) Then
    JsonRaise "Unexpected trailing characters."
  End If

  If formatVersion <> "0.1" Then
    JsonRaise "Unsupported formatVersion: " & formatVersion
  End If

  mapObj.Add "formatVersion", formatVersion
  mapObj.Add "components", components

  Set ParseCoverageMapJson = mapObj
End Function

Function ParseHitsText(hitsText)
  Dim dict, normalized, lines, i, line, eqIndex, pointId, countText, countValue
  Set dict = CreateObject("Scripting.Dictionary")
  normalized = Replace(Replace(CStr(hitsText), vbCrLf, vbLf), vbCr, vbLf)
  lines = Split(normalized, vbLf)

  For i = LBound(lines) To UBound(lines)
    line = Trim(lines(i))
    If line <> "" Then
      eqIndex = InStr(1, line, "=", vbBinaryCompare)
      If eqIndex > 1 Then
        pointId = Trim(Left(line, eqIndex - 1))
        countText = Trim(Mid(line, eqIndex + 1))
        If IsNumeric(countText) Then
          countValue = CLng(countText)
          If dict.Exists(pointId) Then
            dict(pointId) = CLng(dict(pointId)) + countValue
          Else
            dict.Add pointId, countValue
          End If
        End If
      End If
    End If
  Next

  Set ParseHitsText = dict
End Function

Function FormatTestResults(resultLines, totalCount, passCount, failCount)
  Dim sb, i
  sb = "WSCOV Test Results" & vbCrLf
  sb = sb & "==================" & vbCrLf

  If IsList(resultLines) Then
    For i = 0 To ListCount(resultLines) - 1
      sb = sb & CStr(resultLines(CStr(i))) & vbCrLf
    Next
  End If

  sb = sb & vbCrLf
  sb = sb & "TOTAL: " & CStr(totalCount) & vbCrLf
  sb = sb & "PASS : " & CStr(passCount) & vbCrLf
  sb = sb & "FAIL : " & CStr(failCount) & vbCrLf

  FormatTestResults = sb
End Function

Function FormatCoverageSummary(mapObj, hitsDict)
  Dim components, sb, i
  Dim totalCovered, totalPoints
  totalCovered = 0
  totalPoints = 0

  Set components = mapObj("components")
  sb = "WSCOV Coverage Summary" & vbCrLf
  sb = sb & "======================" & vbCrLf

  For i = 0 To ListCount(components) - 1
    Dim component, points, covered, pointCount, j, pointId
    Set component = components(CStr(i))
    Set points = component("points")
    covered = 0
    pointCount = ListCount(points)

    For j = 0 To ListCount(points) - 1
      pointId = CStr(points(CStr(j))("id"))
      If hitsDict.Exists(pointId) Then
        If CLng(hitsDict(pointId)) > 0 Then
          covered = covered + 1
        End If
      End If
    Next

    totalCovered = totalCovered + covered
    totalPoints = totalPoints + pointCount

    sb = sb & component("componentId") _
      & " [scriptIndex=" & CStr(component("scriptIndex")) & "] " _
      & "covered=" & CStr(covered) _
      & " total=" & CStr(pointCount) _
      & " rate=" & FormatPercent2(covered, pointCount) & vbCrLf
  Next

  sb = sb & vbCrLf
  sb = sb & "TOTAL covered=" & CStr(totalCovered) _
    & " total=" & CStr(totalPoints) _
    & " rate=" & FormatPercent2(totalCovered, totalPoints) & vbCrLf

  FormatCoverageSummary = sb
End Function

Function MapErrToExitCode(errNumber, fallbackCode)
  Select Case errNumber
    Case WSCOV_ERR_USAGE
      MapErrToExitCode = WSCOV_EXIT_USAGE
    Case WSCOV_ERR_PARSE
      MapErrToExitCode = WSCOV_EXIT_PARSE
    Case WSCOV_ERR_INSTRUMENT
      MapErrToExitCode = WSCOV_EXIT_INSTRUMENT
    Case WSCOV_ERR_RUNTIME, WSCOV_ERR_JSON
      MapErrToExitCode = WSCOV_EXIT_RUNTIME
    Case Else
      MapErrToExitCode = fallbackCode
  End Select
End Function

Sub RaiseUsage(message)
  Err.Raise WSCOV_ERR_USAGE, "wscov", CStr(message)
End Sub

Sub RaiseParse(message)
  Err.Raise WSCOV_ERR_PARSE, "wscov", CStr(message)
End Sub

Sub RaiseInstrument(message)
  Err.Raise WSCOV_ERR_INSTRUMENT, "wscov", CStr(message)
End Sub

Sub RaiseRuntime(message)
  Err.Raise WSCOV_ERR_RUNTIME, "wscov", CStr(message)
End Sub

Private Sub JsonRaise(message)
  Err.Raise WSCOV_ERR_JSON, "wscov-json", CStr(message)
End Sub

Private Function JsonParseComponentsArray(jsonText, ByRef position)
  Dim items
  Set items = CreateList()

  JsonSkipWhitespace jsonText, position
  JsonExpectChar jsonText, position, "["
  JsonSkipWhitespace jsonText, position

  If JsonTryConsumeChar(jsonText, position, "]") Then
    Set JsonParseComponentsArray = items
    Exit Function
  End If

  Do
    Dim component
    Set component = JsonParseComponentObject(jsonText, position)
    ListAdd items, component

    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, ",") Then
      ' continue
    ElseIf JsonTryConsumeChar(jsonText, position, "]") Then
      Exit Do
    Else
      JsonRaise "Expected ',' or ']' in components array."
    End If
  Loop

  Set JsonParseComponentsArray = items
End Function

Private Function JsonParseComponentObject(jsonText, ByRef position)
  Dim component, componentId, scriptIndex, points
  Set component = CreateObject("Scripting.Dictionary")
  Set points = CreateList()
  componentId = ""
  scriptIndex = 0

  JsonSkipWhitespace jsonText, position
  JsonExpectChar jsonText, position, "{"

  Do
    Dim key
    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, "}") Then
      Exit Do
    End If

    key = JsonReadString(jsonText, position)
    JsonSkipWhitespace jsonText, position
    JsonExpectChar jsonText, position, ":"

    Select Case key
      Case "componentId"
        componentId = JsonReadString(jsonText, position)
      Case "scriptIndex"
        scriptIndex = JsonReadLong(jsonText, position)
      Case "points"
        Set points = JsonParsePointsArray(jsonText, position)
      Case Else
        JsonSkipValue jsonText, position
    End Select

    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, ",") Then
      ' continue
    ElseIf JsonTryConsumeChar(jsonText, position, "}") Then
      Exit Do
    Else
      JsonRaise "Expected ',' or '}' in component object."
    End If
  Loop

  If componentId = "" Then
    JsonRaise "componentId is required."
  End If

  component.Add "componentId", componentId
  component.Add "scriptIndex", CLng(scriptIndex)
  component.Add "points", points

  Set JsonParseComponentObject = component
End Function

Private Function JsonParsePointsArray(jsonText, ByRef position)
  Dim items
  Set items = CreateList()

  JsonSkipWhitespace jsonText, position
  JsonExpectChar jsonText, position, "["
  JsonSkipWhitespace jsonText, position

  If JsonTryConsumeChar(jsonText, position, "]") Then
    Set JsonParsePointsArray = items
    Exit Function
  End If

  Do
    Dim point
    Set point = JsonParsePointObject(jsonText, position)
    ListAdd items, point

    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, ",") Then
      ' continue
    ElseIf JsonTryConsumeChar(jsonText, position, "]") Then
      Exit Do
    Else
      JsonRaise "Expected ',' or ']' in points array."
    End If
  Loop

  Set JsonParsePointsArray = items
End Function

Private Function JsonParsePointObject(jsonText, ByRef position)
  Dim point, pointId
  Set point = CreateObject("Scripting.Dictionary")
  pointId = ""

  JsonSkipWhitespace jsonText, position
  JsonExpectChar jsonText, position, "{"

  Do
    Dim key
    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, "}") Then
      Exit Do
    End If

    key = JsonReadString(jsonText, position)
    JsonSkipWhitespace jsonText, position
    JsonExpectChar jsonText, position, ":"

    Select Case key
      Case "id"
        pointId = JsonReadString(jsonText, position)
      Case Else
        JsonSkipValue jsonText, position
    End Select

    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, ",") Then
      ' continue
    ElseIf JsonTryConsumeChar(jsonText, position, "}") Then
      Exit Do
    Else
      JsonRaise "Expected ',' or '}' in point object."
    End If
  Loop

  If pointId = "" Then
    JsonRaise "point.id is required."
  End If

  point.Add "id", pointId
  Set JsonParsePointObject = point
End Function

Private Sub JsonSkipValue(jsonText, ByRef position)
  Dim ch, literal
  JsonSkipWhitespace jsonText, position
  If position > Len(jsonText) Then
    JsonRaise "Unexpected end while skipping value."
  End If

  ch = Mid(jsonText, position, 1)
  Select Case ch
    Case "{"
      JsonSkipObject jsonText, position
    Case "["
      JsonSkipArray jsonText, position
    Case """"
      literal = JsonReadString(jsonText, position)
    Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
      literal = JsonReadNumberToken(jsonText, position)
    Case "t", "f", "n", "T", "F", "N"
      literal = JsonReadLiteral(jsonText, position)
    Case Else
      JsonRaise "Unexpected character while skipping value: " & ch
  End Select
End Sub

Private Sub JsonSkipObject(jsonText, ByRef position)
  JsonExpectChar jsonText, position, "{"
  JsonSkipWhitespace jsonText, position

  If JsonTryConsumeChar(jsonText, position, "}") Then
    Exit Sub
  End If

  Do
    Dim key
    key = JsonReadString(jsonText, position)
    JsonSkipWhitespace jsonText, position
    JsonExpectChar jsonText, position, ":"
    JsonSkipValue jsonText, position
    JsonSkipWhitespace jsonText, position

    If JsonTryConsumeChar(jsonText, position, ",") Then
      ' continue
    ElseIf JsonTryConsumeChar(jsonText, position, "}") Then
      Exit Do
    Else
      JsonRaise "Expected ',' or '}' while skipping object."
    End If
  Loop
End Sub

Private Sub JsonSkipArray(jsonText, ByRef position)
  JsonExpectChar jsonText, position, "["
  JsonSkipWhitespace jsonText, position

  If JsonTryConsumeChar(jsonText, position, "]") Then
    Exit Sub
  End If

  Do
    JsonSkipValue jsonText, position
    JsonSkipWhitespace jsonText, position
    If JsonTryConsumeChar(jsonText, position, ",") Then
      ' continue
    ElseIf JsonTryConsumeChar(jsonText, position, "]") Then
      Exit Do
    Else
      JsonRaise "Expected ',' or ']' while skipping array."
    End If
  Loop
End Sub

Private Function JsonReadLong(jsonText, ByRef position)
  Dim token
  token = JsonReadNumberToken(jsonText, position)
  If InStr(token, ".") > 0 Or InStr(1, token, "e", vbTextCompare) > 0 Then
    JsonRaise "Expected integer number but got: " & token
  End If
  If Not IsNumeric(token) Then
    JsonRaise "Invalid number token: " & token
  End If
  JsonReadLong = CLng(token)
End Function

Private Function JsonReadNumberToken(jsonText, ByRef position)
  Dim startPos, ch
  JsonSkipWhitespace jsonText, position
  startPos = position

  Do While position <= Len(jsonText)
    ch = Mid(jsonText, position, 1)
    If InStr(1, "0123456789+-.eE", ch, vbBinaryCompare) > 0 Then
      position = position + 1
    Else
      Exit Do
    End If
  Loop

  JsonReadNumberToken = Mid(jsonText, startPos, position - startPos)
  If JsonReadNumberToken = "" Then
    JsonRaise "Expected number token."
  End If
End Function

Private Function JsonReadLiteral(jsonText, ByRef position)
  Dim startPos, ch, token
  JsonSkipWhitespace jsonText, position
  startPos = position

  Do While position <= Len(jsonText)
    ch = Mid(jsonText, position, 1)
    If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Then
      position = position + 1
    Else
      Exit Do
    End If
  Loop

  token = LCase(Mid(jsonText, startPos, position - startPos))
  Select Case token
    Case "true", "false", "null"
      JsonReadLiteral = token
    Case Else
      JsonRaise "Invalid literal token: " & token
  End Select
End Function

Private Function JsonReadString(jsonText, ByRef position)
  Dim sb, ch, esc, hex4
  JsonSkipWhitespace jsonText, position
  JsonExpectChar jsonText, position, """"
  sb = ""

  Do While position <= Len(jsonText)
    ch = Mid(jsonText, position, 1)
    position = position + 1

    If ch = """" Then
      JsonReadString = sb
      Exit Function
    End If

    If ch = "\" Then
      If position > Len(jsonText) Then
        JsonRaise "Invalid escape at end of string."
      End If
      esc = Mid(jsonText, position, 1)
      position = position + 1
      Select Case esc
        Case """"
          sb = sb & """"
        Case "\"
          sb = sb & "\"
        Case "/"
          sb = sb & "/"
        Case "b"
          sb = sb & vbBack
        Case "f"
          sb = sb & vbFormFeed
        Case "n"
          sb = sb & vbLf
        Case "r"
          sb = sb & vbCr
        Case "t"
          sb = sb & vbTab
        Case "u", "U"
          If position + 3 > Len(jsonText) Then
            JsonRaise "Invalid unicode escape length."
          End If
          hex4 = Mid(jsonText, position, 4)
          If Not JsonIsHex4(hex4) Then
            JsonRaise "Invalid unicode escape value: " & hex4
          End If
          sb = sb & ChrW(CLng("&H" & hex4))
          position = position + 4
        Case Else
          JsonRaise "Invalid escape sequence: \" & esc
      End Select
    Else
      sb = sb & ch
    End If
  Loop

  JsonRaise "Unterminated JSON string."
End Function

Private Function JsonTryConsumeChar(jsonText, ByRef position, expectedChar)
  JsonSkipWhitespace jsonText, position
  If position <= Len(jsonText) Then
    If Mid(jsonText, position, 1) = expectedChar Then
      position = position + 1
      JsonTryConsumeChar = True
      Exit Function
    End If
  End If
  JsonTryConsumeChar = False
End Function

Private Sub JsonExpectChar(jsonText, ByRef position, expectedChar)
  JsonSkipWhitespace jsonText, position
  If position > Len(jsonText) Then
    JsonRaise "Expected '" & expectedChar & "' but reached end of input."
  End If
  If Mid(jsonText, position, 1) <> expectedChar Then
    JsonRaise "Expected '" & expectedChar & "' but got '" & Mid(jsonText, position, 1) & "'."
  End If
  position = position + 1
End Sub

Private Sub JsonSkipWhitespace(jsonText, ByRef position)
  Dim ch
  Do While position <= Len(jsonText)
    ch = Mid(jsonText, position, 1)
    If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
      position = position + 1
    Else
      Exit Do
    End If
  Loop
End Sub

Private Function JsonIsHex4(value)
  Dim i, ch
  If Len(value) <> 4 Then
    JsonIsHex4 = False
    Exit Function
  End If

  For i = 1 To 4
    ch = Mid(value, i, 1)
    If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then
      JsonIsHex4 = False
      Exit Function
    End If
  Next

  JsonIsHex4 = True
End Function

Private Function FormatIsoTimestamp(dateValue)
  FormatIsoTimestamp = Right("0000" & CStr(Year(dateValue)), 4) & "-" _
    & Right("00" & CStr(Month(dateValue)), 2) & "-" _
    & Right("00" & CStr(Day(dateValue)), 2) & "T" _
    & Right("00" & CStr(Hour(dateValue)), 2) & ":" _
    & Right("00" & CStr(Minute(dateValue)), 2) & ":" _
    & Right("00" & CStr(Second(dateValue)), 2)
End Function

Private Function FormatPercent2(covered, total)
  Dim scaled, intPart, fracPart
  If CLng(total) <= 0 Then
    FormatPercent2 = "0.00%"
    Exit Function
  End If

  scaled = Int((CDbl(covered) / CDbl(total)) * 10000 + 0.5)
  intPart = scaled \ 100
  fracPart = scaled Mod 100
  FormatPercent2 = CStr(intPart) & "." & Right("00" & CStr(fracPart), 2) & "%"
End Function
