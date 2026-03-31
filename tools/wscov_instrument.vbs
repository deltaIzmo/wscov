Option Explicit

LoadLibrary
WScript.Quit Main()

Function Main()
  Dim exitCode
  exitCode = WSCOV_EXIT_OK

  On Error Resume Next
  RunInstrument
  If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description
    exitCode = MapErrToExitCode(Err.Number, WSCOV_EXIT_INSTRUMENT)
    Err.Clear
  End If
  On Error GoTo 0

  Main = exitCode
End Function

Sub RunInstrument()
  Dim parsed, inputPath, outputPath, mapPath
  Dim fso, outputFolder, mapFolder
  Dim dom, scriptNodes, scriptNode, componentNode
  Dim language, sourceScript, instrumentedScript
  Dim processedCount, mapComponents
  Dim componentCounters, componentId, scriptIndex
  Dim points, mapJson

  Set parsed = ParseArgsInstrument(WScript.Arguments)
  inputPath = parsed("inputWscPath")
  outputPath = parsed("outputWscPath")
  mapPath = parsed("coverageMapPath")

  EnsureFileExists inputPath

  Set fso = CreateObject("Scripting.FileSystemObject")
  outputFolder = fso.GetParentFolderName(outputPath)
  mapFolder = fso.GetParentFolderName(mapPath)
  If outputFolder <> "" Then
    EnsureDirExists outputFolder
  End If
  If mapFolder <> "" Then
    EnsureDirExists mapFolder
  End If

  Set dom = CreateConfiguredDom()
  If dom Is Nothing Then
    RaiseParse "Failed to create MSXML DOM."
  End If

  If Not dom.Load(inputPath) Then
    RaiseParse "Failed to parse input WSC: " & dom.parseError.reason
  End If

  On Error Resume Next
  dom.setProperty "SelectionLanguage", "XPath"
  Err.Clear
  On Error GoTo 0

  Set scriptNodes = dom.selectNodes("//component/script")
  If scriptNodes Is Nothing Then
    RaiseInstrument "No <script> nodes found."
  End If

  Set mapComponents = CreateList()
  Set componentCounters = CreateObject("Scripting.Dictionary")
  processedCount = 0

  For Each scriptNode In scriptNodes
    language = LCase(Trim(CStr(scriptNode.getAttribute("language"))))
    If language = "vbscript" Then
      processedCount = processedCount + 1
      sourceScript = CStr(scriptNode.text)

      If InStr(1, sourceScript, "__WSCOV_RUNTIME__ BEGIN", vbTextCompare) > 0 Then
        RaiseInstrument "Already instrumented script found."
      End If

      Set componentNode = scriptNode.parentNode
      componentId = ResolveComponentId(componentNode, processedCount)
      scriptIndex = NextScriptIndex(componentCounters, componentId)

      Set points = CreateList()
      instrumentedScript = InstrumentScriptBody(sourceScript, points)
      instrumentedScript = InjectRuntime(instrumentedScript)
      scriptNode.text = instrumentedScript

      EnsureCoveragePublicMethods dom, componentNode
      ListAdd mapComponents, CreateMapComponent(componentId, CLng(scriptIndex), points)
    End If
  Next

  If processedCount = 0 Then
    RaiseInstrument "No VBScript <script language=""VBScript""> blocks found."
  End If

  On Error Resume Next
  dom.Save outputPath
  If Err.Number <> 0 Then
    Dim saveMessage
    saveMessage = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseInstrument "Failed to save instrumented WSC: " & saveMessage
  End If
  On Error GoTo 0

  mapJson = BuildCoverageMapJson(inputPath, outputPath, mapComponents)
  WriteTextFileUtf8 mapPath, mapJson

  WScript.Echo "OK instrumented: " & outputPath
  WScript.Echo "OK coverage map: " & mapPath
End Sub

Function CreateConfiguredDom()
  Dim dom
  Set dom = Nothing

  On Error Resume Next
  Set dom = CreateObject("MSXML2.DOMDocument.6.0")
  If (Err.Number <> 0) Or (dom Is Nothing) Then
    Err.Clear
    Set dom = CreateObject("Microsoft.XMLDOM")
  End If
  On Error GoTo 0

  If dom Is Nothing Then
    Set CreateConfiguredDom = Nothing
    Exit Function
  End If

  dom.async = False
  dom.validateOnParse = False
  dom.resolveExternals = False
  dom.preserveWhiteSpace = True

  Set CreateConfiguredDom = dom
End Function

Function ResolveComponentId(componentNode, ordinal)
  Dim componentId, registrationNode
  componentId = Trim(CStr(componentNode.getAttribute("id")))
  If componentId <> "" Then
    ResolveComponentId = componentId
    Exit Function
  End If

  Set registrationNode = componentNode.selectSingleNode("registration")
  If Not registrationNode Is Nothing Then
    componentId = Trim(CStr(registrationNode.getAttribute("progid")))
  End If

  If componentId = "" Then
    componentId = "component" & CStr(ordinal)
  End If

  ResolveComponentId = componentId
End Function

Function NextScriptIndex(counterByComponent, componentId)
  If counterByComponent.Exists(componentId) Then
    counterByComponent(componentId) = CLng(counterByComponent(componentId)) + 1
  Else
    counterByComponent.Add componentId, 0
  End If
  NextScriptIndex = CLng(counterByComponent(componentId))
End Function

Sub EnsureCoveragePublicMethods(dom, componentNode)
  Dim publicNode, firstScript
  Set publicNode = componentNode.selectSingleNode("public")

  If publicNode Is Nothing Then
    Set publicNode = dom.createElement("public")
    Set firstScript = componentNode.selectSingleNode("script")
    If firstScript Is Nothing Then
      componentNode.appendChild publicNode
    Else
      componentNode.insertBefore publicNode, firstScript
    End If
  End If

  EnsurePublicMethod dom, publicNode, "WscovDumpCoverage", "WscovDumpInternal"
  EnsurePublicMethod dom, publicNode, "WscovResetCoverage", "WscovResetInternal"
End Sub

Sub EnsurePublicMethod(dom, publicNode, methodName, internalName)
  Dim child, existingName
  For Each child In publicNode.childNodes
    If LCase(CStr(child.nodeName)) = "method" Then
      existingName = LCase(Trim(CStr(child.getAttribute("name"))))
      If existingName = LCase(methodName) Then
        If Trim(CStr(child.getAttribute("internalName"))) = "" Then
          child.setAttribute "internalName", internalName
        End If
        Exit Sub
      End If
    End If
  Next

  Dim methodNode
  Set methodNode = dom.createElement("method")
  methodNode.setAttribute "name", methodName
  methodNode.setAttribute "internalName", internalName
  publicNode.appendChild methodNode
End Sub

Function CreateMapComponent(componentId, scriptIndex, points)
  Dim item
  Set item = CreateObject("Scripting.Dictionary")
  item.Add "componentId", componentId
  item.Add "language", "VBScript"
  item.Add "scriptIndex", CLng(scriptIndex)
  item.Add "points", points
  Set CreateMapComponent = item
End Function

Function InstrumentScriptBody(scriptText, ByRef points)
  Dim normalized, lines
  Dim outputLines
  Dim i, line, procName, foundProcName
  Dim insideProc, inSignatureContinuation
  Dim inStatementContinuation, continuationPoint
  Dim pointCounter
  Dim pointId, hitLine, indent, point

  normalized = NormalizeNewlines(scriptText)
  lines = Split(normalized, vbLf)
  Set outputLines = CreateList()
  Set points = CreateList()

  insideProc = False
  inSignatureContinuation = False
  inStatementContinuation = False
  Set continuationPoint = Nothing
  procName = ""
  pointCounter = 1

  For i = 0 To UBound(lines)
    line = lines(i)

    If Not insideProc Then
      foundProcName = ProcedureStartName(line)
      If foundProcName <> "" Then
        insideProc = True
        inSignatureContinuation = EndsWithContinuation(line)
        procName = foundProcName
      End If
      ListAdd outputLines, line
    Else
      If inSignatureContinuation Then
        ListAdd outputLines, line
        If Not EndsWithContinuation(line) Then
          inSignatureContinuation = False
        End If
      ElseIf IsProcedureEndLine(line) Then
        ListAdd outputLines, line
        insideProc = False
        inStatementContinuation = False
        Set continuationPoint = Nothing
        procName = ""
      ElseIf inStatementContinuation Then
        ListAdd outputLines, line
        continuationPoint("origScriptLineEnd") = CLng(i + 1)
        If Not EndsWithContinuation(line) Then
          inStatementContinuation = False
          Set continuationPoint = Nothing
        End If
      Else
        If ShouldInsertHitBeforeLine(line) Then
          pointId = FormatPointId(pointCounter)
          pointCounter = pointCounter + 1
          indent = LeadingWhitespace(line)
          hitLine = indent & "WscovHitPoint """ & pointId & """"
          ListAdd outputLines, hitLine

          Set point = CreateCoveragePoint(pointId, procName, CLng(i + 1), CLng(i + 1), Trim(StripInlineComment(line)))
          ListAdd points, point

          If EndsWithContinuation(line) Then
            inStatementContinuation = True
            Set continuationPoint = point
          End If
        End If

        ListAdd outputLines, line
      End If
    End If
  Next

  InstrumentScriptBody = JoinLines(outputLines)
End Function

Function InjectRuntime(scriptText)
  Dim runtimeLines, originalLines, outputLines
  Dim insertAt, i
  Set runtimeLines = RuntimeSnippetLines()

  originalLines = Split(NormalizeNewlines(scriptText), vbLf)
  insertAt = FindRuntimeInsertIndex(originalLines)

  Set outputLines = CreateList()
  For i = 0 To UBound(originalLines)
    If i = insertAt Then
      AppendCollection outputLines, runtimeLines
    End If
    ListAdd outputLines, originalLines(i)
  Next

  If insertAt >= UBound(originalLines) + 1 Then
    AppendCollection outputLines, runtimeLines
  End If

  InjectRuntime = JoinLines(outputLines)
End Function

Function RuntimeSnippetLines()
  Dim lines
  Set lines = CreateList()
  ListAdd lines, "' === __WSCOV_RUNTIME__ BEGIN ==="
  ListAdd lines, "' NOTE: inserted by wscov_instrument.vbs (formatVersion=0.1)"
  ListAdd lines, "Dim wscovDict : Set wscovDict = CreateObject(""Scripting.Dictionary"")"
  ListAdd lines, ""
  ListAdd lines, "Sub WscovHitPoint(id)"
  ListAdd lines, "  Dim k : k = CStr(id)"
  ListAdd lines, "  If wscovDict.Exists(k) Then"
  ListAdd lines, "    wscovDict(k) = CLng(wscovDict(k)) + 1"
  ListAdd lines, "  Else"
  ListAdd lines, "    wscovDict.Add k, 1"
  ListAdd lines, "  End If"
  ListAdd lines, "End Sub"
  ListAdd lines, ""
  ListAdd lines, "Function WscovDumpInternal()"
  ListAdd lines, "  Dim k, out : out = """""
  ListAdd lines, "  For Each k In wscovDict.Keys"
  ListAdd lines, "    out = out & k & ""="" & wscovDict(k) & vbLf"
  ListAdd lines, "  Next"
  ListAdd lines, "  WscovDumpInternal = out"
  ListAdd lines, "End Function"
  ListAdd lines, ""
  ListAdd lines, "Sub WscovResetInternal()"
  ListAdd lines, "  wscovDict.RemoveAll"
  ListAdd lines, "End Sub"
  ListAdd lines, "' === __WSCOV_RUNTIME__ END ==="
  Set RuntimeSnippetLines = lines
End Function

Function FindRuntimeInsertIndex(lines)
  Dim i, code
  FindRuntimeInsertIndex = 0

  For i = 0 To UBound(lines)
    code = Trim(StripInlineComment(lines(i)))
    If code <> "" Then
      If LCase(code) = "option explicit" Then
        FindRuntimeInsertIndex = i + 1
      Else
        FindRuntimeInsertIndex = 0
      End If
      Exit Function
    End If
  Next

  FindRuntimeInsertIndex = 0
End Function

Function ShouldInsertHitBeforeLine(line)
  Dim code
  code = Trim(StripInlineComment(line))
  If code = "" Then
    ShouldInsertHitBeforeLine = False
    Exit Function
  End If

  If IsCommentOnlyLine(line) Then
    ShouldInsertHitBeforeLine = False
    Exit Function
  End If

  If IsStructuralKeywordLine(code) Then
    ShouldInsertHitBeforeLine = False
    Exit Function
  End If

  If ProcedureStartName(code) <> "" Then
    ShouldInsertHitBeforeLine = False
    Exit Function
  End If

  If IsLabelOnlyLine(code) Then
    ShouldInsertHitBeforeLine = False
    Exit Function
  End If

  ShouldInsertHitBeforeLine = True
End Function

Function IsLabelOnlyLine(code)
  Dim value
  value = Trim(code)
  If value = "" Then
    IsLabelOnlyLine = False
    Exit Function
  End If

  If Right(value, 1) = ":" Then
    If InStr(1, value, " ", vbBinaryCompare) = 0 And InStr(1, value, vbTab, vbBinaryCompare) = 0 Then
      IsLabelOnlyLine = True
      Exit Function
    End If
  End If

  IsLabelOnlyLine = False
End Function

Function IsStructuralKeywordLine(code)
  Dim lower
  lower = LCase(Trim(code))

  If lower = "else" Or Left(lower, 5) = "else " Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "elseif" Or Left(lower, 7) = "elseif " Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "end if" Or lower = "endif" Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "case" Or Left(lower, 5) = "case " Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "end select" Or lower = "endselect" Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "next" Or Left(lower, 5) = "next " Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "loop" Or Left(lower, 5) = "loop " Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "wend" Then
    IsStructuralKeywordLine = True
    Exit Function
  End If
  If lower = "end sub" Or lower = "end function" Or lower = "end property" Or lower = "end class" Then
    IsStructuralKeywordLine = True
    Exit Function
  End If

  IsStructuralKeywordLine = False
End Function

Function ProcedureStartName(line)
  Dim text, lower, afterToken
  text = LTrim(line)
  lower = LCase(text)

  Do
    If Left(lower, 7) = "public " Then
      text = Mid(text, 8)
    ElseIf Left(lower, 8) = "private " Then
      text = Mid(text, 9)
    ElseIf Left(lower, 7) = "friend " Then
      text = Mid(text, 8)
    ElseIf Left(lower, 8) = "default " Then
      text = Mid(text, 9)
    Else
      Exit Do
    End If
    text = LTrim(text)
    lower = LCase(text)
  Loop

  If Left(lower, 4) = "sub " Then
    afterToken = Mid(text, 5)
    ProcedureStartName = ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 9) = "function " Then
    afterToken = Mid(text, 10)
    ProcedureStartName = ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 13) = "property get " Then
    afterToken = Mid(text, 14)
    ProcedureStartName = ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 13) = "property let " Then
    afterToken = Mid(text, 14)
    ProcedureStartName = ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 13) = "property set " Then
    afterToken = Mid(text, 14)
    ProcedureStartName = ExtractIdentifier(afterToken)
    Exit Function
  End If

  ProcedureStartName = ""
End Function

Function ExtractIdentifier(fragment)
  Dim text, i, ch
  text = LTrim(fragment)
  If text = "" Then
    ExtractIdentifier = ""
    Exit Function
  End If

  ch = Mid(text, 1, 1)
  If Not IsIdentifierStartChar(ch) Then
    ExtractIdentifier = ""
    Exit Function
  End If

  i = 1
  Do While i <= Len(text)
    ch = Mid(text, i, 1)
    If IsIdentifierChar(ch) Then
      i = i + 1
    Else
      Exit Do
    End If
  Loop

  ExtractIdentifier = Left(text, i - 1)
End Function

Function IsIdentifierStartChar(ch)
  Dim code
  code = AscW(ch)
  IsIdentifierStartChar = ((code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "_")
End Function

Function IsIdentifierChar(ch)
  Dim code
  code = AscW(ch)
  IsIdentifierChar = ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "_")
End Function

Function IsProcedureEndLine(line)
  Dim code
  code = LCase(Trim(StripInlineComment(line)))
  If code = "end sub" Or code = "end function" Or code = "end property" Then
    IsProcedureEndLine = True
  Else
    IsProcedureEndLine = False
  End If
End Function

Function EndsWithContinuation(line)
  Dim code, prevChar
  code = RTrim(StripInlineComment(line))
  If Len(code) < 2 Then
    EndsWithContinuation = False
    Exit Function
  End If

  If Right(code, 1) <> "_" Then
    EndsWithContinuation = False
    Exit Function
  End If

  prevChar = Mid(code, Len(code) - 1, 1)
  If prevChar = " " Or prevChar = vbTab Then
    EndsWithContinuation = True
  Else
    EndsWithContinuation = False
  End If
End Function

Function IsCommentOnlyLine(line)
  Dim text, lower
  text = LTrim(line)
  If text = "" Then
    IsCommentOnlyLine = False
    Exit Function
  End If

  If Left(text, 1) = "'" Then
    IsCommentOnlyLine = True
    Exit Function
  End If

  lower = LCase(text)
  If lower = "rem" Or Left(lower, 4) = "rem " Then
    IsCommentOnlyLine = True
    Exit Function
  End If

  IsCommentOnlyLine = False
End Function

Function StripInlineComment(line)
  Dim i, ch, nextCh, inString, out
  i = 1
  inString = False
  out = ""

  Do While i <= Len(line)
    ch = Mid(line, i, 1)

    If ch = """" Then
      If inString Then
        If i < Len(line) Then
          nextCh = Mid(line, i + 1, 1)
          If nextCh = """" Then
            out = out & """"""
            i = i + 2
          Else
            inString = False
            out = out & ch
            i = i + 1
          End If
        Else
          inString = False
          out = out & ch
          i = i + 1
        End If
      Else
        inString = True
        out = out & ch
        i = i + 1
      End If
    ElseIf ch = "'" And Not inString Then
      Exit Do
    Else
      out = out & ch
      i = i + 1
    End If
  Loop

  StripInlineComment = out
End Function

Function LeadingWhitespace(line)
  Dim i, ch, out
  out = ""
  For i = 1 To Len(line)
    ch = Mid(line, i, 1)
    If ch = " " Or ch = vbTab Then
      out = out & ch
    Else
      Exit For
    End If
  Next
  LeadingWhitespace = out
End Function

Function CreateCoveragePoint(pointId, procName, lineStart, lineEnd, snippet)
  Dim point
  Set point = CreateObject("Scripting.Dictionary")
  point.Add "id", pointId
  point.Add "kind", "line"
  point.Add "proc", procName
  point.Add "origScriptLine", CLng(lineStart)
  point.Add "origScriptLineEnd", CLng(lineEnd)
  point.Add "snippet", snippet
  Set CreateCoveragePoint = point
End Function

Function FormatPointId(pointNumber)
  FormatPointId = "L" & Right("000000" & CStr(pointNumber), 6)
End Function

Sub AppendCollection(target, source)
  Dim i
  For i = 0 To ListCount(source) - 1
    ListAdd target, source(CStr(i))
  Next
End Sub

Function JoinLines(lines)
  Dim i, text
  text = ""
  For i = 0 To ListCount(lines) - 1
    If i > 0 Then
      text = text & vbCrLf
    End If
    text = text & CStr(lines(CStr(i)))
  Next
  JoinLines = text
End Function

Function NormalizeNewlines(text)
  NormalizeNewlines = Replace(Replace(CStr(text), vbCrLf, vbLf), vbCr, vbLf)
End Function

Sub LoadLibrary()
  Dim fso, libPath, ts, code
  Set fso = CreateObject("Scripting.FileSystemObject")
  libPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "wscov_lib.vbs")

  If Not fso.FileExists(libPath) Then
    WScript.Echo "ERROR: Shared library not found: " & libPath
    WScript.Quit 5
  End If

  On Error Resume Next
  Set ts = fso.OpenTextFile(libPath, 1, False)
  code = ts.ReadAll
  ts.Close
  ExecuteGlobal code
  If Err.Number <> 0 Then
    Dim message
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "ERROR: Failed to load shared library: " & message
    WScript.Quit 5
  End If
  On Error GoTo 0
End Sub
