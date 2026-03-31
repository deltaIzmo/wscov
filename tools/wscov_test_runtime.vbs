' Shared test runtime loaded with ExecuteGlobal.
' Keep this file free of Option Explicit so it can be loaded dynamically.

Sub InitializeWscovTestRuntime()
  Set g_tests = CreateList()
  Set g_setupProcs = CreateList()
  Set g_teardownProcs = CreateList()
  Set g_caseDefinitions = CreateObject("Scripting.Dictionary")
End Sub

Function DiscoverTestFiles(testsDir)
  Dim files
  Set files = CreateList()
  CollectTestFilesRecursive testsDir, files
  DiscoverTestFiles = WscovRuntime_SortListToArray(files)
End Function

Sub CollectTestFilesRecursive(folderPath, files)
  Dim fso, folder, subFolder, file
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set folder = fso.GetFolder(folderPath)

  For Each file In folder.Files
    If Right(LCase(file.Name), 9) = ".test.vbs" Then
      ListAdd files, file.Path
    End If
  Next

  For Each subFolder In folder.SubFolders
    CollectTestFilesRecursive subFolder.Path, files
  Next
End Sub

Sub LoadAndExecuteTestFile(testFilePath)
  Dim code
  code = ReadTextFile(testFilePath)
  WscovRuntime_ScanTestCaseDefinitions code

  On Error Resume Next
  ExecuteGlobal code
  If Err.Number <> 0 Then
    Dim message
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseRuntime "Failed to execute test file: " & testFilePath & " (" & message & ")"
  End If
  On Error GoTo 0
End Sub

Sub ExecuteRegisteredTests(mapObj, ByRef resultLines, ByRef coverageLines, ByRef totalCount, ByRef passCount, ByRef failCount)
  Dim i, testEntry
  Dim beforeText, afterText, beforeHits, afterHits, deltaHits, coverageInfo
  Dim failed, reason, testName

  totalCount = ListCount(g_tests)
  passCount = 0
  failCount = 0

  ListAdd coverageLines, "WSCOV Per-Test Coverage"
  ListAdd coverageLines, "======================="

  For i = 0 To ListCount(g_tests) - 1
    Set testEntry = g_tests(CStr(i))
    testName = CStr(testEntry("name"))

    beforeText = DumpCoverageFromSut()
    Set beforeHits = ParseHitsText(beforeText)

    failed = False
    reason = ""
    WscovRuntime_ExecuteSingleTest testEntry, failed, reason

    afterText = DumpCoverageFromSut()
    Set afterHits = ParseHitsText(afterText)
    Set deltaHits = WscovRuntime_DiffHits(beforeHits, afterHits)
    Set coverageInfo = WscovRuntime_BuildCoverageSnapshot(mapObj, deltaHits)

    If failed Then
      failCount = failCount + 1
      ListAdd resultLines, "FAIL " & testName & " - " & reason & " - coverage " & WscovRuntime_FormatCoverageBrief(coverageInfo)
    Else
      passCount = passCount + 1
      ListAdd resultLines, "PASS " & testName & " - coverage " & WscovRuntime_FormatCoverageBrief(coverageInfo)
    End If

    WscovRuntime_AppendCoverageBlock coverageLines, testName, failed, reason, coverageInfo
  Next
End Sub

Sub Wscov_AddTest(name, procRef)
  Dim testEntry

  If Trim(CStr(name)) = "" Then
    Err.Raise 5700, "Wscov_AddTest", "Test name is required."
  End If

  Set testEntry = CreateObject("Scripting.Dictionary")
  testEntry.Add "kind", "proc"
  testEntry.Add "name", CStr(name)
  testEntry.Add "procRef", procRef
  ListAdd g_tests, testEntry
End Sub

Sub Wscov_AddTestMethod(className, methodName)
  WscovRuntime_AddTestMethodCore className, methodName, className & "#" & methodName
End Sub

Sub Wscov_AddNamedTestMethod(className, methodName, displayName)
  WscovRuntime_AddTestMethodCore className, methodName, displayName
End Sub

Sub WscovRuntime_AddTestMethodCore(className, methodName, displayName)
  Dim testEntry, resolvedName

  className = Trim(CStr(className))
  methodName = Trim(CStr(methodName))
  If className = "" Then
    Err.Raise 5701, "Wscov_AddTestMethod", "Test class name is required."
  End If
  If methodName = "" Then
    Err.Raise 5702, "Wscov_AddTestMethod", "Test method name is required."
  End If
  If Not WscovRuntime_IsValidIdentifier(className) Then
    Err.Raise 5703, "Wscov_AddTestMethod", "Invalid class name: " & className
  End If
  If Not WscovRuntime_IsValidIdentifier(methodName) Then
    Err.Raise 5704, "Wscov_AddTestMethod", "Invalid method name: " & methodName
  End If

  resolvedName = className & "#" & methodName
  If Trim(CStr(displayName)) <> "" Then
    resolvedName = CStr(displayName)
  End If

  Set testEntry = CreateObject("Scripting.Dictionary")
  testEntry.Add "kind", "case_method"
  testEntry.Add "name", resolvedName
  testEntry.Add "className", className
  testEntry.Add "methodName", methodName
  ListAdd g_tests, testEntry
End Sub

Sub Wscov_RegisterTestCase(className)
  WscovRuntime_RegisterTestCaseCore className, Empty, False
End Sub

Sub Wscov_RegisterTestCaseMethods(className, methods)
  WscovRuntime_RegisterTestCaseCore className, methods, True
End Sub

Sub WscovRuntime_RegisterTestCaseCore(className, methods, useExplicitMethods)
  Dim methodNames, i, methodName

  className = Trim(CStr(className))
  If className = "" Then
    Err.Raise 5705, "Wscov_RegisterTestCase", "Test class name is required."
  End If
  If Not WscovRuntime_IsValidIdentifier(className) Then
    Err.Raise 5706, "Wscov_RegisterTestCase", "Invalid class name: " & className
  End If

  If CBool(useExplicitMethods) Then
    methodNames = WscovRuntime_ResolveTestMethods(className, methods)
  Else
    methodNames = WscovRuntime_DiscoverTestMethods(className)
  End If
  If WscovRuntime_ArrayCount(methodNames) = 0 Then
    Err.Raise 5707, "Wscov_RegisterTestCase", "No test methods found for class: " & className
  End If

  For i = 0 To UBound(methodNames)
    methodName = CStr(methodNames(i))
    Wscov_AddNamedTestMethod className, methodName, className & "#" & methodName
  Next
End Sub

Sub Wscov_AddSetup(procRef)
  ListAdd g_setupProcs, procRef
End Sub

Sub Wscov_AddTeardown(procRef)
  ListAdd g_teardownProcs, procRef
End Sub

Sub Assert(condition)
  If WscovRuntime_IsTruthy(condition) Then
    Exit Sub
  End If
  Fail "assert failed."
End Sub

Sub AssertTrue(condition)
  Assert condition
End Sub

Sub Refute(condition)
  If Not WscovRuntime_IsTruthy(condition) Then
    Exit Sub
  End If
  Fail "refute failed."
End Sub

Sub AssertEqual(expected, actual)
  If WscovRuntime_ValuesAreEqual(expected, actual) Then
    Exit Sub
  End If
  Fail "Expected " & WscovRuntime_DescribeValue(expected) & " but got " & WscovRuntime_DescribeValue(actual) & "."
End Sub

Sub assert_equal(expected, actual)
  AssertEqual expected, actual
End Sub

Sub RefuteEqual(unexpected, actual)
  If Not WscovRuntime_ValuesAreEqual(unexpected, actual) Then
    Exit Sub
  End If
  Fail "Did not expect " & WscovRuntime_DescribeValue(actual) & "."
End Sub

Sub refute_equal(unexpected, actual)
  RefuteEqual unexpected, actual
End Sub

Sub AssertNil(value)
  If WscovRuntime_IsNil(value) Then
    Exit Sub
  End If
  Fail "Expected nil-like value but got " & WscovRuntime_DescribeValue(value) & "."
End Sub

Sub assert_nil(value)
  AssertNil value
End Sub

Sub RefuteNil(value)
  If Not WscovRuntime_IsNil(value) Then
    Exit Sub
  End If
  Fail "Expected non-nil value."
End Sub

Sub refute_nil(value)
  RefuteNil value
End Sub

Sub AssertEmpty(value)
  If WscovRuntime_IsEmptyValue(value) Then
    Exit Sub
  End If
  Fail "Expected empty value but got " & WscovRuntime_DescribeValue(value) & "."
End Sub

Sub assert_empty(value)
  AssertEmpty value
End Sub

Sub RefuteEmpty(value)
  If Not WscovRuntime_IsEmptyValue(value) Then
    Exit Sub
  End If
  Fail "Expected non-empty value."
End Sub

Sub refute_empty(value)
  RefuteEmpty value
End Sub

Sub AssertMatch(pattern, actual)
  If WscovRuntime_Matches(pattern, actual) Then
    Exit Sub
  End If
  Fail "Expected " & WscovRuntime_DescribeValue(actual) & " to match " & WscovRuntime_DescribeValue(pattern) & "."
End Sub

Sub assert_match(pattern, actual)
  AssertMatch pattern, actual
End Sub

Sub RefuteMatch(pattern, actual)
  If Not WscovRuntime_Matches(pattern, actual) Then
    Exit Sub
  End If
  Fail "Expected " & WscovRuntime_DescribeValue(actual) & " not to match " & WscovRuntime_DescribeValue(pattern) & "."
End Sub

Sub refute_match(pattern, actual)
  RefuteMatch pattern, actual
End Sub

Sub AssertIncludes(container, expected)
  If WscovRuntime_Includes(container, expected) Then
    Exit Sub
  End If
  Fail "Expected container to include " & WscovRuntime_DescribeValue(expected) & "."
End Sub

Sub assert_includes(container, expected)
  AssertIncludes container, expected
End Sub

Sub RefuteIncludes(container, expected)
  If Not WscovRuntime_Includes(container, expected) Then
    Exit Sub
  End If
  Fail "Expected container not to include " & WscovRuntime_DescribeValue(expected) & "."
End Sub

Sub refute_includes(container, expected)
  RefuteIncludes container, expected
End Sub

Sub AssertSame(expected, actual)
  If WscovRuntime_IsSameObject(expected, actual) Then
    Exit Sub
  End If
  Fail "Expected both objects to reference the same instance."
End Sub

Sub assert_same(expected, actual)
  AssertSame expected, actual
End Sub

Sub RefuteSame(expected, actual)
  If Not WscovRuntime_IsSameObject(expected, actual) Then
    Exit Sub
  End If
  Fail "Expected objects to be different instances."
End Sub

Sub refute_same(expected, actual)
  RefuteSame expected, actual
End Sub

Sub Flunk()
  Fail "Flunked."
End Sub

Sub Fail(message)
  Err.Raise 5710, "AssertionFailed", CStr(message)
End Sub

Sub WscovRuntime_ExecuteSingleTest(testEntry, ByRef failed, ByRef reason)
  Dim j, setupRef, teardownRef

  On Error Resume Next
  For j = 0 To ListCount(g_setupProcs) - 1
    Set setupRef = g_setupProcs(CStr(j))
    Call setupRef()
    If Err.Number <> 0 Then
      failed = True
      reason = "setup failed: " & Err.Description
      Err.Clear
      Exit For
    End If
  Next

  If Not failed Then
    Select Case CStr(testEntry("kind"))
      Case "proc"
        WscovRuntime_RunProcTest testEntry("procRef"), failed, reason
      Case "case_method"
        WscovRuntime_RunCaseMethodTest testEntry, failed, reason
      Case Else
        failed = True
        reason = "Unknown test kind: " & CStr(testEntry("kind"))
    End Select
  End If

  For j = 0 To ListCount(g_teardownProcs) - 1
    Set teardownRef = g_teardownProcs(CStr(j))
    Call teardownRef()
    If Err.Number <> 0 Then
      WscovRuntime_AppendFailureReason failed, reason, "teardown failed: " & Err.Description
      Err.Clear
    End If
  Next
  On Error GoTo 0
End Sub

Sub WscovRuntime_RunProcTest(procRef, ByRef failed, ByRef reason)
  On Error Resume Next
  Call procRef()
  If Err.Number <> 0 Then
    failed = True
    reason = Err.Description
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Sub WscovRuntime_RunCaseMethodTest(testEntry, ByRef failed, ByRef reason)
  Dim caseObj, className, methodName

  className = CStr(testEntry("className"))
  methodName = CStr(testEntry("methodName"))

  On Error Resume Next
  Set caseObj = WscovRuntime_CreateCaseInstance(className)
  If Err.Number <> 0 Then
    failed = True
    reason = "test case initialization failed: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0

  WscovRuntime_RunOptionalCaseMethod caseObj, className, "before_setup", "before_setup failed", failed, reason
  If Not failed Then
    WscovRuntime_RunOptionalCaseMethod caseObj, className, "setup", "setup failed", failed, reason
  End If
  If Not failed Then
    WscovRuntime_RunOptionalCaseMethod caseObj, className, "after_setup", "after_setup failed", failed, reason
  End If
  If Not failed Then
    WscovRuntime_RunRequiredCaseMethod caseObj, className, methodName, "test failed", failed, reason
  End If

  WscovRuntime_RunOptionalCaseMethod caseObj, className, "before_teardown", "before_teardown failed", failed, reason
  WscovRuntime_RunOptionalCaseMethod caseObj, className, "teardown", "teardown failed", failed, reason
  WscovRuntime_RunOptionalCaseMethod caseObj, className, "after_teardown", "after_teardown failed", failed, reason
End Sub

Sub WscovRuntime_RunOptionalCaseMethod(caseObj, className, methodName, label, ByRef failed, ByRef reason)
  If Not WscovRuntime_TestCaseMethodExists(className, methodName) Then
    Exit Sub
  End If
  WscovRuntime_RunRequiredCaseMethod caseObj, className, methodName, label, failed, reason
End Sub

Sub WscovRuntime_RunRequiredCaseMethod(caseObj, className, methodName, label, ByRef failed, ByRef reason)
  Dim actualMethodName

  actualMethodName = WscovRuntime_ResolveMethodName(className, methodName)
  On Error Resume Next
  WscovRuntime_InvokeObjectMethod caseObj, actualMethodName
  If Err.Number <> 0 Then
    WscovRuntime_AppendFailureReason failed, reason, label & ": " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0
End Sub

Sub WscovRuntime_AppendFailureReason(ByRef failed, ByRef reason, addition)
  If failed Then
    reason = reason & " | " & addition
  Else
    failed = True
    reason = addition
  End If
End Sub

Function WscovRuntime_CreateCaseInstance(className)
  Dim wscovCaseObjTemp

  If Not WscovRuntime_IsValidIdentifier(className) Then
    Err.Raise 5711, "WscovRuntime_CreateCaseInstance", "Invalid class name: " & className
  End If

  Set wscovCaseObjTemp = Nothing
  Execute "Set wscovCaseObjTemp = New " & className
  Set WscovRuntime_CreateCaseInstance = wscovCaseObjTemp
End Function

Sub WscovRuntime_InvokeObjectMethod(caseObj, methodName)
  Dim wscovCaseObjTemp

  If Not WscovRuntime_IsValidIdentifier(methodName) Then
    Err.Raise 5712, "WscovRuntime_InvokeObjectMethod", "Invalid method name: " & methodName
  End If

  Set wscovCaseObjTemp = caseObj
  Execute "Call wscovCaseObjTemp." & methodName & "()"
End Sub

Function WscovRuntime_BuildCoverageSnapshot(mapObj, hitsDict)
  Dim snapshot, components, componentSummaries
  Dim i, j, pointId, covered, pointCount
  Dim totalCovered, totalPoints, component, points, item

  Set snapshot = CreateObject("Scripting.Dictionary")
  Set componentSummaries = CreateList()
  Set components = mapObj("components")

  totalCovered = 0
  totalPoints = 0

  For i = 0 To ListCount(components) - 1
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

    Set item = CreateObject("Scripting.Dictionary")
    item.Add "componentId", CStr(component("componentId"))
    item.Add "scriptIndex", CLng(component("scriptIndex"))
    item.Add "covered", CLng(covered)
    item.Add "total", CLng(pointCount)
    ListAdd componentSummaries, item
  Next

  snapshot.Add "covered", CLng(totalCovered)
  snapshot.Add "total", CLng(totalPoints)
  snapshot.Add "components", componentSummaries
  Set WscovRuntime_BuildCoverageSnapshot = snapshot
End Function

Sub WscovRuntime_AppendCoverageBlock(coverageLines, testName, failed, reason, coverageInfo)
  Dim components, i, component

  If ListCount(coverageLines) > 2 Then
    ListAdd coverageLines, ""
  End If

  ListAdd coverageLines, "TEST " & testName
  If failed Then
    ListAdd coverageLines, "RESULT FAIL"
    ListAdd coverageLines, "DETAIL " & reason
  Else
    ListAdd coverageLines, "RESULT PASS"
  End If
  ListAdd coverageLines, "TOTAL " & WscovRuntime_FormatCoverageBrief(coverageInfo)

  Set components = coverageInfo("components")
  For i = 0 To ListCount(components) - 1
    Set component = components(CStr(i))
    ListAdd coverageLines, "COMPONENT " & component("componentId") _
      & " [scriptIndex=" & CStr(component("scriptIndex")) & "] " _
      & "covered=" & CStr(component("covered")) _
      & " total=" & CStr(component("total")) _
      & " rate=" & WscovRuntime_FormatPercent2(component("covered"), component("total"))
  Next
End Sub

Function WscovRuntime_FormatCoverageBrief(coverageInfo)
  WscovRuntime_FormatCoverageBrief = "covered=" & CStr(coverageInfo("covered")) _
    & " total=" & CStr(coverageInfo("total")) _
    & " rate=" & WscovRuntime_FormatPercent2(coverageInfo("covered"), coverageInfo("total"))
End Function

Function WscovRuntime_DiffHits(beforeHits, afterHits)
  Dim diff, key, beforeCount, afterCount
  Set diff = CreateObject("Scripting.Dictionary")

  For Each key In afterHits.Keys
    afterCount = CLng(afterHits(key))
    If beforeHits.Exists(key) Then
      beforeCount = CLng(beforeHits(key))
    Else
      beforeCount = 0
    End If

    If afterCount - beforeCount > 0 Then
      diff.Add key, CLng(afterCount - beforeCount)
    End If
  Next

  Set WscovRuntime_DiffHits = diff
End Function

Function WscovRuntime_ResolveTestMethods(className, methods)
  Dim methodNames

  methodNames = WscovRuntime_NormalizeMethodNames(methods)
  If WscovRuntime_ArrayCount(methodNames) = 0 Then
    methodNames = WscovRuntime_DiscoverTestMethods(className)
  End If

  WscovRuntime_ResolveTestMethods = methodNames
End Function

Function WscovRuntime_DiscoverTestMethods(className)
  Dim definition, methods, discovered, key
  Set discovered = CreateList()
  Set definition = WscovRuntime_GetTestCaseDefinition(className)

  If definition Is Nothing Then
    WscovRuntime_DiscoverTestMethods = Array()
    Exit Function
  End If

  Set methods = definition("methods")
  For Each key In methods.Keys
    If Left(LCase(CStr(methods(key))), 4) = "test" Then
      ListAdd discovered, methods(key)
    End If
  Next

  WscovRuntime_DiscoverTestMethods = WscovRuntime_SortListToArray(discovered)
End Function

Function WscovRuntime_NormalizeMethodNames(methods)
  Dim names, i, text, tokens, token, list
  Set list = CreateList()

  If IsArray(methods) Then
    If WscovRuntime_ArrayCount(methods) = 0 Then
      WscovRuntime_NormalizeMethodNames = Array()
      Exit Function
    End If
    For i = LBound(methods) To UBound(methods)
      token = Trim(CStr(methods(i)))
      If token <> "" Then
        ListAdd list, token
      End If
    Next
    WscovRuntime_NormalizeMethodNames = WscovRuntime_SortListToArray(list)
    Exit Function
  End If

  If IsObject(methods) Then
    If IsList(methods) Then
      For i = 0 To ListCount(methods) - 1
        token = Trim(CStr(methods(CStr(i))))
        If token <> "" Then
          ListAdd list, token
        End If
      Next
      WscovRuntime_NormalizeMethodNames = WscovRuntime_SortListToArray(list)
      Exit Function
    End If
  End If

  text = Trim(CStr(methods))
  If text = "" Then
    WscovRuntime_NormalizeMethodNames = Array()
    Exit Function
  End If

  text = Replace(text, vbCrLf, ",")
  text = Replace(text, vbCr, ",")
  text = Replace(text, vbLf, ",")
  text = Replace(text, ";", ",")
  tokens = Split(text, ",")
  For i = LBound(tokens) To UBound(tokens)
    token = Trim(CStr(tokens(i)))
    If token <> "" Then
      ListAdd list, token
    End If
  Next

  WscovRuntime_NormalizeMethodNames = WscovRuntime_SortListToArray(list)
End Function

Sub WscovRuntime_ScanTestCaseDefinitions(code)
  Dim lines, i, codeLine, currentClass, methodName

  lines = Split(WscovRuntime_NormalizeNewlines(code), vbLf)
  currentClass = ""

  For i = 0 To UBound(lines)
    codeLine = Trim(WscovRuntime_StripInlineComment(lines(i)))
    If currentClass = "" Then
      currentClass = WscovRuntime_ParseClassStartName(codeLine)
      If currentClass <> "" Then
        WscovRuntime_EnsureCaseDefinition currentClass
      End If
    Else
      If LCase(codeLine) = "end class" Then
        currentClass = ""
      Else
        methodName = WscovRuntime_ParseClassMethodName(codeLine)
        If methodName <> "" Then
          WscovRuntime_RegisterCaseMethod currentClass, methodName
        End If
      End If
    End If
  Next
End Sub

Sub WscovRuntime_EnsureCaseDefinition(className)
  Dim key, definition
  key = LCase(CStr(className))
  If g_caseDefinitions.Exists(key) Then
    Exit Sub
  End If

  Set definition = CreateObject("Scripting.Dictionary")
  definition.Add "name", CStr(className)
  definition.Add "methods", CreateObject("Scripting.Dictionary")
  g_caseDefinitions.Add key, definition
End Sub

Sub WscovRuntime_RegisterCaseMethod(className, methodName)
  Dim definition, methods, key
  WscovRuntime_EnsureCaseDefinition className
  Set definition = g_caseDefinitions(LCase(CStr(className)))
  Set methods = definition("methods")
  key = LCase(CStr(methodName))

  If methods.Exists(key) Then
    methods(key) = CStr(methodName)
  Else
    methods.Add key, CStr(methodName)
  End If
End Sub

Function WscovRuntime_GetTestCaseDefinition(className)
  Dim key
  key = LCase(Trim(CStr(className)))
  If g_caseDefinitions.Exists(key) Then
    Set WscovRuntime_GetTestCaseDefinition = g_caseDefinitions(key)
  Else
    Set WscovRuntime_GetTestCaseDefinition = Nothing
  End If
End Function

Function WscovRuntime_TestCaseMethodExists(className, methodName)
  Dim definition, methods, key
  Set definition = WscovRuntime_GetTestCaseDefinition(className)
  If definition Is Nothing Then
    WscovRuntime_TestCaseMethodExists = False
    Exit Function
  End If

  Set methods = definition("methods")
  key = LCase(Trim(CStr(methodName)))
  WscovRuntime_TestCaseMethodExists = methods.Exists(key)
End Function

Function WscovRuntime_ResolveMethodName(className, methodName)
  Dim definition, methods, key
  Set definition = WscovRuntime_GetTestCaseDefinition(className)
  If definition Is Nothing Then
    WscovRuntime_ResolveMethodName = CStr(methodName)
    Exit Function
  End If

  Set methods = definition("methods")
  key = LCase(Trim(CStr(methodName)))
  If methods.Exists(key) Then
    WscovRuntime_ResolveMethodName = CStr(methods(key))
  Else
    WscovRuntime_ResolveMethodName = CStr(methodName)
  End If
End Function

Function WscovRuntime_ParseClassStartName(line)
  Dim text, lower
  text = LTrim(CStr(line))
  lower = LCase(text)

  If Left(lower, 6) <> "class " Then
    WscovRuntime_ParseClassStartName = ""
    Exit Function
  End If

  WscovRuntime_ParseClassStartName = WscovRuntime_ExtractIdentifier(Mid(text, 7))
End Function

Function WscovRuntime_ParseClassMethodName(line)
  Dim text, lower, afterToken
  text = LTrim(CStr(line))
  lower = LCase(text)

  If Left(lower, 8) = "private " Then
    WscovRuntime_ParseClassMethodName = ""
    Exit Function
  End If

  Do
    If Left(lower, 7) = "public " Then
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
    WscovRuntime_ParseClassMethodName = WscovRuntime_ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 9) = "function " Then
    afterToken = Mid(text, 10)
    WscovRuntime_ParseClassMethodName = WscovRuntime_ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 13) = "property get " Then
    afterToken = Mid(text, 14)
    WscovRuntime_ParseClassMethodName = WscovRuntime_ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 13) = "property let " Then
    afterToken = Mid(text, 14)
    WscovRuntime_ParseClassMethodName = WscovRuntime_ExtractIdentifier(afterToken)
    Exit Function
  End If

  If Left(lower, 13) = "property set " Then
    afterToken = Mid(text, 14)
    WscovRuntime_ParseClassMethodName = WscovRuntime_ExtractIdentifier(afterToken)
    Exit Function
  End If

  WscovRuntime_ParseClassMethodName = ""
End Function

Function WscovRuntime_ExtractIdentifier(fragment)
  Dim text, i, ch
  text = LTrim(CStr(fragment))
  If text = "" Then
    WscovRuntime_ExtractIdentifier = ""
    Exit Function
  End If

  ch = Mid(text, 1, 1)
  If Not WscovRuntime_IsIdentifierStartChar(ch) Then
    WscovRuntime_ExtractIdentifier = ""
    Exit Function
  End If

  i = 1
  Do While i <= Len(text)
    ch = Mid(text, i, 1)
    If WscovRuntime_IsIdentifierChar(ch) Then
      i = i + 1
    Else
      Exit Do
    End If
  Loop

  WscovRuntime_ExtractIdentifier = Left(text, i - 1)
End Function

Function WscovRuntime_IsValidIdentifier(name)
  Dim text, i, ch
  text = Trim(CStr(name))
  If text = "" Then
    WscovRuntime_IsValidIdentifier = False
    Exit Function
  End If

  ch = Mid(text, 1, 1)
  If Not WscovRuntime_IsIdentifierStartChar(ch) Then
    WscovRuntime_IsValidIdentifier = False
    Exit Function
  End If

  For i = 2 To Len(text)
    ch = Mid(text, i, 1)
    If Not WscovRuntime_IsIdentifierChar(ch) Then
      WscovRuntime_IsValidIdentifier = False
      Exit Function
    End If
  Next

  WscovRuntime_IsValidIdentifier = True
End Function

Function WscovRuntime_IsIdentifierStartChar(ch)
  Dim code
  code = AscW(ch)
  WscovRuntime_IsIdentifierStartChar = ((code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "_")
End Function

Function WscovRuntime_IsIdentifierChar(ch)
  Dim code
  code = AscW(ch)
  WscovRuntime_IsIdentifierChar = ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "_")
End Function

Function WscovRuntime_NormalizeNewlines(text)
  WscovRuntime_NormalizeNewlines = Replace(Replace(CStr(text), vbCrLf, vbLf), vbCr, vbLf)
End Function

Function WscovRuntime_StripInlineComment(line)
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

  WscovRuntime_StripInlineComment = out
End Function

Function WscovRuntime_IsTruthy(value)
  On Error Resume Next

  If IsObject(value) Then
    WscovRuntime_IsTruthy = Not (value Is Nothing)
  ElseIf IsNull(value) Or IsEmpty(value) Then
    WscovRuntime_IsTruthy = False
  ElseIf VarType(value) = vbBoolean Then
    WscovRuntime_IsTruthy = CBool(value)
  ElseIf IsNumeric(value) Then
    WscovRuntime_IsTruthy = (CDbl(value) <> 0)
  Else
    WscovRuntime_IsTruthy = (CStr(value) <> "")
  End If

  If Err.Number <> 0 Then
    WscovRuntime_IsTruthy = False
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function WscovRuntime_ValuesAreEqual(expected, actual)
  On Error Resume Next

  If IsObject(expected) Or IsObject(actual) Then
    If IsObject(expected) And IsObject(actual) Then
      WscovRuntime_ValuesAreEqual = (expected Is actual)
    Else
      WscovRuntime_ValuesAreEqual = False
    End If
  ElseIf IsNull(expected) Or IsNull(actual) Then
    WscovRuntime_ValuesAreEqual = (IsNull(expected) And IsNull(actual))
  ElseIf IsEmpty(expected) Or IsEmpty(actual) Then
    WscovRuntime_ValuesAreEqual = (IsEmpty(expected) And IsEmpty(actual))
  ElseIf IsNumeric(expected) And IsNumeric(actual) Then
    WscovRuntime_ValuesAreEqual = (CDbl(expected) = CDbl(actual))
  Else
    WscovRuntime_ValuesAreEqual = (CStr(expected) = CStr(actual))
  End If

  If Err.Number <> 0 Then
    WscovRuntime_ValuesAreEqual = False
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function WscovRuntime_IsNil(value)
  On Error Resume Next

  If IsObject(value) Then
    WscovRuntime_IsNil = (value Is Nothing)
  Else
    WscovRuntime_IsNil = (IsNull(value) Or IsEmpty(value))
  End If

  If Err.Number <> 0 Then
    WscovRuntime_IsNil = False
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function WscovRuntime_IsEmptyValue(value)
  Dim count

  If WscovRuntime_IsNil(value) Then
    WscovRuntime_IsEmptyValue = True
    Exit Function
  End If

  If IsArray(value) Then
    WscovRuntime_IsEmptyValue = (WscovRuntime_ArrayCount(value) = 0)
    Exit Function
  End If

  If IsObject(value) Then
    If IsList(value) Then
      WscovRuntime_IsEmptyValue = (ListCount(value) = 0)
      Exit Function
    End If

    count = WscovRuntime_ObjectCount(value)
    If count >= 0 Then
      WscovRuntime_IsEmptyValue = (count = 0)
      Exit Function
    End If

    WscovRuntime_IsEmptyValue = False
    Exit Function
  End If

  WscovRuntime_IsEmptyValue = (Len(CStr(value)) = 0)
End Function

Function WscovRuntime_Matches(pattern, actual)
  Dim re
  Set re = WscovRuntime_NormalizeRegExp(pattern)
  WscovRuntime_Matches = re.Test(CStr(actual))
End Function

Function WscovRuntime_Includes(container, expected)
  Dim i, key

  If IsArray(container) Then
    If WscovRuntime_ArrayCount(container) = 0 Then
      WscovRuntime_Includes = False
      Exit Function
    End If
    For i = LBound(container) To UBound(container)
      If WscovRuntime_ValuesAreEqual(container(i), expected) Then
        WscovRuntime_Includes = True
        Exit Function
      End If
    Next
    WscovRuntime_Includes = False
    Exit Function
  End If

  If IsObject(container) Then
    If IsList(container) Then
      For i = 0 To ListCount(container) - 1
        If WscovRuntime_ValuesAreEqual(container(CStr(i)), expected) Then
          WscovRuntime_Includes = True
          Exit Function
        End If
      Next
      WscovRuntime_Includes = False
      Exit Function
    End If

    If TypeName(container) = "Dictionary" Then
      If container.Exists(expected) Then
        WscovRuntime_Includes = True
        Exit Function
      End If

      For Each key In container.Keys
        If WscovRuntime_ValuesAreEqual(container(key), expected) Then
          WscovRuntime_Includes = True
          Exit Function
        End If
      Next

      WscovRuntime_Includes = False
      Exit Function
    End If
  End If

  WscovRuntime_Includes = (InStr(1, CStr(container), CStr(expected), vbBinaryCompare) > 0)
End Function

Function WscovRuntime_IsSameObject(expected, actual)
  On Error Resume Next

  If IsObject(expected) And IsObject(actual) Then
    WscovRuntime_IsSameObject = (expected Is actual)
  Else
    WscovRuntime_IsSameObject = False
  End If

  If Err.Number <> 0 Then
    WscovRuntime_IsSameObject = False
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function WscovRuntime_NormalizeRegExp(pattern)
  Dim re
  If IsObject(pattern) Then
    Set WscovRuntime_NormalizeRegExp = pattern
    Exit Function
  End If

  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = CStr(pattern)
  re.Global = False
  re.IgnoreCase = False
  Set WscovRuntime_NormalizeRegExp = re
End Function

Function WscovRuntime_ObjectCount(value)
  On Error Resume Next
  WscovRuntime_ObjectCount = -1
  WscovRuntime_ObjectCount = CLng(value.Count)
  If Err.Number <> 0 Then
    WscovRuntime_ObjectCount = -1
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function WscovRuntime_DescribeValue(value)
  On Error Resume Next

  If IsObject(value) Then
    If value Is Nothing Then
      WscovRuntime_DescribeValue = "Nothing"
    ElseIf IsList(value) Then
      WscovRuntime_DescribeValue = "List(count=" & CStr(ListCount(value)) & ")"
    Else
      WscovRuntime_DescribeValue = "Object(" & TypeName(value) & ")"
    End If
  ElseIf IsNull(value) Then
    WscovRuntime_DescribeValue = "Null"
  ElseIf IsEmpty(value) Then
    WscovRuntime_DescribeValue = "Empty"
  ElseIf IsArray(value) Then
    WscovRuntime_DescribeValue = "Array(count=" & CStr(WscovRuntime_ArrayCount(value)) & ")"
  Else
    WscovRuntime_DescribeValue = "[" & CStr(value) & "]"
  End If

  If Err.Number <> 0 Then
    WscovRuntime_DescribeValue = "[unprintable]"
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function WscovRuntime_MessageOrDefault(message, defaultMessage)
  Dim text
  On Error Resume Next
  text = Trim(CStr(message))
  If Err.Number <> 0 Then
    text = ""
    Err.Clear
  End If
  On Error GoTo 0

  If text = "" Then
    WscovRuntime_MessageOrDefault = defaultMessage
  Else
    WscovRuntime_MessageOrDefault = text
  End If
End Function

Function WscovRuntime_FormatPercent2(covered, total)
  Dim scaled, intPart, fracPart
  If CLng(total) <= 0 Then
    WscovRuntime_FormatPercent2 = "0.00%"
    Exit Function
  End If

  scaled = Int((CDbl(covered) / CDbl(total)) * 10000 + 0.5)
  intPart = scaled \ 100
  fracPart = scaled Mod 100
  WscovRuntime_FormatPercent2 = CStr(intPart) & "." & Right("00" & CStr(fracPart), 2) & "%"
End Function

Function WscovRuntime_ArrayCount(values)
  On Error Resume Next
  WscovRuntime_ArrayCount = UBound(values) - LBound(values) + 1
  If Err.Number <> 0 Then
    WscovRuntime_ArrayCount = 0
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function WscovRuntime_SortListToArray(items)
  Dim arr, i, j, temp
  If ListCount(items) = 0 Then
    WscovRuntime_SortListToArray = Array()
    Exit Function
  End If

  ReDim arr(ListCount(items) - 1)
  For i = 0 To ListCount(items) - 1
    arr(i) = CStr(items(CStr(i)))
  Next

  For i = 0 To UBound(arr) - 1
    For j = i + 1 To UBound(arr)
      If LCase(arr(j)) < LCase(arr(i)) Then
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
      End If
    Next
  Next

  WscovRuntime_SortListToArray = arr
End Function
