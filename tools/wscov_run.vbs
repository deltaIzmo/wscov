Option Explicit

Dim sut
Dim g_tests
Dim g_setupProcs
Dim g_teardownProcs
Dim g_caseDefinitions

LoadLibrary
LoadTestRuntime
InitializeWscovTestRuntime

WScript.Quit Main()

Function Main()
  Dim exitCode
  exitCode = WSCOV_EXIT_OK

  On Error Resume Next
  exitCode = RunRunner()
  If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description
    exitCode = MapErrToExitCode(Err.Number, WSCOV_EXIT_RUNTIME)
    Err.Clear
  End If
  On Error GoTo 0

  Main = exitCode
End Function

Function RunRunner()
  Dim parsed
  Dim instrumentedWscPath, componentId, testsDir, coverageMapPath, outDir
  Dim fso, testResultsPath, testCoveragePath, hitsPath, summaryPath
  Dim mapText, mapObj
  Dim testFiles, fileCount, i
  Dim resultLines, coverageLines, totalCount, passCount, failCount
  Dim testResultsText, hitsText, hitsDict, summaryText

  Set parsed = ParseArgsRun(WScript.Arguments)
  instrumentedWscPath = parsed("instrumentedWscPath")
  componentId = parsed("componentId")
  testsDir = parsed("testsDir")
  coverageMapPath = parsed("coverageMapPath")
  outDir = parsed("outDir")

  EnsureFileExists instrumentedWscPath
  EnsureFolderExists testsDir
  EnsureFileExists coverageMapPath
  EnsureDirExists outDir

  Set fso = CreateObject("Scripting.FileSystemObject")
  testResultsPath = fso.BuildPath(outDir, "test-results.txt")
  testCoveragePath = fso.BuildPath(outDir, "test-coverage.txt")
  hitsPath = fso.BuildPath(outDir, "hits.txt")
  summaryPath = fso.BuildPath(outDir, "coverage-summary.txt")

  mapText = ReadTextFile(coverageMapPath)
  Set mapObj = ParseCoverageMapJson(mapText)

  Set sut = LoadSut(instrumentedWscPath, componentId)

  On Error Resume Next
  sut.WscovResetCoverage
  If Err.Number <> 0 Then
    Dim resetMessage
    resetMessage = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseRuntime "Failed to call sut.WscovResetCoverage(): " & resetMessage
  End If
  On Error GoTo 0

  testFiles = DiscoverTestFiles(testsDir)
  fileCount = WscovRuntime_ArrayCount(testFiles)
  For i = 0 To fileCount - 1
    LoadAndExecuteTestFile CStr(testFiles(i))
  Next

  Set resultLines = CreateList()
  Set coverageLines = CreateList()
  totalCount = 0
  passCount = 0
  failCount = 0
  ExecuteRegisteredTests mapObj, resultLines, coverageLines, totalCount, passCount, failCount

  If totalCount = 0 Then
    failCount = failCount + 1
    ListAdd resultLines, "FAIL no-tests-registered - no tests were registered via Wscov_AddTest or Wscov_RegisterTestCase."
    If ListCount(coverageLines) > 2 Then
      ListAdd coverageLines, ""
    End If
    ListAdd coverageLines, "TEST no-tests-registered"
    ListAdd coverageLines, "RESULT FAIL"
    ListAdd coverageLines, "DETAIL no tests were registered via Wscov_AddTest or Wscov_RegisterTestCase."
    ListAdd coverageLines, "TOTAL covered=0 total=0 rate=0.00%"
  End If

  testResultsText = FormatTestResults(resultLines, totalCount, passCount, failCount)
  WriteTextFileUtf8 testResultsPath, testResultsText
  WriteAllLines testCoveragePath, coverageLines

  hitsText = DumpCoverageFromSut()
  WriteTextFileUtf8 hitsPath, hitsText

  Set hitsDict = ParseHitsText(hitsText)
  summaryText = FormatCoverageSummary(mapObj, hitsDict)
  WriteTextFileUtf8 summaryPath, summaryText

  WScript.Echo "OK test results: " & testResultsPath
  WScript.Echo "OK test coverage: " & testCoveragePath
  WScript.Echo "OK coverage hits: " & hitsPath
  WScript.Echo "OK coverage summary: " & summaryPath

  If failCount > 0 Then
    RunRunner = WSCOV_EXIT_TEST_FAILED
  Else
    RunRunner = WSCOV_EXIT_OK
  End If
End Function

Function LoadSut(instrumentedWscPath, componentId)
  Dim moniker, obj
  moniker = "script:" & instrumentedWscPath
  If Trim(componentId) <> "" Then
    moniker = moniker & "#" & componentId
  End If

  On Error Resume Next
  Set obj = GetObject(moniker)
  If Err.Number <> 0 Then
    Dim message
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseRuntime "Failed to load WSC moniker: " & moniker & " (" & message & ")"
  End If
  On Error GoTo 0

  Set LoadSut = obj
End Function

Function DumpCoverageFromSut()
  On Error Resume Next
  DumpCoverageFromSut = CStr(sut.WscovDumpCoverage())
  If Err.Number <> 0 Then
    Dim message
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseRuntime "Failed to call sut.WscovDumpCoverage(): " & message
  End If
  On Error GoTo 0
End Function

Sub LoadTestRuntime()
  WscovRuntime_LoadScript "wscov_test_runtime.vbs", "Shared test runtime not found: "
End Sub

Sub WscovRuntime_LoadScript(fileName, missingPrefix)
  Dim fso, scriptPath, ts, code, message
  Set fso = CreateObject("Scripting.FileSystemObject")
  scriptPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), fileName)

  If Not fso.FileExists(scriptPath) Then
    WScript.Echo "ERROR: " & missingPrefix & scriptPath
    WScript.Quit 5
  End If

  On Error Resume Next
  Set ts = fso.OpenTextFile(scriptPath, 1, False)
  code = ts.ReadAll
  ts.Close
  ExecuteGlobal code
  If Err.Number <> 0 Then
    message = Err.Description
    Err.Clear
    On Error GoTo 0
    WScript.Echo "ERROR: Failed to load script file: " & scriptPath & " (" & message & ")"
    WScript.Quit 5
  End If
  On Error GoTo 0
End Sub

Sub LoadLibrary()
  WscovRuntime_LoadScript "wscov_lib.vbs", "Shared library not found: "
End Sub
