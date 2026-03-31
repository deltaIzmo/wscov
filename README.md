# WSCOV

WSCOV is a VBScript-based coverage tool for `.wsc` (Windows Script Component) files.
It instruments a component, runs tests through the WSC moniker loader, and writes both overall and per-test coverage outputs.

The repository already includes:

- core instrument and run scripts under `tools/`
- Windows and WSL wrapper commands for day-to-day use
- sample SUTs under `samples/sut/`
- sample tests under `samples/tests/`

## Repository Layout

- `tools/wscov_instrument.vbs`: core instrumenter
- `tools/wscov_run.vbs`: core runner
- `tools/wscov_test_runtime.vbs`: built-in lightweight test runtime
- `tools/wscov.cmd`: Windows quick start for the Calculator sample
- `tools/wscov_apply_target.cmd`: Windows wrapper for an arbitrary target WSC
- `tools/wscov_regression.cmd`: Windows regression runner for bundled samples
- `tools/wscov_wsl.sh`: WSL wrapper that stages files onto a Windows-local path
- `tools/wscov_regression_wsl.sh`: WSL regression runner for bundled samples
- `samples/sut/`: sample WSC components
- `samples/tests/calculator/`: recommended class-based test example
- `samples/tests/branchy/`: class-based regression sample
- `samples/tests/calculator.test.vbs`: legacy procedural compatibility sample
- `out/`: generated artifacts; tracked only with `.gitkeep`

## Requirements

- Windows 11 with WSH / `cscript.exe` for actual WSC execution
- No external test framework
- For WSL usage: WSL2 with access to `cmd.exe` and a Windows-local writable path

The `.cmd` wrappers should be run from a normal Windows-local checkout, not from `\\wsl$`.
When working from WSL, use the shell wrappers instead.

## Quick Start

### Windows sample run

```bat
tools\wscov.cmd
```

This runs the bundled Calculator sample using:

- input: `samples\sut\Calculator.wsc`
- tests: `samples\tests\calculator`
- output: `out\`

Expected generated files:

- `out\Calculator.__cov__.wsc`
- `out\coverage-map.json`
- `out\test-results.txt`
- `out\test-coverage.txt`
- `out\hits.txt`
- `out\coverage-summary.txt`

### Windows target run

```bat
tools\wscov_apply_target.cmd <input.wsc> <testsDir> <outDir> [componentId]
```

Example:

```bat
tools\wscov_apply_target.cmd samples\sut\Calculator.wsc samples\tests\calculator out\target\calculator Calculator
```

This wrapper runs both instrumentation and test execution, and also keeps:

- `instrument.log`
- `run.log`

### WSL run

```bash
tools/wscov_wsl.sh <input.wsc> <testsDir> <outDir> [componentId]
```

Example:

```bash
tools/wscov_wsl.sh samples/sut/Calculator.wsc samples/tests/calculator out/wsl/calculator Calculator
```

The WSL wrapper:

- copies the minimum required files into a Windows-local workspace
- runs `tools\wscov_apply_target.cmd` there through `cmd.exe`
- copies generated outputs back into the requested WSL `outDir`

Default Windows workspace:

- `%USERPROFILE%\wscov-work\wscov`

Optional override:

```bash
export WSCOV_WINDOWS_WORKROOT='C:\Users\yourname\wscov-work\custom-root'
```

`WSCOV_WINDOWS_WORKROOT` may also be set with a WSL path such as `/mnt/c/Users/yourname/wscov-work/custom-root`.

This copy-based flow is intentional.
The VBScript tooling rejects URL and UNC-style paths, so WSC loading must happen from a normal Windows-local path.
The staged workspace is only for wrapper-driven execution.
Do not run `tools\wscov.cmd` or `tools\wscov_regression.cmd` directly inside that staged copy.

## Core CLI

### Instrument

```bat
cscript //nologo tools\wscov_instrument.vbs <input.wsc> <output.instrumented.wsc> <coverage-map.json>
```

Exit codes:

- `0`: success
- `2`: usage or argument error
- `3`: XML parse or load failure
- `4`: instrumentation failure

### Run

```bat
cscript //nologo tools\wscov_run.vbs <instrumented.wsc> [componentId] <testsDir> <coverage-map.json> <outDir>
```

Exit codes:

- `0`: all tests passed
- `1`: test failures, or no tests were registered
- `2`: usage or argument error
- `5`: runtime error during WSC load, map parsing, or coverage dump

## Wrapper Commands

- `tools\wscov.cmd`: runs the Calculator sample into `out\`
- `tools\wscov_apply_target.cmd`: generic Windows wrapper for one target and one test directory
- `tools\wscov_regression.cmd`: runs bundled sample suites into `out\regression\calculator` and `out\regression\branchy`
- `tools/wscov_wsl.sh`: generic WSL wrapper for one target and one test directory
- `tools/wscov_regression_wsl.sh`: runs bundled sample suites into `out/wsl/regression/calculator` and `out/wsl/regression/branchy`

## Writing Tests

Test files are plain `.test.vbs` files loaded by `tools/wscov_run.vbs`.
The recommended style is a class-based test case with `test_...` auto-discovery:

```vbscript
Class CalculatorTest
  Public Sub test_add()
    assert_equal 3, sut.Add(1, 2)
  End Sub

  Public Sub test_sub()
    assert_equal 1, sut.Sub(3, 2)
  End Sub
End Class

Wscov_RegisterTestCase "CalculatorTest"
```

Supported lifecycle hooks:

- `before_setup`
- `setup`
- `after_setup`
- `before_teardown`
- `teardown`
- `after_teardown`

Each test method runs on a fresh instance of the test class.

Useful assertions:

- `assert`, `refute`
- `assert_equal`, `refute_equal`
- `assert_nil`, `refute_nil`
- `assert_empty`, `refute_empty`
- `assert_match`, `refute_match`
- `assert_includes`, `refute_includes`
- `assert_same`, `refute_same`
- `flunk`

Ruby-like snake_case names and older VB-style names such as `AssertEqual` are both available.

Procedural registration with `Wscov_AddTest` is still supported for compatibility.
See `samples/tests/calculator.test.vbs` for that older style.

## Coverage Outputs

Common run outputs:

- `test-results.txt`: pass/fail summary for each test
- `test-coverage.txt`: per-test coverage detail
- `hits.txt`: raw coverage hits dumped from the instrumented component
- `coverage-summary.txt`: aggregate coverage summary derived from the coverage map and hits
- `coverage-map.json`: instrumentation map used by the runner
- `*.__cov__.wsc`: instrumented WSC generated by the instrument step

`test-results.txt` includes a coverage summary on each result line.
`test-coverage.txt` includes:

- test name and pass/fail result
- failure detail when a test fails
- covered points, total points, and coverage rate for that test
- per-component coverage counts for multi-component `.wsc` files

## License

This project is licensed under the Apache License 2.0.
See `LICENSE` for the authoritative English text.

See `LICENSE.ja` for a non-authoritative Japanese reference translation.