Option Explicit
Dim fso, root, source, checks, failures, parsed, values, errors, referenceTime, ignored
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
Set source = fso.OpenTextFile(root & "\www\includes\validation.asp", 1)
ExecuteGlobal Replace(Replace(source.ReadAll, "<%", ""), "%>", "")
source.Close
ignored = SetLocale(1033)
checks = 0
failures = 0
referenceTime = DateSerial(2026, 11, 16) + TimeSerial(12, 0, 0)

Sub Assert(condition, description)
    checks = checks + 1
    If Not condition Then
        failures = failures + 1
        WScript.Echo "FAIL: " & description
    End If
End Sub

Function Fixture(kind)
    Dim item
    Set item = NewReading()
    item("type") = kind
    item("recorded_at") = "2026-11-15T08:30"
    item("systolic") = "120"
    item("diastolic") = "80"
    item("pulse") = ""
    item("glucose_value") = "5.6"
    item("glucose_unit") = "mmol"
    Set Fixture = item
End Function

Assert ParseInteger("120", 1, 400, parsed), "integer accepted"
Assert parsed = 120, "integer parsed accurately"
Assert Not ParseInteger("120.1", 1, 400, parsed), "fractional pressure rejected"
Assert Not ParseInteger("1e2", 1, 400, parsed), "scientific notation rejected"
Assert Not ParseInteger("-1", 1, 400, parsed), "negative integer rejected"
Assert Not ParseInteger("0", 1, 400, parsed), "zero rejected"
Assert Not ParseInteger("401", 1, 400, parsed), "upper integer bound enforced"
Assert ParseInteger("2147483647", 1, 2147483647, parsed), "maximum ID accepted"
Assert Not ParseInteger("2147483648", 1, 2147483647, parsed), "overflow ID rejected"
Assert Not ParseInteger("1 OR 1=1", 1, 400, parsed), "SQL-shaped integer rejected"
Assert ParseDecimal("5.60", 0.01, 111, parsed), "decimal accepted"
Assert Abs(parsed - 5.6) < 0.000001, "decimal parsed accurately"
Assert Not ParseDecimal("5,6", 0.01, 111, parsed), "decimal comma rejected explicitly"
Assert Not ParseDecimal("5.678", 0.01, 111, parsed), "excess precision rejected"
Assert Not ParseDecimal("NaN", 0.01, 111, parsed), "non-number rejected"
Assert Not ParseDecimal("0", 0.01, 111, parsed), "zero glucose rejected"
Assert ParseDecimal("0.01", 0.01, 111, parsed), "small positive value accepted"
Assert Not ParseDecimal("999999999999999", 0.01, 111, parsed), "huge number rejected safely"
Assert ParseDateOnly("2024-02-29", parsed), "leap day accepted"
Assert Not ParseDateOnly("2025-02-29", parsed), "non-leap February 29 rejected"
Assert Not ParseDateOnly("2026-02-31", parsed), "calendar rollover rejected"
Assert Not ParseDateOnly("2026-13-01", parsed), "invalid month rejected"
Assert Not ParseDateOnly("2026-01-00", parsed), "day zero rejected"
Assert Not ParseDateOnly("11/16/2026", parsed), "ambiguous date rejected"
Assert Not ParseDateOnly("1899-12-31", parsed), "unsupported date rejected"
Assert ParseTimestamp("2026-11-15T23:59", parsed), "valid local timestamp accepted"
Assert IsoTimestamp(parsed) = "2026-11-15T23:59", "timestamp round-trips"
Assert Not ParseTimestamp("2026-11-15T24:00", parsed), "hour rollover rejected"
Assert Not ParseTimestamp("2026-11-15T12:60", parsed), "minute rollover rejected"
Assert Not ParseTimestamp("2026-11-15T12:30Z", parsed), "timezone suffix rejected for local timestamp"
Assert H("<script>&""'") = "&lt;script&gt;&amp;&quot;&#39;", "HTML metacharacters encoded"
Assert H(Null) = "", "null text handled"
Assert CsvCell("a,b") = """a,b""", "CSV comma quoted"
Assert CsvCell("say ""hello""") = """say """"hello""""""", "CSV quotes escaped"
Assert CsvCell("=2+2") = """'=2+2""", "CSV formula neutralized"
Assert CsvCell("  @SUM(1)") = """'  @SUM(1)""", "CSV whitespace-prefixed formula neutralized"
Assert CsvCell(vbTab & "hello") = """'" & vbTab & "hello""", "CSV control prefix neutralized"
Assert CsvCell("normal text") = """normal text""", "ordinary CSV text unchanged"
Assert Abs(ConvertGlucose(108.1092, "mg", "mmol") - 6) < 0.0000001, "mg/dL to mmol/L conversion"
Assert Abs(ConvertGlucose(6, "mmol", "mg") - 108.1092) < 0.0000001, "mmol/L to mg/dL conversion"
Assert ConvertGlucose(5.6, "mmol", "mmol") = 5.6, "same units unchanged"
Assert CompactNumber(5.6) = "5.6", "display removes redundant zeros"
Assert CompactNumber(108) = "108", "whole number display"
ignored = SetLocale(1043)
Assert ParseDecimal("5.6", 0.01, 111, parsed), "decimal parser works in comma-decimal locale"
Assert Abs(parsed - 5.6) < 0.000001, "locale does not change numeric value"
Assert CompactNumber(5.6) = "5.6", "display decimal separator is stable"
ignored = SetLocale(1033)

Set values = Fixture("pressure")
Set errors = ValidateReading(values, referenceTime)
Assert errors.Count = 0, "valid blood pressure accepted"
Assert IsNull(values("parsed_pulse")), "optional pulse stored as null"
Assert IsNull(values("parsed_glucose")), "pressure does not store stale glucose fields"
values("diastolic") = "120"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("diastolic"), "equal systolic and diastolic rejected"
values("diastolic") = "130"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("diastolic"), "reversed blood pressure rejected"
values("systolic") = "400"
values("diastolic") = "300"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Count = 0, "broad upper pressure input limits accepted"
values("systolic") = "401"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("systolic"), "systolic bound enforced"
values("pulse") = "0"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("pulse"), "invalid pulse rejected"
Set values = Fixture("pressure")
values("recorded_at") = "2026-11-16T12:06"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("recorded_at"), "future reading rejected"
values("recorded_at") = "2026-11-16T12:05"
Set errors = ValidateReading(values, referenceTime)
Assert Not errors.Exists("recorded_at"), "five-minute clock tolerance allowed"
Set values = Fixture("glucose")
Set errors = ValidateReading(values, referenceTime)
Assert errors.Count = 0, "valid glucose accepted"
Assert IsNull(values("parsed_systolic")), "glucose does not store stale pressure fields"
values("glucose_value") = "111"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Count = 0, "mmol/L upper limit accepted"
values("glucose_value") = "111.01"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("glucose_value"), "mmol/L upper limit enforced"
values("glucose_unit") = "mg"
values("glucose_value") = "2000"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Count = 0, "mg/dL upper limit accepted"
values("glucose_value") = "2000.01"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("glucose_value"), "mg/dL upper limit enforced"
Set values = Fixture("glucose")
values("glucose_unit") = "invalid"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("glucose_unit"), "unknown units rejected"
Set values = Fixture("glucose")
values("context") = "not-a-context"
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("context"), "unknown context rejected"
Set values = Fixture("pressure")
values("notes") = String(2000, "x")
Set errors = ValidateReading(values, referenceTime)
Assert errors.Count = 0, "2000-character notes accepted"
values("notes") = String(2001, "x")
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("notes"), "oversized notes rejected"
Set values = Fixture("invalid")
Set errors = ValidateReading(values, referenceTime)
Assert errors.Exists("type"), "unknown reading type rejected"

WScript.Echo checks & " validation checks; " & failures & " failures."
If failures > 0 Then WScript.Quit 1
WScript.Quit 0
