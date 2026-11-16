<%
Function Text(value)
    If IsNull(value) Or IsEmpty(value) Then
        Text = ""
    Else
        Text = CStr(value)
    End If
End Function

Function H(value)
    Dim result
    result = Replace(Text(value), "&", "&amp;")
    result = Replace(result, "<", "&lt;")
    result = Replace(result, ">", "&gt;")
    result = Replace(result, """", "&quot;")
    H = Replace(result, "'", "&#39;")
End Function

Function Matches(value, pattern)
    Dim expression
    Set expression = New RegExp
    expression.Pattern = pattern
    Matches = expression.Test(Text(value))
End Function

Function ParseInteger(value, minimum, maximum, ByRef parsed)
    ParseInteger = False
    If Not Matches(value, "^[0-9]{1,10}$") Then Exit Function
    Dim number
    number = CDbl(value)
    If number < minimum Or number > maximum Then Exit Function
    parsed = CLng(number)
    ParseInteger = True
End Function

Function ParseDecimal(value, minimum, maximum, ByRef parsed)
    ParseDecimal = False
    If Not Matches(value, "^[0-9]{1,4}(\.[0-9]{1,2})?$") Then Exit Function
    Dim parts, number
    parts = Split(value, ".")
    number = CDbl(parts(0))
    If UBound(parts) = 1 Then number = number + CDbl(parts(1)) / (10 ^ Len(parts(1)))
    If number < minimum Or number > maximum Then Exit Function
    parsed = number
    ParseDecimal = True
End Function

Function ParseDateOnly(value, ByRef parsed)
    ParseDateOnly = False
    If Not Matches(value, "^[0-9]{4}-[0-9]{2}-[0-9]{2}$") Then Exit Function
    Dim y, m, d, candidate
    y = CInt(Left(value, 4))
    m = CInt(Mid(value, 6, 2))
    d = CInt(Right(value, 2))
    If y < 1900 Or y > 9998 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    candidate = DateSerial(y, m, d)
    If Year(candidate) <> y Or Month(candidate) <> m Or Day(candidate) <> d Then Exit Function
    parsed = candidate
    ParseDateOnly = True
End Function

Function ParseTimestamp(value, ByRef parsed)
    ParseTimestamp = False
    If Not Matches(value, "^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}$") Then Exit Function
    Dim dayValue, hours, minutes
    If Not ParseDateOnly(Left(value, 10), dayValue) Then Exit Function
    hours = CInt(Mid(value, 12, 2))
    minutes = CInt(Right(value, 2))
    If hours > 23 Or minutes > 59 Then Exit Function
    parsed = dayValue + TimeSerial(hours, minutes, 0)
    ParseTimestamp = True
End Function

Function IsoDate(value)
    IsoDate = Right("0000" & Year(value), 4) & "-" & Right("0" & Month(value), 2) & "-" & Right("0" & Day(value), 2)
End Function

Function IsoTimestamp(value)
    IsoTimestamp = IsoDate(value) & "T" & Right("0" & Hour(value), 2) & ":" & Right("0" & Minute(value), 2)
End Function

Function DisplayDate(value)
    DisplayDate = Day(value) & " " & MonthName(Month(value), True) & " " & Year(value)
End Function

Function DisplayTime(value)
    DisplayTime = Right("0" & Hour(value), 2) & ":" & Right("0" & Minute(value), 2)
End Function

Function CompactNumber(value)
    Dim result
    result = Replace(FormatNumber(CDbl(value), 2, -1, 0, 0), ",", ".")
    Do While Right(result, 1) = "0"
        result = Left(result, Len(result) - 1)
    Loop
    If Right(result, 1) = "." Then result = Left(result, Len(result) - 1)
    CompactNumber = result
End Function

Function JsonNumber(value)
    JsonNumber = Replace(CStr(Round(CDbl(value), 4)), ",", ".")
End Function

Function UnitLabel(unit)
    Select Case unit
        Case "mmol": UnitLabel = "mmol/L"
        Case "mg": UnitLabel = "mg/dL"
        Case Else: Err.Raise vbObjectError + 101, "UnitLabel", "Unknown glucose unit."
    End Select
End Function

Function ConvertGlucose(value, originalUnit, targetUnit)
    Dim result
    If (originalUnit <> "mmol" And originalUnit <> "mg") Or (targetUnit <> "mmol" And targetUnit <> "mg") Then
        Err.Raise vbObjectError + 102, "ConvertGlucose", "Unknown glucose unit."
    End If
    result = CDbl(value)
    If originalUnit <> targetUnit Then
        If targetUnit = "mmol" Then
            result = result / 18.0182
        Else
            result = result * 18.0182
        End If
    End If
    ConvertGlucose = result
End Function

Function ValidContext(value)
    Select Case value
        Case "", "fasting", "before-meal", "after-meal", "bedtime", "random": ValidContext = True
        Case Else: ValidContext = False
    End Select
End Function

Function ContextLabel(value)
    Select Case Text(value)
        Case "": ContextLabel = "Not specified"
        Case "fasting": ContextLabel = "Fasting"
        Case "before-meal": ContextLabel = "Before a meal"
        Case "after-meal": ContextLabel = "After a meal"
        Case "bedtime": ContextLabel = "Bedtime"
        Case "random": ContextLabel = "Other / random"
        Case Else: ContextLabel = Text(value)
    End Select
End Function

Function NewReading()
    Dim values, key
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In Array("id", "type", "recorded_at", "systolic", "diastolic", "pulse", "glucose_value", "glucose_unit", "context", "notes")
        values.Add key, ""
    Next
    Set NewReading = values
End Function

Function ValidateReading(values, referenceTime)
    Dim errors, parsed, sysValid, diaValid, maximum
    Set errors = CreateObject("Scripting.Dictionary")
    values("parsed_at") = Null
    values("parsed_systolic") = Null
    values("parsed_diastolic") = Null
    values("parsed_pulse") = Null
    values("parsed_glucose") = Null
    If ParseTimestamp(values("recorded_at"), parsed) Then
        If parsed > DateAdd("n", 5, referenceTime) Then
            errors.Add "recorded_at", "The reading time cannot be more than five minutes in the future."
        Else
            values("parsed_at") = parsed
        End If
    Else
        errors.Add "recorded_at", "Enter a valid local date and time (year 1900 or later)."
    End If
    Select Case values("type")
        Case "pressure"
            sysValid = ParseInteger(values("systolic"), 1, 400, parsed)
            If sysValid Then values("parsed_systolic") = parsed
            diaValid = ParseInteger(values("diastolic"), 1, 300, parsed)
            If diaValid Then values("parsed_diastolic") = parsed
            If Not sysValid Then errors.Add "systolic", "Enter a whole-number systolic value between 1 and 400."
            If Not diaValid Then errors.Add "diastolic", "Enter a whole-number diastolic value between 1 and 300."
            If sysValid And diaValid Then
                If values("parsed_systolic") <= values("parsed_diastolic") Then errors.Add "diastolic", "Diastolic must be lower than systolic."
            End If
            If values("pulse") <> "" Then
                If ParseInteger(values("pulse"), 1, 300, parsed) Then
                    values("parsed_pulse") = parsed
                Else
                    errors.Add "pulse", "Enter a whole-number pulse between 1 and 300, or leave it blank."
                End If
            End If
        Case "glucose"
            maximum = 111
            If values("glucose_unit") = "mg" Then maximum = 2000
            If values("glucose_unit") <> "mmol" And values("glucose_unit") <> "mg" Then errors.Add "glucose_unit", "Choose mmol/L or mg/dL."
            If ParseDecimal(values("glucose_value"), 0.01, maximum, parsed) Then
                values("parsed_glucose") = parsed
            Else
                errors.Add "glucose_value", "Enter a positive glucose value up to " & maximum & ", with at most two decimal places. Use a decimal point."
            End If
            If Not ValidContext(values("context")) Then errors.Add "context", "Choose a context from the list."
        Case Else
            errors.Add "type", "Choose blood pressure or blood glucose."
    End Select
    If Len(values("notes")) > 2000 Then errors.Add "notes", "Keep notes to 2,000 characters or fewer."
    Set ValidateReading = errors
End Function

Function CsvCell(value)
    Dim result, firstCharacter
    result = Text(value)
    firstCharacter = Left(result, 1)
    If Matches(result, "^\s*[=+\-@]") Or firstCharacter = vbTab Or firstCharacter = vbCr Or firstCharacter = vbLf Then result = "'" & result
    CsvCell = """" & Replace(result, """", """""") & """"
End Function
%>
