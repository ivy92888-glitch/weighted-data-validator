Option Explicit
Option Compare Text

' ================================================================
'  Weighted Data Validator
'  Multi-source data validation with weighted scoring + rule engine
'  + color-coded report generation + suggestion analysis
'
'  TWO-PHASE PIPELINE:
'    Phase 1 — Consistency Check (validate + mark)
'    Phase 2 — Process Suggestions + Generate Feedback Report
'
'  Author: ivy92888-glitch  |  License: MIT
'  Disclaimer: Unaffiliated with any spreadsheet vendor.
'              Generic clean-room implementation. Any resemblance to
'              specific business systems is coincidental.
' ================================================================

' --- CONFIG CONSTANTS (see docs/Configuration.md) ---
Private Const WEIGHT_FREQUENCY   As Double = 0.4
Private Const WEIGHT_RECENCY     As Double = 0.6
Private Const MAX_TRUNCATE_LEN   As Integer = 18
Private Const NAME_SEPARATOR     As String  = "/"
Private Const RULE_MARK         As String  = "-"
Private Const TARGET_DOMAIN      As String  = "amazon.example-domain"

' --- COLUMN LETTER CONSTANTS (adjust to match your workbook layout) ---
Private Const COL_LOOKUP         As String = "A"   ' Grouping key (what to look up)
Private Const COL_LINK           As String = "B"   ' External / entity link
Private Const COL_NAME_RAW       As String = "C"   ' Original name (human)
Private Const COL_NAME_EN_RAW    As String = "D"   ' Original English / display name
Private Const COL_CODE_RAW       As String = "E"   ' Original identifier / code
Private Const COL_SUGGESTION     As String = "F"   ' Suggestion column (output of Phase 1)
Private Const COL_SUGGESTION_NAME As String = "G"  ' Suggested name
Private Const COL_SUGGESTION_CODE As String = "H"  ' Suggested identifier / code
Private Const COL_RATE           As String = "I"   ' Percent parameter
Private Const COL_METADATA       As String = "J"   ' Optional metadata / notes

' --- WORKBOOK SHEET CONSTANTS ---
Private Const SHEET_DATA         As String = "Data"
Private Const SHEET_REF_CODES    As String = "ReferenceCodes"
Private Const SHEET_RULES        As String = "Rules"
Private Const SHEET_FEEDBACK     As String = "Feedback"
Private Const SHEET_REF_LOOKUP   As String = "ReferenceLookup"

' --- COLOR PALETTE ---
Private Const CLR_SUGGESTED_BG   As Long = RGB(226, 239, 218)   ' light green — suggested cells
Private Const CLR_WRONG_PART     As Long = RGB(127, 150, 152)   ' slate gray-green — wrong part in description
Private Const CLR_RIGHT_PART     As Long = RGB(198,  89,  17)   ' orange — correct part in description
Private Const CLR_HEADER_BG       As Long = RGB(255, 255,   0)   ' yellow — Feedback sheet header
Private Const CLR_MATCH_OK_FONT   As Long = RGB(  0, 128,   0)   ' dark green — match confirmed
Private Const CLR_MISS_FONT       As Long = RGB(192,   0,   0)   ' dark red — no match / N/A
Private Const CLR_PREVIOUS_MARK   As Long = RGB(146, 208,  80)   ' legacy green — used in ClearPreviousMarks

' --- PRIORITY KEYWORD ARRAYS (lower index = higher priority) ---
Private primaryKeywords As Variant    ' Array("top", "try", "suggest", "recommend")
Private secondaryKeywords As Variant  ' Array("may", "let", "allow", "consider")

' --- RULE DICTIONARY (keyword → annotation) ---
Private ruleDict As Object

' ================================================================
'  MAIN ENTRY: Run both phases in sequence
' ================================================================
Sub RunFullPipeline()
    Call InitializePipeline
    Call RunPipelinePhase1   ' Consistency check + mark suggestions in-place
    Call RunPipelinePhase2   ' Process suggestions + generate Feedback sheet
    MsgBox "Pipeline complete." & vbCrLf & _
           "  Phase 1: Consistency check & suggestion analysis" & vbCrLf & _
           "  Phase 2: Feedback report generated (" & SHEET_FEEDBACK & ")", _
           vbInformation, "Weighted Data Validator"
End Sub

Sub InitializePipeline()
    primaryKeywords   = Array("top", "try", "suggest", "recommend")
    secondaryKeywords = Array("may", "let", "allow", "consider")
    Set ruleDict = CreateObject("Scripting.Dictionary")

    ' Rule engine loads from SHEET_RULES if present
    On Error Resume Next
    Dim rulesWS As Worksheet
    Set rulesWS = ThisWorkbook.Sheets(SHEET_RULES)
    On Error GoTo 0
    If Not rulesWS Is Nothing Then
        Dim lastRow As Long, i As Long
        lastRow = rulesWS.Cells(rulesWS.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            Dim kw As String, ann As String
            kw = Trim(rulesWS.Cells(i, 1).Value)
            ann = Trim(rulesWS.Cells(i, 2).Value)
            If Len(kw) > 0 And Not ruleDict.Exists(kw) Then
                ruleDict.Add kw, ann
            End If
        Next i
    End If
End Sub

' ================================================================
'  PHASE 1 — Consistency Check
'  Walk every row, compute best code from history, mark mismatches,
'  analyze free-text suggestion columns, check rules.
' ================================================================
Sub RunPipelinePhase1()
    Dim ws As Worksheet
    Set ws = ActiveSheet  ' assume user has their Data sheet selected

    ' Validate structure
    Dim refWS As Worksheet, lookupWS As Worksheet
    On Error Resume Next
    Set refWS = ThisWorkbook.Sheets(SHEET_REF_CODES)
    Set lookupWS = ThisWorkbook.Sheets(SHEET_REF_LOOKUP)
    On Error GoTo 0
    If refWS Is Nothing Then
        MsgBox "Warning: '" & SHEET_REF_CODES & "' sheet not found." & vbCrLf & _
               "Weighted scoring will be skipped. Provide a reference table " & _
               "(see docs/Configuration.md).", vbExclamation
    End If

    Dim lastRow As Long, i As Long
    ' Use whichever column exists to find data extent
    lastRow = ws.Cells(ws.Rows.Count, COL_CODE_RAW).End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "No data found. Row 3 is the last expected header row.", vbExclamation
        Exit Sub
    End If

    ' --- STEP 1: Weighted best-code match ---
    For i = 4 To lastRow
        Dim origCode As String
        origCode = Trim(ws.Cells(i, COL_CODE_RAW).Value)
        If Len(origCode) = 0 Then GoTo P1_NextRow

        ' Weighted best code from reference history
        Dim bestCode As String, bestRate As String
        bestCode = ""
        bestRate = ""
        If Not refWS Is Nothing Then
            bestCode = CalculateBestCode(Trim(ws.Cells(i, COL_LOOKUP).Value), refWS, "A", "B")
        End If
        If Len(bestCode) = 0 Then
            bestCode = origCode   ' nothing better — keep original
        End If

        ' Get rate from lookup table if available
        If Not lookupWS Is Nothing And Len(bestCode) > 0 Then
            bestRate = GetRateFromLookup(bestCode, lookupWS)
        End If

        ' Write match result to suggestion column (col F)
        Dim matchResult As String
        If Len(bestCode) > 0 And Len(bestRate) > 0 Then
            matchResult = bestCode & "-" & bestRate
        ElseIf Len(bestCode) > 0 Then
            matchResult = bestCode
        Else
            matchResult = "N/A"
        End If
        ws.Cells(i, COL_SUGGESTION).Value = matchResult

        ' Color-code match vs mismatch
        Dim fCode As String
        fCode = Split(Split(matchResult, "-")(0), "-")(0)
        If Len(bestCode) > 0 And bestCode <> "N/A" Then
            If Trim(origCode) = Trim(bestCode) Then
                ws.Cells(i, COL_SUGGESTION).Font.Color = CLR_MATCH_OK_FONT
            Else
                ws.Cells(i, COL_SUGGESTION).Font.Color = CLR_MISS_FONT
            End If
        Else
            ws.Cells(i, COL_SUGGESTION).Font.Color = CLR_MISS_FONT
        End If

P1_NextRow:
    Next i

    ' --- STEP 2: Rule-engine scan ---
    For i = 4 To lastRow
        Dim rowText As String, annotation As String
        rowText = Trim(ws.Cells(i, COL_NAME_RAW).Value) & " " & _
                  Trim(ws.Cells(i, COL_NAME_EN_RAW).Value) & " " & _
                  Trim(ws.Cells(i, COL_LINK).Value)
        annotation = CheckRuleMatch(rowText, ruleDict)
        If Len(annotation) > 0 Then
            ' Append annotation to notes column (J)
            Dim notes As String
            notes = Trim(ws.Cells(i, COL_METADATA).Value)
            ws.Cells(i, COL_METADATA).Value = _
                IIf(Len(notes) > 0, notes & " | " & annotation, annotation)
        End If
    Next i

    ' --- STEP 3: Analyze free-text suggestions in name columns ---
    For i = 4 To lastRow
        Dim enVal As String, codeVal As String, rateVal As String
        enVal = Trim(ws.Cells(i, COL_NAME_EN_RAW).Value)
        codeVal = Trim(ws.Cells(i, COL_CODE_RAW).Value)

        ' Parse suggestion text if column contains any keyword
        Dim hasSuggestion As Boolean
        hasSuggestion = ContainsSuggestionKeyword(enVal) Or ContainsSuggestionKeyword(codeVal)

        If hasSuggestion Then
            Dim parsedName As String, parsedCode As String, parsedRate As String
            parsedName = ExtractSuggestion(enVal, "name")
            parsedCode = ExtractSuggestion(codeVal, "code")
            parsedRate = ExtractSuggestion(codeVal, "rate")

            ' Write structured suggestions
            If Len(parsedName) > 0 Then ws.Cells(i, COL_SUGGESTION_NAME).Value = parsedName
            If Len(parsedCode) > 0 Then ws.Cells(i, COL_SUGGESTION_CODE).Value = parsedCode
            If Len(parsedRate) > 0 Then ws.Cells(i, COL_RATE).Value = parsedRate

            ' Color-mark suggestion cells
            If Len(parsedName) > 0 Then ws.Cells(i, COL_NAME_EN_RAW).Interior.Color = CLR_SUGGESTED_BG
            If Len(parsedCode) > 0 Then ws.Cells(i, COL_CODE_RAW).Interior.Color = CLR_SUGGESTED_BG
        End If
    Next i
End Sub

' ================================================================
'  PHASE 2 — Process Suggestions + Generate Feedback Report
' ================================================================
Sub RunPipelinePhase2()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_CODE_RAW).End(xlUp).Row
    If lastRow < 4 Then Exit Sub

    ' --- STEP 1: Clear legacy green marks from previous runs ---
    Call ClearPreviousMarks(ws, lastRow)

    ' --- STEP 2: Sanitize entity links (col B) ---
    Dim i As Long
    For i = 4 To lastRow
        Dim linkVal As String
        linkVal = Trim(ws.Cells(i, COL_LINK).Value)
        If Len(linkVal) > 0 Then
            Dim sanitized As String
            sanitized = SanitizeLink(linkVal, TARGET_DOMAIN)
            If sanitized <> linkVal Then
                ws.Cells(i, COL_LINK).Value = sanitized
                ws.Cells(i, COL_LINK).Interior.Color = CLR_SUGGESTED_BG
            End If
        End If
    Next i

    ' --- STEP 3: Generate color-coded Feedback report sheet ---
    Call GenerateFeedbackReport(ws, lastRow)
End Sub

' ================================================================
'  CORE ALGORITHMS
' ================================================================

' --- CalculateBestCode: weighted frequency + recency scoring ---
'   score = freq_score * WEIGHT_FREQUENCY + recency_score * WEIGHT_RECENCY
'   Walks refSheet bottom-up: newer records → higher row number.
'   keyCol   = which column in ref sheet to match against (e.g. "A")
'   valueCol = which column in ref sheet holds the candidate value (e.g. "B")
Function CalculateBestCode(ByVal lookupTerm As String, _
                           ByVal refSheet As Worksheet, _
                           ByVal keyCol As String, _
                           ByVal valueCol As String) As String
    If Len(lookupTerm) = 0 Then CalculateBestCode = "": Exit Function

    Dim lastRow As Long, r As Long
    lastRow = refSheet.Cells(refSheet.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then CalculateBestCode = "": Exit Function

    Dim codeFreq As Object, codeRecency As Object
    Set codeFreq = CreateObject("Scripting.Dictionary")
    Set codeRecency = CreateObject("Scripting.Dictionary")

    ' Bottom-up walk — later records = newer
    For r = lastRow To 2 Step -1
        Dim keyVal As String, candVal As String
        keyVal = Trim(refSheet.Range(keyCol & r).Value)
        If Len(keyVal) = 0 Then GoTo CBN_Next

        ' Match: substring on lookup term OR on key column (fuzzy)
        Dim matched As Boolean
        matched = (InStr(1, keyVal, lookupTerm, vbTextCompare) > 0) Or _
                  (InStr(1, lookupTerm, keyVal, vbTextCompare) > 0)
        If Not matched Then GoTo CBN_Next

        candVal = Trim(refSheet.Range(valueCol & r).Value)
        If Len(candVal) = 0 Then GoTo CBN_Next

        If Not codeFreq.Exists(candVal) Then
            codeFreq.Add candVal, 0
            codeRecency.Add candVal, r
        End If
        codeFreq(candVal) = codeFreq(candVal) + 1

CBN_Next:
    Next r

    If codeFreq.Count = 0 Then CalculateBestCode = "": Exit Function

    ' Find maxes for normalization
    Dim maxFreq As Long, latestRow As Long
    maxFreq = 0: latestRow = 0
    Dim k As Variant
    For Each k In codeFreq.Keys
        If codeFreq(k) > maxFreq Then maxFreq = codeFreq(k)
        If codeRecency(k) > latestRow Then latestRow = codeRecency(k)
    Next k

    ' Weighted scoring
    Dim bestScore As Double, bestCode As String
    bestScore = -1: bestCode = ""
    For Each k In codeFreq.Keys
        Dim fs As Double, rs As Double, score As Double
        fs = (codeFreq(k) / maxFreq) * 100
        rs = ((codeRecency(k) - 2) / (lastRow - 2)) * 100
        score = fs * WEIGHT_FREQUENCY + rs * WEIGHT_RECENCY
        If score > bestScore Then
            bestScore = score
            bestCode = CStr(k)
        End If
    Next k

    CalculateBestCode = bestCode
End Function

' --- ExtractSuggestion: keyword-priority parsing from free text ---
'   type: "code" → extract first number sequence
'         "rate" → extract number before % sign
'         "name" → pick last alternative (split by NAME_SEPARATOR)
Function ExtractSuggestion(ByVal raw As String, ByVal colType As String) As String
    If Len(raw) = 0 Then ExtractSuggestion = "": Exit Function

    ' Find which keyword appeared, pick highest priority
    Dim matchedKw As String, matchedIdx As Long
    matchedKw = "": matchedIdx = 9999

    Dim arr As Variant, j As Long
    arr = primaryKeywords
    For j = LBound(arr) To UBound(arr)
        Dim pos As Long
        pos = InStr(1, raw, CStr(arr(j)), vbTextCompare)
        If pos > 0 And pos < matchedIdx Then
            matchedIdx = pos
            matchedKw = CStr(arr(j))
        End If
    Next j
    If Len(matchedKw) = 0 Then
        arr = secondaryKeywords
        For j = LBound(arr) To UBound(arr)
            pos = InStr(1, raw, CStr(arr(j)), vbTextCompare)
            If pos > 0 And pos < matchedIdx Then
                matchedIdx = pos
                matchedKw = CStr(arr(j))
            End If
        Next j
    End If

    If Len(matchedKw) = 0 Then ExtractSuggestion = "": Exit Function

    ' Extract the content AFTER the keyword
    Dim afterKw As String
    afterKw = Mid(raw, matchedIdx + Len(matchedKw))
    afterKw = Trim(afterKw)

    ' Parse based on column type
    Select Case colType
    Case "code"
        ExtractSuggestion = ExtractFirstNumber(afterKw)
    Case "rate"
        ExtractSuggestion = ExtractPercent(afterKw)
    Case "name"
        ExtractSuggestion = PickLastName(afterKw)
    Case Else
        ExtractSuggestion = afterKw
    End Select
End Function

Function ExtractFirstNumber(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = "\d+"
    If re.Test(s) Then
        ExtractFirstNumber = re.Execute(s)(0).Value
    Else
        ExtractFirstNumber = ""
    End If
End Function

Function ExtractPercent(ByVal s As String) As String
    Dim pctPos As Long
    pctPos = InStr(s, "%")
    If pctPos = 0 Then ExtractPercent = "": Exit Function

    Dim i As Long, start As Long
    start = pctPos - 1
    Do While start > 0 And (IsNumeric(Mid(s, start, 1)) Or Mid(s, start, 1) = ".")
        start = start - 1
    Loop
    ExtractPercent = Trim(Mid(s, start + 1, pctPos - start)) & "%"
End Function

Function PickLastName(ByVal s As String) As String
    If InStr(s, NAME_SEPARATOR) > 0 Then
        Dim parts As Variant
        parts = Split(s, NAME_SEPARATOR)
        PickLastName = Trim(CStr(parts(UBound(parts))))
    Else
        PickLastName = Trim(s)
    End If
End Function

' --- SanitizeLink: strip params → extract entity ID → normalize domain ---
Function SanitizeLink(ByVal url As String, ByVal targetDomain As String) As String
    If Len(url) = 0 Then SanitizeLink = "": Exit Function

    Dim result As String
    result = url

    ' Remove query string
    Dim qPos As Long
    qPos = InStr(result, "?")
    If qPos > 0 Then result = Left(result, qPos - 1)

    ' Remove hash
    Dim hPos As Long
    hPos = InStr(result, "#")
    If hPos > 0 Then result = Left(result, hPos - 1)

    ' Extract entity identifier (pattern: 1 uppercase + 9 alnum)
    Dim id As String
    id = ExtractEntityId(result)
    If Len(id) > 0 Then
        result = "https://" & targetDomain & "/dp/" & id
    End If

    SanitizeLink = result
End Function

Private Function ExtractEntityId(ByVal url As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = "[A-Z][0-9A-Z]{9}"
    If re.Test(url) Then
        ExtractEntityId = re.Execute(url)(0).Value
    Else
        ExtractEntityId = ""
    End If
End Function

' --- CheckRuleMatch: dictionary-based keyword → annotation ---
Function CheckRuleMatch(ByVal rowText As String, ByVal ruleDict As Object) As String
    Dim k As Variant
    For Each k In ruleDict.Keys
        If InStr(1, rowText, CStr(k), vbTextCompare) > 0 Then
            CheckRuleMatch = RULE_MARK & CStr(ruleDict(k))
            Exit Function
        End If
    Next k
    CheckRuleMatch = ""
End Function

' --- ContainsSuggestionKeyword ---
Function ContainsSuggestionKeyword(ByVal text As String) As Boolean
    Dim allKw As Variant, arr As Variant, j As Long
    allKw = Array()

    arr = primaryKeywords
    For j = LBound(arr) To UBound(arr)
        If InStr(1, text, CStr(arr(j)), vbTextCompare) > 0 Then
            ContainsSuggestionKeyword = True: Exit Function
        End If
    Next j
    arr = secondaryKeywords
    For j = LBound(arr) To UBound(arr)
        If InStr(1, text, CStr(arr(j)), vbTextCompare) > 0 Then
            ContainsSuggestionKeyword = True: Exit Function
        End If
    Next j
    ContainsSuggestionKeyword = False
End Function

' --- GetRateFromLookup: look up the percent parameter for a given value ---
Function GetRateFromLookup(ByVal code As String, ByVal lookupSheet As Worksheet) As String
    Dim lastRow As Long, r As Long
    lastRow = lookupSheet.Cells(lookupSheet.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        Dim rowVal As String
        rowVal = Trim(lookupSheet.Cells(r, 2).Value)
        If rowVal = code Then
            Dim rateVal As String
            rateVal = Trim(lookupSheet.Cells(r, 3).Value)
            If IsNumeric(rateVal) Then
                GetRateFromLookup = Format(CDbl(rateVal) * 100, "0.0") & "%"
            Else
                GetRateFromLookup = rateVal
            End If
            Exit Function
        End If
    Next r
    GetRateFromLookup = ""
End Function

' ================================================================
'  DESCRIPTION OPTIMIZATION — collapse / truncate / strip prefix
'  (used when comparing wrong vs right descriptions)
' ================================================================
Function OptimizeDescription(ByVal rawDesc As String, _
                             ByVal wrongDesc As String) As String
    Dim s As String
    s = rawDesc

    ' E1 — collapse repeated "other" keyword variants
    s = CollapseRepeatedOther(s)

    ' E2 — truncate non-numeric long clauses
    If Not ContainsDigit(s) Then
        Dim clauses As Variant
        clauses = Split(s, ";")
        If UBound(clauses) >= 1 Then
            s = clauses(0) & ";" & clauses(1)
        End If
        If Len(s) > MAX_TRUNCATE_LEN Then
            s = Left(s, MAX_TRUNCATE_LEN) & "..."
        End If
    End If

    ' E3 — strip shared prefix when comparing with wrong description
    If Len(wrongDesc) > 0 Then
        s = StripSharedPrefix(s, wrongDesc)
    End If

    OptimizeDescription = s
End Function

Private Function CollapseRepeatedOther(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(其他[\s>\-—]+)+其他"
    If re.Test(s) Then
        CollapseRepeatedOther = re.Replace(s, "其他")
    Else
        CollapseRepeatedOther = s
    End If
End Function

Private Function ContainsDigit(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d"
    ContainsDigit = re.Test(s)
End Function

Private Function StripSharedPrefix(ByVal right As String, ByVal wrong As String) As String
    Dim rightParts As Variant, wrongParts As Variant
    rightParts = Split(right, ">")
    wrongParts = Split(wrong, ">")

    Dim commonEnd As Long, idx As Long
    commonEnd = -1
    For idx = 0 To Application.WorksheetFunction.Min(UBound(rightParts), UBound(wrongParts))
        If Trim(CStr(rightParts(idx))) = Trim(CStr(wrongParts(idx))) Then
            commonEnd = idx
        Else
            Exit For
        End If
    Next idx

    If commonEnd >= 0 And commonEnd < UBound(rightParts) Then
        Dim remaining As String, k As Long
        For k = commonEnd + 1 To UBound(rightParts)
            If Len(remaining) > 0 Then remaining = remaining & " > "
            remaining = remaining & Trim(CStr(rightParts(k)))
        Next k
        StripSharedPrefix = remaining
    Else
        StripSharedPrefix = right
    End If
End Function

' --- NormalizeString: remove spaces, lowercase, strip common punctuation ---
Function NormalizeString(ByVal s As String) As String
    s = LCase(Trim(s))
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, ".", "")
    NormalizeString = s
End Function

' ================================================================
'  PHASE 2 SUPPORT — ClearPreviousMarks + GenerateFeedbackReport
' ================================================================
Sub ClearPreviousMarks(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim i As Long
    For i = 4 To lastRow
        ' Clear legacy green background on suggestion columns
        If ws.Cells(i, COL_NAME_EN_RAW).Interior.Color = CLR_PREVIOUS_MARK Then
            ws.Cells(i, COL_NAME_EN_RAW).Interior.Pattern = xlNone
        End If
        If ws.Cells(i, COL_CODE_RAW).Interior.Color = CLR_PREVIOUS_MARK Then
            ws.Cells(i, COL_CODE_RAW).Interior.Pattern = xlNone
        End If
    Next i
End Sub

Sub GenerateFeedbackReport(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim fbWS As Worksheet

    ' Create or replace Feedback sheet
    On Error Resume Next
    Set fbWS = ThisWorkbook.Sheets(SHEET_FEEDBACK)
    If Not fbWS Is Nothing Then
        Application.DisplayAlerts = False
        fbWS.Delete
        Application.DisplayAlerts = True
    End If
    Set fbWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    fbWS.Name = SHEET_FEEDBACK
    On Error GoTo 0

    ' --- Collect suggestions into 3 category buckets ---
    Dim onlyNameRows() As Variant, onlyCodeRows() As Variant, bothRows() As Variant
    Dim onlyNameCount As Long, onlyCodeCount As Long, bothCount As Long
    Dim estMax As Long
    estMax = lastRow - 3  ' rows 4..lastRow
    If estMax < 100 Then estMax = 100

    ReDim onlyNameRows(1 To estMax, 1 To 5)   ' lookup | suggested name | suggested code | metadata | link
    ReDim onlyCodeRows(1 To estMax, 1 To 5)
    ReDim bothRows(1 To estMax, 1 To 5)
    onlyNameCount = 0: onlyCodeCount = 0: bothCount = 0

    Dim i As Long
    For i = 4 To lastRow
        Dim lookupVal As String, sugName As String, sugCode As String, _
            metaVal As String, linkVal As String
        lookupVal = Trim(ws.Cells(i, COL_LOOKUP).Value)
        sugName   = Trim(ws.Cells(i, COL_SUGGESTION_NAME).Value)
        sugCode   = Trim(ws.Cells(i, COL_SUGGESTION_CODE).Value)
        metaVal   = Trim(ws.Cells(i, COL_METADATA).Value)
        linkVal   = Trim(ws.Cells(i, COL_LINK).Value)

        Dim hasName As Boolean, hasCode As Boolean
        hasName = (Len(sugName) > 0)
        hasCode = (Len(sugCode) > 0)

        If Not (hasName Or hasCode) Then GoTo GFB_Next

        Dim entry(1 To 5) As Variant
        entry(1) = lookupVal: entry(2) = sugName: entry(3) = sugCode
        entry(4) = metaVal:   entry(5) = linkVal

        If hasName And hasCode Then
            bothCount = bothCount + 1
            fbWS.Rows(bothCount + 3).Value = entry   ' +3 to offset from header + category rows
        ElseIf hasName Then
            onlyNameCount = onlyNameCount + 1
            fbWS.Rows(onlyNameCount + 3).Value = entry
        ElseIf hasCode Then
            onlyCodeCount = onlyCodeCount + 1
            fbWS.Rows(onlyCodeCount + 3).Value = entry
        End If

GFB_Next:
    Next i

    ' --- Write header row ---
    fbWS.Range("A1:E1").Value = Array( _
        "Lookup / Group", "Suggested Name", "Suggested Code / Value", _
        "Metadata / Notes", "Sanitized Link")
    fbWS.Range("A1:E1").Font.Bold = True
    fbWS.Range("A1:E1").Interior.Color = CLR_HEADER_BG
    fbWS.Range("A1:E1").HorizontalAlignment = xlCenter

    ' --- Write data in priority order: name-only → code-only → both ---
    Dim writeRow As Long
    writeRow = 2

    ' Name-only
    If onlyNameCount > 0 Then
        fbWS.Cells(writeRow, 1).Value = "=== Suggested Name Only ==="
        fbWS.Cells(writeRow, 1).Font.Bold = True
        writeRow = writeRow + 1
        Dim ni As Long
        For ni = 1 To onlyNameCount
            fbWS.Rows(writeRow).Value = onlyNameRows(ni)
            writeRow = writeRow + 1
        Next ni
    End If

    ' Code-only
    If onlyCodeCount > 0 Then
        fbWS.Cells(writeRow, 1).Value = "=== Suggested Code Only ==="
        fbWS.Cells(writeRow, 1).Font.Bold = True
        writeRow = writeRow + 1
        Dim ci As Long
        For ci = 1 To onlyCodeCount
            fbWS.Rows(writeRow).Value = onlyCodeRows(ci)
            writeRow = writeRow + 1
        Next ci
    End If

    ' Both
    If bothCount > 0 Then
        fbWS.Cells(writeRow, 1).Value = "=== Both Name & Code Suggested ==="
        fbWS.Cells(writeRow, 1).Font.Bold = True
        writeRow = writeRow + 1
        Dim bi As Long
        For bi = 1 To bothCount
            fbWS.Rows(writeRow).Value = bothRows(bi)
            writeRow = writeRow + 1
        Next bi
    End If

    ' --- Deduplicate: exact duplicate rows removed from bottom up ---
    Dim finalLast As Long
    finalLast = writeRow - 1
    If finalLast >= 3 Then
        Dim r As Long
        For r = finalLast To 3 Step -1
            If fbWS.Cells(r, 1).Value = fbWS.Cells(r - 1, 1).Value And _
               fbWS.Cells(r, 2).Value = fbWS.Cells(r - 1, 2).Value And _
               fbWS.Cells(r, 3).Value = fbWS.Cells(r - 1, 3).Value Then
                fbWS.Rows(r).Delete
            End If
        Next r
    End If

    ' --- Feedback sheet formatting ---
    finalLast = fbWS.Cells(fbWS.Rows.Count, 1).End(xlUp).Row
    If finalLast < 2 Then finalLast = 1

    With fbWS
        .Range("A1:E" & finalLast).Borders.LineStyle = xlContinuous
        .Range("A1:E" & finalLast).WrapText = True
        .Range("A1:E" & finalLast).HorizontalAlignment = xlCenter
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 26
        .Columns("C").ColumnWidth = 40
        .Columns("D").ColumnWidth = 20
        .Columns("E").ColumnWidth = 36
        .Rows("1:1").Select
        ActiveWindow.FreezePanes = True
        .Range("A1:E1").AutoFilter
        .Cells(1, 1).Select
    End With

    fbWS.Activate
End Sub
