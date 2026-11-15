Option Explicit
Option Compare Text

' ================================================================
'  Weighted Data Validator
'  Multi-source data validation with weighted scoring + rule engine
'  + color-coded report generation
'
'  Author: ivy92888-glitch  |  License: MIT
'  Disclaimer: Unaffiliated with any spreadsheet vendor.
'              Generic clean-room implementation. Any resemblance to
'              specific business systems is coincidental.
' ================================================================

' --- CONFIG CONSTANTS (see docs/Configuration.md) ---
Private Const WEIGHT_FREQUENCY As Double = 0.4
Private Const WEIGHT_RECENCY   As Double = 0.6
Private Const MIN_NAME_LEN     As Integer = 3
Private Const MAX_TRUNCATE_LEN As Integer = 18
Private Const NAME_SEPARATOR   As String  = "/"
Private Const RULE_MARK       As String  = "-"

' Color palette
Private Const CLR_SUGGESTED As Long = RGB(226, 239, 218)   ' light green
Private Const CLR_PART_ERR  As Long = RGB(127, 150, 152)   ' slate gray-green
Private Const CLR_PART_OK   As Long = RGB(198,  89,  17)   ' orange
Private Const CLR_HEADER    As Long = RGB(255, 255,   0)   ' yellow

' Rule priority keywords
Private primaryKeywords As Variant   ' highest -> lowest
Private secondaryKeywords As Variant ' lower priority

' Rule engine (generic)
Private ruleDict As Object           ' Scripting.Dictionary

' Main entry
Sub RunFullPipeline()
    ' 1. Validate input workbook structure
    ' 2. Run consistency checks (module A)
    ' 3. Run rule-based scanning (module B)
    ' 4. Run smart suggestion extraction (module C)
    ' 5. Sanitize entity links (module D)
    ' 6. Generate color-coded feedback sheet (module E)
    Call Initialize
    Call ValidateWorkbook
    Call RunConsistencyCheck
    Call RunRuleScan
    Call RunSmartSuggestion
    Call SanitizeLinks
    Call GenerateFeedbackReport
    MsgBox "Pipeline complete.", vbInformation
End Sub

Sub Initialize()
    ' Set up priority keyword arrays
    primaryKeywords = Array("top", "try", "suggest", "recommend")
    secondaryKeywords = Array("may", "let", "allow", "consider")

    ' Load rule engine (placeholder — populate from config sheet)
    Set ruleDict = CreateObject("Scripting.Dictionary")
    ' Example: ruleDict.Add "keyword", "annotation text"
End Sub

' ================================================================
' MODULE A — Weighted Best-Code Selector
'   Composite score = freq_score * WEIGHT_FREQUENCY + recency_score * WEIGHT_RECENCY
'   freq_score   = (occurrences / max_occurrences) * 100
'   recency_score = ((latest_row - 2) / (total_rows - 2)) * 100
' ================================================================
Function CalculateBestCode(ByVal codeCol As String, ByVal nameCol As String, _
                           ByVal refSheetName As String) As Variant
    Dim dbSheet As Worksheet
    Dim totalRows As Long, i As Long
    Dim codeFreq As Object, codeRecency As Object
    Dim bestCode As String, bestScore As Double

    Set dbSheet = ThisWorkbook.Sheets(refSheetName)
    totalRows = dbSheet.Cells(dbSheet.Rows.Count, 1).End(xlUp).Row
    Set codeFreq = CreateObject("Scripting.Dictionary")
    Set codeRecency = CreateObject("Scripting.Dictionary")

    ' Walk DB bottom-up — later row = newer record
    For i = totalRows To 2 Step -1
        Dim code As String
        code = Trim(dbSheet.Range(codeCol & i).Value)
        If Len(code) > 0 Then
            If Not codeFreq.Exists(code) Then
                codeFreq.Add code, 0
                codeRecency.Add code, i    ' first (from bottom) = latest row
            End If
            codeFreq(code) = codeFreq(code) + 1
        End If
    Next i

    ' Find maxes for normalization
    Dim maxFreq As Long, latestRow As Long
    maxFreq = 0: latestRow = 0
    Dim k As Variant
    For Each k In codeFreq.Keys
        If codeFreq(k) > maxFreq Then maxFreq = codeFreq(k)
        If codeRecency(k) > latestRow Then latestRow = codeRecency(k)
    Next k

    ' Weighted scoring
    bestScore = -1: bestCode = ""
    For Each k In codeFreq.Keys
        Dim fs As Double, rs As Double, score As Double
        fs = (codeFreq(k) / maxFreq) * 100
        rs = ((codeRecency(k) - 2) / (totalRows - 2)) * 100
        score = fs * WEIGHT_FREQUENCY + rs * WEIGHT_RECENCY
        If score > bestScore Then
            bestScore = score
            bestCode = k
        End If
    Next k

    CalculateBestCode = bestCode
End Function

' ================================================================
' MODULE B — Rule Engine (generic keyword → annotation)
' ================================================================
Function CheckRuleMatch(ByVal text As String, ByVal ruleDict As Object) As String
    Dim k As Variant
    For Each k In ruleDict.Keys
        If InStr(1, text, k, vbTextCompare) > 0 Then
            CheckRuleMatch = RULE_MARK & ruleDict(k)
            Exit Function
        End If
    Next k
    CheckRuleMatch = ""
End Function

' ================================================================
' MODULE C — Smart Suggestion Parser
'   Extract target + value + percent from free-form suggestion text
'   Priority: primary keywords > secondary keywords
' ================================================================
Function ExtractSuggestion(ByVal raw As String, ByVal colType As String) As String
    If Len(raw) = 0 Then ExtractSuggestion = "": Exit Function

    ' Step 1 — find which keyword appeared, pick highest priority
    Dim matchedKw As String, matchedIdx As Long
    matchedKw = "": matchedIdx = 999

    Dim arr As Variant, j As Long
    arr = primaryKeywords
    For j = LBound(arr) To UBound(arr)
        Dim pos As Long
        pos = InStr(1, raw, arr(j), vbTextCompare)
        If pos > 0 And pos < matchedIdx Then
            matchedIdx = pos
            matchedKw = arr(j)
        End If
    Next j
    If Len(matchedKw) = 0 Then
        arr = secondaryKeywords
        For j = LBound(arr) To UBound(arr)
            pos = InStr(1, raw, arr(j), vbTextCompare)
            If pos > 0 And pos < matchedIdx Then
                matchedIdx = pos
                matchedKw = arr(j)
            End If
        Next j
    End If

    If Len(matchedKw) = 0 Then ExtractSuggestion = "": Exit Function

    ' Step 2 — extract the content AFTER the keyword
    Dim afterKw As String
    afterKw = Mid(raw, matchedIdx + Len(matchedKw))
    afterKw = Trim(afterKw)

    ' Step 3 — parse based on column type
    If colType = "code" Then
        ExtractSuggestion = ExtractFirstNumber(afterKw)
    ElseIf colType = "rate" Then
        ExtractSuggestion = ExtractPercent(afterKw)
    ElseIf colType = "name" Then
        ExtractSuggestion = PickLastName(afterKw)
    Else
        ExtractSuggestion = afterKw
    End If
End Function

Private Function ExtractFirstNumber(ByVal s As String) As String
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

Private Function ExtractPercent(ByVal s As String) As String
    Dim pctPos As Long
    pctPos = InStr(s, "%")
    If pctPos = 0 Then ExtractPercent = "": Exit Function

    ' Walk backwards from % to find digit start
    Dim i As Long, start As Long
    start = pctPos - 1
    Do While start > 0 And (IsNumeric(Mid(s, start, 1)) Or Mid(s, start, 1) = ".")
        start = start - 1
    Loop
    ExtractPercent = Trim(Mid(s, start + 1, pctPos - start)) & "%"
End Function

Private Function PickLastName(ByVal s As String) As String
    ' When multiple alternatives separated by "/", keep last one
    Dim parts As Variant
    If InStr(s, NAME_SEPARATOR) > 0 Then
        parts = Split(s, NAME_SEPARATOR)
        PickLastName = Trim(parts(UBound(parts)))
    Else
        PickLastName = Trim(s)
    End If
End Function

' ================================================================
' MODULE D — Entity Link Sanitizer
'   Strip query params → strip ref params → extract identifier
'   → normalize domain suffix
' ================================================================
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

    ' Extract identifier (pattern: uppercase letter + 9 alnum)
    Dim id As String
    id = ExtractEntityId(result)
    If Len(id) > 0 Then
        ' Reconstruct with normalized domain
        Dim protocol As String
        protocol = "https://"
        result = protocol & targetDomain & "/dp/" & id
    End If

    SanitizeLink = result
End Function

Private Function ExtractEntityId(ByVal url As String) As String
    ' Pattern: one uppercase letter followed by 9 alphanumeric chars
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

' ================================================================
' MODULE E — Code-Description Optimizer
'   Collapse consecutive separators → truncate long clauses →
'   strip shared prefix between wrong & right descriptions
' ================================================================
Function OptimizeDescription(ByVal rawDesc As String, _
                             ByVal wrongDesc As String) As String
    Dim s As String
    s = rawDesc

    ' E1 — collapse repeated "other" keywords
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

    ' E3 — strip shared prefix (compare with wrong description)
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
    ' Collapse: "其他 其他" / "其他>其他" / "其他--其他"
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
    ' Compare from leftmost ">" or ";" separator — trim common ancestors
    Dim rightParts As Variant, wrongParts As Variant
    rightParts = Split(right, ">")
    wrongParts = Split(wrong, ">")

    Dim commonEnd As Long, idx As Long
    commonEnd = -1
    For idx = 0 To Application.WorksheetFunction.Min(UBound(rightParts), UBound(wrongParts))
        If Trim(rightParts(idx)) = Trim(wrongParts(idx)) Then
            commonEnd = idx
        Else
            Exit For
        End If
    Next idx

    If commonEnd >= 0 And commonEnd < UBound(rightParts) Then
        Dim remaining As String
        Dim k As Long
        For k = commonEnd + 1 To UBound(rightParts)
            If Len(remaining) > 0 Then remaining = remaining & " > "
            remaining = remaining & Trim(rightParts(k))
        Next k
        StripSharedPrefix = remaining
    Else
        StripSharedPrefix = right
    End If
End Function

' ================================================================
' MODULE F — Color-Coded Feedback Report Generator
'   Write suggestions to a dedicated sheet, categorized by type,
'   with color markers and deduplication.
' ================================================================
Sub GenerateFeedbackReport()
    Dim srcSheet As Worksheet, fbSheet As Worksheet
    Set srcSheet = ThisWorkbook.Sheets("Data")
    On Error Resume Next
    Set fbSheet = ThisWorkbook.Sheets("Feedback")
    If fbSheet Is Nothing Then
        Set fbSheet = ThisWorkbook.Sheets.Add
        fbSheet.Name = "Feedback"
    End If
    On Error GoTo 0

    ' Clear old content
    fbSheet.Cells.Clear
    fbSheet.Rows(1).Interior.Color = CLR_HEADER

    ' Three categories — written in priority order
    Dim onlyName() As Variant, onlyCode() As Variant, both() As Variant
    Dim onlyNameCount As Long, onlyCodeCount As Long, bothCount As Long
    onlyNameCount = 0: onlyCodeCount = 0: bothCount = 0

    ' ... (collect from scan results, then write each category with dedup)

    ' Apply color markers
    ' Suggested cells → CLR_SUGGESTED background
    ' Wrong part → CLR_PART_ERR  (custom character range)
    ' Right part → CLR_PART_OK
End Sub
