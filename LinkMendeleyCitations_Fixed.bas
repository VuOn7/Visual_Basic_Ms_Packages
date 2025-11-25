Option Explicit

' Paste this entire module into a new Module in Word VBA and run LinkMendeleyCitations_Try2
' Run on a copy of your document. Check Immediate window (Ctrl+G) for diagnostics.
' FIXED: Now properly handles multiple citations in a single bracket (e.g., "Aryal et al., 2025; Paudel et al., 2020")

Public Sub LinkMendeleyCitations_Try2()
    Dim doc As Document: Set doc = ActiveDocument
    Dim t0 As Single: t0 = Timer
    Debug.Print "=== Start LinkMendeleyCitations_Try2: " & Now & " ==="

    ' find bibliography cc or References heading
    Dim bibCC As ContentControl: Set bibCC = Nothing
    Dim cc As ContentControl
    For Each cc In doc.ContentControls
        If LCase$(cc.Tag & "") Like "*mendeley_bibliography*" Then
            Set bibCC = cc: Exit For
        End If
    Next cc

    Dim refsRange As Range
    If Not bibCC Is Nothing Then
        Set refsRange = bibCC.Range.Duplicate
        Debug.Print "Using bibliography CC at start=" & refsRange.Start
    Else
        Dim rpos As Long: rpos = FindReferencesStart(doc)
        If rpos > 0 Then
            Set refsRange = doc.Range(rpos, doc.Content.End)
            Debug.Print "Using 'References' heading start=" & rpos
        Else
            Set refsRange = Nothing
            Debug.Print "No explicit bibliography range found — fallback used when building bookmarks."
        End If
    End If

    ' maps
    Dim surnameMap As Object: Set surnameMap = CreateObject("Scripting.Dictionary") ' surname -> Collection(bookmark)
    Dim surnameYearMap As Object: Set surnameYearMap = CreateObject("Scripting.Dictionary") ' surname|year -> bookmark
    Dim bookmarkTextMap As Object: Set bookmarkTextMap = CreateObject("Scripting.Dictionary") ' bookmark -> full ref line
    Dim refKeyMap As Object: Set refKeyMap = CreateObject("Scripting.Dictionary") ' normalized short keys -> bookmark

    BuildReferenceBookmarksAndMaps2 doc, refsRange, surnameMap, surnameYearMap, bookmarkTextMap, refKeyMap
    Debug.Print "Reference bookmarks built: " & bookmarkTextMap.Count & " ; refKeyMap size: " & refKeyMap.Count

    ' collect citation CCs (exclude those inside refsRange)
    Dim citeCCs As New Collection
    For Each cc In doc.ContentControls
        If LCase$(cc.Tag & "") Like "*mendeley_citation*" Or LCase$(cc.Tag & "") Like "*csl_citation*" Or InStr(1, cc.Tag & "", "MENDELEY_CITATION", vbTextCompare) > 0 Then
            If Not refsRange Is Nothing Then
                If cc.Range.Start < refsRange.Start Or cc.Range.Start > refsRange.End Then citeCCs.Add cc
            Else
                citeCCs.Add cc
            End If
        End If
    Next cc
    Debug.Print "Citation CCs to process: " & citeCCs.Count

    ' collect ADDIN fields (fallback)
    Dim fieldList As New Collection
    Dim f As Field
    For Each f In doc.Fields
        If f.Type = wdFieldAddin Then
            If InStr(1, f.Code.Text, "ADDIN CSL_CITATION", vbTextCompare) > 0 Or InStr(1, f.Code.Text, "MENDELEY_CITATION", vbTextCompare) > 0 Then
                If Not refsRange Is Nothing Then
                    If f.Result.Start < refsRange.Start Or f.Result.Start > refsRange.End Then fieldList.Add f
                Else
                    fieldList.Add f
                End If
            End If
        End If
    Next f
    Debug.Print "Citation fields to process: " & fieldList.Count

    Dim created As Long: created = 0
    Dim unmatched As Object: Set unmatched = CreateObject("Scripting.Dictionary")

    ' - Process CCs
    Dim i As Long
    For i = 1 To citeCCs.Count
        Dim cci As ContentControl: Set cci = citeCCs(i)
        Dim raw As String: raw = cci.Range.Text
        If Len(Trim$(raw)) = 0 Then GoTo NextCC
        
        ' Split by semicolon to get individual citations
        Dim tokens() As String: tokens = Split(raw, ";")
        Dim tokenCount As Long: tokenCount = UBound(tokens) - LBound(tokens) + 1
        
        ' Pre-calculate all token info BEFORE adding any hyperlinks
        Dim tokenInfos() As Variant
        ReDim tokenInfos(LBound(tokens) To UBound(tokens))
        
        Dim j As Long
        Dim searchPos As Long: searchPos = 1
        
        ' First pass: gather all token information
        For j = LBound(tokens) To UBound(tokens)
            Dim tokRaw As String: tokRaw = Trim$(tokens(j))
            Dim tokNorm As String: tokNorm = NormalizeToken(tokRaw)
            
            ' Create info array: (0)=tokRaw, (1)=tokNorm, (2)=pos, (3)=surname, (4)=year, (5)=bookmark
            Dim info(0 To 5) As Variant
            info(0) = tokRaw
            info(1) = tokNorm
            info(2) = 0 ' position - will be calculated
            info(3) = "" ' surname
            info(4) = "" ' year
            info(5) = "" ' bookmark
            
            If Len(tokRaw) = 0 Then
                tokenInfos(j) = info
                GoTo NextTokenPrep
            End If
            
            ' attempt tag extraction
            Dim bestSurname As String: bestSurname = ""
            Dim bestYear As String: bestYear = ""
            Dim tagStr As String: tagStr = cci.Tag & ""
            If Len(tagStr) > 0 Then
                bestSurname = TryGetSurnameFromTag(tagStr, j)
                bestYear = TryGetYearFromTag(tagStr, j)
                If Len(bestSurname) > 0 Then Debug.Print "Tag-surname for CCStart=" & cci.Range.Start & " tokenIdx=" & j & " => " & bestSurname
            End If

            ' fallback parsing
            If Len(bestSurname) = 0 Then bestSurname = ExtractFirstAuthorFromCitation(tokNorm)
            If Len(bestYear) = 0 Then bestYear = ExtractYearFromCitationText(tokNorm)
            
            info(3) = bestSurname
            info(4) = bestYear

            If Len(bestSurname) = 0 Then
                unmatched("CC:" & cci.Range.Start & ":" & j) = tokRaw
                Debug.Print "No surname extracted for token [" & tokRaw & "] CCStart=" & cci.Range.Start
                tokenInfos(j) = info
                GoTo NextTokenPrep
            End If

            ' Find candidate bookmark using enhanced resolver
            Dim candBk As String
            candBk = ResolveCandidate2(surnameMap, surnameYearMap, bookmarkTextMap, refKeyMap, bestSurname, bestYear, tokRaw, tokNorm)
            info(5) = candBk

            ' locate token in CC visible text robustly
            Dim pos As Long: pos = FindTokenPosition(searchPos, raw, tokRaw)
            If pos = 0 Then pos = FindTokenPosition(searchPos, raw, tokNorm)
            If pos = 0 Then pos = FindTokenPosition(searchPos, raw, TrimPunctuation(tokNorm))
            
            info(2) = pos
            
            If pos > 0 Then
                searchPos = pos + Len(tokRaw)
            End If
            
            tokenInfos(j) = info
NextTokenPrep:
        Next j
        
        ' Second pass: Add hyperlinks in REVERSE order to avoid position shifts
        For j = UBound(tokens) To LBound(tokens) Step -1
            Dim tInfo As Variant: tInfo = tokenInfos(j)
            Dim tRaw As String: tRaw = tInfo(0)
            Dim tNorm As String: tNorm = tInfo(1)
            Dim tPos As Long: tPos = tInfo(2)
            Dim tSurname As String: tSurname = tInfo(3)
            Dim tYear As String: tYear = tInfo(4)
            Dim tBookmark As String: tBookmark = tInfo(5)
            
            If Len(tRaw) = 0 Then GoTo NextTokenAdd
            If tPos = 0 Then
                If Len(tSurname) > 0 Then
                    unmatched("CCnoloc:" & cci.Range.Start & ":" & j) = tRaw
                    Debug.Print "Could not locate token text [" & tRaw & "] inside CC visible text at CCStart=" & cci.Range.Start
                End If
                GoTo NextTokenAdd
            End If
            
            ' Re-read the CC range start (it may have shifted from previous hyperlinks in this loop)
            Dim currentCCStart As Long: currentCCStart = cci.Range.Start
            
            Dim r As Range: Set r = cci.Range.Duplicate
            r.Start = currentCCStart + tPos - 1
            r.End = r.Start + Len(tRaw)
            
            ' Validate range is within CC
            If r.Start < currentCCStart Or r.End > cci.Range.End Then
                Debug.Print "Range out of bounds for token [" & tRaw & "] - skipping"
                GoTo NextTokenAdd
            End If
            
            If r.Hyperlinks.Count = 0 Then
                If Len(tBookmark) > 0 Then
                    On Error Resume Next
                    doc.Hyperlinks.Add Anchor:=r, Address:="", SubAddress:=tBookmark, TextToDisplay:=r.Text
                    If Err.Number = 0 Then
                        created = created + 1
                        Debug.Print "Linked CC token -> " & tBookmark & " | CCStart=" & currentCCStart & " token=" & Left$(tRaw, 80)
                    Else
                        unmatched("CCaddErr:" & cci.Range.Start & ":" & j) = tRaw
                        Debug.Print "Err adding hyperlink: " & Err.Number & " " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                Else
                    unmatched("CCnomatch:" & cci.Range.Start & ":" & j) = tRaw
                    Debug.Print "No candidate bookmark for seg: [" & tRaw & "] at CCStart=" & currentCCStart
                End If
            Else
                Debug.Print "Already hyperlinked (skip). CCStart=" & currentCCStart & " token=" & Left$(tRaw, 60)
            End If
NextTokenAdd:
        Next j
NextCC:
    Next i

    ' - Process fields (similar logic with reverse processing)
    For i = 1 To fieldList.Count
        Dim fld As Field: Set fld = fieldList(i)
        Dim fRaw As String: fRaw = fld.Result.Text
        If Len(Trim$(fRaw)) = 0 Then GoTo NextField
        
        Dim fTokens() As String: fTokens = Split(fRaw, ";")
        Dim fTokenInfos() As Variant
        ReDim fTokenInfos(LBound(fTokens) To UBound(fTokens))
        
        Dim fSearch As Long: fSearch = 1
        Dim k As Long
        
        ' First pass: gather all token info
        For k = LBound(fTokens) To UBound(fTokens)
            Dim fTokRaw As String: fTokRaw = Trim$(fTokens(k))
            Dim fTokNorm As String: fTokNorm = NormalizeToken(fTokRaw)
            
            Dim fInfo(0 To 5) As Variant
            fInfo(0) = fTokRaw
            fInfo(1) = fTokNorm
            fInfo(2) = 0
            fInfo(3) = ""
            fInfo(4) = ""
            fInfo(5) = ""
            
            If Len(fTokRaw) = 0 Then
                fTokenInfos(k) = fInfo
                GoTo NextFTPrep
            End If
            
            Dim fSurname As String: fSurname = ExtractFirstAuthorFromCitation(fTokNorm)
            Dim fYear As String: fYear = ExtractYearFromCitationText(fTokNorm)
            
            fInfo(3) = fSurname
            fInfo(4) = fYear
            
            If Len(fSurname) = 0 Then
                unmatched("Fldnm:" & fld.Result.Start & ":" & k) = fTokRaw
                fTokenInfos(k) = fInfo
                GoTo NextFTPrep
            End If

            Dim fCand As String
            fCand = ResolveCandidate2(surnameMap, surnameYearMap, bookmarkTextMap, refKeyMap, fSurname, fYear, fTokRaw, fTokNorm)
            fInfo(5) = fCand

            Dim posf As Long: posf = FindTokenPosition(fSearch, fRaw, fTokRaw)
            If posf = 0 Then posf = FindTokenPosition(fSearch, fRaw, fTokNorm)
            
            fInfo(2) = posf
            
            If posf > 0 Then
                fSearch = posf + Len(fTokRaw)
            End If
            
            fTokenInfos(k) = fInfo
NextFTPrep:
        Next k
        
        ' Second pass: Add hyperlinks in REVERSE order
        For k = UBound(fTokens) To LBound(fTokens) Step -1
            Dim fTInfo As Variant: fTInfo = fTokenInfos(k)
            Dim ftRaw As String: ftRaw = fTInfo(0)
            Dim ftNorm As String: ftNorm = fTInfo(1)
            Dim ftPos As Long: ftPos = fTInfo(2)
            Dim ftSurname As String: ftSurname = fTInfo(3)
            Dim ftYear As String: ftYear = fTInfo(4)
            Dim ftBookmark As String: ftBookmark = fTInfo(5)
            
            If Len(ftRaw) = 0 Then GoTo NextFTAdd
            If ftPos = 0 Then
                If Len(ftSurname) > 0 Then
                    unmatched("FldNoLoc:" & fld.Result.Start & ":" & k) = ftRaw
                End If
                GoTo NextFTAdd
            End If
            
            Dim currentFldStart As Long: currentFldStart = fld.Result.Start
            
            Dim rf As Range: Set rf = fld.Result.Duplicate
            rf.Start = currentFldStart + ftPos - 1
            rf.End = rf.Start + Len(ftRaw)
            
            If rf.Hyperlinks.Count = 0 Then
                If Len(ftBookmark) > 0 Then
                    On Error Resume Next
                    doc.Hyperlinks.Add Anchor:=rf, Address:="", SubAddress:=ftBookmark, TextToDisplay:=rf.Text
                    If Err.Number = 0 Then
                        created = created + 1
                    Else
                        unmatched("FldAddErr:" & fld.Result.Start & ":" & k) = ftRaw
                        Err.Clear
                    End If
                    On Error GoTo 0
                Else
                    unmatched("FldNoMatch:" & fld.Result.Start & ":" & k) = ftRaw
                End If
            End If
NextFTAdd:
        Next k
NextField:
    Next i

    ' Summary
    Debug.Print "=== Summary ==="
    Debug.Print "Bookmarks (unique): " & bookmarkTextMap.Count
    Debug.Print "CCs processed: " & citeCCs.Count
    Debug.Print "Fields processed: " & fieldList.Count
    Debug.Print "Hyperlinks created: " & created
    Debug.Print "Unmatched tokens: " & unmatched.Count
    Debug.Print "Elapsed sec: " & Format$(Timer - t0, "0.0")
    Dim msg As String
    msg = "Done." & vbCrLf & _
          "Bookmarks (unique): " & bookmarkTextMap.Count & vbCrLf & _
          "CCs processed: " & citeCCs.Count & vbCrLf & _
          "Fields processed: " & fieldList.Count & vbCrLf & _
          "Hyperlinks created: " & created & vbCrLf & _
          "Unmatched tokens: " & unmatched.Count
    MsgBox msg, vbInformation, "LinkMendeleyCitations_Try2"

    If unmatched.Count > 0 Then
        Debug.Print "---- Unmatched sample keys (up to 80) ----"
        Dim keyOut As Variant, outC As Long: outC = 0
        For Each keyOut In unmatched.Keys
            Debug.Print keyOut & " -> " & Left$(unmatched(keyOut), 240)
            outC = outC + 1
            If outC >= 80 Then Exit For
        Next keyOut
    End If

    Debug.Print "=== End ==="
End Sub


' ----------------------------
' Build bookmarks and extra keys (improved)
' ----------------------------
Private Sub BuildReferenceBookmarksAndMaps2(doc As Document, refsRange As Range, _
    ByRef surnameMap As Object, ByRef surnameYearMap As Object, ByRef bookmarkTextMap As Object, ByRef refKeyMap As Object)

    On Error Resume Next
    Dim r As Range
    If refsRange Is Nothing Then
        Dim posStart As Long: posStart = doc.Content.End - (doc.Content.End \ 4)
        If posStart < 1 Then posStart = 1
        Set r = doc.Range(posStart, doc.Content.End)
        Debug.Print "Fallback refs range used. Start=" & posStart
    Else
        Set r = refsRange.Duplicate
    End If

    Dim p As Paragraph
    For Each p In r.Paragraphs
        Dim line As String: line = Trim$(p.Range.Text)
        If Len(line) < 5 Then GoTo NextPara
        If Len(Replace(line, vbCr, "")) < 5 Then GoTo NextPara

        ' compute surname (first element before comma) or first word
        Dim surname As String
        Dim commaPos As Long: commaPos = InStr(1, line, ",")
        If commaPos > 1 Then
            surname = Trim$(Left$(line, commaPos - 1))
        Else
            surname = Trim$(Split(line, " ")(0))
        End If
        surname = OnlyLetters(surname)
        If Len(surname) = 0 Then GoTo NextPara

        Dim yearVal As String: yearVal = ExtractYearFromCitationText(line)

        Dim baseName As String: baseName = "Ref_" & surname
        If Len(baseName) > 40 Then baseName = Left$(baseName, 40)
        Dim uniqueName As String: uniqueName = baseName
        Dim suffix As Long: suffix = 1
        Do While BookmarkExistsInDoc(doc, uniqueName) Or bookmarkTextMap.Exists(uniqueName)
            suffix = suffix + 1
            uniqueName = Left$(baseName, 30) & "_" & CStr(suffix)
            If Len(uniqueName) > 40 Then uniqueName = Left$(uniqueName, 40)
        Loop

        ' add bookmark at paragraph start
        Dim bR As Range: Set bR = p.Range.Duplicate
        bR.Collapse Direction:=wdCollapseStart
        doc.Bookmarks.Add name:=uniqueName, Range:=bR
        If Err.Number = 0 Then
            bookmarkTextMap.Add uniqueName, line

            ' surnameMap
            If Not surnameMap.Exists(surname) Then
                Dim c As Collection: Set c = New Collection
                c.Add uniqueName
                surnameMap.Add surname, c
            Else
                surnameMap(surname).Add uniqueName
            End If

            ' surname|year
            If Len(yearVal) = 4 Then
                Dim keySY As String: keySY = surname & "|" & yearVal
                If Not surnameYearMap.Exists(keySY) Then surnameYearMap.Add keySY, uniqueName
            End If

            ' create multiple short normalized keys to help matching of corporate/report refs or first-words
            Dim normAll As String: normAll = LCase$(Trim$(TrimPunctuation(OnlyAlphaNumericShort(line))))
            If Len(normAll) >= 6 Then
                If Not refKeyMap.Exists(normAll) Then refKeyMap.Add normAll, uniqueName
            End If

            ' also add first 2-4 words normalized as keys
            Dim w() As String: w = Split(StripPunctuation(line), " ")
            Dim num As Long, kk As Long
            If UBound(w) >= 0 Then
                num = UBound(w) + 1
                If num > 4 Then num = 4
            Else
                num = 0
            End If
            Dim sKey As String: sKey = ""
            For kk = 0 To num - 1
                If Len(w(kk)) > 0 Then
                    If sKey = "" Then sKey = LCase$(w(kk)) Else sKey = sKey & " " & LCase$(w(kk))
                    Dim sKeyNorm As String: sKeyNorm = Trim$(TrimPunctuation(OnlyAlphaNumericShort(sKey)))
                    If Len(sKeyNorm) >= 3 Then
                        If Not refKeyMap.Exists(sKeyNorm) Then refKeyMap.Add sKeyNorm, uniqueName
                    End If
                End If
            Next kk

            Debug.Print "Bookmark: " & uniqueName & " -> " & Left$(line, 140)
        Else
            Debug.Print "Bookmark create failed: " & uniqueName & " Err=" & Err.Number
            Err.Clear
        End If

NextPara:
    Next p
    On Error GoTo 0
End Sub

' ----------------------------
' Enhanced resolver that uses refKeyMap and improved fallbacks
' ----------------------------
Private Function ResolveCandidate2(surnameMap As Object, surnameYearMap As Object, bookmarkTextMap As Object, refKeyMap As Object, _
    surname As String, yearTok As String, visibleToken As String, visibleNorm As String) As String

    Dim sKey As String: sKey = OnlyLetters(surname)
    If Len(sKey) = 0 Then ResolveCandidate2 = "": Exit Function

    ' 1) exact surname|year
    If Len(yearTok) = 4 Then
        Dim ky As String: ky = sKey & "|" & yearTok
        If surnameYearMap.Exists(ky) Then ResolveCandidate2 = surnameYearMap(ky): Exit Function
    End If

    ' 2) direct surname single candidate
    If surnameMap.Exists(sKey) Then
        Dim coll As Collection: Set coll = surnameMap(sKey)
        If coll.Count = 1 Then ResolveCandidate2 = coll(1): Exit Function
    End If

    ' 3) try visible-token substring in any bookmark text
    Dim kk As Variant
    For Each kk In bookmarkTextMap.Keys
        If InStr(1, bookmarkTextMap(kk), visibleToken, vbTextCompare) > 0 Then ResolveCandidate2 = kk: Exit Function
    Next kk

    ' 4) try normalized visibleNorm in refKeyMap
    Dim lookupKey As String: lookupKey = LCase$(Trim$(OnlyAlphaNumericShort(visibleNorm)))
    If Len(lookupKey) >= 3 Then
        If refKeyMap.Exists(lookupKey) Then ResolveCandidate2 = refKeyMap(lookupKey): Exit Function
    End If

    ' 5) try any bookmark which contains the surname (word)
    For Each kk In bookmarkTextMap.Keys
        If InStr(1, " " & LCase$(bookmarkTextMap(kk)) & " ", " " & LCase$(surname) & " ", vbTextCompare) > 0 Then
            ResolveCandidate2 = kk: Exit Function
        End If
    Next kk

    ' 6) fallback: if surnameMap has candidates, return first
    If surnameMap.Exists(sKey) Then
        ResolveCandidate2 = surnameMap(sKey)(1): Exit Function
    End If

    ResolveCandidate2 = ""
End Function

' ----------------------------
' Helper: Find references heading start
' ----------------------------
Private Function FindReferencesStart(doc As Document) As Long
    Dim r As Range: Set r = doc.Content.Duplicate
    With r.Find
        .ClearFormatting
        .Text = "References"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = True
    End With
    If r.Find.Execute Then FindReferencesStart = r.Start Else FindReferencesStart = 0
End Function

' ----------------------------
' Find token position - multiple heuristics
' ----------------------------
Private Function FindTokenPosition(startAt As Long, raw As String, tokRaw As String) As Long
    If Len(Trim$(tokRaw)) = 0 Then FindTokenPosition = 0: Exit Function
    Dim pos As Long

    ' exact
    pos = InStr(startAt, raw, tokRaw, vbTextCompare)
    If pos > 0 Then FindTokenPosition = pos: Exit Function

    ' remove surrounding parentheses
    Dim t As String: t = Trim$(tokRaw)
    If Left$(t, 1) = "(" Then t = Mid$(t, 2)
    If Right$(t, 1) = ")" Then t = Left$(t, Len(t) - 1)
    t = Trim$(t)
    If Len(t) > 0 Then
        pos = InStr(startAt, raw, t, vbTextCompare)
        If pos > 0 Then FindTokenPosition = pos: Exit Function
    End If

    ' punctuation trimmed
    Dim t2 As String: t2 = TrimPunctuation(t)
    If Len(t2) > 0 Then
        pos = InStr(startAt, raw, t2, vbTextCompare)
        If pos > 0 Then FindTokenPosition = pos: Exit Function
    End If

    ' first 1-2 words
    Dim arr() As String: arr = Split(t2, " ")
    Dim candidate As String
    If UBound(arr) >= 0 Then
        candidate = arr(0)
        If Len(candidate) >= 3 Then
            pos = InStr(startAt, raw, candidate, vbTextCompare)
            If pos > 0 Then FindTokenPosition = pos: Exit Function
        End If
    End If
    If UBound(arr) >= 1 Then
        candidate = arr(0) & " " & arr(1)
        pos = InStr(startAt, raw, candidate, vbTextCompare)
        If pos > 0 Then FindTokenPosition = pos: Exit Function
    End If

    FindTokenPosition = 0
End Function

' ----------------------------
' Normalize token text (strip outer parentheses, trailing period)
' ----------------------------
Private Function NormalizeToken(tok As String) As String
    Dim s As String: s = Trim$(tok)
    If Len(s) = 0 Then NormalizeToken = "": Exit Function
    If Left$(s, 1) = "(" Then s = Mid$(s, 2)
    If Right$(s, 1) = ")" Then s = Left$(s, Len(s) - 1)
    s = Trim$(s)
    If Len(s) > 1 Then If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    NormalizeToken = Trim$(s)
End Function

' ----------------------------
' Extract first author surname heuristics
' ----------------------------
Private Function ExtractFirstAuthorFromCitation(token As String) As String
    Dim s As String: s = token
    s = Replace(s, "(", " ")
    s = Replace(s, ")", " ")
    s = Trim$(s)
    Dim etPos As Long: etPos = InStr(1, s, "et al", vbTextCompare)
    If etPos > 0 Then s = Trim$(Left$(s, etPos - 1))
    Dim ampPos As Long: ampPos = InStr(1, s, "&")
    If ampPos > 0 Then s = Trim$(Left$(s, ampPos - 1))
    Dim commaPos As Long: commaPos = InStr(1, s, ",")
    If commaPos > 0 Then
        Dim leftPart As String: leftPart = Trim$(Left$(s, commaPos - 1))
        If Len(OnlyLetters(leftPart)) >= 2 Then ExtractFirstAuthorFromCitation = OnlyLetters(leftPart): Exit Function
    End If
    Dim words() As String: words = Split(s, " ")
    Dim ii As Long
    For ii = UBound(words) To LBound(words) Step -1
        Dim ww As String: ww = Replace(Replace(words(ii), ",", ""), ".", "")
        If Len(OnlyLetters(ww)) >= 2 Then ExtractFirstAuthorFromCitation = OnlyLetters(ww): Exit Function
    Next ii
    If UBound(words) >= LBound(words) Then ExtractFirstAuthorFromCitation = OnlyLetters(words(0)) Else ExtractFirstAuthorFromCitation = ""
End Function

' ----------------------------
' Year extraction (first 4-digit group)
' ----------------------------
Private Function ExtractYearFromCitationText(s As String) As String
    Dim ii As Long, num As String: num = ""
    For ii = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, ii, 1)
        If ch Like "[0-9]" Then num = num & ch Else num = num & " "
    Next ii
    num = Trim$(Replace(num, "  ", " "))
    Dim parts() As String: parts = Split(num, " ")
    Dim pp As Long
    For pp = LBound(parts) To UBound(parts)
        If Len(parts(pp)) = 4 Then ExtractYearFromCitationText = parts(pp): Exit Function
    Next pp
    ExtractYearFromCitationText = ""
End Function

' ----------------------------
' Only letters A-Z
' ----------------------------
Private Function OnlyLetters(s As String) As String
    Dim out As String: out = ""
    Dim ii As Long, ch As String
    For ii = 1 To Len(s)
        ch = Mid$(s, ii, 1)
        If ch Like "[A-Za-z]" Then out = out & ch
    Next ii
    OnlyLetters = out
End Function

' ----------------------------
' Try get surname from CC tag (best-effort)
' ----------------------------
Private Function TryGetSurnameFromTag(tagText As String, tokenIndex As Long) As String
    Dim s As String: s = tagText
    s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
    Dim famKey As String: famKey = """family"":"
    Dim pos As Long: pos = InStr(1, s, famKey, vbTextCompare)
    If pos = 0 Then TryGetSurnameFromTag = "": Exit Function
    Dim names As Collection: Set names = New Collection
    Dim pp As Long: pp = 1
    Do
        pos = InStr(pp, s, famKey, vbTextCompare)
        If pos = 0 Then Exit Do
        Dim startQuote As Long: startQuote = InStr(pos + Len(famKey), s, """")
        If startQuote = 0 Then Exit Do
        Dim endQuote As Long: endQuote = InStr(startQuote + 1, s, """")
        If endQuote = 0 Then Exit Do
        Dim fam As String: fam = Mid$(s, startQuote + 1, endQuote - startQuote - 1)
        names.Add fam
        pp = endQuote + 1
    Loop
    If names.Count = 0 Then TryGetSurnameFromTag = "": Exit Function
    Dim idx As Long: idx = tokenIndex + 1
    If idx >= 1 And idx <= names.Count Then TryGetSurnameFromTag = OnlyLetters(names(idx)) Else TryGetSurnameFromTag = OnlyLetters(names(1))
End Function

' ----------------------------
' Try get year from tag
' ----------------------------
Private Function TryGetYearFromTag(tagText As String, tokenIndex As Long) As String
    Dim s As String: s = tagText
    Dim pos As Long: pos = InStr(1, s, "date-parts", vbTextCompare)
    If pos = 0 Then TryGetYearFromTag = ExtractYearFromCitationText(tagText): Exit Function
    Dim arr() As String: arr = Split(s, "date-parts")
    Dim seg As String
    Dim idx As Long: idx = tokenIndex
    If idx < 0 Then idx = 0
    If idx + 1 <= UBound(arr) Then seg = arr(idx + 1) Else seg = arr(1)
    TryGetYearFromTag = ExtractYearFromCitationText(seg)
End Function

' ----------------------------
' Trim punctuation helper
' ----------------------------
Private Function TrimPunctuation(s As String) As String
    Dim stopChars As String: stopChars = ".,;:()[]{}""'"
    Dim out As String: out = s
    Dim changed As Boolean: changed = True
    Do While changed
        changed = False
        If Len(out) > 0 Then
            If InStr(1, stopChars, Left$(out, 1), vbBinaryCompare) > 0 Then out = Mid$(out, 2): changed = True
        End If
        If Len(out) > 0 Then
            If InStr(1, stopChars, Right$(out, 1), vbBinaryCompare) > 0 Then out = Left$(out, Len(out) - 1): changed = True
        End If
    Loop
    TrimPunctuation = out
End Function

' ----------------------------
' Remove punctuation but keep letters/numbers and single spaces (short normalized key)
' ----------------------------
Private Function OnlyAlphaNumericShort(s As String) As String
    Dim ii As Long, ch As String, out As String
    out = ""
    For ii = 1 To Len(s)
        ch = Mid$(s, ii, 1)
        If ch Like "[A-Za-z0-9]" Then
            out = out & ch
        ElseIf ch = " " Then
            If Right$(out, 1) <> " " Then out = out & " "
        End If
    Next ii
    OnlyAlphaNumericShort = Trim$(LCase$(out))
End Function


' ----------------------------
' Strip punctuation (keeps words separated by single space)
' ----------------------------
Private Function StripPunctuation(s As String) As String
    Dim ii As Long, ch As String, out As String: out = ""
    For ii = 1 To Len(s)
        ch = Mid$(s, ii, 1)
        If ch Like "[A-Za-z0-9&]" Or ch = " " Then
            If ch = " " And Right$(out, 1) = " " Then
                ' skip duplicate
            Else
                out = out & ch
            End If
        Else
            out = out & " "
        End If
    Next ii
    StripPunctuation = Trim$(out)
End Function

' ----------------------------
' Bookmark exists helper
' ----------------------------
Private Function BookmarkExistsInDoc(doc As Document, name As String) As Boolean
    On Error Resume Next
    BookmarkExistsInDoc = doc.Bookmarks.Exists(name)
    On Error GoTo 0
End Function
