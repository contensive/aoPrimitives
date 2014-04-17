VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "htmlToolsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'==============================================================================
'
'==============================================================================
'
'
'
'
Public Function isGuid(ignore As Variant, Source As String) As Boolean
    On Error GoTo ErrorTrap
    '
    '
    ' normalize guid
    '
    If (Len(Source) = 38) And (Left(Source, 1) = "{") And (Right(Source, 1) = "}") Then
        '
        ' Good to go
        '
        isGuid = True
    ElseIf (Len(Source) = 36) And (InStr(1, Source, " ") = 0) Then
        '
        ' might be valid with the brackets, add them
        '
        isGuid = True
        'source = "{" & source & "}"
    ElseIf (Len(Source) = 32) Then
        '
        ' might be valid with the brackets and the dashes, add them
        '
        isGuid = True
        'source = "{" & Mid(source, 1, 8) & "-" & Mid(source, 9, 4) & "-" & Mid(source, 13, 4) & "-" & Mid(source, 17, 4) & "-" & Mid(source, 21) & "}"
    Else
        '
        ' not valid
        '
        isGuid = False
'        source = ""
    End If
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagEndPos", "Unexpected Error [" & Err.Description & "]")
End Function
'
'
'
Public Function insertOuterHTML(ignore As Variant, layout As String, Key As String, textToInsert As String) As String
    On Error GoTo ErrorTrap
    '
    '
    'Dim outerHTML As String
    Dim posStart As Long
    Dim posEnd As Long
    Dim posWorking As Long
    '
    ' short-cut for now, get the outerhtml, find the position, then remove the wrapping tags
    '
    If Key = "" Then
        insertOuterHTML = textToInsert
    Else
    insertOuterHTML = layout
    posStart = getTagStartPos2(ignore, layout, 1, Key)
    'outerHTML = getOuterHTML(ignore, layout, Key, PosStart)
    If posStart <> 0 Then
        posEnd = getTagEndPos(ignore, layout, posStart)
        If posEnd > 0 Then
            '
            ' seems like these are the correct positions here.
            '
            insertOuterHTML = Left(layout, posStart - 1) & textToInsert & Mid(layout, posEnd)
'            posStart = InStr(posStart + 1, layout, ">")
'            If posStart <> 0 Then
'                posStart = posStart + 1
'                posEnd = InStrRev(layout, "<", posEnd)
'                If posEnd <> 0 Then
'                    insertOuterHTML = Left(layout, posStart - 1) & textToInsert & Mid(layout, posEnd)
'                End If
'            End If
        End If
    End If
    End If
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagEndPos", "Unexpected Error [" & Err.Description & "]")
End Function
'
'
'
Public Function insertInnerHTML(ignore As Variant, layout As String, Key As String, textToInsert As String) As String
    On Error GoTo ErrorTrap
    '
    '
    'Dim outerHTML As String
    Dim posStart As Long
    Dim posEnd As Long
    '
    ' short-cut for now, get the outerhtml, find the position, then remove the wrapping tags
    '
    If Key = "" Then
        insertInnerHTML = textToInsert
    Else
    insertInnerHTML = layout
    posStart = getTagStartPos2(ignore, layout, 1, Key)
    'outerHTML = getOuterHTML(ignore, layout, Key, PosStart)
    If posStart <> 0 Then
        posEnd = getTagEndPos(ignore, layout, posStart)
        If posEnd > 0 Then
            posStart = InStr(posStart + 1, layout, ">")
            If posStart <> 0 Then
                posStart = posStart + 1
                posEnd = InStrRev(layout, "<", posEnd - 1)
                If posEnd <> 0 Then
                    ''PosEnd = PosEnd - 1
                    ''
                    '' now move the end forward to skip trailing whitespace
                    ''
                    'Do
                    '    posEnd = posEnd + 1
                    'Loop While (posEnd < Len(layout)) And (InStr(1, "<" & vbTab & vbCr & vbLf & vbTab & " ", Mid(layout, posEnd, 1)) <> 0)
                    'posEnd = posEnd - 1
                    insertInnerHTML = Left(layout, posStart - 1) & textToInsert & Mid(layout, posEnd)
                End If
            End If
        End If
    End If
    End If
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagEndPos", "Unexpected Error [" & Err.Description & "]")
End Function
'
'
'
Public Function getInnerHTML(ignore As Variant, layout As String, Key As String) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim posStart As Long
    Dim posEnd As Long
    '
    ' short-cut for now, get the outerhtml, find the position, then remove the wrapping tags
    '
    If Key = "" Then
        '
        ' inner of nothing is nothing
        '
    Else
        getInnerHTML = layout
        posStart = getTagStartPos2(ignore, layout, 1, Key)
        If posStart <> 0 Then
            posEnd = getTagEndPos(ignore, layout, posStart)
            If posEnd > 0 Then
                posStart = InStr(posStart + 1, layout, ">")
                If posStart <> 0 Then
                    posStart = posStart + 1
                    posEnd = InStrRev(layout, "<", posEnd - 1)
                    If posEnd <> 0 Then
                        '
                        ' now move the end forward to skip trailing whitespace
                        '
                        Do
                            posEnd = posEnd + 1
                        Loop While (posEnd < Len(layout)) And (InStr(1, vbTab & vbCr & vbLf & vbTab & " ", Mid(layout, posEnd, 1)) <> 0)
                        posEnd = posEnd - 1
                        getInnerHTML = Mid(layout, posStart, (posEnd - posStart))
                    End If
                End If
            End If
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagEndPos", "Unexpected Error [" & Err.Description & "]")
End Function
'
'
'
Public Function getOuterHTML(ignore As Variant, layout As String, Key As String) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim posStart As Long
    Dim posEnd As Long
    Dim s As String
    '
Dim c As String

    s = layout
    If s <> "" Then
    posStart = getTagStartPos2(ignore, s, 1, Key)
    If posStart > 0 Then
        '
        ' now backtrack to include the leading whitespace
        '
        Do
            posStart = posStart - 1
        Loop While (posStart > 0) And (InStr(1, vbTab & vbCr & vbLf & vbTab & " ", Mid(s, posStart, 1)) <> 0)
        posStart = posStart + 1
        s = Mid(s, posStart)
        posEnd = getTagEndPos(ignore, s, 1)
        If posEnd > 0 Then
            s = Left(s, posEnd - 1)
            getOuterHTML = s
        End If
    End If
    End If
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagEndPos", "Unexpected Error [" & Err.Description & "]")
End Function
'
'
'
Private Function tagMatch(layout As String, posStartTag As Long, searchId As String, searchClass As String) As Boolean
    Const attrAllowedChars = "abcdefghijklmnopqrstuvwzyz-_"
    Dim Tag As String
    Dim tagLower As String
    Dim Pos As Long
    Dim idMatch As Boolean
    Dim classMatch As Boolean
    Dim Delimiter As String
    Dim testChar As String
    Dim tagLength As Long
    Dim posValueStart As Long
    Dim testValue As String
    Dim testValues() As String
    Dim testCnt As Long
    Dim Ptr As Long
    '
    tagMatch = False
    Pos = InStr(posStartTag, layout, ">")
    If Pos > 0 Then
        tagMatch = True
        Tag = Mid(layout, posStartTag, Pos - posStartTag + 1)
        tagLower = LCase(Tag)
        tagLength = Len(Tag)
        '
        ' check searchId
        '
        If tagMatch And (searchId <> "") Then
            Pos = InStr(1, tagLower, " id=", vbTextCompare)
            If Pos <= 0 Then
                '
                ' id required but this tag has no id attr
                '
                tagMatch = False
            Else
                '
                ' test if the id attr value matches the searchClass
                '
                Pos = Pos + 4
                Delimiter = Mid(tagLower, Pos, 1)
                testValue = ""
                If (Delimiter = """") Or (Delimiter = "'") Then
                    '
                    ' search for end of delimited attribute value
                    '
                    posValueStart = Pos + 1
                    Do
                        Pos = Pos + 1
                        testChar = Mid(tagLower, Pos, 1)
                    Loop While (Pos < tagLength) And (testChar <> Delimiter)
                    If Pos >= tagLength Then
                        '
                        ' delimiter not found, html error
                        '
                        tagMatch = False
                    Else
                        testValue = Mid(Tag, posValueStart, Pos - posValueStart)
                    End If
                Else
                    '
                    ' search for end of non-delimited attribute value
                    '
                    posValueStart = Pos
                    Do While (Pos < tagLength) And (InStr(1, attrAllowedChars, Mid(tagLower, Pos, 1), vbTextCompare))
                        Pos = Pos + 1
                    Loop
                    If Pos >= tagLength Then
                        '
                        ' delimiter not found, html error
                        '
                        tagMatch = False
                    Else
                        testValue = Mid(Tag, posValueStart, Pos - posValueStart)
                    End If
                End If
                If tagMatch And (testValue <> "") Then
                    '
                    '
                    '
                    If searchId <> testValue Then
                        '
                        ' there can only be one id, and this does not match
                        '
                        tagMatch = False
                    End If
                End If
            End If
        End If
        '
        ' check searchClass
        '
        If tagMatch And (searchClass <> "") Then
            Pos = InStr(1, tagLower, " class=", vbTextCompare)
            If Pos <= 0 Then
                '
                ' class required but this tag has no class attr
                '
                tagMatch = False
            Else
                '
                ' test if the class attr value matches the searchClass
                '
                Pos = Pos + 7
                Delimiter = Mid(tagLower, Pos, 1)
                testValue = ""
                If (Delimiter = """") Or (Delimiter = "'") Then
                    '
                    ' search for end of delimited attribute value
                    '
                    posValueStart = Pos + 1
                    Do
                        Pos = Pos + 1
                        testChar = Mid(tagLower, Pos, 1)
                    Loop While (Pos < tagLength) And (testChar <> Delimiter)
                    If Pos >= tagLength Then
                        '
                        ' delimiter not found, html error
                        '
                        tagMatch = False
                    Else
                        testValue = Mid(Tag, posValueStart, Pos - posValueStart)
                    End If
                Else
                    '
                    ' search for end of non-delimited attribute value
                    '
                    posValueStart = Pos
                    Do While (Pos < tagLength) And (InStr(1, attrAllowedChars, Mid(tagLower, Pos, 1), vbTextCompare))
                        Pos = Pos + 1
                    Loop
                    If Pos >= tagLength Then
                        '
                        ' delimiter not found, html error
                        '
                        tagMatch = False
                    Else
                        testValue = Mid(Tag, posValueStart, Pos - posValueStart)
                    End If
                End If
                If tagMatch And (testValue <> "") Then
                    '
                    '
                    '
                    testValues = Split(testValue, " ")
                    testCnt = UBound(testValues) + 1
                    For Ptr = 0 To testCnt - 1
                        If searchClass = testValues(Ptr) Then
                            Exit For
                        End If
                    Next
                    If Ptr >= testCnt Then
                        tagMatch = False
                    End If
                End If
            End If
        End If
    End If
    '
    
End Function
'
'
'
Public Function getTagStartPos2(ignore As Variant, layout As String, layoutStartPos As Long, Key As String) As Long
    On Error GoTo ErrorTrap
    '
    Dim returnPos As Long
    '
    'Dim posMatch As Long
    Dim SegmentStart As Long
    Dim Pos As Long
    'Dim searchIsOver As Boolean
    Dim LoopPtr As Long
    Dim searchKey As String
    Dim lenSearchKey As String
    'Dim s As String
    Dim c As String
    Dim Ptr As Long
    Dim workingKey As String
    Dim workingKeys() As String
    Dim posEnd As Long
    Dim searchClass As String
    Dim searchId As String
    Dim searchTag As String
    Dim posStartTag As Long
    '
's = layout
    returnPos = 0
    workingKey = Key
    If InStr(1, workingKey, ">") <> 0 Then
        '
        ' does not support > yet.
        '
        workingKey = Replace(workingKey, ">", " ")
    End If
    '
    ' eliminate whitespace
    '
    Do While InStr(1, workingKey, vbTab) <> 0
        workingKey = Replace(workingKey, vbTab, " ")
    Loop
    '
    Do While InStr(1, workingKey, vbCr) <> 0
        workingKey = Replace(workingKey, vbCr, " ")
    Loop
    '
    Do While InStr(1, workingKey, vbLf) <> 0
        workingKey = Replace(workingKey, vbLf, " ")
    Loop
    '
    Do While InStr(1, workingKey, "  ") <> 0
        workingKey = Replace(workingKey, "  ", " ")
    Loop
    '
    workingKey = Trim(workingKey)
    '
    If InStr(1, workingKey, " ") <> 0 Then
        '
        ' if there are spaces, do them sequentially
        '
        workingKeys = Split(workingKey, " ")
        SegmentStart = 1
        Do While (layout <> "") And (SegmentStart <> 0) And (Ptr <= UBound(workingKeys))
            SegmentStart = getTagStartPos2(Nothing, layout, SegmentStart, workingKeys(Ptr))
            Ptr = Ptr + 1
        Loop
        returnPos = SegmentStart
    Else
        '
        ' find this single key and get the outerHTML
        '   at this point, the key can be
        '       a class = .xxxx
        '       an id = #xxxx
        '       a tag = xxxx
        '       a compound in either form, xxxx.xxxx or xxxx#xxxx
        '
        '   searchKey = the search pattern to start
        '
        If Left(workingKey, 1) = "." Then
            '
            ' search for a class
            '
            searchClass = Mid(workingKey, 2)
            searchTag = ""
            searchId = ""
            Pos = InStr(1, searchClass, "#")
            If Pos <> 0 Then
                searchId = Mid(searchClass, Pos)
                searchClass = Mid(searchClass, 1, Pos - 1)
            End If
            '
            'workingKey = Mid(workingKey, 2)
            searchKey = "<"
        ElseIf Left(workingKey, 1) = "#" Then
            '
            ' search for an ID
            '
            searchClass = ""
            searchTag = ""
            searchId = Mid(workingKey, 2)
            Pos = InStr(1, searchId, ".")
            If Pos <> 0 Then
                searchClass = Mid(searchId, Pos)
                searchId = Mid(searchId, 1, Pos - 1)
            End If
            '
            'workingKey = Mid(workingKey, 2)
            searchKey = "<"
        Else
            '
            ' search for a tagname
            '
            searchClass = ""
            searchTag = workingKey
            searchId = ""
            '
            Pos = InStr(1, searchTag, "#")
            If Pos <> 0 Then
                searchId = Mid(searchTag, Pos + 1)
                searchTag = Mid(searchTag, 1, Pos - 1)
                Pos = InStr(1, searchId, ".")
                If Pos <> 0 Then
                    searchClass = Mid(searchId, Pos)
                    searchId = Mid(searchId, 1, Pos - 1)
                End If
            End If
            Pos = InStr(1, searchTag, ".")
            If Pos <> 0 Then
                searchClass = Mid(searchTag, Pos + 1)
                searchTag = Mid(searchTag, 1, Pos - 1)
                Pos = InStr(1, searchClass, "#")
                If Pos <> 0 Then
                    searchId = Mid(searchClass, Pos)
                    searchClass = Mid(searchClass, 1, Pos - 1)
                End If
            End If
            '
            searchKey = "<" & searchTag
        End If
        lenSearchKey = Len(searchKey)
        Pos = layoutStartPos
        'posMatch = InStr(layoutStartPos, layout, searchKey)
        'pos = posMatch
        'searchIsOver = False
        Do
            Pos = InStr(Pos, layout, searchKey)
            If Pos = 0 Then
                '
                ' not found, return empty
                '
's = ""
                Exit Do
            Else
                '
                ' string found - go to the start of the tag
                '
                posStartTag = InStrRev(layout, "<", Pos + 1)
                If posStartTag <= 0 Then
                    '
                    ' bad html, no start tag found
                    '
                    Pos = 0
                    returnPos = 0
                ElseIf Mid(layout, posStartTag, 2) = "</" Then
                    '
                    ' this is an end tag, skip it
                    '
                    Pos = Pos + 1
                ElseIf tagMatch(layout, posStartTag, searchId, searchClass) Then
                    '
                    ' match, return with this position
                    '
                    returnPos = Pos
                    Exit Do
                Else
                    '
                    ' no match, skip this and go to the next
                    '
                    Pos = Pos + 1
                End If
            End If
            LoopPtr = LoopPtr + 1
        Loop While LoopPtr < 1000
        '
        '
        '
        If LoopPtr >= 1000 Then
            Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagStartPos2", "Tag limit of 1000 tags per block reached.")
        End If
    End If
    '
    getTagStartPos2 = returnPos
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagStartPos2", "Unexpected Error [" & Err.Description & "]")
End Function
'
'
'
Public Function getTagStartPos(ignore As Variant, layout As String, layoutStartPos As Long, Key As String) As Long
getTagStartPos = getTagStartPos2(ignore, layout, layoutStartPos, Key)
Exit Function

    On Error GoTo ErrorTrap
    '
    '
    Dim posMatch As Long
    Dim SegmentStart As Long
    Dim Pos As Long
    Dim searchIsOver As Boolean
    Dim LoopPtr As Long
    Dim targetName As String
    Dim lenTargetName As String
    Dim s As String
    Dim c As String
    Dim Ptr As Long
    Dim workingKey As String
    Dim workingKeys() As String
    Dim posEnd As Long
    
    '
    s = layout
    getTagStartPos = 0
    workingKey = Key
    If InStr(1, workingKey, ">") <> 0 Then
        '
        ' does not support > yet.
        '
        workingKey = Replace(workingKey, ">", " ")
    End If
    '
    ' eliminate whitespace
    '
    Do While InStr(1, workingKey, vbTab) <> 0
        workingKey = Replace(workingKey, vbTab, " ")
    Loop
    '
    Do While InStr(1, workingKey, vbCr) <> 0
        workingKey = Replace(workingKey, vbCr, " ")
    Loop
    '
    Do While InStr(1, workingKey, vbLf) <> 0
        workingKey = Replace(workingKey, vbLf, " ")
    Loop
    '
    Do While InStr(1, workingKey, "  ") <> 0
        workingKey = Replace(workingKey, "  ", " ")
    Loop
    '
    workingKey = Trim(workingKey)
    '
    If InStr(1, workingKey, " ") <> 0 Then
        '
        ' if there are spaces, do them sequentially
        '
        workingKeys = Split(workingKey, " ")
        SegmentStart = 1
        Do While (s <> "") And (SegmentStart <> 0) And (Ptr <= UBound(workingKeys))
            SegmentStart = getTagStartPos(Nothing, s, SegmentStart, workingKeys(Ptr))
            Ptr = Ptr + 1
        Loop
        getTagStartPos = SegmentStart
            
        'For ptr = 0 To UBound(workingKeys)
        '    s = getOuterHTML(ignore, s, workingKeys(ptr))
        'Next
    Else
        '
        ' find this single key and get the outerHTML
        '   at this point, the key can be
        '       a class = .xxxx
        '       an id = #xxxx
        '       a tag = xxxx
        '       a compound in either form, xxxx.xxxx or xxxx#xxxx
        '
        '   targetName = the search pattern to start
        '
        If Left(workingKey, 1) = "." Then
            '
            ' search for a class
            '
            workingKey = Mid(workingKey, 2)
            targetName = "class="
        ElseIf Left(workingKey, 1) = "#" Then
            '
            ' search for an ID
            '
            workingKey = Mid(workingKey, 2)
            targetName = "id="
        Else
            '
            '   !!!!! might be a compound selector - can not just assume it is a tag
            ' search for a tagname
            '
            targetName = "<"
        End If
        lenTargetName = Len(targetName)
'!!!!! ul.menu_desktop selector fails here. Pos returned 0.
        posMatch = InStr(layoutStartPos, s, workingKey)
        Pos = posMatch
        searchIsOver = False
        Do
            If Pos = 0 Then
                '
                ' not found, return empty
                '
                s = ""
                Exit Do
            Else
                '
                ' string found, verify it ends with either single-quote,double-quote,space,vbcr,vblf, or GT
                '
                c = Mid(s, Pos + Len(workingKey), 1)
                If (InStr(1, ">'"" " & vbCrLf, c) = 0) Then
                    '
                    ' this is not a hit, find the next match
                    '
                    posMatch = InStr(posMatch + 1, s, workingKey)
                    Pos = posMatch
                Else
                    '
                    ' string found, go back to start of string
                    '
                    c = LCase(Mid(s, Pos - 1, 1))
                    If (c >= "a") And (c <= "z") Then
                        '
                        ' this is not a hit, find the next match
                        '
                        posMatch = InStr(posMatch + 1, s, workingKey)
                        Pos = posMatch
                    Else
                        If (c = """") Or (c = "'") Then
                            '
                            ' hit a quote, go back one more
                            '
                            Pos = Pos - 1
                        End If
                        '
                        ' this should be a match, no not
                        '
                        If targetName = Mid(s, Pos - lenTargetName, lenTargetName) Then
                            '
                            ' match found, get from front of tag to the end of this tag
                            '
                            Pos = InStrRev(s, "<", Pos)
                            If Pos = 0 Then
                                s = ""
                            Else
                                getTagStartPos = Pos
                                's = Mid(s, pos)
                                'PosEnd = getTagEndPos(ignore, s, 1)
                                's = Left(s, PosEnd)
                            End If
                            Exit Do
                        Else
                            '
                            ' not a match, go find the next (posMatch is where the string was last found)
                            '
                            posMatch = InStr(posMatch + 1, s, workingKey)
                            Pos = posMatch
                        End If
                    End If
                End If
            End If
            LoopPtr = LoopPtr + 1
        Loop While LoopPtr < 100
    End If
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagEndPos", "Unexpected Error [" & Err.Description & "]")
End Function
'
'=================================================================================================
'   return the position following the tag which closes the tag that starts the string
'       starting postion<div><div><p>this and that</p><!-- </div> --></div></div>And a whole lot more
'       returns the position of the "A" following the last /div
'       string 123<img>456 returns pointing to "4"
'       string 123456 returns pointing to "6"
'       returns 0 if the end was not found
'=================================================================================================
'
Public Function getTagEndPos(ignore As Variant, Source, startPos) As Long
    On Error GoTo ErrorTrap
    '
    '
    Dim Pos As Long
    Dim TagName As String
    Dim endTag As String
    Dim startTag As String
    Dim posNest As String
    Dim posEnd As Long
    Dim posComment As Long
    Dim c As String
    '
    Pos = InStr(startPos, Source, "<")
    TagName = ""
    getTagEndPos = 0
    If Pos <> 0 Then
        Pos = Pos + 1
        Do While Pos < Len(Source)
            c = LCase(Mid(Source, Pos, 1))
            If (c >= "a") And (c <= "z") Then
                TagName = TagName & c
            Else
                Exit Do
            End If
            Pos = Pos + 1
        Loop
        If TagName <> "" Then
            endTag = "</" & TagName
            startTag = "<" & TagName
            Do While (Pos <> 0)
                posEnd = InStr(Pos + 1, Source, endTag, vbTextCompare)
                If posEnd = 0 Then
                    '
                    ' no end was found, return the tag or rest of the string
                    '
                    getTagEndPos = InStr(Pos + 1, Source, ">") + 1
                    If posEnd = 1 Then
                        getTagEndPos = Len(Source)
                    End If
                    Exit Do
                Else
                    posNest = InStr(Pos + 1, Source, startTag, vbTextCompare)
                    If posNest = 0 Then
                        '
                        ' no nest found, set to end
                        '
                        posNest = Len(Source)
                    End If
                    posComment = InStr(Pos + 1, Source, "<!--")
                    If posComment = 0 Then
                        '
                        ' no comment found, set to end
                        '
                        posComment = Len(Source)
                    End If
                    If (posNest < posEnd) And (posNest < posComment) Then
                        '
                        ' ----- the tag is nested, find the end of the nest
                        '
                        Pos = getTagEndPos(ignore, Source, posNest)
                        ' 8/28/2012, if there is a nested tag right before the correct end tag, it skips the end:
                        ' <div class=a>a<div class=b>b</div></div>
                        ' the second /div is missed because getTagEndPos returns one past the >, then the
                        ' next search starts +1 that position
                        If (Pos > 0) Then
                            Pos = Pos - 1
                        End If
                    ElseIf (posComment < posEnd) Then
                        '
                        ' ----- there is a comment between the tag and the first tagend, skip it
                        '
                        Pos = InStr(posComment, Source, "-->")
                        If Pos = 0 Then
                            '
                            ' start comment with no end, exit now
                            '
                            getTagEndPos = Len(Source)
                            Exit Do
                        End If
                    Else
                        '
                        ' ----- end position is here, go to the end of it and exit
                        '
                        Pos = InStr(posEnd, Source, ">")
                        If Pos = 0 Then
                            '
                            ' no end was found, just exit
                            '
                            getTagEndPos = Len(Source)
                            Exit Do
                        Else
                            '
                            ' ----- end was found
                            '
                            getTagEndPos = Pos + 1
                            Exit Do
                        End If
                    End If
                End If
            Loop
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call Err.Raise(KmaErrorInternal, "aoPrimitives.htmlToolsClass.getTagEndPos", "Unexpected Error [" & Err.Description & "]")
End Function
'
Private Sub Class_Initialize()
    Dim a As Long
    a = a
End Sub

Private Sub Class_Terminate()
    Dim a As Long
    a = a
End Sub
