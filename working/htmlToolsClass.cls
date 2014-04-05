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
    posStart = getTagStartPos(ignore, layout, 1, Key)
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
    posStart = getTagStartPos(ignore, layout, 1, Key)
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
        posStart = getTagStartPos(ignore, layout, 1, Key)
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
    posStart = getTagStartPos(ignore, s, 1, Key)
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
Public Function getTagStartPos(ignore As Variant, layout As String, layoutStartPos As Long, Key As String) As Long
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
        ' find this key and get the outerHTML
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
            ' search for a tagname
            '
            targetName = "<"
        End If
        lenTargetName = Len(targetName)
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
                ' string found, verify it ends with either single-quote,double-quote,space,vbcr,vblf
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
