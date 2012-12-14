VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "contentCmdClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'========================================================================
'   Prototype for aoPimitives.StructuredDataClass
'
'   Elerts Newsletter creates <?contensive> processing instructions
'   and passes them directly here to execute.
'   After Structured Data is inf=corporated in Contensive, this class
'   can be removed and the <?contensive> data will be executed when
'   it is returned from the Elerts Newsletter.
'========================================================================
'
Private Enum AddonContextEnum
    '
    ' this is a copy of what is in mainclass. It is public so it will not change, but can not be referenced here
    '
    ContextPage = 1
    ContextAdmin = 2
    ContextTemplate = 3
    ContextEmail = 4
    ContextRemoteMethod = 5
    ContextOnNewVisit = 6
    ContextOnPageEnd = 7
    ContextOnPageStart = 8
    ContextEditor = 9
    ContextHelpUser = 10
    ContextHelpAdmin = 11
    ContextHelpDeveloper = 12
    ContextOnContentChange = 13
    ContextFilter = 14
    ContextSimple = 15
    ContextOnBodyStart = 16
    ContextOnBodyEnd = 17
End Enum
'
'=================================================================================
'   Execute Method, v3.4 Interface
'
'   Contensive calls this add-on if any add-on returns
'       an xml document starting with a processing instruction "<?Contensive"
'       In this case, the xml document will be passed in the "data" var
'
'   Alternatively, calling the addon with the var "addon" will execute that addon
'       and parse the resulting output.
'=================================================================================
'
Public Function execute(CsvObject As Object, mainObject As Object, optionString As String, filterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim src As String
    Dim Context As Long
    Dim personalizationPeopleId As Long
    Dim personalizationIsAuthenticated As Boolean
    '
    src = CsvObject.GetAddonOption("data", optionString)
    Context = kmaEncodeInteger(CsvObject.GetAddonOption("context", optionString))
    personalizationPeopleId = kmaEncodeInteger(CsvObject.GetAddonOption("personalizationPeopleId", optionString))
    personalizationIsAuthenticated = kmaEncodeBoolean(CsvObject.GetAddonOption("personalizationIsAuthenticated", optionString))
    '
    ' compatibility with old Contensive
    '
    If (personalizationPeopleId = 0) And (Not (mainObject Is Nothing)) Then
        personalizationPeopleId = mainObject.MemberID
        personalizationIsAuthenticated = mainObject.isAuthenticated
    End If
    If src <> "" Then
        '
        ' test for Contensive processign instruction
        '
        execute = ExecuteAllCmdLists(CsvObject, mainObject, src, Context, personalizationPeopleId, personalizationIsAuthenticated)
    End If
    '
    Exit Function
ErrorTrap:
    Call CsvObject.reportError("trap", "aoPrimitives.execute", True)
End Function
'
'============================================================================================
'
'   Content Replacements
'
'   A list of commands that create, modify and return strings
'   the start and end with escape sequences contentReplaceEscapeStart/contentReplaceEscapeEnd
'       {{ and }} previously
'       {% and %} right now
'
'   format:
'       {% commands %}
'
'    commands
'       a single command or a JSON array of commands.
'       if a command has arguments, the command should be a JSON object
'           openLayout layoutName
'
'       one command, no arguments -- non JSON
'               {% user %}
'       one command, one argument -- non JSON
'               {% user "firstname" %}
'
'       one command, no arguments -- JSON command array of one
'               {% [ "user" ] %}
'               cmdList[0] = "user"
'
'       two commands, no arguments -- JSON command array
'               {% [
'                       "user",
'                       "user"
'                   ] %}
'               cmdList[0] = "user"
'               cmdList[1] = "user"
'
'       one command, one argument -- JSON object for command
'               {% [
'                       {
'                           "cmd": "layout",
'                           "arg": "mylayout"
'                       }
'                   ] %}
'               cmdList[0].cmd = layout
'               cmdList[0].arg = "mylayout"
'
'       one command, two arguments
'               {% [
'                       {
'                           "cmd": "set",
'                           "arg": {
'                               "find":"$fpo$",
'                               "replace":"Some Content"
'                       }
'                   ] %}
'               cmdList[0].cmd = "replace"
'               cmdList[0].arg.find = "$fpo$"
'               cmdList[0].arg.replace = "Some Content"
'
'       two commands, two arguments
'               {% [
'                       {
'                           "cmd": "import",
'                           "arg": "myTemplate.html"
'                       },
'                       {
'                           "cmd": "setInner",
'                           "arg": {
'                               "find":".contentBoxClass",
'                               "replace":"{% addon contentBox %}"
'                       }
'                   ] %}
'               cmdList[0].cmd = "import"
'               cmdList[0].arg = "myTemplate.html"
'               cmdList[1].cmd = "setInner"
'               cmdList[0].arg.find = ".contntBoxClass"
'               cmdList[0].arg.replace = "{% addon contentBox %}"
'
'           import htmlFile
'           importVirtual htmlFile
'           open textFile
'           openVirtual webfilename
'           addon contentbox( JSON-Object-optionstring-list )
'           set find replace
'           setInner findLocation replace
'           setOuter findLocation replace
'           user firstname
'           site propertyname
'
Private Function ExecuteAllCmdLists(csv As Object, mainOrNothing As Object, src As String, Context As AddonContextEnum, personalizationPeopleId As Long, personalizationIsAuthenticated As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim badCmd As Boolean
    Dim notFound As Boolean
    Dim AddonOptionString As String
    Dim pairNames As Variant
    Dim i As Long
    Dim pairName As String
    Dim pairValue As String
    Dim ValueVariant As Variant
    Dim posNextStart As Long
    Dim posArgOpen As Long
    Dim posArgClose As Long
    Dim posOpen As Long
    Dim posClose As Long
    Dim cmd As String
    Dim cmdArgs As String
    Dim cmdResult As String
    '
    Dim ACGuid As String
    Dim addonName As String
    Dim ACInstanceID As String
    Dim AddonOptionStringHTMLEncoded As String
    Dim CSPeople As Long
    Dim CSPeopleSet As Long
    Dim cmdArgJSON As Object
    'Dim peopleId As Long
    Dim isJSON As Boolean
    '
    Dim posDq As Long
    Dim posSq As Long
    Dim Ptr As Long
    Dim ptrLast As Long
    Dim dst As String
    Dim escape As String
    Dim dataVersion As String
    '
    dataVersion = csv.GetDataBuildVersion()
    dst = ""
    ptrLast = 1
    Do
        cmd = ""
        posOpen = InStr(ptrLast, src, contentReplaceEscapeStart)
        Ptr = posOpen
        If Ptr = 0 Then
            '
            ' not found, copy the rest of src to dst
            '
        Else
            '
            ' scan until we have passed all double and single quotes that are before the next
            '
            notFound = True
            Do
                posClose = InStr(Ptr, src, contentReplaceEscapeEnd)
                If posClose = 0 Then
                    '
                    ' brace opened but no close, forget the open and exit
                    '
                    posOpen = 0
                    notFound = False
                Else
                    posDq = Ptr
                    Do
                        posDq = InStr(posDq + 1, src, """")
                        escape = ""
                        If posDq > 0 Then
                            escape = Mid(src, posDq - 1, 1)
                        End If
                    Loop While (escape = "\")
                    posSq = Ptr
                    Do
                        posSq = InStr(posSq + 1, src, "'")
                        escape = ""
                        If posSq > 0 Then
                            escape = Mid(src, posSq - 1, 1)
                        End If
                    Loop While (escape = "\")
                    Select Case kmaGetFirstNonZeroLong(posSq, posDq)
                        Case 0
                            '
                            ' both 0, posClose is OK as-is
                            '
                            notFound = False
                        Case 1
                            '
                            ' posSq is before posDq
                            '
                            If posSq > posClose Then
                                notFound = False
                            Else
                                '
                                ' skip forward to the next non-escaped sq
                                '
                                Do
                                    posSq = InStr(posSq + 1, src, "'")
                                    escape = ""
                                    If posSq > 0 Then
                                        escape = Mid(src, posSq - 1, 1)
                                    End If
                                Loop While (escape = "\")
                                Ptr = posSq + 1
                                'notFound = False
                            End If
                        Case Else
                            '
                            ' posDq is before posSq
                            '
                            If posDq > posClose Then
                                notFound = False
                            Else
                                '
                                ' skip forward to the next non-escaped dq
                                '
                                Do
                                    'Ptr = posDq + 1
                                    posDq = InStr(posDq + 1, src, """")
                                    escape = ""
                                    If posDq > 0 Then
                                        escape = Mid(src, posDq - 1, 1)
                                    End If
                                Loop While (escape = "\")
                                Ptr = posDq + 1
                                'notFound = False
                            End If
                    End Select
                End If
            Loop While notFound
        End If
        If posOpen <= 0 Then
            '
            ' no cmd found, add from the last ptr to the end
            '
            dst = dst & Mid(src, ptrLast)
            Ptr = -1
        Else
            '
            ' cmd found, process it and add the results to the dst
            '
            cmd = Mid(src, posOpen + 2, (posClose - posOpen - 2))
            cmdResult = ExecuteAllCmdLists_Execute(csv, mainOrNothing, cmd, badCmd, Context, personalizationPeopleId, personalizationIsAuthenticated, dataVersion)
            If badCmd Then
                '
                ' the command was bad, put it back in place (?) in case it was not a command
                '
                cmdResult = contentReplaceEscapeStart & cmd & contentReplaceEscapeEnd
            End If
            dst = dst & Mid(src, ptrLast, posOpen - ptrLast) & cmdResult
            Ptr = posClose + 2
        End If
        ptrLast = Ptr
    Loop While (Ptr > 1)
    '
    ExecuteAllCmdLists = dst
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call csv.reportError("trap", "aoPrimitives.contentCmdClass", True)
End Function
'
'=================================================================================================================
'   EncodeActiveContent - Execute Content Command Source
'=================================================================================================================
'
Private Function ExecuteAllCmdLists_Execute(csv As Object, mainOrNothing As Object, cmdSrc As String, return_BadCmd As Boolean, Context As AddonContextEnum, personalizationPeopleId As Long, personalizationIsAuthenticated As Boolean, dataVersion As String) As String
    On Error GoTo ErrorTrap
    '
    ' accumulator gets the result of each cmd, then is passed to the next command to filter
    '
    Dim CmdAccumulator As String
    '
    Dim CS As Long
    Dim htmlTools As htmlToolsClass
    Dim argField As String
    Dim importHead As String
    Dim argFind As String
    Dim argReplace As String
    Dim ArgName As String
    Dim ArgInstanceId As String
    Dim ArgGuid As String
    Dim ArgOptionString As String
    Dim ArgOptionStringHTMLEncoded As String
    '
    Dim cmd As String
    Dim cmdArg As String
    Dim Pos As Long
    Dim AddonOptionString As String
    Dim pairNames As Variant
    Dim i As Long
    Dim pairName As String
    Dim pairValue As String
    Dim ValueVariant As Variant
    Dim posNextStart As Long
    Dim posArgOpen As Long
    Dim posArgClose As Long
    Dim posBraceOpen As Long
    Dim posBraceClose As Long
    Dim cmdArgs As String
    Dim ACGuid As String
    Dim addonName As String
    Dim ACInstanceID As String
    Dim AddonOptionStringHTMLEncoded As String
    Dim CSPeople As Long
    Dim CSPeopleSet As Long
    Dim Ptr As Long
    Dim WorkingContent As String
    'Dim cmdDict As Dictionary
    'Dim cmdColl As Collection
    Dim cmdStringOrDictionary As Object
    Dim cmdDictionaryOrCollection As Object
    Dim cmdCollection As Collection
    Dim cmdDef As Dictionary
    Dim cmdArgDef As Dictionary
    Dim cmdArgTest As Object
    Dim triming As Boolean
    Dim leftChr As String
    Dim rightChr As String
    Dim trimLen As Long
    Dim whiteChrs As String
    Dim trimming As Boolean
    '
    Set htmlTools = New htmlToolsClass
    '
    cmdSrc = Trim(cmdSrc)
    whiteChrs = vbCr & vbLf & vbTab & " "
    Do
        trimming = False
        trimLen = Len(cmdSrc)
        If trimLen > 0 Then
            leftChr = Left(cmdSrc, 1)
            rightChr = Right(cmdSrc, 1)
            If InStr(1, whiteChrs, leftChr) <> 0 Then
                cmdSrc = Mid(cmdSrc, 2)
                trimming = True
            End If
            If InStr(1, whiteChrs, rightChr) <> 0 Then
                cmdSrc = Mid(cmdSrc, 1, Len(cmdSrc) - 1)
                trimming = True
            End If
        End If
    Loop While trimming
    CmdAccumulator = ""
    If cmdSrc <> "" Then
        '
        ' convert cmdSrc to cmdCollection
        '   cmdCollection is a collection of
        '       1) dictionary objects
        '       2) strings
        '
        ' the cmdSrc can be one of three things:
        '   - [a,b,c,d] a JSON array - parseJSON returns a collection
        '       - leave as collection
        '   - {a:b,c:d} a JSON object - parseJSON returns a dictionary
        '       - convert to a collection of each dictionaries
        '   - a "b" - do not use the parseJSON
        '       - just make a collection
        '
        If (Left(cmdSrc, 1) = "{") And (Right(cmdSrc, 1) = "}") Then
            '
            ' JSON is a single command in the form of an object, like:
            '   { "import": "test.html" }
            '
            On Error Resume Next
            Set cmdDictionaryOrCollection = csv.parseJSON(cmdSrc)
            If Err.Number <> 0 Then
                Call csv.reportError("Error parsing JSON command list [" & GetErrString() & "]", "ExecuteAllCmdLists_Execute")
                Err.Clear
                Exit Function
            ElseIf (LCase(TypeName(cmdDictionaryOrCollection)) <> "dictionary") Then
                Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing JSON command list, expected a single command, command list [" & cmdSrc & "]")
                Exit Function
            Else
                '
                '   Top leve should be an array of commands (or objects) [a,b,c,d]
                '       if it is an object {a:b,c:d,}, convert to arry [{a:b},{c:d}]
                ' convert dictionary of objects to a collection of objects
                '
                Dim dictionaryKeys() As Variant
                Dim key As Variant
                Dim itemObject As Object
                Dim itemVariant As Variant
                '
                Set cmdCollection = New Collection
                Dim cmdObject As Dictionary
                '
                dictionaryKeys = cmdDictionaryOrCollection.Keys
                For Each key In dictionaryKeys
                    If IsObject(cmdDictionaryOrCollection.Item(key)) Then
                        Set cmdObject = New Dictionary
                        Set itemObject = cmdDictionaryOrCollection.Item(key)
                        Call cmdObject.Add(key, itemObject)
                        Call cmdCollection.Add(cmdObject)
                    Else
                        Set cmdObject = New Dictionary
                        itemVariant = cmdDictionaryOrCollection.Item(key)
                        Call cmdObject.Add(key, itemVariant)
                        Call cmdCollection.Add(cmdObject)
                    End If
                Next
'                If (cmdDictionaryOrCollection.Count = 1) Then
'                    Set cmdCollection = New Collection
'                    Call cmdCollection.Add(cmdDictionaryOrCollection.Item(0), cmdDictionaryOrCollection.key(0))
'                Else
'                    For Each cmdString In cmdDictionaryOrCollection
'                        cmdCollection.Add (cmdString)
'                    Next
'                End If
            End If
        ElseIf (Left(cmdSrc, 1) = "[") And (Right(cmdSrc, 1) = "]") Then
            '
            ' JSON is a command list in the form of an array, like:
            '   [ "clear" , { "import": "test.html" },{ "open" : "myfile.txt" }]
            '
            On Error Resume Next
            Set cmdDictionaryOrCollection = csv.parseJSON(cmdSrc)
            If Err.Number <> 0 Then
                Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing JSON command list, error [" & GetErrString() & "], command list [" & cmdSrc & "]")
                Err.Clear
                Exit Function
            ElseIf (LCase(TypeName(cmdDictionaryOrCollection)) <> "collection") Then
                Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing JSON command list, expected a command list but parser did not return list, command list [" & cmdSrc & "]")
                Exit Function
            Else
                '
                ' assign command array
                '
                Set cmdCollection = cmdDictionaryOrCollection
            End If
        Else
            '
            ' a single text command without JSON wrapper, like
            '   open myfile.html
            '   open "myfile.html"
            '   "open" "myfile.html"
            '   "content box"
            '   all other posibilities are syntax errors
            '
            cmd = Trim(cmdSrc)
            cmdArg = ""
            If Mid(cmd, 1, 1) = """" Then
                '
                'cmd is quoted
                '   "open"
                '   "Open" file
                '   "Open" "file"
                '
                Pos = InStr(2, cmd, """")
                If Pos <= 1 Then
                    Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing content command [" & cmdSrc & "], expected a close quote around position " & Pos)
                    Exit Function
                Else
                    If Pos = Len(cmd) Then
                        '
                        ' cmd like "open"
                        '
                        cmdArg = ""
                        cmd = Mid(cmd, 2, Pos - 2)
                    ElseIf Mid(cmd, Pos + 1, 1) <> " " Then
                        '
                        ' syntax error, must be a space between cmd and argument
                        '
                        Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing content command [" & cmdSrc & "], expected a space between command and argument around position " & Pos)
                        Exit Function
                    Else
                        cmdArg = Trim(Mid(cmd, Pos + 1))
                        cmd = Mid(cmd, 2, Pos - 2)
                    End If
                End If
                
            Else
                '
                ' no quotes, can be
                '   open
                '   open file
                '
                Pos = InStr(1, cmd, " ")
                If Pos > 0 Then
                    cmdArg = Mid(cmdSrc, Pos + 1)
                    cmd = Trim(Mid(cmdSrc, 1, Pos - 1))
                End If
            End If
            If Mid(cmdArg, 1, 1) = """" Then
                '
                'cmdarg is quoted
                '
                Pos = InStr(2, cmdArg, """")
                If Pos <= 1 Then
                    Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing JSON command list, expected a quoted command argument, command list [" & cmdSrc & "]")
                    Exit Function
                Else
                    cmdArg = Mid(cmdArg, 2, Pos - 2)
                End If
            End If
            If (Left(cmdArg, 1) = "{") And (Right(cmdArg, 1) = "}") Then
                '
                ' argument is in the form of an object, like:
                '   { "text name": "my text" }
                '
                On Error Resume Next
                Set cmdDictionaryOrCollection = csv.parseJSON(cmdArg)
                If Err.Number <> 0 Then
                    Call csv.reportError("Error parsing JSON command argument list [" & GetErrString() & "]", "ExecuteAllCmdLists_Execute")
                    Err.Clear
                    Exit Function
                ElseIf (LCase(TypeName(cmdDictionaryOrCollection)) <> "dictionary") Then
                    Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing JSON command argument list, expected a single command, command list [" & cmdSrc & "]")
                    Exit Function
                Else
                    '
                    ' create command array of one command
                    '
                    Set cmdCollection = New Collection
                    Call cmdCollection.Add(cmdDictionaryOrCollection)
                End If
                Set cmdDef = New Dictionary
                Call cmdDef.Add(cmd, cmdDictionaryOrCollection)
                Set cmdCollection = New Collection
                Call cmdCollection.Add(cmdDef)
            Else
                '
                ' command and arguments are strings
                '
                Set cmdDef = New Dictionary
                Call cmdDef.Add(cmd, cmdArg)
                Set cmdCollection = New Collection
                Call cmdCollection.Add(cmdDef)
            End If
        End If
        '
        ' execute the commands in the JSON cmdCollection
        '
        'Dim cmdVariant As Variant
        
        For Each cmdStringOrDictionary In cmdCollection
            '
            ' repeat for all commands in the collection:
            ' convert each command in the command array to a cmd string, and a cmdArgDef dictionary
            ' each cmdStringOrDictionary is a command. It may be:
            '   A - "command"
            '   B - { "command" }
            '   C - { "command" : "single-default-argument" }
            '   D - { "command" : { "name" : "The Name"} }
            '   E - { "command" : { "name" : "The Name" , "secondArgument" : "secondValue" } }
            '
            If LCase(TypeName(cmdStringOrDictionary)) = "string" Then
                '
                ' case A & B, the cmdDef is a string
                '
                cmd = cmdStringOrDictionary
                Set cmdArgDef = New Dictionary
            ElseIf LCase(TypeName(cmdStringOrDictionary)) = "dictionary" Then
                '
                ' cases C-E, (0).key=cmd, (0).value = argument (might be string or object)
                '
                Set cmdDef = cmdStringOrDictionary
                If cmdDef.Count <> 1 Then
                    '
                    ' syntax error
                    '
                Else
                    cmd = cmdDef.Keys(0)
                    If LCase(TypeName(cmdDef.Item(cmd))) = "string" Then
                        '
                        ' command definition with default argument
                        '
                        Set cmdArgDef = New Dictionary
                        Call cmdArgDef.Add("default", cmdDef.Item(cmd))
                    ElseIf LCase(TypeName(cmdDef.Item(cmd))) = "dictionary" Then
                        Set cmdArgDef = cmdDef.Item(cmd)
                    Else
                        '
                        ' syntax error, bad command
                        '
                        Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing JSON command list, , command list [" & cmdSrc & "]")
                        Err.Clear
                        Exit Function
                    End If
                End If
            Else
                '
                ' syntax error
                '
                Call HandleClassErrorAndResume(csv, "ExecuteAllCmdLists_Execute", "Error parsing JSON command list, , command list [" & cmdSrc & "]")
                Err.Clear
                Exit Function
            End If
            '
            ' execute the cmd with cmdArgDef dictionary
            '
            Select Case LCase(cmd)
                Case "textbox"
                    '
                    ' Opens a textbox addon (patch for text box name being "text name" so it requies json)copy content record
                    '
                    ' arguments
                    '   name: copy content record
                    ' default
                    '   name
                    '
                    CmdAccumulator = ""
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "name", "default"
                                ArgName = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If ArgName <> "" Then
                        On Error Resume Next
                        If (dataVersion >= "4.1.456") Then
                            CmdAccumulator = csv.GetContentCopy(ArgName, "copy content", "", personalizationPeopleId, mainOrNothing, False)
                        ElseIf Not (mainOrNothing Is Nothing) Then
                            CmdAccumulator = mainOrNothing.GetContentCopy(ArgName)
                        End If
                    End If
                Case "opencopy"
                    '
                    ' Opens a copy content record
                    '
                    ' arguments
                    '   name: layout record name
                    ' default
                    '   name
                    '
                    CmdAccumulator = ""
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "name", "default"
                                ArgName = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If ArgName <> "" Then
                        On Error Resume Next
                        CS = csv.OpenCSContent("copy content", "name=" & KmaEncodeSQLText(ArgName), "id", , , , , "copy")
                        If csv.IsCSOK(CS) Then
                            CmdAccumulator = csv.GetCS(CS, "copy")
                        End If
                        Call csv.CloseCS(CS)
                    End If
                Case "openlayout"
                    '
                    ' Opens a layout record
                    '
                    ' arguments
                    '   name: layout record name
                    ' default
                    '   name
                    '
                    CmdAccumulator = ""
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "name", "default"
                                ArgName = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If ArgName <> "" Then
                        On Error Resume Next
                        CS = csv.OpenCSContent("layouts", "name=" & KmaEncodeSQLText(ArgName), "id", , , , , "layout")
                        If csv.IsCSOK(CS) Then
                            CmdAccumulator = csv.GetCS(CS, "layout")
                        End If
                        Call csv.CloseCS(CS)
                    End If
                Case "open"
                    '
                    ' Opens a file in the wwwPath
                    '
                    ' arguments
                    '   name: filename
                    ' default
                    '   name
                    '
                    CmdAccumulator = ""
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "name", "default"
                                ArgName = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If ArgName <> "" Then
                        On Error Resume Next
                        CmdAccumulator = csv.ReadFile(csv.PhysicalWWWPath & ArgName)
                        Err.Clear
                    End If
                Case "import"
                    '
                    ' Opens an html file in the wwwPath and imports the head and returns the body
                    '
                    ' arguments
                    '   name: filename
                    ' default argument
                    '   name
                    '
                    CmdAccumulator = ""
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "name", "default"
                                ArgName = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If ArgName <> "" Then
                        On Error Resume Next
                        CmdAccumulator = csv.ReadFile(csv.PhysicalWWWPath & ArgName)
                        Err.Clear
                        If CmdAccumulator <> "" Then
                            importHead = GetTagInnerHTML(CmdAccumulator, "head", False)
                            If importHead <> "" Then
                                ' try this, but it may not be implemented yet
                                On Error Resume Next
                                Call csv.addHeadTags(importHead)
                                Err.Clear
                                On Error GoTo ErrorTrap
                                'EncodeContent_HeadTags = EncodeContent_HeadTags & vbCrLf & vbTab & importHead
                            End If
                            CmdAccumulator = GetTagInnerHTML(CmdAccumulator, "body", False)
                        End If
                    End If
                Case "user"
                    '
                    ' returns the value of the current users field
                    '
                    ' arguments
                    '   field: fieldName
                    ' default
                    '   field
                    '
                    CmdAccumulator = ""
                    argField = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "field", "default"
                                argField = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If argField = "" Then
                        argField = "name"
                    End If
                    If Not CSPeopleSet Then
                        CSPeople = csv.OpenCSContentRecord("People", personalizationPeopleId)
                        CSPeopleSet = True
                    End If
                    If csv.IsCSOK(CSPeople) And csv.IsCSFieldSupported(CSPeople, argField) Then
                        CmdAccumulator = csv.GetCS(CSPeople, argField)
                    End If
                Case "site"
                    '
                    ' returns a site property
                    '
                    ' arguments
                    '   name: the site property name
                    ' default argument
                    '   name
                    '
                    CmdAccumulator = ""
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "name", "default"
                                ArgName = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If ArgName <> "" Then
                        On Error Resume Next
                        CmdAccumulator = csv.GetSiteProperty(ArgName, "", 0)
                    End If
                Case "set"
                    '
                    ' does a find and replace
                    '
                    ' arguments
                    '   find: what to search for in teh accumulator
                    '   replace: what to replace it with
                    ' default argument
                    '   find
                    '
                    'CmdAccumulator = ""
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "find"
                                argFind = cmdArgDef.Items(Ptr)
                            Case "replace"
                                argReplace = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If argFind <> "" Then
                        On Error Resume Next
                        CmdAccumulator = Replace(CmdAccumulator, argFind, argReplace, , , vbTextCompare)
                    End If
                Case "setinner"
                    '
                    ' does a find and replace on the inner HTML of an element identified by its class selector
                    '
                    ' arguments
                    '   find: what to search for in teh accumulator
                    '   replace: what to replace it with
                    ' default argument
                    '   find
                    '
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "find"
                                argFind = cmdArgDef.Items(Ptr)
                            Case "replace"
                                argReplace = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If argFind <> "" Then
                        On Error Resume Next
                        CmdAccumulator = htmlTools.insertInnerHTML(Nothing, CmdAccumulator, argFind, argReplace)
                    End If
                Case "getinner"
                    '
                    ' returns the inner HTML of an element identified by its class selector
                    '
                    ' arguments
                    '   find: what to search for in teh accumulator
                    ' default argument
                    '   find
                    '
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "find"
                                argFind = cmdArgDef.Items(Ptr)
                            Case "replace"
                                argReplace = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If argFind <> "" Then
                        On Error Resume Next
                        CmdAccumulator = htmlTools.getInnerHTML(Nothing, CmdAccumulator, argFind)
                    End If
                Case "setouter"
                    '
                    ' does a find and replace on the outer HTML of an element identified by its class selector
                    '
                    ' arguments
                    '   find: what to search for in teh accumulator
                    '   replace: what to replace it with
                    ' default argument
                    '   find
                    '
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "find"
                                argFind = cmdArgDef.Items(Ptr)
                            Case "replace"
                                argReplace = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If argFind <> "" Then
                        On Error Resume Next
                        CmdAccumulator = htmlTools.insertOuterHTML(Nothing, CmdAccumulator, argFind, argReplace)
                    End If
                Case "getouter"
                    '
                    ' returns the outer HTML of an element identified by its class selector
                    '
                    ' arguments
                    '   find: what to search for in teh accumulator
                    ' default argument
                    '   find
                    '
                    ArgName = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "find"
                                argFind = cmdArgDef.Items(Ptr)
                            Case "replace"
                                argReplace = cmdArgDef.Items(Ptr)
                        End Select
                    Next
                    If argFind <> "" Then
                        On Error Resume Next
                        CmdAccumulator = htmlTools.getOuterHTML(Nothing, CmdAccumulator, argFind)
                    End If
                Case "addon"
                    '
                    ' execute an add-on
                    '
                    addonName = ""
                    ArgInstanceId = ""
                    ArgGuid = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "addon"
                                addonName = cmdArgDef.Items(Ptr)
                            Case "instanceid"
                                ArgInstanceId = cmdArgDef.Items(Ptr)
                            Case "guid"
                                ArgGuid = cmdArgDef.Items(Ptr)
                            Case Else
                                ArgOptionString = ArgOptionString & "&" & csv.EncodeAddonOptionArgument(kmaEncodeText(cmdArgDef.Keys(Ptr))) & "=" & csv.EncodeAddonOptionArgument(kmaEncodeText(cmdArgDef.Items(Ptr)))
                        End Select
                    Next
                    ArgOptionString = ArgOptionString & "&cmdAccumulator=" & csv.EncodeAddonOptionArgument(CmdAccumulator)
                    ArgOptionString = Mid(ArgOptionString, 2)
Dim addonStatusOK As Boolean
                    CmdAccumulator = csv.ExecuteAddon2(0, addonName, ArgOptionString, Context, "", 0, "", "", False, 0, "", addonStatusOK, Nothing, "", mainOrNothing, "", personalizationPeopleId, personalizationIsAuthenticated)
                    'CmdAccumulator = mainOrNothing.ExecuteAddon3(addonName, ArgOptionString, Context)
                Case Else
                    '
                    ' attempts to execute an add-on with the command name
                    '
                    addonName = cmd
                    ArgInstanceId = ""
                    ArgGuid = ""
                    For Ptr = 0 To cmdArgDef.Count - 1
                        Select Case LCase(cmdArgDef.Keys(Ptr))
                            Case "instanceid"
                                ArgInstanceId = cmdArgDef.Items(Ptr)
                            Case "guid"
                                ArgGuid = cmdArgDef.Items(Ptr)
                            Case Else
                                ArgOptionString = ArgOptionString & "&" & csv.EncodeAddonOptionArgument(kmaEncodeText(cmdArgDef.Keys(Ptr))) & "=" & csv.EncodeAddonOptionArgument(kmaEncodeText(cmdArgDef.Items(Ptr)))
                        End Select
                    Next
                    ArgOptionString = ArgOptionString & "&cmdAccumulator=" & csv.EncodeAddonOptionArgument(CmdAccumulator)
                    ArgOptionString = Mid(ArgOptionString, 2)
                    CmdAccumulator = csv.ExecuteAddon2(0, addonName, ArgOptionString, Context, "", 0, "", "", False, 0, "", addonStatusOK, Nothing, "", mainOrNothing, "", personalizationPeopleId, personalizationIsAuthenticated)
                    'CmdAccumulator = mainOrNothing.ExecuteAddon3(addonName, ArgOptionString, Context)
            End Select
        Next
    End If
    '
    ExecuteAllCmdLists_Execute = CmdAccumulator
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call csv.reportError("trap", "aoPrimitives.execute", True)
End Function
'
'
'
Private Sub HandleClassError(csv As Object, method As String, cause As String, ResumeNext As Boolean)
    Call csv.reportError(cause, "aoPrimitives.contentCmdClass." & method, "aoPrimitives.execute", ResumeNext)
End Sub
'
'
'
Private Sub HandleClassErrorAndResume(csv As Object, method As String, cause As String)
    Call csv.reportError(cause, "aoPrimitives.contentCmdClass." & method, "aoPrimitives.execute", True)
End Sub


