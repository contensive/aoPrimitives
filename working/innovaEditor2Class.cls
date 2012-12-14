VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "innovaEditor2Class"
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
Private main As Object
'Private Main As ccWeb3.MainClass
Private csv As Object
'
'=================================================================================
'
'=================================================================================
'
Public Function execute(CsvObject As Object, MainObject As Object, optionString As String, filterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim editorName As String
    Dim editorValue As String
    Dim editorFieldType As Long
    Dim editorReadOnly As Boolean
    Dim editorWidth As String
    Dim editorHeight As String
    Dim editorAllowActiveContent As Boolean
    Dim editorAddonList As String
    Dim editorStyles As String
    Dim editorStyleOptions As String
    '
    Set main = MainObject
    Set csv = CsvObject
    '
    editorName = main.GetAddonOption("editorname", optionString)
    editorValue = main.GetAddonOption("editorValue", optionString)
    editorFieldType = kmaEncodeInteger(main.GetAddonOption("editorFieldType", optionString))
    editorReadOnly = kmaEncodeBoolean(main.GetAddonOption("editorReadOnly", optionString))
    editorWidth = main.GetAddonOption("editorWidth", optionString)
    editorHeight = main.GetAddonOption("editorHeight", optionString)
    editorAllowActiveContent = kmaEncodeBoolean(main.GetAddonOption("editorAllowActiveContent", optionString))
    editorAddonList = main.GetAddonOption("editorAddonList", optionString)
    editorStyles = main.GetAddonOption("editorStyles", optionString)
    editorStyleOptions = main.GetAddonOption("editorStyleOptions", optionString)
    '
    execute = getInnovaEditorAddon(main, editorName, editorStyles, editorValue, editorHeight, editorWidth, editorAllowActiveContent, editorReadOnly, editorAddonList)
    '
    Exit Function
ErrorTrap:
    HandleClassTrapError ("execute")
End Function
'
'========================================================================
'   The addon interface for html editors
'========================================================================
'
Friend Function getInnovaEditorAddon(main As Object, fieldname As String, styleList As String, DefaultValue As String, Height As String, Width As String, AllowActiveContent As Boolean, ReadOnly As Boolean, addonListJSON As String)
    On Error GoTo ErrorTrap
    '
    Const DefaultValueFPO = "<DefaultValue>"
    Const WidthFPO = "<Width>"
    Const HeightFPO = "<Height>"
    Const FieldNameFPO = "<FieldName>"
    Const EditorNameFPO = "<EditorName>"
    '
    Dim styleRules As String
    Dim EditorStyleRulesFilename As String
    'Dim fso As New FileSystemObject
    Dim SourceLeft As String
    Dim SourceRight As String
    Dim FeatureName As String
    Dim Files() As String
    Dim Filename As String
    Dim Ptr As Long
    Dim BS() As String
    Dim BPtr As Long
    Dim BCnt As Long
    Dim cmd As String
    Dim BodyCmd As String
    'Dim LCaseStyleName As String
    Dim StyleCaption As String
    Dim FeatureFile As String
    Dim FeatureRows() As String
    Dim FeatureList As String
    Dim posStart As Long
    Dim posEnd As Long
    Dim SourceFeatureList As String
    Dim SourceFeatures() As String
    Dim StyleNames() As String
    Dim StyleNameCnt As Long
    Dim StylePtr As Long
    Dim FirstChr As String
    Dim IncludeInEditor As Boolean
    Dim SourceFilename As String
    Dim SourceEditor As String
    Dim CutPosition As Long
    Dim iDefaultValue As String
    Dim EditorObjectName As String
    Dim iWidth As String
    Dim iHeight As String
    Dim SourceMenu As String
    Dim EditorStyleRules As String
    Dim styleSource As String
    Dim StyleLines() As String
    Dim LinePtr As Long
    Dim LineCnt As Long
    Dim Splits() As String
    Dim StyleName As String
    Dim StyleCommands As String
    Dim LoopPtr As Long
    'Dim EditorAddonMenu As String
    Dim MediaStartPtr As Long
    Dim MediaEndPtr As Long
    Dim BraceCnt As Long
    Dim BracePtr As Long
    Dim BraceOpenPtr As Long
    Dim BraceClosedPtr As Long
    Dim BraceOpenCnt As Long
    Dim BraceClosedCnt As Long
    Dim AbortParse As Boolean
    Dim ParsePtr As Long
    Dim Raw As String
    Dim innovaCnt As Long
    '
    ' Setup replacements for Raw
    '
    innovaCnt = kmaEncodeInteger(main.GetViewingProperty("innovaCnt", "0"))
    iDefaultValue = DefaultValue
    If AllowActiveContent And (iDefaultValue <> "") Then
        iDefaultValue = main.EncodeContent5(iDefaultValue, 0, "", 0, 0, False, False, False, True, True, False, "", "", False, 0)
    End If
    iDefaultValue = Replace(iDefaultValue, "&", "&amp;")
    iDefaultValue = Replace(iDefaultValue, "<", "&lt;")
    iDefaultValue = Replace(iDefaultValue, ">", "&gt;")
    If Width <> "" Then
        iWidth = Width
    Else
        iWidth = main.SiteProperty_DefaultFormInputWidth
    End If
    'iWidth = KmaEncodeMissingText(Width, Main.SiteProperty_DefaultFormInputWidth)
    If (iWidth = "") Or (iWidth = "0") Then
        iWidth = main.SiteProperty_DefaultFormInputWidth
    End If
    If IsNumeric(iWidth) Then
        iWidth = iWidth & "px"
    End If
    If Height <> "" Then
        iHeight = Height
    Else
        iHeight = main.SiteProperty_DefaultFormInputHTMLHeight
    End If
    'iHeight = KmaEncodeMissingText(Height, Main.SiteProperty_DefaultFormInputHTMLHeight)
    If IsNumeric(iHeight) Then
        iHeight = iHeight & "px"
    End If
    '
    EditorObjectName = "oEdit" & innovaCnt
    '
    Raw = GetInnovaEditorCache(main, AllowActiveContent)
    If Raw = "" Then
        '
        ' Innova editor
        '
        If AllowActiveContent Then
            SourceEditor = main.ReadFile(main.PhysicalccLibPath & "\Editor\ActiveEditor.js")
        Else
            SourceEditor = main.ReadFile(main.PhysicalccLibPath & "\Editor\HTMLEditor.js")
        End If
        '
        SourceEditor = Replace(SourceEditor, "##ELEMENTNAME##", FieldNameFPO)
        SourceEditor = Replace(SourceEditor, "##EDITORNAME##", EditorNameFPO)
        SourceEditor = Replace(SourceEditor, "##WIDTH##", WidthFPO)
        SourceEditor = Replace(SourceEditor, "##HEIGHT##", HeightFPO)
        SourceEditor = Replace(SourceEditor, "##CONTENT##", DefaultValueFPO)
        SourceEditor = Replace(SourceEditor, "##CUSTOMTAG##", "EditorCustomTag")
        SourceEditor = Replace(SourceEditor, "##RESOURCELIBRARYURL##", main.SiteProperty_AdminURL)
        SourceEditor = Replace(SourceEditor, "##STYLERULES##", "EditorStyleRules")
        SourceEditor = Replace(SourceEditor, "##PUBLISHINGPATH##", main.ServerProtocol & main.ServerHost & main.ServerAppRootPath)
        '
        ' Block features based on InnovaEditorFeaturefilename
        '
        FeatureFile = main.ReadVirtualFile(InnovaEditorFeaturefilename)
        If FeatureFile = "" Then
            FeatureFile = main.ReadFile(main.PhysicalccLibPath & "\Config\DefaultEditorConfig.txt")
            Call main.SaveVirtualFile(InnovaEditorFeaturefilename, FeatureFile)
        End If
        If FeatureFile <> "" Then
            FeatureRows = Split(FeatureFile, vbCrLf)
            If UBound(FeatureRows) < 2 Then
                FeatureFile = ""
            Else
                If Not AllowActiveContent Then
                    '
                    ' public editor
                    '
                    FeatureList = FeatureRows(2)
                ElseIf Not main.IsAdmin Then
                    '
                    ' content manager
                    '
                    FeatureList = FeatureRows(1)
                Else
                    '
                    ' administrator
                    '
                    FeatureList = FeatureRows(0)
                End If
                
            End If
        End If
        If FeatureFile = "" Then
            If Not AllowActiveContent Then
                FeatureList = InnovaEditorPublicFeatureList
            ElseIf Not main.IsAdmin Then
                FeatureList = InnovaEditorFeatureList
            Else
                FeatureList = InnovaEditorFeatureList
            End If
        End If
        '
        posStart = InStr(1, SourceEditor, ".features=[", vbTextCompare)
        If posStart > 0 Then
            posStart = posStart + 11
            posEnd = InStr(posStart, SourceEditor, "]", vbTextCompare)
            If posEnd > 0 Then
                posEnd = posEnd - 1
                SourceRight = Mid(SourceEditor, posEnd + 1)
                SourceLeft = Mid(SourceEditor, 1, posStart - 1)
                SourceFeatureList = Mid(SourceEditor, posStart, posEnd - posStart + 1)
                SourceFeatureList = Replace(SourceFeatureList, vbCrLf, "")
                SourceFeatures = Split(SourceFeatureList, ",")
                For Ptr = 0 To UBound(SourceFeatures)
                    FeatureName = LCase(SourceFeatures(Ptr))
                    If Mid(FeatureName, 1, 1) = """" Then
                        FeatureName = Mid(FeatureName, 2, Len(FeatureName) - 2)
                    End If
                    Select Case FeatureName
                        Case "|"
                        Case "brk"
                        Case Else
                            If InStr(1, "," & FeatureList & ",", "," & FeatureName & ",", vbTextCompare) = 0 Then
                                SourceFeatures(Ptr) = ""
                            End If
                    End Select
                Next
                SourceFeatureList = Join(SourceFeatures, ",")
                Do While InStr(1, SourceFeatureList, ",,") <> 0
                    SourceFeatureList = Replace(SourceFeatureList, ",,", ",")
                Loop
                Do While InStr(1, SourceFeatureList, "|"",""|") <> 0
                    SourceFeatureList = Replace(SourceFeatureList, "|"",""|", "|")
                Loop
                'SourceFeatureList = Replace(SourceFeatureList, "|"",""|", "|")
                SourceEditor = SourceLeft & SourceFeatureList & SourceRight
            End If
        End If
        '
        ' Build the final table
        '
        Raw = Raw _
            & "<table border=0 cellpadding=0 cellspacing=0 width=""100%"">" _
            & "<tr><td colspan=2><image src=""/ccLib/images/spacer.gif"" width=400 height=1></td></tr>" _
            & "<tr><td><image src=""/ccLib/images/spacer.gif"" width=1 height=505></td><td>" & SourceEditor & "</td></tr>" _
            & "</table>" _
            & ""
        '
        ' Save the Raw template back
        '
        Call SetInnovaEditorCache(main, Raw, AllowActiveContent)
    End If
    '
    ' replace custom parts back into the raw input
    '
    If innovaCnt = 0 Then
        Call main.AddHeadScriptLink("/ccLib/Editor/scripts/innovaeditor.js", "wysiwyg editor")
        '
        ' Add Styles code to page
        '
        styleRules = GetInnovaEditorStyleRules(main, styleList)
        Call main.AddHeadJavascript("var EditorStyleRules=" & styleRules)
        'EditorStyleRulesFilename = Replace(EditorStyleRulesFilenamePattern, "$templateid$", TemplateIDForStyles, , , vbTextCompare)
        'If Not fso.FileExists(Main.PhysicalFilePath & EditorStyleRulesFilename) Then
        '    Call Main.SaveVirtualFile(EditorStyleRulesFilename, "var EditorStyleRules=" & Main.GetInnovaEditorStyleRules(contenttypeweb, TemplateIDForStyles, emailId))
        'End If
        'Call Main.AddHeadScriptLink(Main.ServerFilePath & EditorStyleRulesFilename, "wysiwyg editor")
'
' Can not save Editor Addon Menu to file because the cache has to flush every time anything is edited -- because
'   any addon can have a selector that selects any content -- the answer might be ajaxing the bubble edits. That way,
'   only the ones that are needed are built, and they are current.
'
        Call main.AddHeadScriptCode("var EditorCustomTag=" & addonListJSON, "wysiwyg editor")
    End If
    Raw = Replace(Raw, DefaultValueFPO, iDefaultValue)
    Raw = Replace(Raw, WidthFPO, iWidth)
    Raw = Replace(Raw, HeightFPO, iHeight)
    Raw = Replace(Raw, FieldNameFPO, fieldname)
    Raw = Replace(Raw, EditorNameFPO, EditorObjectName)
    '
    getInnovaEditorAddon = Raw
    'getInnovaEditorAddon = "<div class=""ccEditorTextCon"">" & Raw & "</div>"
    Call main.SetViewingProperty("innovaCnt", CStr(innovaCnt + 1))
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set main = Nothing
    Call HandleError("innovaEditorAddonClass", "getInnovaEditorAddon", Err.Number, Err.source, Err.Description, True, False)
End Function
'
'
'
Friend Function GetInnovaEditorStyleRules(main As Object, styleSheet As String) As String
    On Error GoTo ErrorTrap
    '
    Const styleRulesKey = "innovaEditorStyleRules"
    '
    Dim s As String
    '
    Dim SourceLeft As String
    Dim SourceRight As String
    Dim FeatureName As String
    Dim Files() As String
    Dim Filename As String
    Dim Ptr As Long
    Dim BS() As String
    Dim BPtr As Long
    Dim BCnt As Long
    Dim cmd As String
    Dim BodyCmd As String
    'Dim LCaseStyleName As String
    Dim StyleCaption As String
    Dim FeatureFile As String
    Dim FeatureRows() As String
    Dim FeatureList As String
    Dim posStart As Long
    Dim posEnd As Long
    Dim SourceFeatureList As String
    Dim SourceFeatures() As String
    Dim StyleNames() As String
    Dim StyleNameCnt As Long
    Dim StylePtr As Long
    Dim FirstChr As String
    Dim IncludeInEditor As Boolean
    Dim SourceFilename As String
    Dim SourceEditor As String
    Dim CutPosition As Long
    Dim iDefaultValue As String
    Dim EditorObjectName As String
    Dim iWidth As String
    Dim iHeight As String
    Dim SourceMenu As String
    Dim EditorStyleRules As String
    Dim styleSource As String
    Dim StyleLines() As String
    Dim LinePtr As Long
    Dim LineCnt As Long
    Dim Splits() As String
    Dim StyleName As String
    Dim StyleCommands As String
    Dim LoopPtr As Long
    Dim AddonMenuJson As String
    Dim MediaStartPtr As Long
    Dim MediaEndPtr As Long
    Dim BraceCnt As Long
    Dim BracePtr As Long
    Dim BraceOpenPtr As Long
    Dim BraceClosedPtr As Long
    Dim BraceOpenCnt As Long
    Dim BraceClosedCnt As Long
    Dim AbortParse As Boolean
    Dim ParsePtr As Long
    Dim Raw As String
    Dim CS As Long
    Dim SQL As String
    Dim OtherHeadTags As String
    '
    EditorStyleRules = main.GetViewingProperty(styleRulesKey, "")
    If EditorStyleRules = "" Then
        styleSource = styleSheet
        If InStr(1, styleSource, "{") <> 0 Then
            '
            ' remove white space at the front of each line
            '
            Do While InStr(1, styleSource, vbCrLf & " ") <> 0
                styleSource = Replace(styleSource, vbCrLf & " ", vbCrLf)
            Loop
            '
            ' remove comments
            '
            posStart = InStr(1, styleSource, "/*")
            Do While posStart <> 0 And LoopPtr < 1000
                posEnd = InStr(posStart + 2, styleSource, "*/")
                If posEnd = 0 Then
                    posStart = 0
                Else
                    styleSource = Mid(styleSource, 1, posStart - 1) & Mid(styleSource, posEnd + 2)
                    posStart = InStr(1, styleSource, "/*")
                End If
                LoopPtr = LoopPtr + 1
            Loop
            '
            ' remove @import
            '
            posStart = InStr(1, styleSource, "@import", vbTextCompare)
            Do While posStart <> 0 And LoopPtr < 1000
                posEnd = InStr(posStart + 3, styleSource, vbCrLf)
                If posEnd = 0 Then
                    posStart = 0
                Else
                    styleSource = Mid(styleSource, 1, posStart - 1) & Mid(styleSource, posEnd + 2)
                    posStart = InStr(1, styleSource, vbCrLf & "@")
                End If
                LoopPtr = LoopPtr + 1
            Loop
'            '
'            ' Convert special case .ccBodyWeb or .ccBodyEmail styles to Body
'            '
'            Select Case contentType
'                Case contenttypeweb
'                    styleSource = Replace(styleSource, ".ccBodyWeb", "body", , , vbTextCompare)
'                    ' -- plus cover this old typo --
'                    styleSource = Replace(styleSource, ".bodyweb", "body", , , vbTextCompare)
'                Case contentTypeEmail
'                    styleSource = Replace(styleSource, ".ccbodyemail", "body", , , vbTextCompare)
'            End Select
            '
            ' remove disfunctional tags (embedded tag screws with editor)
            '
            styleSource = GetInnovaEditorStyleRules_FilterStyles(styleSource, "@media")
            styleSource = GetInnovaEditorStyleRules_FilterStyles(styleSource, ">")
            '
            ' remove contensive tags
            '      ok - back to the old way - there is something wrong with the innova editor, it will not hide the class selectors
            '      and, the .cc selectors should be contensive only, and should not be used anymore.
            styleSource = GetInnovaEditorStyleRules_FilterStyles(styleSource, ".cc")
            styleSource = GetInnovaEditorStyleRules_FilterStyles(styleSource, "A.cc")
            '
            ' remove illegal characters
            '
            ' double quotes are not illegal, they just need to be handled in the javascript lines for the editor
            'StyleSource = Replace(StyleSource, """", "")
            ' hope I am right... removed these because sometimes there are things like // on a line by themselves
            ' if crlf are gone, these things go to the beginning of the next stylename, which messes it up
            'StyleSource = Replace(StyleSource, vbCrLf, " ")
            'StyleSource = Replace(StyleSource, vbCr, " ")
            'StyleSource = Replace(StyleSource, vbLf, " ")
            styleSource = Replace(styleSource, vbTab, " ")
            '
            ' Convert to Innova format
            '
            StyleLines = Split(styleSource, "}")
            LineCnt = UBound(StyleLines) + 1
            s = ""
            BodyCmd = ""
            For LinePtr = 0 To LineCnt - 1
                Splits = Split(StyleLines(LinePtr), "{")
                If UBound(Splits) > 0 Then
                    StyleName = Trim(Splits(0))
                    StyleCommands = Trim(Splits(1))
                    'If InStr(1, StyleName, "ccTextAreaHeadOpened", vbTextCompare) <> 0 Then
                    'StyleName = StyleName
                    'End If
                    '
                    ' keep name only back to the first crlf
                    '
Dim Pos As Long
                    Pos = InStrRev(StyleName, vbCr)
                    If Pos > 0 Then
                        StyleName = Mid(StyleName, Pos + 1)
                    End If
                    Pos = InStrRev(StyleName, vbLf)
                    If Pos > 0 Then
                        StyleName = Mid(StyleName, Pos + 1)
                    End If
                    StyleName = Trim(StyleName)
                    '
                    ' remove spaces around colon in name
                    '
                    LoopPtr = 0
                    Do While (InStr(1, StyleName, " :") <> 0) And (LoopPtr < 100)
                        StyleName = Replace(StyleName, " :", ":")
                        LoopPtr = LoopPtr + 1
                    Loop
                    LoopPtr = 0
                    Do While (InStr(1, StyleName, ": ") <> 0) And (LoopPtr < 100)
                        StyleName = Replace(StyleName, ": ", ":")
                        LoopPtr = LoopPtr + 1
                    Loop
                    '
                    ' set empty style to a space
                    '
                    If StyleCommands = "" Then
                        StyleCommands = " "
                    End If
                    StyleCommands = kmaEncodeJavascript(StyleCommands)
                    StyleCommands = Replace(StyleCommands, """", "\""")
                    If StyleName <> "" Then
                        StyleNames = Split(StyleName, ",")
                        For StylePtr = 0 To UBound(StyleNames)
                            StyleName = Trim(StyleNames(StylePtr))
                            If LCase(StyleName) = "body" Then
                                '
                                ' Handle body at end
                                '
                                If Trim(StyleCommands) <> "" Then
                                    'StyleName = StyleName
                                    BodyCmd = BodyCmd & ";" & StyleCommands
                                End If
                            Else
                                If StyleName <> "" Then
                                    'LCaseStyleName = LCase(StyleName)
                                    FirstChr = Mid(StyleName, 1, 1)
                                    IncludeInEditor = (FirstChr = ".")
                                    If IncludeInEditor Then
                                        IncludeInEditor = IncludeInEditor And (InStr(1, StyleName, " ") = 0) And (InStr(1, StyleName, ",") = 0) And (InStr(1, StyleName, ">") = 0)
                                    End If
                                    'StyleCaption = ""
                                    If InStr(1, StyleCommands, """") <> 0 Then
                                        StyleCommands = StyleCommands
                                    End If
                                    StyleName = kmaEncodeJavascript(StyleName)
                                    StyleName = Replace(StyleName, """", "\""")
                                    StyleCaption = StyleName
                                    If IncludeInEditor Then
                                        s = s & vbCrLf & ",[""" & StyleName & """,true,""" & StyleCaption & """,""" & StyleCommands & """]"
                                    Else
                                        s = s & vbCrLf & ",[""" & StyleName & """,false,""" & StyleCaption & """,""" & StyleCommands & """]"
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
            If BodyCmd <> "" Then
                '
                ' Restore buffered body styles together at the front
                '
                BS = Split(BodyCmd, ";")
                BodyCmd = ""
                'BodyCmd = " padding:0px;"
                BCnt = UBound(BS) + 1
                For BPtr = 0 To BCnt - 1
                    cmd = LCase(Trim(BS(BPtr)))
                    If cmd = "" Then
                    ElseIf InStr(1, cmd, "background", vbTextCompare) <> 0 Then
                    ElseIf InStr(1, cmd, "padding", vbTextCompare) <> 0 Then
                    ElseIf InStr(1, cmd, "margin", vbTextCompare) <> 0 Then
                    Else
                        BodyCmd = BodyCmd & ";" & cmd
                    End If
                Next
            End If
            BodyCmd = BodyCmd & ";background-color:" & main.GetSiteProperty("Editor Background Color", "white")
            If BodyCmd <> "" Then
                If Mid(BodyCmd, 1, 1) = ";" Then
                    BodyCmd = Mid(BodyCmd, 2)
                End If
                s = vbCrLf & ",[""body"",false,""body"",""" & BodyCmd & """]" & s
            End If
        End If
        EditorStyleRules = s
        Call main.SetViewingProperty(styleRulesKey, EditorStyleRules)
    End If
    GetInnovaEditorStyleRules = EditorStyleRules
    '
    ' clean up
    '
    If GetInnovaEditorStyleRules = "" Then
        '
        ' If empty, create a single dummy style
        '
        GetInnovaEditorStyleRules = "[["".NoStyleSheet"",""false"",""No Style Sheet"",""color:black;""]]"
    Else
        '
        ' not empty, clean up
        '
        GetInnovaEditorStyleRules = "[" & Mid(GetInnovaEditorStyleRules, 4) & "]"
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set main = Nothing
    Call HandleError("innovaEditorAddonClass", "getInnovaEditorStyleRules", Err.Number, Err.source, Err.Description, True, False)
End Function
'
'
'
Private Function GetInnovaEditorStyleRules_FilterStyles(styleSource As String, TargetString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim MediaEndPtr As Long
    Dim BraceClosedPtr As Long
    Dim ParsePtr As Long
    Dim BraceClosedCnt As Long
    Dim BraceOpenCnt As Long
    Dim BraceOpenPtr As Long
    Dim LoopPtr As Long
    Dim MediaStartPtr As Long
    Dim AbortParse As Boolean
    '
    GetInnovaEditorStyleRules_FilterStyles = styleSource

    
    MediaStartPtr = InStr(1, GetInnovaEditorStyleRules_FilterStyles, TargetString, vbTextCompare)
    If MediaStartPtr <> 0 Then
        MediaStartPtr = InStrRev(GetInnovaEditorStyleRules_FilterStyles, vbCrLf, MediaStartPtr)
    End If
    AbortParse = False
    Do While MediaStartPtr <> 0 And Not AbortParse And LoopPtr < 1000
        BraceOpenPtr = InStr(MediaStartPtr, GetInnovaEditorStyleRules_FilterStyles, "{")
        If BraceOpenPtr = 0 Then
            '
            ' problem - no braces, exit the process
            '
            Exit Do
        Else
            BraceOpenCnt = 1
            BraceClosedCnt = 0
            ParsePtr = BraceOpenPtr + 1
            Do While BraceOpenCnt > BraceClosedCnt And Not AbortParse
                BraceOpenPtr = InStr(ParsePtr, GetInnovaEditorStyleRules_FilterStyles, "{")
                BraceClosedPtr = InStr(ParsePtr, GetInnovaEditorStyleRules_FilterStyles, "}")
                If BraceOpenPtr = 0 And BraceClosedPtr = 0 Then
                    '
                    ' Problem, abort Parse
                    '
                    AbortParse = True
                ElseIf BraceOpenPtr = 0 And BraceClosedPtr <> 0 Then
                    '
                    ' A close was found, and there are no more opens
                    '
                    BraceClosedCnt = BraceClosedCnt + 1
                    ParsePtr = BraceClosedPtr + 1
                ElseIf BraceOpenPtr < BraceClosedPtr Then
                    '
                    ' Another open found
                    '
                    BraceOpenCnt = BraceOpenCnt + 1
                    ParsePtr = BraceOpenPtr + 1
                Else
                    '
                    ' Another Closed found
                    '
                    BraceClosedCnt = BraceClosedCnt + 1
                    ParsePtr = BraceClosedPtr + 1
                End If
            Loop
        End If
        MediaEndPtr = ParsePtr
        GetInnovaEditorStyleRules_FilterStyles = Mid(GetInnovaEditorStyleRules_FilterStyles, 1, MediaStartPtr - 1) & Mid(GetInnovaEditorStyleRules_FilterStyles, MediaEndPtr)
        MediaStartPtr = InStr(MediaStartPtr, GetInnovaEditorStyleRules_FilterStyles, TargetString, vbTextCompare)
        If MediaStartPtr <> 0 Then
            MediaStartPtr = InStrRev(GetInnovaEditorStyleRules_FilterStyles, vbCrLf, MediaStartPtr)
        End If
        LoopPtr = LoopPtr + 1
    Loop
    If LoopPtr >= 1000 Then
        Call HandleErrorAndResumeNext("innovaEditorAddonClass", "getInnovaEditorStyleRules_FilterStyles", "Maximum loop count exceeded removing @media from stylesheets", 0)
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("innovaEditorAddonClass", "getInnovaEditorStyleRules_FilterStyles", Err.Number, Err.source, Err.Description, True, False)
End Function
'
'
'
Friend Function GetInnovaEditorCache(main As Object, AllowActiveContent As Boolean) As String
    On Error GoTo ErrorTrap
    '
    If AllowActiveContent Then
        GetInnovaEditorCache = main.GetViewingProperty("InnovaEditorCacheActiveContent", "")
    Else
        GetInnovaEditorCache = main.GetViewingProperty("InnovaEditorCache", "")
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("InnovaEditorAddonClass", "GetInnovaEditorCache", Err.Number, Err.source, Err.Description, True, False)
End Function
'
'
'
Friend Sub SetInnovaEditorCache(main As Object, Raw As String, AllowActiveContent As Boolean)
    On Error GoTo ErrorTrap
    '
    If AllowActiveContent Then
        Call main.SetViewingProperty("InnovaEditorCacheActiveContent", Raw)
    Else
        Call main.SetViewingProperty("InnovaEditorCache", Raw)
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleError("InnovaEditorAddonClass", "SetInnovaEditorCache", Err.Number, Err.source, Err.Description, True, False)
End Sub
'
'===========================================================================
'
'===========================================================================
'
Private Sub HandleClassTrapError(MethodName As String, Optional Context As String)
    '
    If main Is Nothing Then
        Call HandleError2("unknown", Context, App.EXEName, "PreviewClass", MethodName, Err.Number, Err.source, Err.Description, True, False, "unknown")
    Else
        Call HandleError2(main.ApplicationName, Context, App.EXEName, "PreviewClass", MethodName, Err.Number, Err.source, Err.Description, True, False, main.ServerLink)
    End If
    '
End Sub
'
'===========================================================================
'
'===========================================================================
'
Private Sub HandleClassAppendLogfile(MethodName As String, Context As String)
    If main Is Nothing Then
        Call AppendLogFile2("", Context, App.EXEName, "PreviewClass", MethodName, 0, "", "", False, True, "", "", "trace")
    Else
        Call AppendLogFile2(main.ApplicationName, Context, App.EXEName, "PreviewClass", MethodName, 0, "", "", False, True, main.ServerLink, "", "trace")
    End If
    
End Sub





