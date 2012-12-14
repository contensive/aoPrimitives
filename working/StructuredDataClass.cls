VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StructuredDataClass"
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
'Private Main As ccWeb3.object
'Private csv As Object
'Private Main As Object
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
Public Function execute(CsvObject As Object, MainObject As Object, optionString As String, filterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim dst As Long
    Dim Ptr As Long
    Dim ptrLast As Long
    Dim indent As Long
    Dim d As String
    Dim OutputIsEndTag As Boolean
    Dim ptrSelfClose As Long
    Dim ReturnProcessingInstructions As Boolean
    Dim csv As Object
    Dim main As Object
    Dim src As String
    Dim AddonIDNameOrGuid As String
    Dim addonId As Long
    Dim InnerOptionString As String
    Dim doc As DOMDocument60
    Dim node As IXMLDOMElement
    Dim Pos As Long
    Dim BuildVersion As String
    Dim htmlTools As htmlToolsClass
    '
    Set htmlTools = New htmlToolsClass
    '
    Set csv = CsvObject
    Set main = MainObject
    '
    src = main.getaddonoption("data", optionString)
    BuildVersion = main.SiteProperty_BuildVersion
    If src <> "" Then
        '
        ' test for Contensive processign instruction
        '
        Pos = InStr(1, src, "<?contensive", vbTextCompare)
        If Pos = 0 Then
            '
            ' no instruction, just return result
            '
            execute = src
        Else
            '
            ' Parse the xml structureddata
            '
            Set doc = New DOMDocument60
            Call doc.loadXML(src)
            If doc.hasChildNodes Then
                execute = processNode(main, csv, doc.documentElement, "", 0, "", BuildVersion, htmlTools)
            End If
        End If
        ReturnProcessingInstructions = False
        If ReturnProcessingInstructions Then
            '
            ' auto indent the processing instructions
            '
            Do While (InStr(1, src, vbTab & "<") <> 0)
                src = Replace(src, vbTab & "<", "<")
            Loop
            '
            Do While (InStr(1, src, vbCr & "<") <> 0)
                src = Replace(src, vbCr & "<", "<")
            Loop
            '
            Do While (InStr(1, src, vbLf & "<") <> 0)
                src = Replace(src, vbLf & "<", "<")
            Loop
            '
            d = ""
            Ptr = 1
            ptrLast = 1
            indent = 1
            Do
                Ptr = InStr(ptrLast + 1, src, "<")
                ptrSelfClose = InStr(ptrLast + 1, src, "/>")
                If Ptr = 0 Then
                    '
                    ' end of source
                    '
                    d = d & vbCrLf & String(indent, vbTab) & Mid(src, ptrLast)
                ElseIf (ptrSelfClose <> 0) And (ptrSelfClose < Ptr) Then
                    '
                    ' last tag was self-closing (/>)
                    '
                    d = d & vbCrLf & String(indent, vbTab) & Mid(src, ptrLast, Ptr - ptrLast)
                    indent = indent - 1
                    If OutputIsEndTag Then
                        indent = indent - 1
                    End If
                    OutputIsEndTag = False
                Else
                    '
                    '
                    '
                    If Mid(src, Ptr + 1, 1) = "/" Then
                        '
                        ' tag end
                        '
                        d = d & vbCrLf & String(indent, vbTab) & Mid(src, ptrLast, Ptr - ptrLast)
                        If OutputIsEndTag Then
                            indent = indent - 1
                        End If
                        OutputIsEndTag = True
                    Else
                        '
                        ' tag start
                        '
                        d = d & vbCrLf & String(indent, vbTab) & Mid(src, ptrLast, Ptr - ptrLast)
                        If OutputIsEndTag Then
                            indent = indent - 1
                        End If
                        OutputIsEndTag = False
                        indent = indent + 1
                    End If
                End If
                ptrLast = Ptr
            Loop While Ptr > 0
            '
            execute = execute _
                & CR & "<div>" _
                & CR & "<div><b>Processing instructions</b></div>" _
                & "<pre>" & kmaEncodeHTML(d) & "</pre>" _
                & CR & "</div>"
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call main.ReportError2(Err, "Error in prototypeProcessClass.Execute")
End Function
'
'=================================================================================
' Process a single instruction node
'   returns the modified layout
'   parentItemNodeLayout = a copy of the layout created by the first item child node
'       found in a parent node. It should be passed in blank. When the first
'       item node is found, the layout is copied to this argument. As the
'       parent routine iterates though it's child nodes, it passes back the
'       save parentItemNodeLayout each time
'
'   parentItemNodeCnt = the count of item nodes found in the parent of this node
'       if this is an item node, the parentItemNodeCnt is incremented
'=================================================================================
'
Private Function processNode(main As Object, csv As Object, node As IXMLDOMElement, layout As String, ByRef parentItemNodeCnt As Long, ByRef parentItemNodeLayout As String, BuildVersion As String, htmlTools As htmlToolsClass) As String
    On Error GoTo ErrorTrap
    '
    Dim LayoutID As Long
    Dim StylesFilename As String
    Dim SharedStylesIDList As String
    Dim Pos2 As Long
    Dim Ptr As Long
    Dim StyleID As Long
    Dim ListSplit() As String
    Dim tagClassPos As Long
    Dim tagEndPos As Long
    Dim tagStartPos As Long
    Dim newClass As String
    Dim childItemNodeLayout As String
    Dim childItemNodeCnt As Long
    Dim posStart As Long
    'Dim parentItemNodeLayout As String
    'Dim parentItemNodeCnt As Long
    Dim childNode As IXMLDOMElement
    Dim CS As Long
    Dim s As String
    Dim layoutLoad As String
    Dim layoutName As String
    Dim layoutDestinationKey As String
    Dim layoutSourceKey As String
    Dim sourceLayout As String
    Dim sourceLayoutSave As String
    Dim appendedResults As String
    Dim parentItemNodeLayoutModified As String
    Dim itemForGroupsList As String

    '
    ' execute this node's instruction
    '
    sourceLayout = layout
    Select Case LCase(node.baseName)
        Case "layout"
            sourceLayoutSave = sourceLayout
            layoutLoad = kmaEncodeText(node.getAttribute("load"))
            layoutSourceKey = kmaEncodeText(node.getAttribute("source"))
            layoutDestinationKey = kmaEncodeText(node.getAttribute("destination"))
            If layoutLoad <> "" Then
                '
                ' open layout record
                '
                If htmlTools.isGuid(main, layoutLoad) Then
                    CS = csv.OpenCSContent("Layouts", "ccguid=" & KmaEncodeSQLText(layoutLoad), "ID")
                ElseIf IsNumeric(layoutLoad) Then
                    CS = csv.OpenCSContent("Layouts", "id=" & layoutLoad, "ID")
                Else
                    CS = csv.OpenCSContent("Layouts", "name=" & KmaEncodeSQLText(layoutLoad), "ID")
                End If
                If csv.IsCSOK(CS) Then
                    LayoutID = csv.GetCSInteger(CS, "id")
                    sourceLayout = csv.GetCS(CS, "layout")
                    If BuildVersion < "4.1.159" Then
                        StylesFilename = ""
                        SharedStylesIDList = ""
                    Else
                        StylesFilename = csv.GetCS(CS, "StylesFilename")
                        SharedStylesIDList = csv.GetCS(CS, "SharedStyles")
                    End If
                End If
                Call csv.CloseCS(CS)
                '
                ' Add exclusive styles
                '
                If StylesFilename <> "" Then
                    Call main.AddStylesheetLink2(main.ServerProtocol & main.ServerHost & main.ServerFilePath & StylesFilename, "layout")
                End If
                '
                ' Add shared styles
                '
                If SharedStylesIDList <> "" Then
                    ListSplit = Split(SharedStylesIDList, ",")
                    For Ptr = 0 To UBound(ListSplit)
                        StyleID = kmaEncodeInteger(ListSplit(Ptr))
                        If StyleID <> 0 Then
                            If BuildVersion < "4.1.159" Then
                                Call main.AddSharedStyleID(kmaEncodeInteger(ListSplit(Ptr)))
                            Else
                                Call main.AddSharedStyleID2(kmaEncodeInteger(ListSplit(Ptr)), "layout")
                            End If
                        End If
                    Next
                End If
                Pos2 = InStr(1, sourceLayout, "<body", vbTextCompare)
                If Pos2 > 0 Then
                    Pos2 = InStr(Pos2, sourceLayout, ">")
                    If Pos2 > 0 Then
                        sourceLayout = Mid(sourceLayout, Pos2 + 1)
                        Pos2 = InStr(Pos2, sourceLayout, "</body", vbTextCompare)
                        If Pos2 > 0 Then
                            sourceLayout = Mid(sourceLayout, 1, Pos2 - 1)
                        End If
                    End If
                End If
            End If
            If layoutSourceKey <> "" Then
                '
                ' if the source is specified, grab the outerHTML
                ' note: the source should include the leading white-space
                '
                sourceLayout = htmlTools.getOuterHTML(main, sourceLayout, layoutSourceKey)
                'sourceLayout = getOuterHTML(main, sourceLayout, layoutSourceKey, PosStart)
            ElseIf (layoutDestinationKey <> "") Then
                '
                ' destination key with no source - use innerHTML of the destination
                '
                sourceLayout = htmlTools.getInnerHTML(main, sourceLayout, layoutDestinationKey)
            End If
            '
            ' process all the node's child nodes
            '
            If node.ChildNodes.length > 0 Then
                childItemNodeCnt = 0
                childItemNodeLayout = ""
                For Each childNode In node.ChildNodes
                    sourceLayout = processNode(main, csv, childNode, sourceLayout, childItemNodeCnt, childItemNodeLayout, BuildVersion, htmlTools)
                Next
            End If
            '
            If layoutDestinationKey <> "" Then
                '
                ' if destination is specified, insert the resulting layout into the sourcelayout at the destination specified
                ' note: the destination should include the trailing white-space
                '
                sourceLayout = htmlTools.insertInnerHTML(main, sourceLayoutSave, layoutDestinationKey, sourceLayout)
            Else
                '
                ' no destination specified, return with just the sourcelayout
                '
                sourceLayout = sourceLayout
            End If
        Case "item"
            '
            ' process all the node's child nodes, each with a fresh source layout
            '
            itemForGroupsList = kmaEncodeText(node.getAttribute("forgroups"))
            If parentItemNodeCnt = 0 Then
                '
                ' if this is the first item, create the item source from the currently modified
                ' and return just the item as the sourceLayout. This way common changes to an item can
                ' be made before th item is repeated over and over.
                ' So any modifications to the layout are 'absorbed' into the item layout, and the
                ' modified layout is deleted from the queue.
                ' BUT - changes made within the item should NOT be returned back to the source layout
                ' so class changes like 'hit' will not be padded along to future item lists
                '
                parentItemNodeLayout = sourceLayout
                sourceLayout = ""
            End If
            If node.ChildNodes.length > 0 Then
                '
                ' modify the source layout with it's child nodes
                '
                parentItemNodeLayoutModified = parentItemNodeLayout
                childItemNodeCnt = 0
                childItemNodeLayout = ""
                For Each childNode In node.ChildNodes
                    parentItemNodeLayoutModified = processNode(main, csv, childNode, parentItemNodeLayoutModified, childItemNodeCnt, childItemNodeLayout, BuildVersion, htmlTools)
                Next
            End If
            '
            ' if blocking is set, wrap this item node in a Block Text addon
            '
            If itemForGroupsList <> "" Then
                parentItemNodeLayoutModified = "" _
                    & CR & "<ac type=""aggregatefunction"" name=""Block Text"" querystring=""AllowGroups=" & csv.EncodeAddonOptionArgument(itemForGroupsList) & """ >" _
                    & kmaIndent(parentItemNodeLayoutModified) _
                    & CR & "<ac type=""aggregatefunction"" name=""Block Text End"" >"
            End If
            '
            ' return the node or append the node, depending on if it is the first
            '
            If parentItemNodeCnt = 0 Then
                '
                ' first item just returns the modified source
                '
                sourceLayout = parentItemNodeLayoutModified
            Else
                '
                ' addl items append modified source to the previous layout
                '
                sourceLayout = sourceLayout & parentItemNodeLayoutModified
            End If
            parentItemNodeCnt = parentItemNodeCnt + 1
        Case "text"
            '
            ' set innerHTML in source to value provided in instructions
            '
            layoutDestinationKey = kmaEncodeText(node.getAttribute("destination"))
            sourceLayout = htmlTools.insertInnerHTML(main, sourceLayout, layoutDestinationKey, node.Text)
            If parentItemNodeLayout <> "" Then
                '
                ' repeat the change on the item source layout
                '
                parentItemNodeLayout = htmlTools.insertInnerHTML(main, parentItemNodeLayout, layoutDestinationKey, node.xml)
            End If
        Case "addclass"
            '
            ' add a class to the destination
            '
            layoutDestinationKey = kmaEncodeText(node.getAttribute("destination"))
            newClass = kmaEncodeText(node.Text)
            If layoutDestinationKey <> "" And newClass <> "" Then
                tagStartPos = htmlTools.getTagStartPos(main, sourceLayout, 1, layoutDestinationKey)
                If tagStartPos > 0 Then
                tagEndPos = InStr(tagStartPos, sourceLayout, ">")
                If tagEndPos > 0 Then
                tagClassPos = InStr(tagStartPos, sourceLayout, "class=", vbTextCompare)
                If tagClassPos = 0 Or (tagClassPos > tagEndPos) Then
                    '
                    ' no class, add it to end of tag
                    '
                    sourceLayout = Mid(sourceLayout, 1, tagEndPos - 1) & " class=""" & newClass & """" & Mid(sourceLayout, tagEndPos)
                Else
                    '
                    ' modify current class attribute
                    '
                    sourceLayout = sourceLayout
                End If
                End If
                End If
            End If
        Case "setid"
            '
            ' set the id of the destination
            '
            layoutDestinationKey = kmaEncodeText(node.getAttribute("destination"))
    End Select
    
    '
    ' return the result string
    '
    processNode = sourceLayout
    '
    Exit Function
ErrorTrap:
    Call main.ReportError2(Err, "Error in prototypeProcessClass.processNode")
End Function




