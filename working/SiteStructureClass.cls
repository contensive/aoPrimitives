VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiteStructureClass"
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
Private Main As Object
'Private Main As ccWeb3.MainClass
Private Csv As Object
'
'=================================================================================
'
'=================================================================================
'
Public Function execute(CsvObject As Object, MainObject As Object, optionString As String, filterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Set Main = MainObject
    Set Csv = CsvObject
    '
    ' For Page Addons, return the result
    '
    execute = GetSiteStructure(False)
    '
    Exit Function
ErrorTrap:
    HandleClassTrapError ("execute")
End Function
'
'
'
Private Function GetSiteStructure(IsWorkflowRendering As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim PCC As Variant
    Dim BuildVersion As String
    '
    PCC = Main.GetPCC(False, False)
    BuildVersion = Main.SiteProperty_BuildVersion
    '
    GetSiteStructure = "" _
        & vbCrLf & vbTab & "<sitestructure version=""" & App.Major & "." & App.Minor & "." & Format(App.Revision, "000") & """>" _
        & kmaIndent(GetMenusNode(IsWorkflowRendering, PCC, BuildVersion)) _
        & kmaIndent(GetSectionsNode(IsWorkflowRendering, PCC, BuildVersion)) _
        & kmaIndent(GetPagesNode(IsWorkflowRendering, PCC, BuildVersion)) _
        & vbCrLf & vbTab & "</sitestructure>"
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetSiteStructure")
End Function
'
'
'
Private Function GetMenusNode(IsWorkflowRendering As Boolean, PCC As Variant, BuildVersion As String) As String
    On Error GoTo ErrorTrap
    '
    GetMenusNode = "" _
        & vbCrLf & vbTab & "<menus>" _
        & kmaIndent(GetMenuNodes(IsWorkflowRendering, PCC, BuildVersion)) _
        & vbCrLf & vbTab & "</menus>"
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetMenusNode")
End Function
'
'=============================================================================
'   Get the Menu Node
'=============================================================================
'
Friend Function GetMenuNodes(IsWorkflowRendering As Boolean, PCC As Variant, BuildVersion As String) As String
    On Error GoTo ErrorTrap
    '
    Dim SubNodes As String
    Dim DefaultTemplateID As Long
    Dim IsEditingMenus As Boolean
    Dim PageActive As Boolean
    Dim TCPtr As Long
    Dim PCCPtr As Long
    Dim RootPageID As Long
    Dim CSMenus As Long
    Dim CSTemplates As Long
    Dim CSPage As Long
    Dim MenuName As String
    Dim TemplateID As Long
    Dim ContentID As Long
    Dim ContentName As String
    Dim PageList_ParentBranchPointer As Long
    Dim Link As String
    Dim MenuID As Long
    Dim AuthoringTag As String
    Dim MenuImage As String
    Dim MenuImageOver As String
    Dim LandingLink As String
    Dim MenuString As String
    Dim MenuCaption As String
    Dim MenuTemplateID As Long
    Dim Criteria As String
    Dim SelectFieldList As String
    Dim ShowHiddenMenu As Boolean
    Dim HideMenu As Boolean
    Dim PageContentCID As Long
    Dim BlockPage As Boolean
    Dim BlockMenu As Boolean
    Dim SQL As String
    Dim IsAllMenusMenuMode As Boolean
    Dim CS As Long
    '
    '
    SelectFieldList = "ID, Name,Depth,Layout,Delimiter,FlyoutOnHover,FlyoutDirection,StylePrefix,StylesFilename"
    CSMenus = Main.OpenCSContent("Dynamic Menus", , , , , , SelectFieldList)
    Do While Main.IsCSOK(CSMenus)
        MenuID = Csv.GetCSInteger(CSMenus, "ID")
        If True Then
            MenuName = Trim(Csv.GetCSText(CSMenus, "Name"))
            If MenuName = "" Then
                MenuName = "Menu " & MenuID
                Call Main.ExecuteSQL("default", "update ccMenus set Name=" & KmaEncodeSQLText(MenuName) & " where ID=" & MenuID)
            End If
            '
            ' Get MenuSection Nodes
            '
            SubNodes = ""
            CS = Main.OpenCSContent("Dynamic Menu Section Rules", "(DynamicMenuID=" & MenuID & ")and(sectionid is not null)", , , , , "SectionID")
            Do While Main.IsCSOK(CS)
                '
                SubNodes = SubNodes & vbCrLf & vbTab & "<menusection sectionid=""" & Main.GetCSInteger(CS, "sectionid") & """/>"
                Main.NextCSRecord (CS)
            Loop
            '
            ' Get Menu, remove crlf, and parse the line with crlf
            '
            'SubNodes = GetPageNode(RootPageID, Link, 99, 0, "", "", "", MenuCaption, MenuID, False, BuildVersion, PCC, IsWorkflowRendering)
            GetMenuNodes = GetMenuNodes _
                & vbCrLf & vbTab & "<menu" _
                & " id=""m" & MenuID & """" _
                & " menuid=""" & MenuID & """" _
                & " name=""" & MenuName & """" _
                & " depth=""" & Csv.GetCSInteger(CSMenus, "depth") & """" _
                & " layout=""" & Csv.GetCSInteger(CSMenus, "Layout") & """" _
                & " delimiter=""" & kmaEncodeHTML(Csv.GetCSText(CSMenus, "Delimiter")) & """" _
                & " flyoutonhover=""" & Csv.GetCSBoolean(CSMenus, "FlyoutOnHover") & """" _
                & " flyoutdirection=""" & Csv.GetCSInteger(CSMenus, "FlyoutDirection") & """" _
                & " styleprefix=""" & kmaEncodeHTML(Csv.GetCSText(CSMenus, "StylePrefix")) & """" _
                & " stylesfilename=""" & kmaEncodeHTML(Csv.GetCSText(CSMenus, "StylesFilename")) & """" _
                & ""
            If SubNodes <> "" Then
                GetMenuNodes = GetMenuNodes & ">" _
                    & kmaIndent(SubNodes) _
                    & vbCrLf & vbTab & "</menu>"
            Else
                GetMenuNodes = GetMenuNodes & "/>"
            End If
        End If
        Call Main.NextCSRecord(CSMenus)
    Loop
    Call Main.CloseCS(CSMenus)
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetMenuNodes", "Trap")
End Function
'
'
'
Private Function GetSectionsNode(IsWorkflowRendering As Boolean, PCC As Variant, BuildVersion As String) As String
    On Error GoTo ErrorTrap
    '
    GetSectionsNode = "" _
        & vbCrLf & vbTab & "<sections>" _
        & kmaIndent(GetSectionNodes(IsWorkflowRendering, PCC, BuildVersion)) _
        & vbCrLf & vbTab & "</sections>"
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetSectionsNode")
End Function
'
'=============================================================================
'   Get the Section Menu
'   MenuName blank reverse menu to legacy mode (all sections on menu)
'=============================================================================
'
Friend Function GetSectionNodes(IsWorkflowRendering As Boolean, PCC As Variant, BuildVersion As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Ptr As Long
    Dim Rows2 As Variant
    Dim Rows() As String
    Dim GroupIDAccessList As String
    Dim CS As Long
    Dim SubNodes As String
    Dim DefaultTemplateID As Long
    Dim IsEditingSections As Boolean
    Dim PageActive As Boolean
    Dim TCPtr As Long
    Dim PCCPtr As Long
    Dim RootPageID As Long
    Dim CSSections As Long
    Dim CSTemplates As Long
    Dim CSPage As Long
    Dim SectionName As String
    Dim TemplateID As Long
    Dim ContentID As Long
    'Dim ContentName As String
    Dim PageList_ParentBranchPointer As Long
    Dim Link As String
    Dim SectionID As Long
    Dim AuthoringTag As String
    Dim MenuImage As String
    Dim MenuImageOver As String
    Dim LandingLink As String
    Dim MenuString As String
    Dim SectionCaption As String
    'Dim SectionTemplateID As Long
    Dim Criteria As String
    Dim SelectFieldList As String
    Dim ShowHiddenMenu As Boolean
    Dim HideMenu As Boolean
    Dim PageContentCID As Long
    Dim BlockPage As Boolean
    Dim BlockSection As Boolean
    Dim SQL As String
    Dim IsAllSectionsMenuMode As Boolean
    '
    '
    PageContentCID = Main.GetContentID("Page Content")
    SelectFieldList = "ID, Name,TemplateID,ContentID,MenuImageFilename,Caption,MenuImageOverFilename,HideMenu,BlockSection,RootPageID"
    ShowHiddenMenu = Main.IsEditingAnything()
    CSSections = Main.OpenCSContent("Site Sections", , , , , , SelectFieldList)
    Do While Main.IsCSOK(CSSections)
        HideMenu = Csv.GetCSBoolean(CSSections, "HideMenu")
        BlockSection = Csv.GetCSBoolean(CSSections, "BlockSection")
        SectionID = Csv.GetCSInteger(CSSections, "ID")
        SectionName = Trim(Csv.GetCSText(CSSections, "Name"))
        If SectionName = "" Then
            SectionName = "Section " & SectionID
        End If
        SectionCaption = Csv.GetCSText(CSSections, "Caption")
        If SectionCaption = "" Then
            SectionCaption = SectionName
        End If
        GroupIDAccessList = ""
        If BlockSection Then
            CS = Main.OpenCSSQL("default", "select groupid from ccSectionBlockRules where sectionid=" & SectionID)
            If Main.IsCSOK(CS) Then
                Rows2 = Main.GetCSRows(CS)
                For Ptr = 0 To UBound(Rows2, 2)
                    GroupIDAccessList = GroupIDAccessList & "," & Rows2(0, Ptr)
                Next
                GroupIDAccessList = Mid(GroupIDAccessList, 2)
            End If
            Call Main.CloseCS(CS)
        End If
        'SectionTemplateID = Csv.GetCSInteger(CSSections, "TemplateID")
        'If SectionTemplateID = 0 Then
        '    If DefaultTemplateID = 0 Then
        '        DefaultTemplateID = Main.GetRecordID("Page Templates", "default")
        '        If DefaultTemplateID = 0 Then
        '            DefaultTemplateID = Main.InsertContentRecordGetID("Page Templates")
        '            Call Main.ExecuteSQL("default", "update cctemplates set name='default' where id=" & DefaultTemplateID)
        '        End If
        '    End If
        '    SectionTemplateID = DefaultTemplateID
        'End If
        'ContentID = Csv.GetCSInteger(CSSections, "ContentID")
        'If (ContentID <> PageContentCID) And (Not Main.IsWithinContent(ContentID, PageContentCID)) Then
        '    ContentID = PageContentCID
        '    Call Csv.SetCS(CSSections, "ContentID", ContentID)
        'End If
        'If ContentID = PageContentCID Then
        '    ContentName = "Page Content"
        'Else
        '    ContentName = Main.GetContentNameByID(ContentID)
        '    If ContentName = "" Then
        '        ContentName = "Page Content"
        '        ContentID = Main.GetContentID(ContentName)
        '        Call Main.ExecuteSQL("default", "update ccSections set ContentID=" & ContentID & " where ID=" & SectionID)
        '    End If
        'End If
        MenuImage = Csv.GetCSText(CSSections, "MenuImageFilename")
        If MenuImage <> "" Then
            MenuImage = Main.ServerFilePath & MenuImage
        End If
        MenuImageOver = Csv.GetCSText(CSSections, "MenuImageOverFilename")
        If MenuImageOver <> "" Then
            MenuImageOver = Main.ServerFilePath & MenuImageOver
        End If
        '
        ' Get Root Page for templateID
        '
        TemplateID = 0
        BlockPage = False
        Link = ""
        RootPageID = Csv.GetCSInteger(CSSections, "rootpageid")
        PCCPtr = Main.GetPCCPtr(RootPageID, IsWorkflowRendering, False)
        If PCCPtr < 0 Then
            RootPageID = 0
        End If
        '
        ' Get MenuSection Nodes
        '
        SubNodes = ""
        CS = Main.OpenCSContent("Dynamic Menu Section Rules", "(SectionID=" & SectionID & ")and(DynamicMenuID is not null)", , , , , "DynamicMenuID")
        Do While Main.IsCSOK(CS)
            '
            SubNodes = SubNodes & vbCrLf & vbTab & "<menusection menuid=""" & Main.GetCSInteger(CS, "DynamicMenuID") & """/>"
            Main.NextCSRecord (CS)
        Loop
        '
        ' Get Menu, remove crlf, and parse the line with crlf
        '
        'SubNodes = GetPageNode(RootPageID, Link, 99, 0, "", "", "", SectionCaption, SectionID, False, BuildVersion, PCC, IsWorkflowRendering)
        GetSectionNodes = GetSectionNodes _
            & vbCrLf & vbTab & "<section" _
            & " id=""s" & SectionID & """" _
            & " sectionid=""" & SectionID & """" _
            & " name=""" & SectionName & """" _
            & " caption=""" & SectionCaption & """" _
            & " hide=""" & HideMenu & """" _
            & " block=""" & BlockSection & """" _
            & " groupaccesslist=""" & GroupIDAccessList & """" _
            & " menuimage=""" & kmaEncodeHTML(MenuImage) & """" _
            & " menuimageover=""" & kmaEncodeHTML(MenuImageOver) & """" _
            & " pageid=""" & RootPageID & """"
        If SubNodes <> "" Then
            GetSectionNodes = GetSectionNodes & ">" _
                & kmaIndent(SubNodes) _
                & vbCrLf & vbTab & "</section>"
        Else
            GetSectionNodes = GetSectionNodes & "/>"
        End If
        Call Main.NextCSRecord(CSSections)
    Loop
    Call Main.CloseCS(CSSections)
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetSectionNodes", "Trap")
End Function
'
'======================================================================================
'   Get all the root page nodes (with their children)
'======================================================================================
'
Private Function GetPagesNode(IsWorkflowRendering As Boolean, PCC As Variant, BuildVersion As String) As String
    On Error GoTo ErrorTrap
    '
    '
    Dim Overview As String
    Dim Active As Boolean
    Dim PseudoChildChildPagesFound As Boolean
    Dim PCCRowPtr As Long
    Dim SortForward As Boolean
    Dim SortFieldName As String
    Dim SortPtr As Long
    Dim Ptr As Long
    Dim ChildPageCount As Long
    Dim ChildPagesFoundTest As String
    Dim FieldList As String
    Dim ChildCountWithNoPubs As Long
    Dim MenuID As Long
    Dim MenuCaption As String
    Dim ChildCount As Long
    Dim ChildSize As Long
    Dim ChildPointer As Long
    Dim ChildID() As Long
    Dim ChildAllowChild() As Boolean
    Dim ChildCaption() As String
    Dim ChildLink() As String
    Dim ChildOverview() As String
    Dim ChildSortMethodID() As Long
    Dim ChildChildPagesFound() As Boolean
    Dim ContentID As Long
    Dim MenuLinkOverRide As String
    Dim PageID As Long
    Dim UsedPageIDStringLocal As String
    Dim Criteria As String
    Dim MenuDepthLocal As Long
    Dim OrderByCriteria As String
    Dim WorkingLink As String
    Dim TemplateID As Long
    Dim ContentControlID As Long
    Dim Link As String
    Dim PubDate As Date
    Dim PCCPtr As Long
    Dim DateExpires As Date
    Dim DateArchive As Date
    Dim IsIncludedInMenu As Boolean
    Dim PCCPtrs() As Long
    Dim PtrCnt As Long
    Dim SortSplit() As String
    Dim SortSplitCnt As Long
    Dim Index As FastIndexClass
    Dim PCCColPtr As Long
    Dim PCCPtrsSorted As Variant
    Dim AllowInMenus As Boolean
    '
    '   Determine default orderby for pages
    '
    OrderByCriteria = Main.GetContentProperty("page content", "defaultsortmethod")
    If OrderByCriteria = "" Then
        OrderByCriteria = "ID"
    End If
    '
    '   Get list of root pages and sort them
    '
    PCCPtr = Main.GetPCCFirstChildPtr(0, IsWorkflowRendering, False)
    PtrCnt = 0
    Do While PCCPtr >= 0
        ReDim Preserve PCCPtrs(PtrCnt)
        PCCPtrs(PtrCnt) = PCCPtr
        PtrCnt = PtrCnt + 1
        PCCPtr = Main.PCCParentIDIndex.GetNextPointerMatch("0")
    Loop
    If PtrCnt > 0 Then
        PCCPtrsSorted = Main.GetPCCPtrsSorted(PCCPtrs, OrderByCriteria)
    End If
    '
    '   Get Nodes from list of root pages
    '
    Ptr = 0
    Do While Ptr < PtrCnt
        PCCPtr = PCCPtrsSorted(Ptr)
        GetPagesNode = GetPagesNode & GetPageNode(PCCPtr, IsWorkflowRendering, PCC, BuildVersion, OrderByCriteria)
        Ptr = Ptr + 1
    Loop
    '
    ' wrap it in the pages node
    '
    If GetPagesNode = "" Then
        GetPagesNode = "<pages/>"
    Else
        GetPagesNode = "" _
            & vbCrLf & vbTab & "<pages>" _
            & kmaIndent(GetPagesNode) _
            & vbCrLf & vbTab & "</pages>"
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetPagesNode")
End Function
'
'======================================================================================
'   Get a PageNode for the PCCPtr given with its child nodes
'======================================================================================
'
Private Function GetPageNode(PCCPtr As Long, IsWorkflowRendering As Boolean, PCC As Variant, BuildVersion As String, DefaultChildListOrderByCriteria As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Ptr As Long
    Dim Rows2 As Variant
    Dim CS As Long
    Dim GroupIDAccessList As String
    Dim PageContentBlock As Boolean
    Dim PagePageBlock As Boolean
Dim PCCChildPtr As Long
    Dim PageMenuImage As String
    Dim PageMenuImageOver As String
    Dim PageMenuImageDown As String
    Dim PageMenuImageDownOver As String
    Dim PageMenuNewWindow As Boolean
    Dim PageOverview As String
    Dim ChildNodes As String
    Dim PCCPtrsSorted As Variant
    Dim ChildNodePtr As Long
    Dim ChildListOrderByCriteria As String
    Dim PCCPtrs() As Long
    Dim ChildNodeCnt As Long
    Dim PageName As String
    Dim Overview As String
    Dim CSSection As Long
    Dim PseudoChildPagesFound As Boolean
    Dim IsValidPage As Boolean
    Dim DateExpires As Date
    Dim DateArchive As Date
    Dim PubDate As Date
    Dim ChildPageCount As Long
    Dim ContentName As String
    Dim AddRootButton As Boolean
    Dim TopMenuCaption As String
    Dim PageMenuCaption As String
    '
    Dim PageID As Long
    Dim BakeName As String
    Dim Criteria As String
    Dim NodeIDPrefix As String
    Dim PageChildListSortMethodID As Long
    Dim PageLink As String
    Dim PageLinkNoRedirect As String
    Dim PageParentID As Long
    Dim PageTemplateID As Long
    Dim PageCCID As Long
    Dim PageAllowChildListDisplay As Boolean
    Dim PageMenuLinkOverRide As String
    Dim ChildPagesFound As Boolean
    Dim FieldList As String
    Dim ChildPagesFoundTest As String
    '
    ContentName = "Page Content"
    '
    ' Skip if expired, archive or non-published
    '
    DateExpires = KmaEncodeDate(PCC(PCC_DateExpires, PCCPtr))
    DateArchive = KmaEncodeDate(PCC(PCC_DateArchive, PCCPtr))
    PubDate = KmaEncodeDate(PCC(PCC_PubDate, PCCPtr))
    IsValidPage = ((DateExpires = CDate(0)) Or (DateExpires > Now)) And ((PubDate = CDate(0)) Or (PubDate < Now))
    If IsValidPage Then
        '
        ' Get page values
        '
        PageID = kmaEncodeInteger(PCC(PCC_ID, PCCPtr))
        PageName = kmaEncodeText(PCC(PCC_Name, PCCPtr))
        PageChildListSortMethodID = kmaEncodeInteger(PCC(PCC_ChildListSortMethodID, PCCPtr))
        PageMenuCaption = kmaEncodeText(PCC(PCC_MenuHeadline, PCCPtr))
        If PageMenuCaption = "" Then
            PageMenuCaption = PageName
            If PageMenuCaption = "" Then
                PageMenuCaption = "Page " & CStr(PageID)
            End If
        End If
        PageCCID = kmaEncodeInteger(PCC(PCC_ContentControlID, PCCPtr))
        PageTemplateID = kmaEncodeInteger(PCC(PCC_TemplateID, PCCPtr))
        PageAllowChildListDisplay = kmaEncodeBoolean(PCC(PCC_AllowChildListDisplay, PCCPtr))
        PageMenuLinkOverRide = kmaEncodeText(PCC(PCC_Link, PCCPtr))
        PageParentID = kmaEncodeInteger(PCC(PCC_ParentID, PCCPtr))
        PageLink = Main.GetPageLink(PageID)
        PageContentBlock = kmaEncodeBoolean(PCC(PCC_BlockContent, PCCPtr))
        PagePageBlock = kmaEncodeBoolean(PCC(PCC_BlockPage, PCCPtr))
        
        GroupIDAccessList = ""
        If PageContentBlock Or PagePageBlock Then
            CS = Main.OpenCSSQL("default", "select groupid from ccPageContentBlockRules where recordid=" & PageID)
            If Main.IsCSOK(CS) Then
                Rows2 = Main.GetCSRows(CS)
                For Ptr = 0 To UBound(Rows2, 2)
                    GroupIDAccessList = GroupIDAccessList & "," & Rows2(0, Ptr)
                Next
                GroupIDAccessList = Mid(GroupIDAccessList, 2)
            End If
            Call Main.CloseCS(CS)
        End If
        
        '
        '   Child Nodes
        '
        ChildListOrderByCriteria = ""
        If PageChildListSortMethodID > 0 Then
            ChildListOrderByCriteria = Main.GetSortMethodByID(PageChildListSortMethodID)
        End If
        If ChildListOrderByCriteria = "" Then
            ChildListOrderByCriteria = DefaultChildListOrderByCriteria
        End If
        If ChildListOrderByCriteria = "" Then
            ChildListOrderByCriteria = "ID"
        End If
        PCCChildPtr = Main.GetPCCFirstChildPtr(PageID, IsWorkflowRendering, False)
        ChildNodeCnt = 0
        Do While PCCChildPtr >= 0
            ReDim Preserve PCCPtrs(ChildNodeCnt)
            PCCPtrs(ChildNodeCnt) = PCCChildPtr
            ChildNodeCnt = ChildNodeCnt + 1
            PCCChildPtr = Main.PCCParentIDIndex.GetNextPointerMatch(CStr(PageID))
        Loop
        If ChildNodeCnt > 0 Then
            PCCPtrsSorted = Main.GetPCCPtrsSorted(PCCPtrs, ChildListOrderByCriteria)
            '
            ChildNodePtr = 0
            Do While ChildNodePtr < ChildNodeCnt
                PCCPtr = PCCPtrsSorted(ChildNodePtr)
                ChildNodes = ChildNodes & GetPageNode(PCCPtr, IsWorkflowRendering, PCC, BuildVersion, ChildListOrderByCriteria)
                ChildNodePtr = ChildNodePtr + 1
            Loop
        End If
        '
        ' Create Page Node
        '
        GetPageNode = "" _
            & vbCrLf & vbTab & "<page" _
            & " id=""p" & CStr(PageID) & """" _
            & " PageID=""" & CStr(PageID) & """" _
            & " Caption=""" & PageMenuCaption & """" _
            & " link=""" & kmaEncodeHTML(PageLink) & """" _
            & " name=""" & kmaEncodeHTML(PageName) & """" _
            & " newwindow=""" & kmaEncodeBoolean(PageMenuNewWindow) & """" _
            & " overview=""" & kmaEncodeHTML(PageOverview) & """" _
            & " contentblock=""" & PageContentBlock & """" _
            & " groupaccesslist=""" & kmaEncodeHTML(GroupIDAccessList) & """" _
            & ""
            '& " ImageSrc=""" & kmaEncodeHTML(PageMenuImage) & """" _
            '& " ImageOverSrc=""" & kmaEncodeHTML(PageMenuImageOver) & """" _
            '& " ImageDownSrc=""" & kmaEncodeHTML(PageMenuImageDown) & """" _
            '& " ImageDownOverSrc=""" & kmaEncodeHTML(PageMenuImageDownOver) & """" _
            '& " pageblock=""" & PagePageBlock & """" _

        If ChildNodes = "" Then
            GetPageNode = GetPageNode & "/>"
        Else
            GetPageNode = GetPageNode & ">" _
                & kmaIndent(ChildNodes) _
                & vbCrLf & vbTab & "</page>"
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetPageNode")
End Function
'
'===========================================================================
'
'===========================================================================
'
Private Sub HandleClassTrapError(MethodName As String, Optional Context As String)
    '
    If Main Is Nothing Then
        Call HandleError2("unknown", Context, "aoDynamicMenu", "MenuClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, "unknown")
    Else
        Call HandleError2(Main.ApplicationName, Context, "aoDynamicMenu", "MenuClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, Main.ServerLink)
    End If
    '
End Sub
'
'===========================================================================
'
'===========================================================================
'
Private Sub HandleClassAppendLogfile(MethodName As String, Context As String)
    If Main Is Nothing Then
        Call AppendLogFile2("", Context, "aoDynamicMenu", "MenuClass", MethodName, 0, "", "", False, True, "", "", "trace")
    Else
        Call AppendLogFile2(Main.ApplicationName, Context, "aoDynamicMenu", "MenuClass", MethodName, 0, "", "", False, True, Main.ServerLink, "", "trace")
    End If
    
End Sub


