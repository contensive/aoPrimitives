VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PreviewClass"
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
Private csv As Object
'
'=================================================================================
'
'=================================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, optionString As String, filterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CS As Long
    Dim LayoutIDNameOrGuid As String
    Dim AddonIDNameOrGuid As String
    Dim PageIDNameOrGuid As String
    Dim dataOptionstring As String
    Dim src As String
    Dim PageID As Long
    Dim Body As String
    '
    Set Main = MainObject
    Set csv = CsvObject
    '
    LayoutIDNameOrGuid = Main.GetStreamText("layout")
    If LayoutIDNameOrGuid <> "" Then
        src = "<?contensive type=""presentation"" ?>" _
            & "<layout load=""" & LayoutIDNameOrGuid & """ />" _
            & ""
            dataOptionstring = "data=" & csv.EncodeAddonOptionArgument(src)
            Body = Main.ExecuteAddon3(StructuredDataProcessorGuid, dataOptionstring)
            Execute = "" _
                & Main.DocType _
                & CR & "<html>" _
                & CR & "<head>" & kmaIndent(Main.GetHTMLHead()) & CR & "</head>" _
                & CR & TemplateDefaultBodyTag & kmaIndent(Body) & CR & "</body>" _
                & CR & "</html>"
    ElseIf PageIDNameOrGuid <> "" Then
        '
        '
        '
        If IsNumeric(PageIDNameOrGuid) Then
            PageID = kmaEncodeInteger(PageIDNameOrGuid)
        Else
            CS = Main.OpenCSContent("page content", "ccguid=" & KmaEncodeSQLText(PageIDNameOrGuid))
            If Main.IsCSOK(CS) Then
                PageID = Main.GetCSInteger(CS, "id")
            End If
            Call Main.CloseCS(CS)
            If PageID = 0 Then
                CS = Main.OpenCSContent("page content", "name=" & KmaEncodeSQLText(PageIDNameOrGuid))
                If Main.IsCSOK(CS) Then
                    PageID = Main.GetCSInteger(CS, "id")
                End If
                Call Main.CloseCS(CS)
            End If
        End If
        Execute = Main.GetContentPage("", , , , , PageIDNameOrGuid)
        
    Else
        Execute = "<div>no objects to preview</div>"
    End If
    '
    Exit Function
ErrorTrap:
    HandleClassTrapError ("execute")
End Function
'
'===========================================================================
'
'===========================================================================
'
Private Sub HandleClassTrapError(MethodName As String, Optional Context As String)
    '
    If Main Is Nothing Then
        Call HandleError2("unknown", Context, App.EXEName, "PreviewClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, "unknown")
    Else
        Call HandleError2(Main.ApplicationName, Context, App.EXEName, "PreviewClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, Main.ServerLink)
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
        Call AppendLogFile2("", Context, App.EXEName, "PreviewClass", MethodName, 0, "", "", False, True, "", "", "trace")
    Else
        Call AppendLogFile2(Main.ApplicationName, Context, App.EXEName, "PreviewClass", MethodName, 0, "", "", False, True, Main.ServerLink, "", "trace")
    End If
    
End Sub



