
Imports System.Exception
Imports System.Xml

Namespace Contensive.Addons

    Public Class resizeImage
        Inherits BaseClasses.AddonBaseClass

        Public Overrides Function Execute(ByVal CP As Contensive.BaseClasses.CPBaseClass) As Object

            CP.site.testPoint("resize, 100")
            '
            Dim imgSrc As String = ""
            Dim imgWidth As Integer = 0
            Dim imgHeight As Integer = 0
            Dim imgDst As String = ""
            Dim fileName As String = ""
            Dim filePath As String = ""
            Dim fileExtension As String = ""
            Dim objSiz As New SfImageResize.ImageResize
            Dim useWidth As String = ""
            Dim useHeight As String = ""

            Dim retDoc As New XmlDocument
            Dim decStatement As XmlDeclaration = retDoc.CreateXmlDeclaration("1.0", Nothing, Nothing)
            Dim docRoot As XmlElement = retDoc.CreateElement("namevalues")
            Dim nodeImage As XmlElement = retDoc.CreateElement("newImage")
            Dim nodeWidth As XmlElement = retDoc.CreateElement("imageWidth")
            Dim nodeHeight As XmlElement = retDoc.CreateElement("imageHeight")
            'Dim nodeWidth As XmlElement = retDoc.CreateElement("imageHeight")
            'Dim nodeHeight As XmlElement = retDoc.CreateElement("imageWidth")
            '
            CP.site.testPoint("resize, 200")
            '
            Try
                imgSrc = CP.Doc.Var("Image Source")
                imgWidth = CP.Utils.EncodeInteger(CP.Doc.Var("Image Width"))
                imgHeight = CP.Utils.EncodeInteger(CP.Doc.Var("Image Height"))

                fileName = Right(imgSrc, imgSrc.Length - (InStrRev(imgSrc, "/", , vbTextCompare)))

                CP.site.testPoint("resize, 300")
                If fileName <> "" Then
                    objSiz.LoadFromFile(imgSrc)
                    filePath = imgSrc.Replace(fileName, "")
                    fileExtension = Right(fileName, fileName.Length - (InStrRev(fileName, ".", , vbTextCompare) - 1))

                    CP.site.testPoint("resize, 400")
                    If imgWidth = 0 Or imgHeight = 0 Then
                        objSiz.Proportional = True
                        If imgWidth = 0 Then
                            objSiz.Height = imgHeight
                        Else
                            objSiz.Width = imgWidth
                        End If
                    Else
                        objSiz.Proportional = False

                        objSiz.Width = imgWidth
                        objSiz.Height = imgHeight
                    End If

                    CP.site.testPoint("resize, 500")
                    imgWidth = objSiz.Width
                    imgHeight = objSiz.Height

                    imgDst = filePath & fileName.Replace(fileExtension, "-" & imgWidth & "x" & imgHeight & fileExtension)

                    CP.site.testPoint("resize, 600")
                    Call objSiz.DoResize()
                    Call objSiz.SaveToFile(imgDst)

                    retDoc.AppendChild(decStatement)
                    CP.site.testPoint("resize, 700")
                    retDoc.AppendChild(docRoot)

                    nodeImage.InnerText = imgDst
                    nodeHeight.InnerText = imgHeight
                    nodeWidth.InnerText = imgWidth

                    docRoot.AppendChild(nodeImage)
                    docRoot.AppendChild(nodeHeight)
                    docRoot.AppendChild(nodeWidth)

                    CP.site.testPoint("resize, 800")
                End If

            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try

            Return retDoc.InnerXml

        End Function

    End Class

End Namespace
