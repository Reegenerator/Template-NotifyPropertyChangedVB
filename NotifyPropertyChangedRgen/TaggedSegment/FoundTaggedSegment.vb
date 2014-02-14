
Imports System.Xml.Linq
Imports System.Text.RegularExpressions
Imports EnvDTE

Partial Class TagManager(Of T As {New, GeneratorAttribute})


    ''' <summary>
    ''' Stores information parsed by TagManager
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FoundTaggedSegment
        Inherits TextRange
        ''' <summary>
        ''' Attribute generated from the found xml tag
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Property FoundTag As T
        ''' <summary>
        ''' Actual attribute declared on containing property or class
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Property DeclaredAttribute As T
        Property Manager As TagManager(Of T)
        Private Regex As Regex
        Sub New(mgr As TagManager(Of T), sourceAttr As T, start As EnvDTE.TextPoint, endP As EnvDTE.TextPoint)
            Manager = mgr
            StartPoint = start
            EndPoint = endP
            DetectSegmentType()
            Parse(sourceAttr)
        End Sub

        Property GenerateDate As Date
        Property SegmentType As SegmentTypes




        Public Function IsOutdated() As Boolean
            Select Case FoundTag.RegenMode
                Case GeneratorAttribute.RegenModes.Always
                    Return True
                Case Else
                    Dim diffProperties = Not FoundTag.AreArgumentsEquals(DeclaredAttribute)
                    Return diffProperties

            End Select

        End Function

        Private Sub DetectSegmentType()
            Dim firstline = StartPoint.CreateEditPoint.GetLineText()
            If firstline.Trim.StartsWith("#Region") Then
                Me.SegmentType = SegmentTypes.Region
                Regex = Manager.RegionRegex

            Else
                Me.SegmentType = SegmentTypes.Statements
                Regex = Manager.CommentRegex
            End If
        End Sub

#Region "Find Segment"

        ''' <summary>
        ''' Extract valid xml inside Region Name and within inline comment
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Function ExtractXmlContent() As String
            Dim text = Me.GetText
            Dim xmlContent = ""
            Select Case Me.SegmentType
                Case SegmentTypes.Region
                    xmlContent = Manager.RegionRegex.Replace(text, "${xml}")
                Case SegmentTypes.Statements
                    xmlContent = Manager.CommentRegex.Replace(text, "${tag}${content}${tagend}")
            End Select

            Return xmlContent
        End Function

        Sub Parse(parentAttr As T)

            DeclaredAttribute = parentAttr

            If Not IsValid Then Return
            Dim xml = ExtractXmlContent()
            Try
                Dim xdoc = XDocument.Parse(xml)
                Dim xr = xdoc.Root


                GenerateDate = If(xr.@Date IsNot Nothing, CDate(xr.@Date), Nothing)
                FoundTag = New T()
                FoundTag.CopyPropertyFromTag(xr)
            Catch ex As Exception
                DebugHere()
            End Try



        End Sub





#End Region

    End Class
End Class