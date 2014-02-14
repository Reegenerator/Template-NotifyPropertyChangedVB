Imports EnvDTE
Imports EnvDTE80
Imports System.Text

Imports System.Xml.Linq
Imports System.Reflection


Partial Class TagManager(Of T As {New, GeneratorAttribute})

    ''' <summary>
    ''' Holds information required to generate code segments
    ''' </summary>
    ''' <remarks></remarks>
    Public Class TaggedSegmentWriter



        Sub New()
            'Do nothing
            GenAttribute = New T
        End Sub

        ''' <summary>
        ''' Create a new writer with the same Class, TriggeringBaseClass and GeneratorAttribute
        ''' </summary>
        ''' <param name="parentWriter">
        ''' source of properties to be copied
        ''' </param>
        ''' <remarks></remarks>
        Sub New(parentWriter As TaggedSegmentWriter)
            [Class] = parentWriter.Class
            TriggeringBaseClass = parentWriter.TriggeringBaseClass
            'Clone instead of reusing parent's attribute, because they may have different property values
            GenAttribute = CType(parentWriter.GenAttribute.MemberwiseClone, T)

        End Sub



        Property TriggeringBaseClass As CodeClass2
        Property [Class] As CodeClass2
        Property GenAttribute As T
        Property SearchStart As TextPoint
        Property SearchEnd As TextPoint
        Property InsertStart As TextPoint
        Property InsertedEnd As TextPoint
        Property SegmentType As SegmentTypes
        Property Content As String
        Property ProcessedContent As String
        Property TagComment As String
        Property OpenFileOnGenerated As Boolean = True
        Public Property HasError() As Boolean

        Private _Status As StringBuilder
        ReadOnly Property Status As StringBuilder
            Get
                If _Status Is Nothing Then _Status = New StringBuilder
                Return _Status
            End Get
        End Property

        ''' <summary>
        ''' True if the code generation was triggered by the base of current class
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ReadOnly Property IsTriggeredByBaseClass As Boolean
            Get
                Return TriggeringBaseClass IsNot Nothing AndAlso TriggeringBaseClass IsNot [Class]
            End Get
        End Property


        Sub OutlineText()

            Dim endWithoutNewline = InsertedEnd.CreateEditPoint
            endWithoutNewline.CharLeft()
            InsertStart.CreateEditPoint.OutlineSection(endWithoutNewline)
        End Sub

        Function GetContentEndPoint() As TextPoint
            Dim endP As EditPoint = InsertStart.CreateEditPoint
            endP.CharRightExact(Content.Length)
            Return endP
        End Function

        Function GetSearchText() As String
            Return SearchStart.CreateEditPoint.GetText(SearchEnd)
        End Function

        Function InsertAndFormat() As TextPoint
            Dim text = Me.GenText()
            InsertedEnd = InsertStart.InsertAndFormat(text)
            Return InsertedEnd
        End Function

#Region "Tag Generation" '!――――――――――――――――――――――――――――――――――――――――――――――――――――――

        Function GenXmlTag() As XElement
            'set to nothing if it's by Attribute(default) so Trigger attribute is not written out
            Dim triggerType As GeneratorAttribute.TriggerTypes?
            If IsTriggeredByBaseClass Then
                triggerType = GeneratorAttribute.TriggerTypes.BaseClassAttribute
            End If

            Dim triggerInfo = If(triggerType = GeneratorAttribute.TriggerTypes.BaseClassAttribute, TriggeringBaseClass.Name, Nothing)

            Dim xml = New XElement(GenAttribute.TagPrototype)
            If triggerType IsNot Nothing Then xml.@Trigger = triggerType.ToString
            If triggerInfo IsNot Nothing Then xml.@TriggerInfo = triggerInfo
            xml.@Date = Now.ToString

            Dim xmlNameType = GetType(XmlPropertyAttribute)
            Dim membersWithXmlName = From m In GenAttribute.TypeCache.GetMembers
                                    Select Member = m, XmlName = m.GetCustomAttributes(Of XmlPropertyAttribute).FirstOrDefault
                                    Where XmlName IsNot Nothing


            For Each p In GenAttribute.GetXmlProperties

                Dim propValue = p.Value.GetValue(GenAttribute)
                If propValue IsNot Nothing Then xml.Add(New XAttribute(p.Key, propValue))

            Next
            Return xml
        End Function

        Public Function CreateTaggedCommentText() As String
            '?Newline is added surrounding the text because we can't figure out how to add newline in TagXmlWriter
            Dim xml = GenXmlTag()
            xml.Add(vbNewLine & Content & vbNewLine)
            Return TagXmlWriter.ToCommentedString(xml)
        End Function

        Public Function CreateTaggedRegionName() As String
            Dim xml = GenXmlTag()
            Dim regionNameXml = TagXmlWriter.ToRegionNameString(xml)
            Return TagComment.Conjoin(vbTab, regionNameXml)
        End Function

        Public Function GenTaggedRegionText() As String
            Dim res = String.Format("#Region ""{0}""{1}{2}{1}{3}{1}", CreateTaggedRegionName(), Environment.NewLine, Content, "#End Region")
            Return res
        End Function

        Function GenText() As String
            Select Case SegmentType
                Case SegmentTypes.Region
                    Return GenTaggedRegionText()
                Case SegmentTypes.Statements
                    Return CreateTaggedCommentText()
                Case Else
                    Throw New Exception("Unknown SegmentType")
            End Select
        End Function



#End Region '!―――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――――

    End Class

End Class