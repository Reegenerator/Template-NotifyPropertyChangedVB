Imports EnvDTE
Imports System.Text.RegularExpressions
Imports EnvDTE80
Imports System.Xml.Linq
Imports System.IO

Imports System.Xml

''' <summary>
''' Parse and generate code wrapped with xml information, so it can be easily found and replaced 
''' </summary>
''' <remarks></remarks>
Public Class TagManager(Of T As {New, GeneratorAttribute})


    Public Sub New()
        _Attribute = New T
        Init()

    End Sub

#Region "Regex"
    Private _CommentRegex As Regex
    Public ReadOnly Property CommentRegex As Text.RegularExpressions.Regex
        Get
            Return _CommentRegex
        End Get
    End Property

    Private _RegionRegex As Regex

    Public ReadOnly Property RegionRegex As Text.RegularExpressions.Regex
        Get
            Return _RegionRegex
        End Get
    End Property

    Private Sub Init()
        'initialize regex
        Dim regionPatternFormat As String = <String><![CDATA[                           
        (\#Region\s*"(?<textinfo>[^<\r\n]*?)(?<xml><Gen\s*Renderer='NotifyPropertyChanged'.*?/>)")\s*
            (?<content> 
                (?>
		        (?! \#Region | \#End\sRegion ) .
	        |
		        \#Region (?<Depth>)
	        |
		        \#End\sRegion (?<-Depth>)
	        )*
	        (?(Depth)(?!))
        
            )
        \#End\sRegion
        ]]></String>.Value

        Dim commentPatternFormat As String = <String><![CDATA[                           
            (
            '(?<tag><Gen\s*Renderer='NotifyPropertyChanged'\s*[^<>]*/>)
            )
            |           
            (
                '(?<tag><Gen\s*Renderer='NotifyPropertyChanged'\s*
                    [^<>]*#Match everything but tag symbols
                    (?<!/)>)\s*#Match only > but not />
                (?<content>.*?)(?<!</Gen>)
                '(?<tagend></Gen>)\s*
            )
        ]]></String>.Value


        Dim rendererAttr = _Attribute.TagPrototype.Attribute("Renderer")
        Dim pattern = String.Format(commentPatternFormat, _Attribute.TagPrototype.Name.LocalName, rendererAttr.Value, rendererAttr.Name)
        _CommentRegex = New Text.RegularExpressions.Regex(pattern, DefaultRegexOption)
        Dim regPattern As String = String.Format(regionPatternFormat, _Attribute.TagPrototype.Name.LocalName, rendererAttr.Value, rendererAttr.Name)
        _RegionRegex = New Text.RegularExpressions.Regex(regPattern, DefaultRegexOption)

    End Sub

    Private Function GetRegexByType(segmentType As SegmentTypes) As Regex
        Select Case segmentType
            Case SegmentTypes.Region
                Return RegionRegex
            Case SegmentTypes.Statements
                Return CommentRegex
            Case Else
                Return Nothing
        End Select

    End Function
#End Region


#Region "Properties"

    Private Shared _Attribute As T
    ''' <summary>
    ''' Generator Attribute
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared ReadOnly Property GenAttribute As T
        Get
            If _Attribute Is Nothing Then _Attribute = New T
            Return _Attribute
        End Get

    End Property


#End Region

    ''' <summary>
    ''' Find textPoint marked with '<code>'<Gen Type="InsertPoint" /></code>
    ''' </summary>
    ''' <param name="writer"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    Function FindInsertionPoint(writer As TaggedSegmentWriter) As FoundTaggedSegment
     
        Return FindSegments(writer, GeneratorAttribute.TagTypes.InsertPoint).FirstOrDefault
    End Function
    Function FindGeneratedSegments(writer As TaggedSegmentWriter) As FoundTaggedSegment()

        Return FindSegments(writer, GeneratorAttribute.TagTypes.Generated).
                Where(Function(x) x.FoundTag.SegmentClass = writer.GenAttribute.SegmentClass).ToArray

    End Function
    Function FindSegments(writer As TaggedSegmentWriter, tagType As GeneratorAttribute.TagTypes) As IEnumerable(Of FoundTaggedSegment)
        Return FindSegments(writer).Where(Function(x) x.FoundTag.Type = tagType)
    End Function

    ''' <summary>
    ''' Find tagged segment within GenInfo.SearchStart and GenInfo.SearchEnd
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Not using EditPoint.FindPattern because it can only search from startpoint to end of doc, no way to limit to selection
    ''' Not using DTE Find because it has to change params of current find dialog, might screw up normal find usage
    '''  </remarks>
    Public Function FindSegments(info As TaggedSegmentWriter) As FoundTaggedSegment()

        Dim regex = GetRegexByType(info.SegmentType)

        'Using regex in FindPattern does
        Dim text = info.GetSearchText
        Dim matches = regex.Matches(text)
        Dim segments = New List(Of FoundTaggedSegment)
        For Each m In matches.Cast(Of Match)()
            Dim matchStart As EditPoint = Nothing, matchEnd As EditPoint = Nothing

            If m.Success Then
                'Convert match into start and end TextPoints
                matchStart = info.SearchStart.CreateEditPoint
                matchStart.CharRightExact(m.Index)
                matchEnd = matchStart.CreateEditPoint
                matchEnd.CharRightExact(m.Length)

            End If

            Dim segment = New FoundTaggedSegment(Me, info.GenAttribute, matchStart, matchEnd)
            If segment.IsValid Then segments.Add(segment)
        Next
        Return segments.ToArray
    End Function
    Public Function IsAnyOutdated(info As TaggedSegmentWriter) As Boolean
        Dim segments = FindGeneratedSegments(info)
        Return Not segments.Any OrElse segments.Any(Function(x) x.IsOutdated)
    End Function


    ''' <summary>
    ''' Insert or Replace text in taggedRange if outdated (or set to always generate)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertOrReplace(info As TaggedSegmentWriter) As Boolean
        Dim taggedRanges = FindGeneratedSegments(info)
        Dim needInsert = False
        If taggedRanges.Length = 0 Then
            'if none found, then insert
            needInsert = True
        Else
            'if any is outdated, delete, and reinsert
            For Each t In From t1 In taggedRanges Where t1.IsOutdated

                t.Delete()
                needInsert = True
            Next
        End If
        If Not needInsert Then Return False

        info.InsertAndFormat()
        '!Open file if requested
        If info.OpenFileOnGenerated AndAlso info.Class IsNot Nothing Then
            If Not info.Class.ProjectItem.IsOpen Then info.Class.ProjectItem.Open()
        End If
        Return True
    End Function
    Public Function Remove(info As TaggedSegmentWriter) As Boolean
        Dim taggedRanges = FindSegments(info)
        For Each t In taggedRanges
            t.Delete()
        Next
    End Function
End Class