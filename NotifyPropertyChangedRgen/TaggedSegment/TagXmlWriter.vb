Imports System.Xml
Imports System.Text
Imports System.IO
Imports System.Xml.Linq


''' <summary>
''' A custom XmlTextWriter that generate xml embedded in comment or region name
''' </summary>
''' <remarks>
''' Setting WriteStartElement("", localname, "") only removes namespace for element, but not the xmlns attribute
''' Overriding WriteAttributes doesn't work, never gets called, so XML has to be stripped off of namespace beforehand
''' 
''' newline before content cannot be done by overriding WriteEndElement(this writes blah end tag)
''' What we need is a private function in XmlTextWriter, called WriteEndStartTag
''' </remarks>
Public Class TagXmlWriter
    Inherits XmlTextWriter
    Const CodeCommentPrefix As String = "'"
    Property SegmentType As SegmentTypes

    'Property IsRegion As Boolean
    Public Sub New(writer As StringWriter)
        MyBase.New(writer)
        Me.QuoteChar = "'"c
    End Sub

    Public Overloads Overrides Sub WriteStartElement(ByVal prefix As String, ByVal localname As String, ByVal ns As String)
        'insert inline comment character before the start tag
        If SegmentType = SegmentTypes.Statements Then Me.WriteString(CodeCommentPrefix)
        MyBase.WriteStartElement(prefix, localname, ns)
    End Sub

    Public Overrides Sub WriteFullEndElement()
        'insert inline comment character before the end tag
        If SegmentType = SegmentTypes.Statements Then Me.WriteString(CodeCommentPrefix)
        MyBase.WriteFullEndElement()
        'add new line
        Me.WriteString(Environment.NewLine)
    End Sub
    Public Shared Function ToCommentedString(x As XElement) As String
        Return InternalToString(x, SegmentTypes.Statements)
    End Function

    ''' <summary>
    ''' Write xml based on segment type
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="segmentType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function InternalToString(ByVal x As XElement, segmentType As SegmentTypes) As String
        Dim writer As New StringWriter()
        Dim cw As New TagXmlWriter(writer) With {.SegmentType = segmentType}
        'Strip Namespace if it's an Xelement
        x = StripNS(x)
        'write
        x.WriteTo(cw)
        Return writer.GetStringBuilder.ToString
    End Function
    Public Shared Function EscapeQuote(s As String) As String
        Const Quote As String = """"
        Const DoubleQuote As String = Quote & Quote
        Return s.Replace(Quote, DoubleQuote)
    End Function

    Public Shared Function ToRegionNameString(x As XElement) As String

        Dim xml = InternalToString(x, SegmentTypes.Region)
        'Escape quote to double quote, so it will be valid as region name
        Dim res = EscapeQuote(xml)
        Return res
    End Function

    Public Shared Function ToStringNoNS(xmlDocument As XElement) As String
        Return StripNS(xmlDocument).ToString
    End Function
    Public Shared Function StripNS(root As XElement) As XElement
        Dim res = New XElement(root)
        res.ReplaceAttributes(root.Attributes().Where(Function(attr) (Not attr.IsNamespaceDeclaration)))
        Return res
    End Function


 
End Class
