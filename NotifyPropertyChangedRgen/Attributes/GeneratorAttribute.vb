Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Xml.Linq
Imports Microsoft.VisualBasic.CompilerServices
Imports EnvDTE
Imports EnvDTE80
Imports System.Reflection

''' <summary>
''' BaseClass for Generator Attributes
''' </summary>
''' <remarks>
''' We cannot use nested class for the attribute (e.g. inside the renderer, because then other project would require 
''' reference to Kodeo.Regenerator.dll, since the renderer is derived from CodeRenderer
''' </remarks>
Public Class GeneratorAttribute
    Inherits Attribute
    Const NameSuffix As String = "_GenAttribute"

    Public Enum RegenModes
        OnVersionChanged
        Once
        Always
    End Enum

    ''' <summary>
    ''' Cause of code generation
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum TriggerTypes
        ''' <summary>
        ''' Code generation is triggered because the class is marked with a GeneratorAttribute 
        ''' </summary>
        ''' <remarks></remarks>
        Attribute
        ''' <summary>
        ''' Code generation is triggered because the baseClass is marked with a GeneratorAttribute
        ''' </summary>
        ''' <remarks></remarks>
        BaseClassAttribute
    End Enum
    Public Enum TagTypes
        Generated
        InsertPoint
    End Enum
    Public Property ParentProperty As CodeProperty2
    Public Property ParentFunction As CodeFunction2
    Public Property ParentElement As CodeElement2

    Public Property IsInsertionPoint As Boolean
    Public Property TypeCache As TypeCache

    Public Property Type As TagTypes

    ''' <summary>
    ''' Use to differentiate segments when we are calling FindSegments
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' Without a class differentiator, when searching for a class level segment, it will match all segments within the class
    ''' and will cause unintended deletion when the segment needs to be updated
    ''' </remarks>
    <XmlProperty("Class")>
    Public Property SegmentClass As String
    Shared Property XmlPropertiesByType As New Dictionary(Of Type, Dictionary(Of String, PropertyInfo))

    Overridable ReadOnly Property TagPrototype As XElement
        Get
            Return Nothing
        End Get
    End Property
#Region "Constructors"

    Sub New()
        '?+Do not create public constructors with parameters
        '?It will be hard to determine, using CodeAttribute (In CodeAttributeArgument Name is empty, only Value is shown), which constructor is being used (the info is design time)
        '?Use named parameter instead when trying to set the properties (Name and Value are shown)
        Init()
    End Sub

    Friend Sub New(f As CodeFunction2)
        ParentFunction = f
        Init(f.AsElement)
    End Sub
    Friend Sub New(p As CodeProperty2)
        ParentProperty = p
        Init(p.AsElement)
    End Sub

    Friend Sub New(cc As CodeClass)
        Init(cc.AsElement)
        CopyPropertyFromAttributeArguments(cc.AsElement.GetCustomAttribute(Me.GetType).GetCodeAttributeArguments)
    End Sub

    Private Sub Init(ByVal ele As CodeElement2)
        Init()
        ParentElement = ele
        Dim attrs = ele.GetCustomAttribute(Me.GetType).GetCodeAttributeArguments()
        CopyPropertyFromAttributeArguments(attrs)
    End Sub
    Overridable Sub Init()
        'Do once per application run. Can't make shared because we need to get actual derived class
        Dim type = Me.GetType
        TypeCache = TypeResolver.ByType(type)
        'Store members with XmlPropertyAttribute into a dictionary, to be used when writing the xml
        If Not XmlPropertiesByType.ContainsKey(type) Then
            Dim xmlmembers = New Dictionary(Of String, PropertyInfo)
            Dim members = TypeCache.GetMembers.ToArray
            For Each m In members
                Dim xmlName = m.GetCustomAttribute(Of XmlPropertyAttribute)()
                If xmlName Is Nothing Then Continue For
                xmlmembers.Add(xmlName.Name, CType(TypeCache(m.Name), PropertyInfo))
            Next
            XmlPropertiesByType.Add(type, xmlmembers)
        End If


    End Sub


    Friend Sub New(xele As XElement)
        Init()
        CopyPropertyFromTag(xele)

    End Sub

#End Region

    Function GetXmlProperties() As Dictionary(Of String, PropertyInfo)
        Return XmlPropertiesByType(Me.GetType)
    End Function
    Overridable Overloads Function MemberwiseClone() As GeneratorAttribute
        Return CType(MyBase.MemberwiseClone(), GeneratorAttribute)
    End Function
    Public Sub CopyPropertyFromTag(xele As XElement)
        Dim xmlProps = GetXmlProperties()
        For Each attr In xele.Attributes
            Dim name = attr.Name.LocalName
            If name = "Renderer" Then Continue For
            Dim propInfo As PropertyInfo = Nothing
            'If a property has an XmlProperty attribute, it will be rendered using that name, instead of the property name
            'Check XmlProperties first
            If Not xmlProps.TryGetValue(name, propInfo) Then
                'if not found, get by property name
                propInfo = TryCast(TypeCache.TryGetMember(attr.Name.LocalName), PropertyInfo)
            End If

            If propInfo IsNot Nothing Then
                SetPropertyFromAttributeArgumentString(propInfo, attr.Value)
            End If
        Next
    End Sub

    Private Sub CopyPropertyFromAttributeArguments(ByVal args As IEnumerable(Of CodeAttributeArgument))
        For Each arg In args
            Dim propInfo As PropertyInfo

            '?if enum, strip qualifier in value
            propInfo = CType(TypeCache(arg.Name), PropertyInfo)

            SetPropertyFromAttributeArgumentString(propInfo, arg.Value)
        Next
    End Sub
    ''' <summary>
    ''' Parse Attribute Argument into the actual string value
    ''' </summary>
    ''' <param name="propInfo"></param>
    ''' <param name="value"></param>
    ''' <remarks>
    ''' Attribute argument is presented exactly as it was typed
    ''' Ex: SomeArg:="Test" would result in the Argument.Value "Test" (with quote)
    ''' Ex: SomeArg:=("Test") would result in the Argument.Value ("Test") (with parentheses and quote)
    ''' </remarks>
    Private Sub SetPropertyFromAttributeArgumentString(propInfo As PropertyInfo, value As String)
        Dim stringValue As String
        Dim propType = propInfo.PropertyType
        If propType.IsEnum Then
            stringValue = value.StripQualifier()
        ElseIf propType Is GetType(String) Then
            stringValue = value.Trim(""""c)
        Else
            stringValue = value
        End If
        propInfo.SetValueFromString(Me, stringValue)
    End Sub

    <XmlProperty("Ver")>
    Overridable Property Version As Version

    ''' <summary>
    ''' Mode to be written in xml tag
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Can be overridden by GenInfo.RegenMode</remarks>
    <XmlPropertyAttribute("Mode")>
    Overridable Property RegenMode As RegenModes


    Private _TagName As String

    Public Overridable ReadOnly Property TagName() As String
        Get
            If _TagName = "" Then
                Dim nm = Me.GetType.Name
                _TagName = nm.Substring(0, nm.Length - NameSuffix.Length)
            End If
            Return _TagName
        End Get
    End Property

    Public Shared Function GetTypeFromTagName(tag As String) As Type
        Static assm As Assembly = GetType(GeneratorAttribute).Assembly
        Return assm.GetType(tag + NameSuffix)
    End Function


    Overridable Property IsIgnored As Boolean
    Overridable Property ApplyToDerivedClasses As Boolean = True

    Overridable Function AreArgumentsEquals(other As GeneratorAttribute) As Boolean
        Return Me.Version = other.Version
    End Function


End Class
