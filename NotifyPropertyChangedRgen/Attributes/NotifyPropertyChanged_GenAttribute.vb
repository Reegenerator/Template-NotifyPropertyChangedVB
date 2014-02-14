Imports EnvDTE80
Imports System.Reflection
Imports ThisClass = NotifyPropertyChangedRgen.NotifyPropertyChanged_GenAttribute
Imports <xmlns:rgn="http://tempuri.org/NotifyPropertyChanged.xsd">
Imports System.Xml.Linq

Public Class NotifyPropertyChanged_GenAttribute
    Inherits GeneratorAttribute

    Sub New()
        Init()
    End Sub
    Sub New(p As CodeProperty2)
        MyBase.New(p)
        Init()
    End Sub
    Sub New(cc As CodeClass2)
        MyBase.New(cc)
        Init()
    End Sub
    Overrides Sub Init()
        MyBase.Init()
        'To regenerate all OnVersionChanged generated code, increment the version number
        Version = New Version(1, 1, 0, 16)
    End Sub
    Friend Sub New(f As CodeFunction2)
        MyBase.New(f)
        Init()
    End Sub

 
    Private Shared PrototypeXElement As XElement
    Public Overrides ReadOnly Property TagPrototype As Xml.Linq.XElement
        Get

            If PrototypeXElement Is Nothing Then PrototypeXElement = <rgn:Gen Renderer=<%= TagName %>/>
            Return PrototypeXElement
        End Get
    End Property

    Public Enum GenerationTypes
        DefaultType
        SetAndNotify
        NotifyOnly
    End Enum

    Shared ReadOnly ThisType As Type = GetType(ThisClass)

    Property GenerationType As GenerationTypes


    ''' <summary>
    ''' A simple comma delimited string, since its in attribute we cannot use expression as parameters
    ''' But it will be checked against the type during generation, if the property does not exists there will be a warning
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlPropertyAttribute("ExtraNotifications")>
    Property ExtraNotifications As String


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns>Array of invalid properties</returns>
    ''' <remarks></remarks>
    Function ValidateExtraNotifications(cc As CodeClass2, extras() As String) As String()

        Dim propNames = New HashSet(Of String)(cc.GetProperties.Select(Function(x) x.Name))
        Dim invalids = extras.Where(Function(x) Not propNames.Contains(x))
        If invalids.Count > 0 Then
            Return invalids.ToArray()
        End If

        ExtraNotifications = String.Join(", ", extras)
        Return New String() {} 'return empty array
    End Function

    Overrides Function AreArgumentsEquals(other As GeneratorAttribute) As Boolean
        Dim otherNPC = DirectCast(other, NotifyPropertyChanged_GenAttribute)
        Return MyBase.AreArgumentsEquals(other) AndAlso
            Me.ExtraNotifications = otherNPC.ExtraNotifications
    End Function
End Class