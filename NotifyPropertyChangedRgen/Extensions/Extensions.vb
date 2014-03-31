Imports System.Runtime.CompilerServices
Imports System.ComponentModel.Design
Imports System.Linq.Expressions
Imports System.IO
Imports System.Collections.Concurrent
Imports Microsoft.VisualBasic.CompilerServices
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell.Design
Imports EnvDTE
Imports EnvDTE80
Imports Kodeo.Reegenerator.Generators
Imports Kodeo.Reegenerator
Imports System.Xml.Linq
Imports System.Xml
Imports System.Text
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.ComponentModel

Module Extensions
    Public Const DefaultRegexOption As RegexOptions = RegexOptions.Compiled Or RegexOptions.IgnoreCase Or RegexOptions.IgnorePatternWhitespace Or RegexOptions.Singleline

#Region "Code element helpers"

    <Extension>
    Public Function GetClassesEx(item As Wrappers.ProjectItem) As IEnumerable(Of CodeClass2)
        Dim classes = item.GetClasses().Values.SelectMany(Function(x) x).Cast(Of CodeClass2)()
        Return classes
    End Function


    <Extension>
    Public Function GetClassesWithAttributes(item As Wrappers.ProjectItem, attributes As System.Type()) _
        As IEnumerable(Of CodeClass2)
        'Replace nested class + delimiter into . as the format used in CodeAttribute.FullName
        Dim fullNames = attributes.Select(Function(x) x.DottedFullName).ToArray
        Dim res = item.GetClassesEx.Where(
            Function(cclass) fullNames.All(
                Function(attrName) cclass.Attributes.Cast(Of CodeAttribute).Any(
                    Function(cAttr) cAttr.FullName = attrName))
            )

        Return res
    End Function



    ''' <summary>
    ''' Returns full name delimited by only dots (and no +(plus sign))
    ''' </summary>
    ''' <param name="x"></param>
    ''' <returns></returns>
    ''' <remarks>Nested class is separated with +, while CodeClass delimit them using dots</remarks>
    <Extension>
    Public Function DottedFullName(x As Type) As String

        Return x.FullName.Replace("+", ".")
    End Function

    <Extension>
    Public Function GetClassesWithAttribute(item As Wrappers.ProjectItem, attribute As System.Type) _
        As IEnumerable(Of CodeClass2)
        Dim fullName = attribute.DottedFullName
        '   all attributes is in class attribute
        Dim res = item.GetClassesEx.Where(
            Function(cclass) cclass.Attributes.Cast(Of CodeAttribute).Any(Function(x) x.FullName = fullName)
            )

        Return res
    End Function

    <Extension>
    Public Function GetClassesWithAttribute(dte As EnvDTE.DTE, attribute As System.Type) As IEnumerable(Of CodeClass2)
        Dim projects = Wrappers.Solution.GetSolutionProjects(dte.Solution).Values
        Dim res = From p In projects
                From eleList In p.GetCodeElements(Of CodeClass2).Values
                From ele In eleList
                Where ele.Attributes.Cast(Of CodeAttribute).Any(Function(x) x.AsElement.IsEqual(attribute))
                Select ele

        Return res
    End Function

    ''' <summary>
    ''' Use this to convert Code element into a more generic CodeElement and get CodeElement based extensions
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="cc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function AsElement(Of T)(cc As T) As CodeElement2
        Return CType(cc, CodeElement2)
    End Function
    <Extension>
    Public Function AsCode(Of T)(cc As T) As CodeType

        Return CType(cc, CodeType)
    End Function

    <Extension>
    Public Function HasAttribute(ct As CodeType, attrType As Type) As Boolean
        Return ct.GetCustomAttribute(attrType) IsNot Nothing
    End Function


#Region "CodeElement2 Attribute Functions"

    <Extension>
    Public Function GetCustomAttribute(cc As CodeElement2, attrType As Type) As CodeAttribute2

        Dim res = cc.GetCustomAttributes.Where(Function(x) x.AsElement.IsEqual(attrType)).FirstOrDefault
        Return res
    End Function

    ''' <summary>
    ''' Get Custom Attributes
    ''' </summary>
    ''' <param name="ce"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Requires Named Argument when declaring the Custom Attribute, otherwise Name will be empty.
    ''' Not using reflection because it requires successful build
    ''' </remarks>
    <Extension>
    Public Function GetCustomAttributes(ce As CodeElement2) As IEnumerable(Of CodeAttribute2)
        '!Property
        Dim prop = TryCast(ce, CodeProperty2)
        If prop IsNot Nothing Then Return prop.Attributes.Cast(Of CodeAttribute2)()

        '!Function
        Dim func = TryCast(ce, CodeFunction2)
        If func IsNot Nothing Then Return func.Attributes.Cast(Of CodeAttribute2)()

        '!Class
        Dim cc = TryCast(ce, CodeClass2)
        If cc IsNot Nothing Then
            Return cc.Attributes.Cast(Of CodeAttribute2)()
        End If

        Throw New Exception("CodeElement not recognized")
        Return Enumerable.Empty(Of CodeAttribute2)()
    End Function

    <Extension>
    Public Function HasAttribute(ct As CodeElement2, attrType As Type) As Boolean
        Return ct.GetCustomAttribute(attrType) IsNot Nothing
    End Function

#End Region
#Region "GetCustomAttributes of CodeType/CodeClass "


    <Extension>
    Public Function GetCustomAttributes(ct As CodeType) As IEnumerable(Of CustomAttributeData)

        Return Type.GetType(ct.FullName).CustomAttributes
    End Function
    <Extension>
    Public Function GetCustomAttributes(cc As CodeClass) As IEnumerable(Of CustomAttributeData)
        Return Type.GetType(cc.FullName).CustomAttributes
    End Function
    <Extension>
    Public Function GetCustomAttribute(ct As CodeType, attrType As Type) As CustomAttributeData
        Return ct.GetCustomAttributes.Where(Function(x) x.AttributeType Is attrType).FirstOrDefault
    End Function
    <Extension>
    Public Function GetCustomAttribute(cc As CodeClass, attrType As Type) As CustomAttributeData
        Return cc.GetCustomAttributes.Where(Function(x) x.AttributeType Is attrType).FirstOrDefault
    End Function



    <Extension>
    Function GetCodeAttributeArguments(ByVal cattr As CodeAttribute2) As IEnumerable(Of CodeAttributeArgument)
        If cattr Is Nothing Then Return Enumerable.Empty(Of CodeAttributeArgument)()
        Return cattr.Arguments.Cast(Of CodeAttributeArgument)()
    End Function
    <Extension>
    Public Function IsEqual(ele As CodeElement2, type As Type) As Boolean
        Static attrType As Type = GetType(Attribute)
        Return ele.FullName = type.FullName OrElse ele.Name = type.Name OrElse
                (type.IsSubclassOf(attrType) AndAlso
                    (ele.FullName & "Attribute" = type.FullName OrElse ele.Name & "Attribute" = type.Name)
                )
    End Function

#End Region
#Region "CodeClass members(property, function, variable) helper"

    <Extension>
    Public Function GetProperties(cls As CodeClass) As IEnumerable(Of CodeProperty2)
        Return cls.Children.OfType(Of CodeProperty2)()
    End Function
    <Extension>
    Public Function GetFunctions(cls As CodeClass) As IEnumerable(Of CodeFunction2)
        Return cls.Children.OfType(Of CodeFunction2)()
    End Function

    <Extension>
    Public Function GetAutoProperties(cls As CodeClass2) As CodeProperty2()

        Dim props = cls.GetProperties
        Return props.Where(Function(x) x.ReadWrite = EnvDTE80.vsCMPropertyKind.vsCMPropertyKindReadWrite AndAlso
                                       x.Setter Is Nothing AndAlso
                                       x.OverrideKind <> EnvDTE80.vsCMOverrideKind.vsCMOverrideKindAbstract).ToArray
    End Function

    <Extension>
    Public Function GetVariables(cls As CodeClass) As IEnumerable(Of CodeVariable)
        Return cls.Children.OfType(Of CodeVariable)()
    End Function
    <Extension>
    Public Function GetDependencyProperties(cls As CodeClass) As IEnumerable(Of CodeVariable)
        Try

            Dim sharedFields = cls.GetVariables.Where(Function(x) x.IsShared AndAlso x.Type.CodeType IsNot Nothing)
            Return sharedFields.Where(Function(x) x.Type.CodeType.FullName = "System.Windows.DependencyProperty")
        Catch ex As Exception

            MsgBox(ex)
        End Try
        Return Nothing
    End Function



    ''' <summary>
    ''' Get Bases recursively
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function GetAncestorClasses(cc As CodeClass2) As IEnumerable(Of CodeClass2)
        Dim bases = cc.Bases.Cast(Of CodeClass2).ToArray
        If bases.FirstOrDefault Is Nothing Then Return bases
        Dim grandBases = bases.SelectMany(Function(x) x.GetAncestorClasses)

        Return bases.Concat(grandBases)

    End Function
#End Region
#End Region

    <Extension>
    Public Function GetText(prop As CodeProperty2, Optional part As vsCMPart = vsCMPart.vsCMPartWholeWithAttributes) As String
        Dim p = prop.GetStartPoint(part)
        If p Is Nothing Then Return ""
        Return p.CreateEditPoint.GetText(prop.GetEndPoint(part))

    End Function
    <Extension>
    Public Function GetText(cls As CodeClass2, Optional part As vsCMPart = vsCMPart.vsCMPartWholeWithAttributes) As String
        Dim p = cls.GetStartPoint(part)
        If p Is Nothing Then Return ""
        Return p.CreateEditPoint.GetText(cls.GetEndPoint(part))

    End Function

    ReadOnly InterfaceImplementationPattern As String =
        <String><![CDATA[                           
            ^.*?\sAs\s.*?(?<impl>Implements\s.*?)$
        ]]></String>.Value
    <Extension>
    Public Function GetInterfaceImplementation(prop As CodeProperty2) As String
        Static regex As Regex = New Regex(InterfaceImplementationPattern, DefaultRegexOption)
        Dim g = regex.Match(prop.GetText(vsCMPart.vsCMPartHeader)).Groups("impl")

        'add space to separate
        If g.Success Then Return " " & g.Value
        Return Nothing
    End Function
    <Extension>
    Public Function GetAttributeStartPoint(prop As CodeProperty2) As TextPoint
        Return prop.GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes)
    End Function


    Private _DocCommentRegex As Text.RegularExpressions.Regex
    ''' <summary>
    ''' Lazy Regex property to match doc comments
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property DocCommentRegex As Text.RegularExpressions.Regex
        Get
            Const docCommentPattern As String = "\s'''"
            If _DocCommentRegex Is Nothing Then
                _DocCommentRegex = New Text.RegularExpressions.Regex(docCommentPattern)
            End If
            Return _DocCommentRegex
        End Get
    End Property

    <Extension>
    Public Function GetCommentStartPoint(ce As CodeElement) As EditPoint
        Return ce.GetStartPoint(vsCMPart.vsCMPartHeaderWithAttributes).GetCommentStartPoint()
    End Function
    <Extension>
    Public Function GetCommentStartPoint(ce As CodeProperty2) As EditPoint
        Return ce.GetStartPoint(vsCMPart.vsCMPartHeaderWithAttributes).GetCommentStartPoint()
    End Function

    ''' <summary>
    ''' Get to the beginning of doc comments for startPoint
    ''' </summary>
    ''' <param name="startPoint"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' EnvDte does not have a way to get to the starting point of a code element doc comment. 
    ''' If we need to insert some text before a code element that has doc comments we need to go to the beggining of the comments.
    ''' </remarks>
    <Extension>
    Public Function GetCommentStartPoint(startPoint As TextPoint) As EditPoint

        Dim sp = startPoint.CreateEditPoint
        'keep going 1 line up until the line does not start with doc comment prefix
        Do
            sp.LineUp()
        Loop While DocCommentRegex.IsMatch(sp.GetLineText)
        'Go to the beginning of first line of comment, or element itself
        sp.LineDown()
        sp.StartOfLine()
        Return sp
    End Function

    <Extension> _
    Public Function ToStringFormatted(xml As XElement) As String
        Dim settings As New XmlWriterSettings()
        settings.OmitXmlDeclaration = True

        Dim result As New StringBuilder()
        Using writer As XmlWriter = XmlWriter.Create(result, settings)

            xml.WriteTo(writer)
        End Using
        Return result.ToString()
    End Function

#Region "ExprToString. Convert Member Expression to string."

    Public Function ExprsToString(Of T)(ParamArray exprs() As Expression(Of Func(Of T, Object))) As String()

        Dim strings = (From x In exprs Select DirectCast(x, LambdaExpression).ExprToString()).ToArray
        Return strings
    End Function

    <Extension>
    Public Function ExprToString(Of T, T2)(ByVal expr As Expression(Of Func(Of T, T2))) As String
        Return DirectCast(expr, LambdaExpression).ExprToString()
    End Function

    <Extension>
    Public Function ExprToString(ByVal memberExpr As LambdaExpression) As String
        If memberExpr Is Nothing Then Return ""
        Dim currExpr As System.Linq.Expressions.Expression
        'when T2 is object, the expression will be wrapped in UnaryExpression of Convert{}
        Dim convertedToObject = TryCast(memberExpr.Body, UnaryExpression)
        If convertedToObject IsNot Nothing Then
            'unwrap
            currExpr = convertedToObject.Operand
        Else
            currExpr = memberExpr.Body
        End If
        Select Case currExpr.NodeType
            Case ExpressionType.MemberAccess
                Dim ex = DirectCast(currExpr, MemberExpression)
                Return ex.Member.Name
        End Select

        Throw New Exception("Expression ToString() extension only processes MemberExpression")
    End Function

#End Region

    Public Function FindAllDerivedTypes(Of T)() As List(Of Type)
        Return FindAllDerivedTypes(Of T)(Assembly.GetAssembly(GetType(T)))
    End Function

    Public Function FindAllDerivedTypes(Of T)(assembly As Assembly) As List(Of Type)
        Dim derivedType = GetType(T)
        Return assembly.GetTypes().Where(Function(x) x <> derivedType AndAlso derivedType.IsAssignableFrom(x)).ToList()

    End Function

    <Extension>
    Public Function GetSubclasses(cc As CodeClass2) As IEnumerable(Of CodeClass2)
        Dim fullname = cc.FullName
        Dim list As New List(Of CodeClass2)
        Kodeo.Reegenerator.Wrappers.CodeElement.TraverseSolutionForCodeElements(Of CodeClass2)(cc.DTE.Solution,
                                                                                Sub(x) list.Add(x),
                                                                                Function(x) x.FullName <> fullname AndAlso x.IsDerivedFrom(fullname))
        Return list.ToArray
    End Function

    <Extension>
    Public Function RemoveEmptyLines(s As String) As String
        Static regex As Regex = New Regex("^\s+$[\r\n]*", RegexOptions.Multiline)
        Return regex.Replace(s, "")
    End Function


    <Extension>
    Function SelectOrDefault(Of T, TResult)(obj As T, selectFunc As Func(Of T, TResult), Optional defaultValue As TResult = Nothing) As TResult
        If obj Is Nothing Then Return defaultValue
        Return selectFunc(obj)
    End Function

    ''' <summary>
    ''' Returns a type from an assembly reference by ProjectItem.Project. Cached.
    ''' </summary>
    ''' <param name="pi"></param>
    ''' <param name="typeName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Function GetTypeFromProject(pi As ProjectItem, typeName As String) As Type
        Static cache As New Dictionary(Of String, Assembly)
        Dim path = pi.GetAssemblyPath()
        If Not cache.ContainsKey(path) Then
            cache.Add(path, Assembly.LoadFrom(path))
        End If

        Dim asm = cache(path)
        Return asm.GetType(typeName)
    End Function
    <Extension>
    Function ToType(cc As CodeClass) As Type
        Return cc.ProjectItem.GetTypeFromProject(cc.FullName)
    End Function

    ''' <summary>
    ''' Convert CodeProperty2 to PropertyInfo. Cached
    ''' </summary>
    ''' <param name="prop"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Function ToPropertyInfo(prop As CodeProperty2) As PropertyInfo

        Static classCache As New ConcurrentDictionary(Of CodeClass, Type)
        Dim classType = classCache.GetOrAdd(prop.Parent, Function() prop.Parent.ToType)
        Return classType.GetProperty(prop.Name, BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
    End Function

    <Extension>
    Function GetGeneratorAttribute(mi As MemberInfo) As GeneratorAttribute
        Static type As Type = GetType(GeneratorAttribute)
        Dim genAttr = mi.GetCustomAttributes.FirstOrDefault(Function(x) x.GetType.IsSubclassOf(type))
        Return CType(genAttr, GeneratorAttribute)
    End Function

    <Extension>
    Function GetAssemblyOfProjectItem(pi As ProjectItem) As Assembly
        Dim path As String = pi.GetAssemblyPath()

        If path <> "" Then
            Return Assembly.LoadFrom(path)
        Else
            Return Nothing
        End If

    End Function

    <Extension>
    Public Function GetAssemblyPath(ByVal pi As ProjectItem) As String

        Dim assemblyName = pi.ContainingProject.Properties.Cast(Of [Property]).
                FirstOrDefault(Function(x) x.Name = "AssemblyName").
                SelectOrDefault(Function(x) x.Value)

        Return pi.ContainingProject.GetAssemblyPath
    End Function

    ''' <summary>
    ''' Currently unused. If we require a succesful build, a project that requires succesful generation would never build, catch-22
    ''' </summary>
    ''' <param name="vsProject"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function GetAssemblyPath(vsProject As EnvDTE.Project) As String
        Dim fullPath As String = vsProject.Properties.Item("FullPath").Value.ToString()
        Dim outputPath As String = vsProject.ConfigurationManager.ActiveConfiguration.Properties.Item("OutputPath").Value.ToString()
        Dim outputDir As String = Path.Combine(fullPath, outputPath)
        Dim outputFileName As String = vsProject.Properties.Item("OutputFileName").Value.ToString()
        Dim assemblyPath As String = Path.Combine(outputDir, outputFileName)
        Return assemblyPath
    End Function

    ReadOnly TypeWithoutQualifierPattern As String =
        <String><![CDATA[    
            (?<=\.?)[^\.]+?$
        ]]></String>.Value
    ReadOnly TypeWithoutQualifierRegex As Regex = New Regex(TypeWithoutQualifierPattern, RegexOptions.IgnoreCase Or RegexOptions.IgnorePatternWhitespace)
    <Extension>
    Public Function StripQualifier(s As String) As String
        Dim stripped = TypeWithoutQualifierRegex.Match(s).Value
        Return stripped
    End Function

    <Extension>
    Public Function ParseAsEnum(Of T As Structure)(qualifiedName As String, defaultValue As T) As T
        If qualifiedName Is Nothing Then Return defaultValue
        Dim res As T = defaultValue
        [Enum].TryParse(qualifiedName.StripQualifier, res)
        Return res
    End Function

    <Extension>
    Public Function GetOrInit(Of T)(ByRef x As T, initFunc As Func(Of T)) As T
        If x Is Nothing Then
            x = initFunc()
        End If
        Return x
    End Function

    ''' <summary>
    ''' Create type instance from string
    ''' </summary>
    ''' <param name="type"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Function ConvertFromString(type As Type, value As String) As Object
        Return TypeDescriptor.GetConverter(type).ConvertFromString(value)
    End Function

    ''' <summary>
    ''' Set property value from string representation
    ''' </summary>
    ''' <param name="propInfo"></param>
    ''' <param name="obj"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    <Extension>
    Sub SetValueFromString(propInfo As PropertyInfo, obj As Object, value As String)
        Dim setValue As Object
        If propInfo.PropertyType Is GetType(Version) Then
            setValue = Version.Parse(value)
        Else
            setValue = propInfo.PropertyType.ConvertFromString(value)

        End If
        propInfo.SetValue(obj, setValue)
    End Sub

    <Extension>
    Sub AddInterfaceIfNotExists(cls As CodeClass2, interfaceName As String)
        If Not cls.ImplementedInterfaces.OfType(Of CodeInterface).Any(Function(x) x.FullName = interfaceName) Then
            cls.AddImplementedInterface(interfaceName)
        End If
    End Sub

    <Extension>
    Function DotJoin(s As String, ParamArray segments As String()) As String
        Dim all = {s}.Concat(segments).ToArray
        Return String.Join(".", all)
    End Function


    ''' <summary>
    ''' Returns CodeTypeRef.AsFullName, if null, returns CodeTypeRef.AsString
    ''' </summary>
    ''' <param name="ctr"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If there's compile error AsFullName will be null
    ''' </remarks>
    <Extension>
    Public Function SafeFullName(ctr As CodeTypeRef) As String
        Return If(ctr.AsFullName, ctr.AsString)
    End Function

End Module
