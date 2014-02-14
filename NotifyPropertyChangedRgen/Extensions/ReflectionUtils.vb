Imports System.Collections.Concurrent
Imports System.Reflection
Imports Microsoft.VisualStudio.Shell.Interop
Imports System.ComponentModel.Design
Imports Microsoft.VisualStudio.Shell.Design
Imports Microsoft.VisualStudio.Shell
Imports Kodeo.Reegenerator.Wrappers
Imports EnvDTE80

Public Class TypeCacheList
    ReadOnly ByNameCache As New Dictionary(Of String, TypeCache)
    ReadOnly ByTypeCache As New Dictionary(Of Type, TypeCache)

    Function Contains(type As Type) As Boolean
        Return ByTypeCache.ContainsKey(type)
    End Function
    Function ByName(cc As CodeClass2) As TypeCache

        If Not ByNameCache.ContainsKey(cc.FullName) Then
            Dim tc = New TypeCache(cc.FullName)
            ByNameCache.Add(cc.FullName, tc)
            ByTypeCache.Add(tc.TypeInfo.AsType, tc)
        End If

        Return ByNameCache(cc.FullName)
    End Function

    Function ByType(type As Type) As TypeCache
        If Not ByTypeCache.ContainsKey(type) Then
            Dim tc = New TypeCache(type)
            ByNameCache.Add(tc.TypeInfo.Name, tc)
            ByTypeCache.Add(tc.TypeInfo.AsType, tc)
        End If
        Return ByTypeCache(type)
    End Function

End Class

Public Class TypeCache

    ReadOnly Cache As Dictionary(Of String, MemberInfo)
    Property TypeInfo As TypeInfo

    Sub New(typeName As String, Optional caseSensitiveMembers As Boolean = False)
        Me.New(System.Type.GetType(typeName))

    End Sub


    Sub New(type As Type, Optional caseSensitiveMembers As Boolean = False)
        Dim comparer = If(caseSensitiveMembers, StringComparer.Ordinal, StringComparer.OrdinalIgnoreCase)
        Cache = New Dictionary(Of String, MemberInfo)(comparer)

        Try
            TypeInfo = type.GetTypeInfo()
            For Each m In TypeInfo.GetMembers(BindingFlags.Instance Or BindingFlags.Public)
                'Prevent error on multiple cctor
                If Not Cache.ContainsKey(m.Name) Then Cache.Add(m.Name, m)

            Next

        Catch ex As Exception
            DebugHere()
        End Try
    End Sub
    Default ReadOnly Property Item(name As String) As MemberInfo
        Get
            If Not Cache.ContainsKey(name) Then
                Throw New Exception(String.Format("Member {0} not found in {1}", name, TypeInfo.Name))
            End If
            Return Cache(name)
        End Get
    End Property
    Public Function GetMembers() As IEnumerable(Of MemberInfo)
        Return Cache.Values
    End Function

    Function GetMember(name As String) As MemberInfo
        Return Cache(name)

    End Function

    Function Contains(name As String) As Boolean
        Return Cache.ContainsKey(name)
    End Function

    Function TryGetMember(key As String) As MemberInfo
        Dim value As MemberInfo = Nothing
        Cache.TryGetValue(key, value)
        Return value
    End Function

    Sub AddAlias(name As String, alternateName As String)
        Try
            Cache.Add(alternateName, Me(name))

        Catch ex As Exception
            System.Diagnostics.Debugger.Launch()
        End Try
    End Sub
End Class

Public Class TypeResolver
    Public Shared TypeCacheList As New TypeCacheList
    Shared Function ByType(type As System.Type) As TypeCache
        Return TypeCacheList.ByType(type)
    End Function
    Shared Function Contains(type As System.Type) As Boolean
        Return TypeCacheList.Contains(type)
    End Function
    Shared Function ByName(cc As EnvDTE80.CodeClass2) As TypeCache
        Return TypeCacheList.ByName(cc)
    End Function
    'Shared DteServiceProvider As IServiceProvider
    'Overloads Shared Function GetService(Of T)(dte As EnvDTE.DTE) As T

    '    If DteServiceProvider Is Nothing Then DteServiceProvider = New ServiceProvider(CType(dte, Microsoft.VisualStudio.OLE.Interop.IServiceProvider))
    '    Return CType(ServiceProvider.GlobalProvider.GetService(GetType(T)), T)
    'End Function

    ' ''' <summary>
    ' ''' 
    ' ''' </summary>
    ' ''' <returns></returns>
    ' ''' <remarks>http://blogs.clariusconsulting.net/kzu/how-to-get-a-system-type-from-an-envdte-codetyperef-or-envdte-codeclass/</remarks>
    'Private Shared Function GetResolutionService(project As EnvDTE.Project) As ITypeResolutionService
    '    ''If DteServiceProvider Is Nothing Then DteServiceProvider = 
    '    ''New ServiceProvider(CType(project.DTE, Microsoft.VisualStudio.OLE.Interop.IServiceProvider))
    '    ''Debugger.Launch()
    '    'Dim typeService As Object = ServiceProvider.GlobalProvider.GetService(GetType(DynamicTypeService)) ' CType(DteServiceProvider.GetService(GetType(DynamicTypeService)), DynamicTypeService)
    '    'Const vsDesign As String = "Microsoft.VisualStudio.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    '    'Dim assm = Assembly.Load(vsDesign)
    '    'Dim dtsType = assm.GetType("Microsoft.VisualStudio.Design.VSDynamicTypeService")
    '    ''Dim constructor = dtsType.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, {GetType(IServiceProvider)}, Nothing)
    '    ''Dim dts = constructor.Invoke({ServiceProvider.GlobalProvider})
    '    ''Dim methods = dtsType.GetMethods()
    '    ''Dim GetTypeResolutionServiceMethod = dtsType.GetMethod("GetTypeResolutionService", BindingFlags.Public Or BindingFlags.Instance)
    '    ''Dim x = GetTypeResolutionServiceMethod.Invoke(typeService, Nothing)

    '    'Dim dts2 = Activator.CreateInstance(dtsType, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, {ServiceProvider.GlobalProvider}, Nothing)
    '    ''Dim dts As Object = Activator.CreateInstanceFrom(, )
    '    ''GetService(Of DynamicTypeService)(project.DTE)

    '    'Dim sln As IVsSolution = GetService(Of IVsSolution)(project.DTE)
    '    'Dim hier As IVsHierarchy = Nothing
    '    'sln.GetProjectOfUniqueName(project.UniqueName, hier)

    '    'Debug.Assert(hier IsNot Nothing, "No active hierarchy is selected.")
    '    'Dim memberInfo = typeService.GetType.GetMethod("GetTypeResolutionService")
    '    'Return CType(typeService, DynamicTypeService).GetTypeResolutionService(hier)
    'End Function

    'Public Shared Function ResolveType(cc As CodeClass2) As Type
    '    Dim resolver = GetResolutionService(cc.ProjectItem.ContainingProject)

    '    Return resolver.GetType(cc.FullName, True)
    'End Function

End Class
