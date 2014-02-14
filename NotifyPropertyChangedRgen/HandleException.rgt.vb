Imports System
Imports System.Collections.Generic
Imports Wrappers = Kodeo.Reegenerator.Wrappers
Imports Generators = Kodeo.Reegenerator.Generators
Imports EnvDTE
Imports System.Text


    Public Partial Class HandleException
    Private _interfaces As List(Of CodeInterface)

    Public Overrides Sub PreRender()
        MyBase.PreRender()
        ' System.Diagnostics.Debugger.Break()
        Me._interfaces = Me.GetInterfaces()
    End Sub

    Private Function GetInterfaces() As List(Of CodeInterface)
        Dim interfaces As List(Of CodeInterface) = New List(Of CodeInterface)()
        If MyBase.ProjectItem.DteObject.FileCodeModel Is Nothing Then
            Return interfaces
        End If
        AddInterfaces(interfaces, MyBase.ProjectItem.DteObject.FileCodeModel.CodeElements)
        Return interfaces
    End Function

    Private Sub AddInterfaces(ByVal interfaces As List(Of CodeInterface), ByVal codeElements As EnvDTE.CodeElements)
        For Each codeElement As CodeElement In codeElements
            If codeElement.Kind = EnvDTE.vsCMElement.vsCMElementInterface Then
                interfaces.Add(CType(codeElement, CodeInterface))
            End If
            If codeElement.Kind = EnvDTE.vsCMElement.vsCMElementNamespace Then
                AddInterfaces(interfaces, CType(codeElement, EnvDTE.CodeNamespace).Members)
            End If
        Next codeElement
    End Sub

    Public Function GetMethodCall(ByVal codeFunction As CodeFunction) As String
        Dim result As String = codeFunction.Prototype(CType(vsCMPrototype.vsCMPrototypeParamNames, Integer))
        result = result.Replace("ByVal ", String.Empty)
        result = result.Replace("ByRef ", String.Empty)
        result = result.Replace(" )", ")")
        Return result
    End Function
End Class
