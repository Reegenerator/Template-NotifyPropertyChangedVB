Imports System.ComponentModel.Design
Imports EnvDTE80
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell.Design
Imports Kodeo.Reegenerator.Generators
Imports EnvDTE
Imports System.Text
Imports System.Reflection
Imports Microsoft.VisualStudio.Shell

Public MustInherit Class CodeRendererEx
    Inherits CodeRenderer
    Private SavedOutput As String
    Public ReadOnly Newline As String = Environment.NewLine
    ReadOnly Property DTE As EnvDTE.DTE
        Get
            Return MyBase.ProjectItem.DteObject.DTE
        End Get
    End Property

    ReadOnly Property DTE2 As EnvDTE80.DTE2
        Get
            Return DirectCast(DTE, EnvDTE80.DTE2)
        End Get
    End Property

    Public _OutputBuilder As StringBuilder
    ReadOnly Property OutputBuilder As StringBuilder
        Get
            If _OutputBuilder Is Nothing Then
                _OutputBuilder = Me.Output.GetStringBuilder
            End If
            Return _OutputBuilder

        End Get
    End Property


    Public Function GetCodeElementsAtCursor(Optional kind? As vsCMElement = Nothing) As CodeElement()

        Dim sel As TextSelection = CType(DTE.ActiveDocument.Selection, TextSelection)
        Dim pnt As TextPoint = CType(sel.ActivePoint, TextPoint)

        ' Discover every code element containing the insertion point.
        Dim fcm As FileCodeModel = _
            DTE.ActiveDocument.ProjectItem.FileCodeModel
        Dim res = GetCodeElementsAtPoint(fcm, pnt)
        If kind.HasValue Then
            res = res.Where(Function(x) x.Kind = vsCMElement.vsCMElementProperty).ToArray
        End If
        Return res
    End Function
    Public Function GetCodeElementsAtCursor(Of T As Class)() As IEnumerable(Of T)
        Dim kind As vsCMElement
        If GetType(T) Is GetType(CodeProperty) Then
            kind = vsCMElement.vsCMElementProperty
        ElseIf GetType(T) Is GetType(CodeClass) Then
            kind = vsCMElement.vsCMElementClass
        End If

        Dim ce = GetCodeElementsAtCursor(kind)
        Return ce.Cast(Of T)()

    End Function

    Public Function RenderToString() As String
        Return Text.ASCIIEncoding.ASCII.GetString(Me.Render().GeneratedCode)
    End Function

    Public Function GetCodeElementsAtPoint(ByVal fcm As FileCodeModel, ByVal pnt As TextPoint) As CodeElement()
        Dim res = New List(Of CodeElement)
        Dim elem As CodeElement
        Dim scope As vsCMElement

        For Each scope In [Enum].GetValues(scope.GetType())
            Try
                elem = fcm.CodeElementFromPoint(pnt, scope)
                If elem IsNot Nothing Then res.Add(elem)
            Catch ex As Exception
                'don’t do anything -
                'this is expected when no code elements are in scope
            End Try
        Next
        Return res.ToArray
    End Function
    Public Function GetTextSelection() As TextSelection
        Return CType(DTE.ActiveDocument.Selection, TextSelection)
    End Function

    Public Function SaveAndClearOutput() As String
        Me.OutputBuilder.Insert(0, SavedOutput) ' combine with savedoutput in front
        SavedOutput = Me.Output.ToString()
        Me.OutputBuilder.Clear()
        Return SavedOutput


    End Function
    ''' <summary>
    ''' Instead of generating to a file. This is a workaround to return the value as string
    ''' </summary>
    ''' <param name="action"></param>
    ''' <param name="removeEmptyLines"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetIsolatedOutput(action As action, Optional removeEmptyLines As Boolean = True) As String
        Me.SaveAndClearOutput()
        action()
        Dim s = Me.RestoreOutput()
        If (removeEmptyLines) Then s = s.RemoveEmptyLines
        Return s
    End Function

    ''' <summary>
    ''' Restore saved output while returning current output
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RestoreOutput() As String
        Dim saved = SavedOutput
        SavedOutput = ""
        Dim curr = Me.Output.ToString

        OutputBuilder.Clear()
        OutputBuilder.Append(saved)
        Return curr
    End Function


    Public Function IsIgnored(ce As CodeProperty2) As Boolean
        DebugWriteLine(ce.Name)
        Dim value = ce.ToPropertyInfo.GetGeneratorAttribute()
        Return (value IsNot Nothing) AndAlso value.IsIgnored

    End Function

    Public Sub DebugWrite(text As String)
        Me.OutputPaneTraceListener.Write(text)
    End Sub
    Public Sub DebugWriteLine(text As String)
        Me.OutputPaneTraceListener.WriteLine(text)
    End Sub

End Class
