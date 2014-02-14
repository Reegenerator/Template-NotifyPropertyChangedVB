Imports System.Runtime.CompilerServices
Imports System.Text
Imports EnvDTE
Imports Kodeo.Reegenerator.Generators
Imports Microsoft.VisualBasic.CompilerServices

Module DebugExtensions

#Region "Debug helpers"
    ''' <summary>
    ''' Get string representation of Type.Member value
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="list"></param>
    ''' <param name="memberName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function DebugMember(Of T)(list As IEnumerable(Of T), memberName As String) As String
        Dim sb As New StringBuilder

        For Each x In list
            Dim text = Versioned.CallByName(x, memberName, CallType.Get).ToString()
            sb.AppendLine(text)
        Next
        Return sb.ToString
    End Function

    ''' <summary>
    ''' Get string value of specified member of each item in a list as type T
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="list"></param>
    ''' <param name="memberName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function DebugMembers(Of T)(list As IEnumerable, memberName As String) As IEnumerable(Of T)
        Dim sb As New List(Of T)
        Dim res = From x In list
                  Select Versioned.CallByName(x, memberName, CallType.Get)

        Return res.Cast(Of T)()
    End Function

    ''' <summary>
    ''' Get value of specified member of each item in a list as string array
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="memberName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function DebugMembers(list As IEnumerable, memberName As String) As String()
        Return list.DebugMembers(Of String)(memberName).ToArray()
    End Function

    ''' <summary>
    ''' Like DebugMember, but with two members
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="memberName"></param>
    ''' <param name="secondMemberName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function DebugMembers(list As IEnumerable, memberName As String, secondMemberName As String) As Object
        Dim sb As New List(Of String)
        Dim res = From x In list
                  Select Member1 = Versioned.CallByName(x, memberName, CallType.Get),
                    Member2 = Versioned.CallByName(x, secondMemberName, CallType.Get)

        Return res.ToArray

    End Function

    ''' <summary>
    ''' Alias for DebugPosition
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="charCount"></param>
    ''' <param name="listener"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function DP(point As TextPoint, Optional charCount As Integer = 10, Optional listener As OutputPaneTraceListener = Nothing) As String
        Return DebugPosition(point, charCount, listener)
    End Function

    ''' <summary>
    ''' Show the position of a textpoint by printing the surrounding text
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="charCount"></param>
    ''' <param name="listener"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function DebugPosition(point As TextPoint, Optional charCount As Integer = 50, Optional listener As OutputPaneTraceListener = Nothing) As String

        Dim start = point.CreateEditPoint
        Dim text = String.Format("{0}>|<{1}", start.GetText(-charCount), start.GetText(charCount))
        If listener IsNot Nothing Then listener.WriteLine(text)
        Return text

    End Function

    Private DebugSkipped As Boolean
    ''' <summary>
    ''' Launch debugger or Break if it's already attached
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DebugHere()
        If System.Diagnostics.Debugger.IsAttached Then
            System.Diagnostics.Debugger.Break()
        Else
            'If debug is cancelled once, stop trying to launch
            If DebugSkipped Then Return
            Dim launched = System.Diagnostics.Debugger.Launch()
            If Not launched Then
                DebugSkipped = True
            End If

        End If
    End Sub
#End Region

End Module
