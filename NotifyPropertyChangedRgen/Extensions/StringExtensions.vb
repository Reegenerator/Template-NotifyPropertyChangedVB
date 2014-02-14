Imports System.Runtime.CompilerServices
Imports System.Text

Public Module StringExtensions

#Region "StringBuilder"

    <Extension>
    Public Function AppendFormatLine(sb As StringBuilder, format As String, ParamArray values() As Object) _
        As StringBuilder
        Return sb.AppendFormat(format, values).AppendLine()
    End Function

    <Extension>
    Public Function AppendIndentFormat(sb As StringBuilder, tabCount As Integer, format As String,
                                       ParamArray values() As Object) As StringBuilder
        Static tab As Char = vbTab.First
        Return sb.AppendFormat("{0}{1}", New String(tab, tabCount), String.Format(format, values))
    End Function

    <Extension>
    Public Function AppendIndent(sb As StringBuilder, tabCount As Integer, text As String) As StringBuilder
        Static tab As Char = vbTab.First
        Return sb.AppendFormat("{0}{1}", New String(tab, tabCount), text)
    End Function
    ''' <summary>
    ''' Join two strings , only if both are not empty strings
    ''' </summary>
    ''' <param name="leftSide"></param>
    ''' <param name="conjunction"></param>
    ''' <param name="rightSide"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Function Conjoin(ByVal leftSide As String, ByVal conjunction As String, ByVal rightSide As String) As String
        Return leftSide & If(leftSide <> "" AndAlso rightSide <> "", conjunction, "") & rightSide
    End Function
#End Region
    <Extension>
    Function Quote(s As String) As String
        Return """" & s & """"
    End Function
End Module
