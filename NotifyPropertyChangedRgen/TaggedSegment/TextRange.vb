Imports EnvDTE

''' <summary>
''' Creat own class because the built in TextRange is not usable (used for regular expression search result)
''' </summary>
''' <remarks></remarks>
Public Class TextRange
    Property StartPoint As TextPoint
    Property EndPoint As TextPoint

    ''' <summary>
    ''' Valid if both StartPoint and EndPoint are not null
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property IsValid As Boolean
        Get
            Return StartPoint IsNot Nothing AndAlso EndPoint IsNot Nothing
        End Get
    End Property
    Sub New()

    End Sub
    Sub New(sp As TextPoint, ep As TextPoint)
        StartPoint = sp
        EndPoint = ep
    End Sub
    Sub ReplaceText(text As String)
        StartPoint.CreateEditPoint.ReplaceText(EndPoint, text, EnvDTE.vsEPReplaceTextOptions.vsEPReplaceTextAutoformat Or vsEPReplaceTextOptions.vsEPReplaceTextNormalizeNewlines)
    End Sub

    Public Sub Delete()
        If IsValid Then
            Dim ep = StartPoint.CreateEditPoint
            ep.Delete(EndPoint)
           
        End If
    End Sub

    Public Function GetText() As String
        Return StartPoint.CreateEditPoint.GetText(EndPoint)
    End Function
End Class