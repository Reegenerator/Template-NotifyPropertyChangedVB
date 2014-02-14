
Partial Public Class LibraryRenderer
    Public Const DefaultClassName As String = "NotifyPropertyChanged_Gen_Extensions"
    Property IsNet45 As Boolean
    Public Property ClassName As String = DefaultClassName
    Public Const INotifierName As String = "INotifier"
    'Public Property GeneratorTag As String
    Public Overrides Sub PreRender()
        MyBase.PreRender()

    End Sub

    Public Overloads Function RenderToString(classNm As String) As String
        Me.ClassName = classNm
        Return Me.RenderToString
    End Function
End Class
