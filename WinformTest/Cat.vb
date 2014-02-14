Imports NotifyPropertyChangedRgen
'<NotifyPropertyChanged_Gen()>
Public Class Cat
    Property PlaysPiano As Boolean

    Private _MaxLife As Integer = 9
    ReadOnly Property MaxLife As Integer
        Get
            Return _MaxLife
        End Get
    End Property

    <NotifyPropertyChanged_Gen(ExtraNotifications:="MaxLife")>
    Sub Kill()
        _MaxLife -= 1
    End Sub
End Class
