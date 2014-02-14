Imports NotifyPropertyChangedRgen
<NotifyPropertyChangedRgen.NotifyPropertyChanged_Gen>
Public Class Person
#Region "INotifier Functions	<Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.13' Date='14/02/2014 12:55:14' Mode='OnVersionChanged' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd' />"
    Implements INotifier
    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) _
        Implements WinformTest.INotifier.PropertyChanged
    Sub NotifyPropertyChanged(ByVal propertyName As String) Implements INotifier.Notify
        RaiseEvent PropertyChanged(Me, New System.ComponentModel.PropertyChangedEventArgs(propertyName))
    End Sub

#End Region



#Region "FirstName auto expanded by	<Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.13' Date='14/02/2014 13:03:23' Mode='Once' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd' />"
    Private _FirstName As System.String
    <NotifyPropertyChanged_Gen(ExtraNotifications:="Name")>
    Property FirstName() As System.String
        Get
            Return _FirstName
        End Get
        Set(ByVal value As System.String)
            '<Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.13' Date='14/02/2014 13:03:23' ExtraNotifications='Name' Mode='OnVersionChanged' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd'>
            Me.SetPropertyAndNotify(_FirstName, value)
            Me.NotifyChanged("Name")
            '</Gen>
        End Set
    End Property
#End Region



    <NotifyPropertyChanged_Gen(IsIgnored:=True)>
    Property LastName() As System.String






    ReadOnly Property Name As String
        Get
            Return String.Format("{0} {1}", FirstName, LastName)
        End Get
    End Property




#Region "Age auto expanded by   <Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.12' Mode='Once' Date='2014-02-11T20:13:58.5114179+08:00' xmlns='http://tempuri.org/Reegenerator.xsd' />"
    Private _Age As System.Int32
    <NotifyPropertyChanged_Gen(ExtraNotifications:="AgeString")>
    Property Age() As System.Int32
        Get
            Return _Age
        End Get
        Set(ByVal value As System.Int32)

            If value > 0 Then
                'Use InsertPoint to place the generated code in a position different from the default (last line of setter)

                '<Gen Renderer='NotifyPropertyChanged' Type='InsertPoint'/>
                '<Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.13' Date='14/02/2014 14:40:13' ExtraNotifications='AgeString' Mode='OnVersionChanged' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd'>
                Me.SetPropertyAndNotify(_Age, value)
                Me.NotifyChanged("AgeString")
                '</Gen>

            End If
           
        End Set
    End Property
#End Region
#Region "Address auto expanded by	<Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.13' Date='14/02/2014 13:04:29' Mode='Once' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd' />"
    Private _Address As System.String
    Property Address() As System.String
        Get
            Return _Address
        End Get
        Set(ByVal value As System.String)
            '<Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.13' Date='14/02/2014 13:04:29' Mode='OnVersionChanged' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd'>
            Me.SetPropertyAndNotify(_Address, value)
            '</Gen>
        End Set
    End Property
#End Region




    ReadOnly Property AgeString As String
        Get
            Return String.Format("{0} years old", Age)
        End Get
    End Property


    <NotifyPropertyChangedRgen.NotifyPropertyChanged_Gen(ExtraNotifications:="LastName, Name")>
    Sub ChangeLastName(newLastname As String)
        Me.LastName = newLastname

        '<Gen Renderer='NotifyPropertyChanged' Ver='1.1.0.13' Date='14/02/2014 13:17:05' ExtraNotifications='LastName, Name' Mode='OnVersionChanged' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd'>
        Me.NotifyChanged("LastName", "Name")
        '</Gen>
    End Sub


End Class
