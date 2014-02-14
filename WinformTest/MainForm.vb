Public Class MainForm
    Dim Person As Person
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Init()
    End Sub

    Private Sub Init()
        Person = New Person With {.FirstName = "Bill", .LastName = "Gates", .Address = "Earth", .Age = "42"}
        PersonBindingSource.DataSource = Person
    End Sub

    Private Sub incrementAgeButton_Click(sender As Object, e As EventArgs) Handles changeLastnameButton.Click
        Person.ChangeLastName(Person.LastName + "Changed")
    End Sub
End Class
