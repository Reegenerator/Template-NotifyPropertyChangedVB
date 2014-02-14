Public Class XmlPropertyAttribute
    Inherits Attribute
    Property Name As String
    Sub New(attrName As String)
        Name = attrName
    End Sub
End Class
