#Region "<Gen Renderer='NotifyPropertyChanged' Date='15/02/2014 15:22:06' Ver='1.1.0.16' Mode='OnVersionChanged' xmlns='http://tempuri.org/NotifyPropertyChanged.xsd' />"
Imports System.Runtime.CompilerServices
Imports System.Linq.Expressions

Interface INotifier
    Inherits System.ComponentModel.INotifyPropertyChanged
    Sub Notify(propertyName As String)
End Interface

Module NotifyPropertyChanged_Gen_Extensions
    <Extension>
    Public Sub NotifyChanged(notifier As INotifier, propertyName As String)
        notifier.Notify(propertyName)
    End Sub



    <Extension>
    Function SetPropertyAndNotify(Of T)(notifier As INotifier, ByRef field As T, newValue As T, Optional ByVal propertyName As String = Nothing) As Boolean
        If EqualityComparer(Of T).[Default].Equals(field, newValue) Then
            Return False
        End If
        field = newValue
        notifier.NotifyChanged(propertyName)
        Return True
    End Function


    <Extension>
    Sub NotifyChanged(Of T)(notifier As INotifier, memberExpr As System.Linq.Expressions.Expression(Of Func(Of T, Object)))
        notifier.Notify(ExprToString(memberExpr))
    End Sub

    <Extension>
    Sub NotifyChanged(Of T)(notifier As INotifier, ParamArray propExpressions() As System.Linq.Expressions.Expression(Of Func(Of T, Object)))
        For Each p In propExpressions
            notifier.NotifyChanged(p)
        Next
    End Sub

    <Extension>
    Sub NotifyChanged(notifier As INotifier, ParamArray props() As String)
        For Each p In props
            notifier.NotifyChanged(p)
        Next
    End Sub

#Region "ExprToString"

    Public Function ExprsToString(Of T)(ParamArray exprs() As Expression(Of Func(Of T, Object))) As String()

        Dim strings = (From x In exprs Select DirectCast(x, LambdaExpression).ExprToString()).ToArray
        Return strings
    End Function

    <Extension>
    Public Function ExprToString(Of T, T2)(ByVal expr As Expression(Of Func(Of T, T2))) As String
        Return DirectCast(expr, LambdaExpression).ExprToString()
    End Function

    <Extension>
    Public Function ExprToString(ByVal memberExpr As LambdaExpression) As String
        If memberExpr Is Nothing Then Return ""
        Dim currExpr As System.Linq.Expressions.Expression
        'when T2 is object, the expression will be wrapped in UnaryExpression of Convert{}
        Dim convertedToObject = TryCast(memberExpr.Body, UnaryExpression)
        If convertedToObject IsNot Nothing Then
            'unwrap
            currExpr = convertedToObject.Operand
        Else
            currExpr = memberExpr.Body
        End If
        Select Case currExpr.NodeType
            Case ExpressionType.MemberAccess
                Dim ex = DirectCast(currExpr, MemberExpression)
                Return ex.Member.Name
        End Select

        Throw New Exception("Expression ToString() extension only processes MemberExpression")
    End Function

#End Region
End Module




#End Region
