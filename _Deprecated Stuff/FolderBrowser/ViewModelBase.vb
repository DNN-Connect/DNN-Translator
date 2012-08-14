Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.ComponentModel
Imports System.Linq.Expressions

Namespace Controls.FolderBrowser
 Public Class ViewModelBase
  Implements INotifyPropertyChanged
  Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

  Public Sub New()
  End Sub

  Protected Overridable Sub OnPropertyChanged(propName As String)
   RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propName))
  End Sub

  Protected Sub OnPropertyChanged(Of T)(propertyExpression As Expression(Of Func(Of T)))
   If propertyExpression.Body.NodeType = ExpressionType.MemberAccess Then
    Dim memberExpr = TryCast(propertyExpression.Body, MemberExpression)
    Dim propertyName As String = memberExpr.Member.Name
    Me.OnPropertyChanged(propertyName)
   End If
  End Sub

 End Class
End Namespace
